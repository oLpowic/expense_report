"""Testy raportów xlsx: miesięczny i roczny (struktura, formuły, wartości po przeliczeniu)."""
import shutil
import subprocess
import sys

import pandas as pd
import pytest
from openpyxl import load_workbook

from src.report_monthly import create_monthly_report, month_key, month_status
from src.report_yearly import create_yearly_report

NOV = pd.Period("2026-11", freq="M")
TOTAL_LABEL = "Razem (bez inwestycji)"

# sumy z danych testowych (conftest.sample_rows)
NOV_SPENDING = 2100 + 230.5 - 30 + 89.9 + 45 + 77 + 50      # bez inwestycji (500), z odjętym zwrotem
OCT_SPENDING = 1500 + 100
INVESTMENTS = 500


def monthly(p, month=NOV, filter_name=None, df=None):
    return create_monthly_report(p["base"], p["df"] if df is None else df, p["categories"], p["settings"],
                                 month, filter_name, p["data_start"], p["data_end"])


def _recalc_with_excel(path, out):
    """Przelicza plik w prawdziwym Excelu (Windows + pywin32). Zwraca ścieżkę wyniku albo None."""
    if sys.platform != "win32":
        return None
    try:
        import win32com.client
    except ImportError:
        return None
    try:
        app = win32com.client.DispatchEx("Excel.Application")
    except Exception:           # Excel nie jest zainstalowany
        return None
    app.Visible = False
    app.DisplayAlerts = False
    target = out / path.name
    try:
        wb = app.Workbooks.Open(str(path.resolve()))
        app.CalculateFull()
        wb.SaveAs(str(target.resolve()), FileFormat=51)     # 51 = xlsx
        wb.Close(False)
    finally:
        app.Quit()
    return target


def _recalc_with_libreoffice(path, out):
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return None
    subprocess.run([soffice, "--headless", "--convert-to", "xlsx", "--outdir", str(out), str(path)],
                   check=True, capture_output=True, timeout=180)
    return out / path.name


def recalculated(path, tmp_path):
    """Przelicza formuły (Excel, a jeśli go brak - LibreOffice) i zwraca skoroszyt z wartościami.
    Gdy nie ma żadnego z nich, test jest pomijany."""
    out = tmp_path / "recalc"
    out.mkdir(exist_ok=True)
    result = _recalc_with_excel(path, out) or _recalc_with_libreoffice(path, out)
    if result is None:
        pytest.skip("brak Excela (pywin32) i LibreOffice - pomijam test wartości formuł")
    return load_workbook(result, data_only=True)


def find_row(ws, label, col=1):
    for row in range(1, ws.max_row + 1):
        if ws.cell(row, col).value == label:
            return row
    raise AssertionError(f"Nie znaleziono '{label}' w arkuszu {ws.title}")


# ---------------------------------------------------------------- pomocnicze
@pytest.mark.parametrize("month, expected", [
    ("2026-11", "niepełny"),     # dane kończą się 12.11 (blokada pominięta, ostatnia transakcja 11.11)
    ("2026-10", "niepełny"),     # dane zaczynają się 28.10
    ("2026-05", "brak danych"),
    ("2027-01", "brak danych"),
])
def test_month_status(pipeline, month, expected):
    assert month_status(pd.Period(month, freq="M"), pipeline["data_start"], pipeline["data_end"]) == expected


def test_month_status_full_month():
    start, end = pd.Timestamp(2026, 9, 1), pd.Timestamp(2026, 12, 31)
    assert month_status(pd.Period("2026-10", freq="M"), start, end) == "pełny"


def test_month_key():
    assert month_key(pd.Period("2026-01", freq="M")) == 202601


# ---------------------------------------------------------------- raport miesięczny
def test_monthly_creates_file_in_year_folder(pipeline):
    path = monthly(pipeline)
    assert path == pipeline["base"] / "output" / "2026" / "2026-11.xlsx"
    assert path.exists()


def test_monthly_filter_is_in_file_name(pipeline):
    assert monthly(pipeline, filter_name="everyday").name == "2026-11_everyday.xlsx"


def test_monthly_sheets_and_charts(pipeline):
    wb = load_workbook(monthly(pipeline))
    assert wb.sheetnames == ["Podsumowanie", "Porównanie", "Transakcje", "Do przypisania"]
    assert len(wb["Podsumowanie"]._charts) == 1
    assert len(wb["Porównanie"]._charts) == 1


def test_monthly_transactions_sheet_has_window_rows(pipeline):
    wb = load_workbook(monthly(pipeline))
    # październik: 2 wydatki; listopad: 8 pozycji (czynsz, Lidl, zwrot, Wolt, inwestycja, Netflix,
    # sklep, Koleo). Bez wypłaty (wpływ), odrzuconego Zalando i blokady.
    assert wb["Transakcje"].max_row - 1 == 10


def test_monthly_summary_uses_formulas_and_month_cell(pipeline):
    ws = load_workbook(monthly(pipeline))["Podsumowanie"]
    assert ws["B2"].value == 202611
    assert str(ws["C8"].value).startswith("=SUMIFS(")
    assert "niepełny" in ws["B5"].value


def test_monthly_savings_are_not_in_category_table(pipeline):
    ws = load_workbook(monthly(pipeline))["Podsumowanie"]
    names = [ws.cell(r, 1).value for r in range(8, 30)]
    assert "Inwestycje" not in names
    assert any(isinstance(n, str) and n.startswith("Inwestycje (osobno") for n in names)


def test_monthly_comparison_skips_months_without_data(pipeline):
    ws = load_workbook(monthly(pipeline))["Porównanie"]
    headers = [c.value for c in ws[1] if c.value]
    # kolumna z kategoriami + październik + listopad (maj-wrzesień bez danych są pominięte)
    assert len(headers) == 3
    assert headers[1].startswith("2026-10") and headers[2].startswith("2026-11")


def test_monthly_unassigned_sheet_lists_other_category(pipeline):
    ws = load_workbook(monthly(pipeline))["Do przypisania"]
    texts = [ws.cell(r, 3).value for r in range(2, 6)]
    assert any(t and "sklep xyz" in t for t in texts)


def test_monthly_computed_values(pipeline, tmp_path):
    wb = recalculated(monthly(pipeline), tmp_path)
    ws = wb["Podsumowanie"]
    total_row = find_row(ws, TOTAL_LABEL)
    assert ws.cell(total_row, 3).value == pytest.approx(NOV_SPENDING)        # suma listopada
    # październikowe wydatki są po 11. dniu, więc "te same dni" (1-11) dają 0
    assert ws.cell(total_row, 6).value == pytest.approx(0)
    assert ws.cell(total_row, 7).value == pytest.approx(NOV_SPENDING)
    assert ws.cell(total_row, 9).value == pytest.approx(OCT_SPENDING)        # poprzedni okres w całości
    assert ws.cell(total_row, 4).value == pytest.approx(1.0)                 # udziały sumują się do 100%
    inv_row = find_row(ws, "Inwestycje (osobno, poza wydatkami)")
    assert ws.cell(inv_row, 3).value == pytest.approx(INVESTMENTS)
    fixed, variable = ws.cell(find_row(ws, "Stałe"), 3).value, ws.cell(find_row(ws, "Zmienne"), 3).value
    assert fixed + variable == pytest.approx(NOV_SPENDING)


def test_monthly_refund_reduces_category(pipeline, tmp_path):
    ws = recalculated(monthly(pipeline), tmp_path)["Podsumowanie"]
    assert ws.cell(find_row(ws, "Ubrania i zakupy"), 3).value == pytest.approx(-30.0)


def test_monthly_empty_window_raises(pipeline):
    empty = pipeline["df"].iloc[0:0]
    with pytest.raises(ValueError):
        monthly(pipeline, df=empty)


# ---------------------------------------------------------------- raport roczny
def yearly(p, year=2026, filter_name=None, df=None, start=None, end=None):
    return create_yearly_report(p["base"], p["df"] if df is None else df, p["categories"], p["settings"],
                                year, filter_name, start or p["data_start"], end or p["data_end"])


def test_yearly_file_location_and_name(pipeline):
    path = yearly(pipeline)
    assert path == pipeline["base"] / "output" / "2026_roczny.xlsx"
    assert yearly(pipeline, filter_name="everyday").name == "2026_roczny_everyday.xlsx"


def test_yearly_sheets_and_charts(pipeline):
    wb = load_workbook(yearly(pipeline))
    assert wb.sheetnames == ["Podsumowanie", "Miesiące", "Transakcje", "Do przypisania"]
    assert len(wb["Podsumowanie"]._charts) == 1
    assert len(wb["Miesiące"]._charts) == 2


def test_yearly_marks_incomplete_year(pipeline):
    ws = load_workbook(yearly(pipeline))["Podsumowanie"]
    assert "NIEPEŁNY" in ws["B3"].value


def test_yearly_marks_complete_year(pipeline):
    path = yearly(pipeline, start=pd.Timestamp(2026, 1, 1), end=pd.Timestamp(2026, 12, 31))
    assert "pełny" in load_workbook(path)["Podsumowanie"]["B3"].value


def test_yearly_no_data_for_year_raises(pipeline):
    with pytest.raises(ValueError):
        yearly(pipeline, year=2020)


def test_yearly_computed_values(pipeline, tmp_path):
    wb = recalculated(yearly(pipeline), tmp_path)
    ws = wb["Podsumowanie"]
    total_row = find_row(ws, TOTAL_LABEL)
    assert ws.cell(total_row, 3).value == pytest.approx(NOV_SPENDING + OCT_SPENDING)
    assert ws["B4"].value == 2                                              # 2 miesiące z wydatkami
    assert ws.cell(total_row, 5).value == pytest.approx((NOV_SPENDING + OCT_SPENDING) / 2)
    inv_row = find_row(ws, "Inwestycje (osobno, poza wydatkami)")
    assert ws.cell(inv_row, 3).value == pytest.approx(INVESTMENTS)
    # rekordy miesięcy
    assert ws.cell(find_row(ws, "Najdroższy miesiąc"), 3).value == pytest.approx(NOV_SPENDING)
    assert ws.cell(find_row(ws, "Najtańszy miesiąc (z wydatkami)"), 3).value == pytest.approx(OCT_SPENDING)


def test_yearly_months_sheet_matches_summary(pipeline, tmp_path):
    wb = recalculated(yearly(pipeline), tmp_path)
    months = wb["Miesiące"]
    total_row = find_row(months, TOTAL_LABEL)
    assert months.cell(total_row, 11).value == pytest.approx(OCT_SPENDING)   # kolumna K = październik
    assert months.cell(total_row, 12).value == pytest.approx(NOV_SPENDING)   # kolumna L = listopad
    assert months.cell(total_row, 14).value == pytest.approx(OCT_SPENDING + NOV_SPENDING)
