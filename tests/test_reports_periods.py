"""Testy porównania "te same dni", prognozy i trybu od wypłaty w raporcie miesięcznym."""
import pandas as pd
import pytest
from openpyxl import load_workbook

from src.categorize import categorize
from src.loader import load_expenses
from src.periods import PeriodMap, find_paydays
from src.report_monthly import create_monthly_report

from tests.conftest import d, make_row, write_export
from tests.test_reports import find_row, recalculated

NOV = pd.Period("2026-11", freq="M")
TOTAL = "Razem (bez inwestycji)"


def build(tmp_path, settings, categories, rules, rows, mode):
    write_export(tmp_path / "data" / "exports" / "e.xlsx", rows)
    raw = load_expenses(tmp_path, settings)
    df = categorize(raw, rules, categories, settings)
    start, end = raw["date"].min().normalize(), raw["date"].max().normalize()
    paydays = find_paydays(raw, settings) if mode == "payday" else None
    pmap = PeriodMap(mode, start, end, paydays)
    path = create_monthly_report(tmp_path, df, categories, settings, NOV, None, start, end, pmap=pmap)
    return path, pmap


def shop(day, amount, balance):
    return make_row(d(day), amount, "LIDL Szczecin", balance=balance)


def test_calendar_same_days_comparison(tmp_path, settings, categories, rules):
    rows = [shop("2026-10-02", -100, 900), shop("2026-10-20", -300, 800),
            shop("2026-11-02", -50, 700), shop("2026-11-05", -70, 600)]
    path, _ = build(tmp_path, settings, categories, rules, rows, "calendar")
    ws = recalculated(path, tmp_path)["Podsumowanie"]
    assert ws["B6"].value == 5                                   # dni, które minęły
    r = find_row(ws, TOTAL)
    assert ws.cell(r, 3).value == pytest.approx(120)             # listopad
    assert ws.cell(r, 6).value == pytest.approx(100)             # październik, dni 1-5
    assert ws.cell(r, 7).value == pytest.approx(20)              # zmiana
    assert ws.cell(r, 9).value == pytest.approx(400)             # październik w całości
    assert ws.cell(r, 8).value == pytest.approx(0.2)             # zmiana %


def test_forecast_for_open_calendar_month(tmp_path, settings, categories, rules):
    rows = [shop("2026-10-02", -100, 900), shop("2026-11-02", -50, 700), shop("2026-11-05", -70, 600)]
    path, _ = build(tmp_path, settings, categories, rules, rows, "calendar")
    ws = recalculated(path, tmp_path)["Podsumowanie"]
    r = find_row(ws, "Prognoza wydatków (PLN)")
    assert ws.cell(r, 3).value == pytest.approx(120 / 5 * 30)    # listopad ma 30 dni
    assert ws.cell(find_row(ws, "Długość okresu (dni)"), 3).value == 30


def test_forecast_equals_total_when_period_complete(tmp_path, settings, categories, rules):
    rows = [shop("2026-10-02", -100, 900), shop("2026-11-02", -50, 700), shop("2026-11-30", -70, 600)]
    path, _ = build(tmp_path, settings, categories, rules, rows, "calendar")
    ws = recalculated(path, tmp_path)["Podsumowanie"]
    assert ws.cell(find_row(ws, "Prognoza wydatków (PLN)"), 3).value == pytest.approx(120)


def payday_rows():
    pay = lambda day, bal: make_row(d(day), 5400, "Wypłata", "Firma", typ="Przelew przychodzący", balance=bal)
    return [pay("2026-10-05", 5000), shop("2026-10-06", -100, 4900), shop("2026-10-20", -300, 4600),
            pay("2026-11-03", 9000), shop("2026-11-04", -50, 8950), shop("2026-11-08", -70, 8880)]


def test_payday_mode_assigns_periods_and_compares_same_days(tmp_path, settings, categories, rules):
    path, _ = build(tmp_path, settings, categories, rules, payday_rows(), "payday")
    wb = recalculated(path, tmp_path)
    ws = wb["Podsumowanie"]
    assert ws["B6"].value == 6                                   # 3-8 listopada
    r = find_row(ws, TOTAL)
    assert ws.cell(r, 3).value == pytest.approx(120)
    assert ws.cell(r, 6).value == pytest.approx(100)             # okres październikowy, dni 1-6 (6.10 = dzień 2)
    assert ws.cell(r, 9).value == pytest.approx(400)
    assert "od wypłaty" in ws["B5"].value
    # transakcje z 2.11 i wcześniej należą do okresu październikowego
    tr = load_workbook(path)["Transakcje"]
    periods = {tr.cell(i, 1).value.strftime("%Y-%m-%d"): tr.cell(i, 2).value for i in range(2, tr.max_row + 1)}
    assert periods["2026-10-20"] == 202610 and periods["2026-11-04"] == 202611


def test_payday_forecast_uses_previous_period_length(tmp_path, settings, categories, rules):
    path, _ = build(tmp_path, settings, categories, rules, payday_rows(), "payday")
    ws = recalculated(path, tmp_path)["Podsumowanie"]
    # poprzedni okres: 5.10 - 2.11 = 29 dni
    assert ws.cell(find_row(ws, "Prognoza wydatków (PLN)"), 3).value == pytest.approx(120 / 6 * 29)


def test_payday_mode_without_payday_fails(tmp_path, settings, categories, rules):
    rows = [shop("2026-10-02", -100, 900), shop("2026-11-02", -50, 700)]
    with pytest.raises(ValueError, match="nie znaleziono żadnej wypłaty"):
        build(tmp_path, settings, categories, rules, rows, "payday")
