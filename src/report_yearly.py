"""
report_yearly.py - raport roczny w xlsx (z wykresami).

Co robi ten plik:
1. bierze wszystkie transakcje z wybranego roku
2. zapisuje je w arkuszu "Transakcje"
3. buduje arkusze "Podsumowanie" i "Miesiące", które LICZĄ SIĘ FORMUŁAMI z arkusza
   "Transakcje" (SUMIFS) - tak samo jak w raporcie miesięcznym
4. dodaje arkusz "Do przypisania" z największymi transakcjami z kategorii "Inne"

Plik: output/<rok>_roczny.xlsx (z filtrem: output/<rok>_roczny_<filtr>.xlsx)
Rok niepełny (np. bieżący) jest wyraźnie oznaczony i nadpisuje się przy kolejnym uruchomieniu.
"""
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from src.report_monthly import (FONT, INPUT_FILL, NUM_FMT, PCT_FMT, TOTAL_FILL, TYPE_PL,
                                header_row, month_status, put, style_chart, write_transactions)

MONTHS_SHEET = "Miesiące"


def write_months(ws, cats, year, labels, last_row):
    """Arkusz 'Miesiące': kategorie × 12 miesięcy (SUMIFS) + wykresy."""
    col_month = f"Transakcje!$B$2:$B${last_row}"
    col_cat = f"Transakcje!$C$2:$C${last_row}"
    col_sum = f"Transakcje!$E$2:$E${last_row}"

    header_row(ws, 1, ["Kategoria"] + labels + ["Razem"])
    put(ws, "A2", "klucz miesiąca (używany w formułach)", color="808080")
    for m in range(1, 13):
        col = get_column_letter(1 + m)
        put(ws, f"{col}2", f"=Podsumowanie!$B$2*100+{m}", fmt='0000"-"00', color="808080", align="center")

    first = 3
    for i, (name, _) in enumerate(cats):
        r = first + i
        put(ws, f"A{r}", name)
        for m in range(1, 13):
            col = get_column_letter(1 + m)
            put(ws, f"{col}{r}", f"=SUMIFS({col_sum},{col_cat},$A{r},{col_month},{col}$2)", fmt=NUM_FMT)
        put(ws, f"N{r}", f"=SUM(B{r}:M{r})", bold=True, fmt=NUM_FMT)

    last_cat = first + len(cats) - 1
    total_row = last_cat + 1
    put(ws, f"A{total_row}", "Razem (bez inwestycji)", bold=True, fill=TOTAL_FILL)
    for m in range(1, 14):
        col = get_column_letter(1 + m)
        put(ws, f"{col}{total_row}", f"=SUM({col}{first}:{col}{last_cat})", bold=True,
            fmt=NUM_FMT, fill=TOTAL_FILL)

    put(ws, f"A{total_row + 2}",
        "Miesiące oznaczone jako niepełne mają tylko część danych, więc ich sum nie porównuj 1:1 z pełnymi.",
        color="808080")

    # wykres 1: skumulowany, kategorie w kolorach
    stacked = BarChart()
    stacked.type = "col"
    stacked.grouping = "stacked"
    stacked.overlap = 100
    stacked.add_data(Reference(ws, min_col=1, max_col=13, min_row=first, max_row=last_cat),
                     from_rows=True, titles_from_data=True)
    stacked.set_categories(Reference(ws, min_col=2, max_col=13, min_row=1, max_row=1))
    stacked.y_axis.title = "PLN"
    style_chart(stacked, "Wydatki w kolejnych miesiącach (wg kategorii)", 26, 11)
    ws.add_chart(stacked, f"A{total_row + 4}")

    # wykres 2: łączna kwota w miesiącu
    line = LineChart()
    line.add_data(Reference(ws, min_col=1, max_col=13, min_row=total_row, max_row=total_row),
                  from_rows=True, titles_from_data=True)
    line.set_categories(Reference(ws, min_col=2, max_col=13, min_row=1, max_row=1))
    line.y_axis.title = "PLN"
    line.legend = None
    style_chart(line, "Łączne wydatki w miesiącu", 26, 9)
    ws.add_chart(line, f"A{total_row + 27}")

    ws.column_dimensions["A"].width = 34
    for m in range(1, 14):
        ws.column_dimensions[get_column_letter(1 + m)].width = 17
    ws.row_dimensions[1].height = 30
    return first, last_cat, total_row


def write_summary(ws, cats, year, status_text, last_row, m_first, m_total):
    """Arkusz 'Podsumowanie': tabela kategorii za cały rok, stałe/zmienne, rekordy miesięcy."""
    col_month = f"Transakcje!$B$2:$B${last_row}"
    col_cat = f"Transakcje!$C$2:$C${last_row}"
    col_type = f"Transakcje!$D$2:$D${last_row}"
    col_sum = f"Transakcje!$E$2:$E${last_row}"
    in_year = f'{col_month},">="&($B$2*100+1),{col_month},"<="&($B$2*100+12)'
    months_total = f"'{MONTHS_SHEET}'!$B${m_total}:$M${m_total}"

    put(ws, "A1", "Raport roczny", bold=True).font = Font(name=FONT, size=14, bold=True)
    put(ws, "A2", "Rok", bold=True)
    put(ws, "B2", year, fmt="0", color="0000FF", fill=INPUT_FILL, align="center")
    put(ws, "C2", "← możesz zmienić rok, jeśli w arkuszu Transakcje są dane z innego roku", color="808080")
    put(ws, "A3", "Kompletność danych", bold=True)
    put(ws, "B3", status_text)
    put(ws, "A4", "Miesięcy z wydatkami", bold=True)
    put(ws, "B4", f'=COUNTIF({months_total},"<>0")', fmt="0", align="center")

    header_row(ws, 6, ["Kategoria", "Typ", "Suma roku (PLN)", "Udział", "Średnio / miesiąc (PLN)", "Transakcje"])
    first = 7
    total_row = first + len(cats)
    for i, (name, typ) in enumerate(cats):
        r = first + i
        put(ws, f"A{r}", name)
        put(ws, f"B{r}", TYPE_PL.get(typ, typ))
        put(ws, f"C{r}", f"='{MONTHS_SHEET}'!N{m_first + i}", fmt=NUM_FMT)
        put(ws, f"D{r}", f"=IF($C${total_row}=0,0,C{r}/$C${total_row})", fmt=PCT_FMT)
        put(ws, f"E{r}", f"=IF($B$4=0,0,C{r}/$B$4)", fmt=NUM_FMT)
        put(ws, f"F{r}", f"=COUNTIFS({col_cat},$A{r},{in_year})", fmt="0")

    last_cat = total_row - 1
    put(ws, f"A{total_row}", "Razem (bez inwestycji)", bold=True, fill=TOTAL_FILL)
    put(ws, f"B{total_row}", None, fill=TOTAL_FILL)
    put(ws, f"C{total_row}", f"=SUM(C{first}:C{last_cat})", bold=True, fmt=NUM_FMT, fill=TOTAL_FILL)
    put(ws, f"D{total_row}", f"=SUM(D{first}:D{last_cat})", bold=True, fmt=PCT_FMT, fill=TOTAL_FILL)
    put(ws, f"E{total_row}", f"=IF($B$4=0,0,C{total_row}/$B$4)", bold=True, fmt=NUM_FMT, fill=TOTAL_FILL)
    put(ws, f"F{total_row}", f"=SUM(F{first}:F{last_cat})", bold=True, fmt="0", fill=TOTAL_FILL)

    # stałe vs zmienne i inwestycje
    r = total_row + 2
    header_row(ws, r, ["Podział", "", "Suma roku (PLN)", "Udział"])
    for j, (label, typ_pl) in enumerate((("Stałe", "stały"), ("Zmienne", "zmienny"))):
        rr = r + 1 + j
        put(ws, f"A{rr}", label)
        put(ws, f"C{rr}", f'=SUMIFS({col_sum},{col_type},"{typ_pl}",{in_year})', fmt=NUM_FMT)
        put(ws, f"D{rr}", f"=IF($C${total_row}=0,0,C{rr}/$C${total_row})", fmt=PCT_FMT)
    rr = r + 3
    put(ws, f"A{rr}", "Inwestycje (osobno, poza wydatkami)")
    put(ws, f"C{rr}", f'=SUMIFS({col_sum},{col_type},"inwestycje",{in_year})', fmt=NUM_FMT)

    # najdroższy i najtańszy miesiąc (z wydatkami)
    r = rr + 2
    header_row(ws, r, ["Rekordy", "", "Suma (PLN)", "Miesiąc"])
    labels_rng = f"'{MONTHS_SHEET}'!$B$1:$M$1"
    put(ws, f"A{r + 1}", "Najdroższy miesiąc")
    put(ws, f"C{r + 1}", f"=MAX({months_total})", fmt=NUM_FMT)
    put(ws, f"D{r + 1}", f'=IFERROR(INDEX({labels_rng},MATCH(C{r + 1},{months_total},0)),"-")')
    put(ws, f"A{r + 2}", "Najtańszy miesiąc (z wydatkami)")
    put(ws, f"C{r + 2}", f'=IFERROR(SMALL({months_total},COUNTIF({months_total},0)+1),0)', fmt=NUM_FMT)
    put(ws, f"D{r + 2}", f'=IFERROR(INDEX({labels_rng},MATCH(C{r + 2},{months_total},0)),"-")')

    # wykres: suma roku po kategoriach
    chart = BarChart()
    chart.type = "bar"
    chart.legend = None
    chart.add_data(Reference(ws, min_col=3, min_row=6, max_row=last_cat), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=first, max_row=last_cat))
    chart.y_axis.title = "PLN"
    chart.x_axis.scaling.orientation = "maxMin"
    chart.y_axis.crosses = "max"
    style_chart(chart, "Wydatki w roku wg kategorii", 18, 9)
    ws.add_chart(chart, "H6")

    for col, width in zip("ABCDEF", (34, 22, 18, 22, 22, 12)):
        ws.column_dimensions[col].width = width


def write_unassigned(ws, data, default_category, limit=50):
    """Arkusz 'Do przypisania': największe transakcje z kategorii 'Inne' w roku."""
    header_row(ws, 1, ["Data", "Kwota (PLN)", "Opis"])
    rows = data[data["category"] == default_category].sort_values("spent", ascending=False)
    total = len(rows)
    rows = rows.head(limit)
    if rows.empty:
        put(ws, "A2", "Wszystkie transakcje zostały rozpoznane.")
    for i, r in enumerate(rows.itertuples(), start=2):
        put(ws, f"A{i}", r.date.strftime("%Y-%m-%d"))
        put(ws, f"B{i}", float(r.spent), fmt=NUM_FMT)
        put(ws, f"C{i}", r.text[:120])
    note = f"Pokazano {len(rows)} największych z {total}. Dopisz słowa kluczowe do config/rules.json."
    put(ws, f"A{len(rows) + 4}", note, color="808080")
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 90


def create_yearly_report(base_dir, data, categories, settings, year, filter_name, data_start, data_end):
    """
    Tworzy plik output/<rok>_roczny[_filtr].xlsx i zwraca jego ścieżkę.

    data       - skategoryzowane transakcje (po filtrze), wszystkie lata
    year       - rok (liczba, np. 2026)
    data_start, data_end - zakres dat w eksportach (do oznaczania niepełnych miesięcy i roku)
    """
    window = data[data["date"].dt.year == year].sort_values("date", kind="stable")
    if window.empty:
        raise ValueError(f"Brak transakcji z roku {year}.")

    present = set(window["category"])
    cats = [(c["name"], c["type"]) for c in categories["categories"]
            if c["name"] in present and c["type"] != "savings"]

    # etykiety miesięcy ze statusem kompletności
    periods = [pd.Period(year=year, month=m, freq="M") for m in range(1, 13)]
    labels = [f"{p.strftime('%Y-%m')} ({month_status(p, data_start, data_end)})" for p in periods]

    year_start, year_end = pd.Timestamp(year, 1, 1), pd.Timestamp(year, 12, 31)
    if data_start <= year_start and data_end >= year_end:
        status_text = f"rok {year}: pełny"
    else:
        status_text = (f"rok {year}: NIEPEŁNY (dane w eksporcie: "
                       f"{data_start:%Y-%m-%d} - {data_end:%Y-%m-%d})")

    last_row = len(window) + 1

    wb = Workbook()
    ws_sum = wb.active
    ws_sum.title = "Podsumowanie"
    ws_m = wb.create_sheet(MONTHS_SHEET)
    ws_tr = wb.create_sheet("Transakcje")
    ws_un = wb.create_sheet("Do przypisania")

    write_transactions(ws_tr, window)
    m_first, _, m_total = write_months(ws_m, cats, year, labels, last_row)
    write_summary(ws_sum, cats, year, status_text, last_row, m_first, m_total)
    write_unassigned(ws_un, window, settings["default_category"])

    out_dir = Path(base_dir) / "output"
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / (f"{year}_roczny" + (f"_{filter_name}" if filter_name else "") + ".xlsx")
    wb.save(out_path)
    return out_path
