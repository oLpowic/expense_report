"""
report_monthly.py - raport miesięczny w xlsx (z wykresami).

Co robi ten plik:
1. wybiera transakcje z wybranego miesiąca i z poprzednich miesięcy (do porównania)
2. zapisuje je w arkuszu "Transakcje"
3. buduje arkusze "Podsumowanie" i "Porównanie", które LICZĄ SIĘ FORMUŁAMI z arkusza
   "Transakcje" (SUMIFS). Dzięki temu, jeśli w Excelu zmienisz kategorię transakcji
   albo miesiąc w komórce B2, liczby i wykresy przeliczą się same.
4. dodaje arkusz "Do przypisania" z transakcjami z kategorii "Inne"

Okres może być kalendarzowy albo liczony od dnia wypłaty (patrz src/periods.py).
Porównanie z poprzednim okresem robimy "na te same dni": jeśli bieżący okres trwa np. 21 dni,
to poprzedni też liczymy tylko za jego pierwsze 21 dni (obok jest też suma za cały poprzedni okres).

Uwaga: okres w arkuszach jest zapisany jako liczba rrrrmm (np. 202611),
wyświetlana jako 2026-11. Dzięki temu formuły działają tak samo w każdym języku Excela.
"""
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from src.periods import PeriodMap, assign_periods

FONT = "Arial"
NUM_FMT = '#,##0.00;-#,##0.00;"-"'     # zero wyświetla się jako "-"
PCT_FMT = '0.0%;-0.0%;"-"'
KEY_FMT = '0000"-"00'                  # 202611 -> 2026-11
TYPE_PL = {"fixed": "stały", "variable": "zmienny", "savings": "inwestycje"}

HEADER_FILL = PatternFill("solid", start_color="1F3864")
INPUT_FILL = PatternFill("solid", start_color="FFFF00")
TOTAL_FILL = PatternFill("solid", start_color="D9E1F2")


# ---------------------------------------------------------------- pomocnicze
def month_key(period):
    """Period('2026-11') -> 202611."""
    return period.year * 100 + period.month


def month_status(period, data_start, data_end):
    """Czy dane obejmują cały miesiąc? Zwraca: pełny / niepełny / brak danych."""
    first = period.start_time.normalize()
    last = period.end_time.normalize()
    if data_end < first or data_start > last:
        return "brak danych"
    if data_start <= first and data_end >= last:
        return "pełny"
    return "niepełny"


def put(ws, cell, value, bold=False, fmt=None, color="000000", fill=None, align=None):
    """Wpisuje wartość do komórki i ustawia czcionkę."""
    c = ws[cell]
    c.value = value
    c.font = Font(name=FONT, size=10, bold=bold, color=color)
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    if align:
        c.alignment = Alignment(horizontal=align)
    return c


def header_row(ws, row, labels, first_col=1):
    """Ciemny nagłówek tabeli."""
    for i, label in enumerate(labels):
        c = ws.cell(row=row, column=first_col + i, value=label)
        c.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", wrap_text=True)


def style_chart(chart, title, width, height):
    """Ustawienia wykresu, dzięki którym Excel rysuje go poprawnie."""
    chart.title = title
    chart.title.overlay = False
    chart.roundedCorners = False
    chart.width, chart.height = width, height
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    if chart.legend is not None:
        chart.legend.overlay = False
        chart.legend.position = "b"


# ---------------------------------------------------------------- arkusze
def write_transactions(ws, data):
    """Arkusz 'Transakcje': surowe wiersze, na których liczą się formuły."""
    header_row(ws, 1, ["Data", "Okres", "Kategoria", "Typ", "Kwota (PLN)", "Rodzaj", "Opis", "Dzień okresu"])
    for i, r in enumerate(data.itertuples(), start=2):
        ws.cell(i, 1, r.date.to_pydatetime()).number_format = "yyyy-mm-dd"
        ws.cell(i, 2, int(r.period)).number_format = KEY_FMT
        ws.cell(i, 3, r.category)
        ws.cell(i, 4, TYPE_PL.get(r.type, r.type))
        ws.cell(i, 5, float(r.spent)).number_format = NUM_FMT
        ws.cell(i, 6, "wydatek" if r.kind == "expense" else "zwrot")
        ws.cell(i, 7, r.text[:120])
        ws.cell(i, 8, int(r.period_day))
        for col in range(1, 9):
            ws.cell(i, col).font = Font(name=FONT, size=10)
    for col, width in zip("ABCDEFGH", (12, 10, 24, 12, 14, 10, 70, 12)):
        ws.column_dimensions[col].width = width
    ws.auto_filter.ref = f"A1:H{len(data) + 1}"
    ws.freeze_panes = "A2"


def write_summary(ws, data, cats, month, filter_name, status_text, last_row, elapsed, length):
    """Arkusz 'Podsumowanie': tabela kategorii, stałe/zmienne, TOP 5, wykres."""
    # zakresy z arkusza Transakcje, używane w formułach
    col_month = f"Transakcje!$B$2:$B${last_row}"
    col_cat = f"Transakcje!$C$2:$C${last_row}"
    col_type = f"Transakcje!$D$2:$D${last_row}"
    col_sum = f"Transakcje!$E$2:$E${last_row}"
    col_day = f"Transakcje!$H$2:$H${last_row}"

    put(ws, "A1", "Raport wydatków", bold=True).font = Font(name=FONT, size=14, bold=True)
    put(ws, "A2", "Miesiąc", bold=True)
    put(ws, "B2", month_key(month), fmt=KEY_FMT, color="0000FF", fill=INPUT_FILL, align="center")
    put(ws, "C2", "← możesz zmienić (np. na poprzedni miesiąc, zapisany jako rrrrmm), a tabele i wykresy przeliczą się same",
        color="808080")
    put(ws, "A3", "Poprzedni miesiąc", bold=True)
    put(ws, "B3", "=IF(MOD(B2,100)=1,B2-89,B2-1)", fmt=KEY_FMT, align="center")
    put(ws, "A4", "Filtr", bold=True)
    put(ws, "B4", filter_name or "brak")
    put(ws, "A5", "Kompletność danych", bold=True)
    put(ws, "B5", status_text)
    put(ws, "A6", "Dni w porównaniu", bold=True)
    put(ws, "B6", elapsed, fmt="0", color="0000FF", fill=INPUT_FILL, align="center")
    put(ws, "C6", "← ile pierwszych dni okresu porównujemy (poprzedni okres liczony tylko za tyle dni)",
        color="808080")

    # --- tabela kategorii
    header_row(ws, 7, ["Kategoria", "Typ", "Suma (PLN)", "Udział", "Transakcje",
                       "Poprzedni okres, te same dni (PLN)", "Zmiana (PLN)", "Zmiana %",
                       "Poprzedni okres, cały (PLN)"])
    first = 8
    total_row = first + len(cats)
    for i, (name, typ) in enumerate(cats):
        r = first + i
        put(ws, f"A{r}", name)
        put(ws, f"B{r}", TYPE_PL.get(typ, typ))
        put(ws, f"C{r}", f'=SUMIFS({col_sum},{col_cat},$A{r},{col_month},$B$2)', fmt=NUM_FMT)
        put(ws, f"D{r}", f"=IF($C${total_row}=0,0,C{r}/$C${total_row})", fmt=PCT_FMT)
        put(ws, f"E{r}", f'=COUNTIFS({col_cat},$A{r},{col_month},$B$2)', fmt="0")
        put(ws, f"F{r}", f'=SUMIFS({col_sum},{col_cat},$A{r},{col_month},$B$3,{col_day},"<="&$B$6)', fmt=NUM_FMT)
        put(ws, f"I{r}", f'=SUMIFS({col_sum},{col_cat},$A{r},{col_month},$B$3)', fmt=NUM_FMT)
        put(ws, f"G{r}", f"=C{r}-F{r}", fmt=NUM_FMT)
        put(ws, f"H{r}", f'=IF(F{r}=0,"-",G{r}/F{r})', fmt=PCT_FMT, align="right")

    last_cat = total_row - 1
    put(ws, f"A{total_row}", "Razem (bez inwestycji)", bold=True, fill=TOTAL_FILL)
    put(ws, f"B{total_row}", None, fill=TOTAL_FILL)
    for col in "CEFGI":
        fmt = "0" if col == "E" else NUM_FMT
        put(ws, f"{col}{total_row}", f"=SUM({col}{first}:{col}{last_cat})", bold=True, fmt=fmt, fill=TOTAL_FILL)
    put(ws, f"D{total_row}", f"=SUM(D{first}:D{last_cat})", bold=True, fmt=PCT_FMT, fill=TOTAL_FILL)
    put(ws, f"H{total_row}", f'=IF(F{total_row}=0,"-",G{total_row}/F{total_row})', bold=True,
        fmt=PCT_FMT, fill=TOTAL_FILL, align="right")

    # --- stałe vs zmienne i inwestycje (formuły SUMIFS po kolumnie Typ)
    r = total_row + 2
    header_row(ws, r, ["Podział", "", "Suma (PLN)", "Udział"])
    for j, (label, typ_pl) in enumerate((("Stałe", "stały"), ("Zmienne", "zmienny"))):
        rr = r + 1 + j
        put(ws, f"A{rr}", label)
        put(ws, f"C{rr}", f'=SUMIFS({col_sum},{col_type},"{typ_pl}",{col_month},$B$2)', fmt=NUM_FMT)
        put(ws, f"D{rr}", f"=IF($C${total_row}=0,0,C{rr}/$C${total_row})", fmt=PCT_FMT)
    rr = r + 3
    put(ws, f"A{rr}", "Inwestycje (osobno, poza wydatkami)")
    put(ws, f"C{rr}", f'=SUMIFS({col_sum},{col_type},"inwestycje",{col_month},$B$2)', fmt=NUM_FMT)

    # --- prognoza na koniec okresu (proste tempo: dotychczasowe wydatki / dni * długość okresu)
    r = rr + 2
    header_row(ws, r, ["Prognoza na koniec okresu", "", "Wartość"])
    put(ws, f"A{r + 1}", "Dni okresu, które minęły")
    put(ws, f"C{r + 1}", "=$B$6", fmt="0")
    put(ws, f"A{r + 2}", "Długość okresu (dni)")
    put(ws, f"C{r + 2}", length, fmt="0", color="0000FF")
    put(ws, f"A{r + 3}", "Prognoza wydatków (PLN)", bold=True)
    put(ws, f"C{r + 3}", f"=IF(C{r + 1}>=C{r + 2},C{total_row},C{total_row}/C{r + 1}*C{r + 2})",
        bold=True, fmt=NUM_FMT)
    rr = r + 3

    # --- TOP 5 (wartości z chwili generowania raportu, nie formuły)
    r = rr + 2
    put(ws, f"A{r}", "5 największych wydatków (stan z chwili generowania raportu)", bold=True)
    header_row(ws, r + 1, ["Data", "Kategoria", "Kwota (PLN)", "Opis"])
    month_rows = data[(data["period"] == month_key(month))
                      & (data["kind"] == "expense") & (data["type"] != "savings")]
    for j, row in enumerate(month_rows.nlargest(5, "spent").itertuples()):
        rr = r + 2 + j
        put(ws, f"A{rr}", row.date.strftime("%Y-%m-%d"))
        put(ws, f"B{rr}", row.category)
        put(ws, f"C{rr}", float(row.spent), fmt=NUM_FMT)
        put(ws, f"D{rr}", row.text[:70])

    # --- wykres: suma po kategoriach
    chart = BarChart()
    chart.type = "bar"
    chart.legend = None
    chart.add_data(Reference(ws, min_col=3, min_row=7, max_row=last_cat), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=first, max_row=last_cat))
    chart.y_axis.title = "PLN"
    chart.x_axis.scaling.orientation = "maxMin"   # kategorie od góry, tak jak w tabeli
    chart.y_axis.crosses = "max"                  # oś kwot na dole
    style_chart(chart, "Wydatki wg kategorii", 18, 9)
    ws.add_chart(chart, "K7")

    for col, width in zip("ABCDEFGHI", (34, 22, 16, 10, 12, 24, 16, 12, 22)):
        ws.column_dimensions[col].width = width


def write_comparison(ws, cats, periods, statuses, last_row):
    """Arkusz 'Porównanie': kategorie × miesiące (formuły SUMIFS) + wykres."""
    col_month = f"Transakcje!$B$2:$B${last_row}"
    col_cat = f"Transakcje!$C$2:$C${last_row}"
    col_sum = f"Transakcje!$E$2:$E${last_row}"

    labels = [f"{p.strftime('%Y-%m')} ({s})" for p, s in zip(periods, statuses)]
    header_row(ws, 1, ["Kategoria"] + labels)
    put(ws, "A2", "klucz miesiąca (używany w formułach)", color="808080")
    for j, p in enumerate(periods):
        put(ws, f"{get_column_letter(2 + j)}2", month_key(p), fmt=KEY_FMT, color="808080", align="center")

    first = 3
    for i, (name, _) in enumerate(cats):
        r = first + i
        put(ws, f"A{r}", name)
        for j in range(len(periods)):
            col = get_column_letter(2 + j)
            put(ws, f"{col}{r}", f"=SUMIFS({col_sum},{col_cat},$A{r},{col_month},{col}$2)", fmt=NUM_FMT)

    last_cat = first + len(cats) - 1
    total_row = last_cat + 1
    put(ws, f"A{total_row}", "Razem (bez inwestycji)", bold=True, fill=TOTAL_FILL)
    for j in range(len(periods)):
        col = get_column_letter(2 + j)
        put(ws, f"{col}{total_row}", f"=SUM({col}{first}:{col}{last_cat})", bold=True,
            fmt=NUM_FMT, fill=TOTAL_FILL)

    put(ws, f"A{total_row + 2}",
        "Miesiące oznaczone jako niepełne mają tylko część danych, więc ich sum nie porównuj 1:1 z pełnymi.",
        color="808080")

    # wykres skumulowany: każda kategoria to osobny kolor, słupki = miesiące
    chart = BarChart()
    chart.type = "col"
    chart.grouping = "stacked"
    chart.overlap = 100
    chart.add_data(Reference(ws, min_col=1, max_col=1 + len(periods), min_row=first, max_row=last_cat),
                   from_rows=True, titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=2, max_col=1 + len(periods), min_row=1, max_row=1))
    chart.y_axis.title = "PLN"
    style_chart(chart, "Wydatki w kolejnych miesiącach", 24, 11)
    ws.add_chart(chart, f"A{total_row + 4}")

    ws.column_dimensions["A"].width = 34
    for j in range(len(periods)):
        ws.column_dimensions[get_column_letter(2 + j)].width = 20
    ws.row_dimensions[1].height = 30


def write_unassigned(ws, data, month, default_category):
    """Arkusz 'Do przypisania': transakcje z kategorii 'Inne' w wybranym okresie."""
    header_row(ws, 1, ["Data", "Kwota (PLN)", "Opis"])
    rows = data[(data["period"] == month_key(month)) & (data["category"] == default_category)]
    rows = rows.sort_values("spent", ascending=False)
    if rows.empty:
        put(ws, "A2", "Wszystkie transakcje zostały rozpoznane.")
    for i, r in enumerate(rows.itertuples(), start=2):
        put(ws, f"A{i}", r.date.strftime("%Y-%m-%d"))
        put(ws, f"B{i}", float(r.spent), fmt=NUM_FMT)
        put(ws, f"C{i}", r.text[:120])
    put(ws, f"A{len(rows) + 4}", "Dopisz słowa kluczowe do config/rules.json, żeby je skategoryzować.", color="808080")
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 90


# ---------------------------------------------------------------- główna funkcja
def create_monthly_report(base_dir, data, categories, settings, month, filter_name, data_start, data_end,
                          pmap=None):
    """
    Tworzy plik output/<rok>/<rrrr-mm>[_filtr].xlsx i zwraca jego ścieżkę.

    data       - skategoryzowane transakcje (po filtrze), wszystkie miesiące
    month      - pd.Period, wybrany miesiąc
    data_start, data_end - zakres dat w eksporcie (do oznaczania niepełnych okresów)
    pmap       - PeriodMap (kalendarzowy / od wypłaty); domyślnie kalendarzowy
    """
    if pmap is None:
        pmap = PeriodMap("calendar", data_start, data_end)
    data = assign_periods(data, pmap)

    # miesiące w oknie porównania: N poprzednich + wybrany
    n = settings["compare_months"]
    periods = [month - k for k in range(n, -1, -1)]
    window = data[data["period"].isin([month_key(p) for p in periods])].sort_values("date", kind="stable")
    if window.empty:
        raise ValueError("Brak transakcji do raportu.")

    # kategorie do pokazania: z config, tylko te z danymi, bez inwestycji
    present = set(window["category"])
    cats = [(c["name"], c["type"]) for c in categories["categories"]
            if c["name"] in present and c["type"] != "savings"]
    statuses = [pmap.status(p) for p in periods]

    # opis kompletności danych: wybrany i poprzedni okres
    prev_status = pmap.status(month - 1)
    status_text = (f"{month}: {statuses[-1]}, poprzedni okres: {prev_status} "
                   f"(dane w eksporcie: {data_start:%Y-%m-%d} - {data_end:%Y-%m-%d}; "
                   f"tryb: {pmap.mode_name}, okres {pmap.describe(month)})")

    # w porównaniu pomijamy miesiące bez żadnych danych (poza wybranym)
    keep = [i for i, (p, s) in enumerate(zip(periods, statuses)) if s != "brak danych" or p == month]
    periods = [periods[i] for i in keep]
    statuses = [statuses[i] for i in keep]

    last_row = len(window) + 1  # ostatni wiersz danych w arkuszu Transakcje

    wb = Workbook()
    ws_sum = wb.active
    ws_sum.title = "Podsumowanie"
    ws_cmp = wb.create_sheet("Porównanie")
    ws_tr = wb.create_sheet("Transakcje")
    ws_un = wb.create_sheet("Do przypisania")

    write_transactions(ws_tr, window)
    write_summary(ws_sum, window, cats, month, filter_name, status_text, last_row,
                  pmap.elapsed(month), pmap.length(month))
    write_comparison(ws_cmp, cats, periods, statuses, last_row)
    write_unassigned(ws_un, window, month, settings["default_category"])

    out_dir = Path(base_dir) / "output" / str(month.year)
    out_dir.mkdir(parents=True, exist_ok=True)
    name = f"{month.strftime('%Y-%m')}" + (f"_{filter_name}" if filter_name else "") + ".xlsx"
    out_path = out_dir / name
    wb.save(out_path)
    return out_path
