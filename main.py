"""
main.py - jedyny plik, który uruchamiasz.

Użycie:
    python main.py                      -> raport miesięczny za ostatni miesiąc z danych
    python main.py 2026-11              -> raport miesięczny za wybrany miesiąc
    python main.py 2026-11 everyday     -> raport miesięczny + filtr z config/filters/
    python main.py year                 -> raport roczny za rok ostatniej transakcji
    python main.py year 2026            -> raport roczny za wybrany rok
    python main.py year 2026 everyday   -> raport roczny + filtr
    python main.py help                 -> ta pomoc

Raporty miesięczne: output/<rok>/<rrrr-mm>.xlsx
Raporty roczne:     output/<rok>_roczny.xlsx
W styczniu raport roczny za poprzedni rok powstaje sam, jeśli jeszcze go nie ma
(można to wyłączyć w config/settings.json: "auto_yearly_in_january").
"""
import sys
from datetime import date
from pathlib import Path

import pandas as pd

from src.categorize import apply_filter, categorize, load_filter
from src.loader import load_expenses, load_json
from src.report_monthly import create_monthly_report
from src.report_yearly import create_yearly_report

BASE = Path(__file__).parent  # folder, w którym leży main.py


def maybe_auto_yearly(df, categories, settings, data_start, data_end):
    """W styczniu generuje raport roczny za poprzedni rok, jeśli jeszcze go nie ma."""
    today = date.today()
    if not settings["auto_yearly_in_january"] or today.month != 1:
        return
    year = today.year - 1
    path = BASE / "output" / f"{year}_roczny.xlsx"
    if path.exists() or not (df["date"].dt.year == year).any():
        return
    path = create_yearly_report(BASE, df, categories, settings, year, None, data_start, data_end)
    print(f"Styczeń: automatycznie utworzono raport roczny za {year}: {path}")


def main():
    args = sys.argv[1:]

    # pomoc: python main.py help
    if args and args[0] in ("help", "-h", "--help"):
        print(__doc__)
        return

    yearly = bool(args) and args[0] == "year"

    # 1) argumenty: "year [rok] [filtr]" albo "[miesiąc] [filtr]"
    if yearly:
        year_arg = args[1] if len(args) > 1 else None
        filter_name = args[2] if len(args) > 2 else None
        month_arg = None
    else:
        month_arg = args[0] if args else None
        filter_name = args[1] if len(args) > 1 else None

    # 2) konfiguracja (przy pierwszym uruchomieniu trzeba skopiować szablon)
    if not (BASE / "config" / "settings.json").exists():
        sys.exit("Brak folderu config/. Skopiuj folder config.example/ jako config/ "
                 "(cp -r config.example config) i dostosuj reguły do siebie.")
    settings = load_json(BASE / "config" / "settings.json")
    categories = load_json(BASE / "config" / "categories.json")
    rules = load_json(BASE / "config" / "rules.json")

    # 3) wczytanie danych; zapamiętujemy zakres dat całego eksportu
    raw = load_expenses(BASE, settings)
    data_start, data_end = raw["date"].min().normalize(), raw["date"].max().normalize()

    # 4) kategoryzacja i filtr (na wszystkich danych, bo potrzebne do porównań)
    df = categorize(raw, rules, categories, settings)
    unfiltered = df
    if filter_name:
        df = apply_filter(df, load_filter(BASE, filter_name))

    # 5a) raport roczny
    if yearly:
        year = int(year_arg) if year_arg else data_end.year
        path = create_yearly_report(BASE, df, categories, settings, year, filter_name, data_start, data_end)
        spending = df[(df["date"].dt.year == year) & (df["type"] != "savings")]
        print(f"\nRok {year}: wydatki {spending['spent'].sum():.2f} {categories['currency']} "
              f"({len(spending)} transakcji)")
        print(f"Zapisano: {path}")
        return

    # 5b) raport miesięczny (domyślnie ostatni miesiąc z danych)
    month = pd.Period(month_arg or data_end.strftime("%Y-%m"), freq="M")
    if df[df["date"].dt.to_period("M") == month].empty:
        sys.exit(f"Brak transakcji w miesiącu {month} (dla podanych kryteriów).")
    path = create_monthly_report(BASE, df, categories, settings, month, filter_name, data_start, data_end)

    this_month = df[df["date"].dt.to_period("M") == month]
    spending = this_month[this_month["type"] != "savings"]
    other = (spending["category"] == settings["default_category"]).sum()
    print(f"\nMiesiąc {month}: wydatki {spending['spent'].sum():.2f} {categories['currency']} "
          f"({len(spending)} transakcji, w kategorii '{settings['default_category']}': {other})")
    print(f"Zapisano: {path}")

    # 6) w styczniu: automatyczny raport roczny za poprzedni rok (bez filtra)
    maybe_auto_yearly(unfiltered, categories, settings, data_start, data_end)


if __name__ == "__main__":
    main()
