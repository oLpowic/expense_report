"""
conftest.py - wspólne narzędzia do testów.

Wszystkie dane w testach są SYNTETYCZNE (zmyślone) - nie ma tu żadnych prawdziwych transakcji.
Konfigurację bierzemy z config/template/, więc testy sprawdzają też szablon.
"""
import shutil
from datetime import datetime
from pathlib import Path

import pandas as pd
import pytest

from src.categorize import categorize
from src.loader import load_expenses, load_json

ROOT = Path(__file__).resolve().parent.parent
EXAMPLE = ROOT / "config" / "template"

# Testy korzystają z szablonu config/template/. Bez niego każdy test kończyłby się osobnym,
# mało czytelnym błędem, więc przerywamy od razu z jasnym komunikatem.
_REQUIRED = ["settings.json", "categories.json", "rules.json", "filters"]
_missing = [name for name in _REQUIRED if not (EXAMPLE / name).exists()]
if _missing:
    pytest.exit(
        f"\nBrak szablonu konfiguracji: {EXAMPLE}\n"
        f"Nie znaleziono: {', '.join(_missing)}\n"
        "Folder config/template/ musi zawierać settings.json, categories.json, rules.json "
        "oraz podfolder filters/ z plikami .json (są w repozytorium / w paczce expense_report.zip).\n"
        "Nie kopiuj tu swojego config/, bo trafią do szablonu Twoje osobiste reguły.",
        returncode=2,
    )

# nagłówki takie jak w eksporcie z banku
COLUMNS = ["Data transakcji", "Data zaksięgowania", "Data odrzucenia", "Kwota", "Waluta", "Nadawca",
           "Odbiorca", "Opis", "Produkt", "Typ transakcji", "Kwota zlecenia", "Waluta zlecenia",
           "Status", "Saldo po transakcji"]


def make_row(date, amount, desc, recipient=None, typ="Transakcja kartą", status="Zrealizowana",
             rejected=None, balance=1000.0, currency="PLN"):
    """Jeden wiersz eksportu (słownik z nagłówkami jak w banku)."""
    return {
        "Data transakcji": date, "Data zaksięgowania": date, "Data odrzucenia": rejected,
        "Kwota": amount, "Waluta": currency, "Nadawca": "Jan Testowy", "Odbiorca": recipient,
        "Opis": desc, "Produkt": "Konto", "Typ transakcji": typ, "Kwota zlecenia": amount,
        "Waluta zlecenia": currency, "Status": status, "Saldo po transakcji": balance,
    }


def write_export(path, rows):
    """Zapisuje listę wierszy jako plik xlsx z nagłówkami jak w banku."""
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    pd.DataFrame(rows, columns=COLUMNS).to_excel(path, index=False)


def d(text):
    """'2026-11-04' -> datetime (Excel zapisze to jako prawdziwą datę, jak w eksporcie banku)."""
    return datetime.strptime(text, "%Y-%m-%d")


def sample_rows():
    """Zestaw transakcji za październik i listopad 2026 z różnymi przypadkami brzegowymi."""
    rows = [
        # październik
        make_row(d("2026-10-28"), -1500.00, "Czesne za semestr", "Uczelnia X", typ="Przelew wychodzący"),
        make_row(d("2026-10-29"), -100.00, "BIEDRONKA 123 Szczecin"),
        # listopad
        make_row(d("2026-11-02"), -2100.00, "Czynsz listopad", "Wspólnota", typ="Przelew wychodzący"),
        make_row(d("2026-11-03"), 5000.00, "Wypłata", "Pracodawca", typ="Przelew przychodzący"),
        make_row(d("2026-11-04"), -230.50, "LIDL Szczecin"),
        make_row(d("2026-11-05"), 30.00, "RESERVED zwrot"),                    # zwrot kartą
        make_row(d("2026-11-06"), -89.90, "WOLT dostawa", typ="Transakcja BLIK"),
        make_row(d("2026-11-07"), -500.00, "BOSSA wpłata", typ="Przelew wychodzący"),  # inwestycje
        make_row(d("2026-11-08"), -45.00, "Netflix"),
        make_row(d("2026-11-09"), -77.00, "SKLEP XYZ"),                        # nierozpoznany
        make_row(d("2026-11-10"), -300.00, "ZALANDO", status="Odrzucona"),     # pomijany
        make_row(d("2026-11-11"), -50.00, "KOLEO bilety kolejowe"),
        make_row(d("2026-11-12"), -20.00, "Blokada kartą", typ="Blokada środków",
                 status="W trakcie realizacji", balance=None),                  # blokada, pomijana
    ]
    # różne salda, żeby każdy wiersz miał inny "odcisk palca"
    for i, row in enumerate(rows):
        if row["Saldo po transakcji"] is not None:
            row["Saldo po transakcji"] = 5000.0 - 10 * i
    return rows


@pytest.fixture
def settings():
    return load_json(EXAMPLE / "settings.json")


@pytest.fixture
def categories():
    return load_json(EXAMPLE / "categories.json")


@pytest.fixture
def rules():
    return load_json(EXAMPLE / "rules.json")


@pytest.fixture
def project(tmp_path):
    """Tymczasowy projekt: config skopiowany z szablonu + jeden eksport z danymi testowymi."""
    shutil.copytree(EXAMPLE, tmp_path / "config")
    write_export(tmp_path / "data" / "exports" / "export.xlsx", sample_rows())
    return tmp_path


@pytest.fixture
def pipeline(project, settings, categories, rules):
    """Wczytane i skategoryzowane dane testowe + zakres dat eksportu."""
    raw = load_expenses(project, settings)
    df = categorize(raw, rules, categories, settings)
    return {
        "base": project, "raw": raw, "df": df, "settings": settings, "categories": categories,
        "data_start": raw["date"].min().normalize(), "data_end": raw["date"].max().normalize(),
    }
