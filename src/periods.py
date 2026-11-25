"""
periods.py - okresy rozliczeniowe.

Raport może liczyć "miesiąc" na dwa sposoby (settings.json -> "period_mode"
albo flaga --calendar / --payday w main.py):

    calendar - zwykły miesiąc kalendarzowy (1. - ostatni dzień miesiąca)
    payday   - okres od dnia wypłaty do dnia przed następną wypłatą

Okres ma etykietę w postaci miesiąca (Period 2026-11). W trybie "payday" jest to miesiąc,
w którym przyszła wypłata: okres "2026-11" trwa od wypłaty z listopada do dnia przed wypłatą
z grudnia.

Co robi ten plik:
1. find_paydays   - wykrywa wypłaty w danych (duży przelew przychodzący)
2. PeriodMap      - zna granice okresów, ich długość, kompletność i liczbę dni, które minęły
3. assign_periods - dopisuje do transakcji kolumny:
       period      - klucz okresu jako liczba rrrrmm (np. 202611)
       period_day  - który dzień okresu to jest (1 = pierwszy dzień okresu)
"""
import bisect

import pandas as pd

# domyślne kryteria wypłaty (można nadpisać w settings.json -> "payday")
DEFAULT_PAYDAY = {
    "min_amount": 2000,                       # wypłata to przelew od tej kwoty wzwyż
    "bank_types": ["Przelew przychodzący"],   # typ transakcji z banku (pusta lista = dowolny)
    "keywords": [],                           # opcjonalnie: słowa w opisie, np. ["wynagrodzenie"]
}

MODES = {"calendar": "kalendarzowy", "payday": "od wypłaty"}


def month_key(period):
    """Period('2026-11') -> 202611."""
    return period.year * 100 + period.month


def find_paydays(raw, settings):
    """
    Zwraca słownik {Period miesiąca: data wypłaty}.
    W każdym miesiącu kalendarzowym bierzemy NAJWIĘKSZY pasujący przelew.
    """
    cfg = {**DEFAULT_PAYDAY, **settings.get("payday", {})}
    income = raw[raw["amount"] >= cfg["min_amount"]]
    if cfg["bank_types"]:
        income = income[income["bank_type"].isin(cfg["bank_types"])]
    if cfg["keywords"]:
        income = income[income["text"].map(lambda t: any(k in t for k in cfg["keywords"]))]
    if income.empty:
        return {}

    income = income.assign(month=income["date"].dt.to_period("M"))
    best = income.sort_values("amount", ascending=False).drop_duplicates("month")
    return {row["month"]: row["date"].normalize() for _, row in best.iterrows()}


class PeriodMap:
    """Wie, kiedy zaczyna się i kończy każdy okres oraz do którego okresu należy dana data."""

    def __init__(self, mode, data_start, data_end, paydays=None):
        if mode not in MODES:
            raise ValueError(f"Nieznany tryb okresu '{mode}'. Dozwolone: {list(MODES)}")
        self.mode = mode
        self.data_start = pd.Timestamp(data_start).normalize()
        self.data_end = pd.Timestamp(data_end).normalize()
        self.paydays = dict(sorted((paydays or {}).items()))
        self.starts = sorted(self.paydays.values())
        if mode == "payday" and not self.starts:
            raise ValueError(
                "Tryb 'od wypłaty': nie znaleziono żadnej wypłaty w danych. Sprawdź w settings.json "
                "sekcję 'payday' (min_amount, bank_types, keywords) albo użyj trybu kalendarzowego."
            )

    @property
    def mode_name(self):
        return MODES[self.mode]

    # ------------------------------------------------------------ granice
    def label_of(self, date):
        """Do jakiego okresu (Period) należy data."""
        date = pd.Timestamp(date).normalize()
        if self.mode == "calendar":
            return date.to_period("M")
        i = bisect.bisect_right(self.starts, date) - 1
        if i < 0:                                   # przed pierwszą znaną wypłatą
            return self.starts[0].to_period("M") - 1
        return self.starts[i].to_period("M")

    def bounds(self, label):
        """
        (początek, koniec) okresu jako daty włącznie.
        None oznacza: nieznany początek albo otwarty koniec (okres jeszcze trwa).
        (None, None) = tego okresu nie ma w danych.
        """
        if self.mode == "calendar":
            return label.start_time.normalize(), label.end_time.normalize()

        start = self.paydays.get(label)
        if start is None:
            if label == self.starts[0].to_period("M") - 1:   # okres sprzed pierwszej wypłaty
                return None, self.starts[0] - pd.Timedelta(days=1)
            return None, None
        i = self.starts.index(start)
        end = self.starts[i + 1] - pd.Timedelta(days=1) if i + 1 < len(self.starts) else None
        return start, end

    # ------------------------------------------------------------ miary
    def status(self, label):
        """pełny / niepełny / brak danych (czy eksport obejmuje cały okres)."""
        start, end = self.bounds(label)
        if start is None and end is None:
            return "brak danych"
        if (start is not None and self.data_end < start) or (end is not None and self.data_start > end):
            return "brak danych"
        if start is not None and end is not None and self.data_start <= start and self.data_end >= end:
            return "pełny"
        return "niepełny"

    def elapsed(self, label):
        """Ile dni okresu już minęło (do końca danych albo do końca okresu)."""
        start, end = self.bounds(label)
        first = start if start is not None else self.data_start
        last = min(end, self.data_end) if end is not None else self.data_end
        return max(1, (last - first).days + 1)

    def length(self, label):
        """Długość okresu w dniach. Dla okresu otwartego: długość poprzedniego (albo 30)."""
        if self.mode == "calendar":
            return label.days_in_month
        start, end = self.bounds(label)
        if start is not None and end is not None:
            return (end - start).days + 1
        prev_start, prev_end = self.bounds(label - 1)
        if prev_start is not None and prev_end is not None:
            return (prev_end - prev_start).days + 1
        return 30

    def day_index(self, date, label):
        """Który dzień okresu to jest (1 = pierwszy dzień)."""
        start, _ = self.bounds(label)
        first = start if start is not None else self.data_start
        return (pd.Timestamp(date).normalize() - first).days + 1

    def describe(self, label):
        """Tekst typu '2026-11-03 - otwarty' do pokazania w raporcie."""
        start, end = self.bounds(label)
        s = f"{start:%Y-%m-%d}" if start is not None else "?"
        e = f"{end:%Y-%m-%d}" if end is not None else "otwarty"
        return f"{s} - {e}"


def assign_periods(df, pmap):
    """Dopisuje do transakcji kolumny 'period' (rrrrmm) i 'period_day' (dzień okresu)."""
    df = df.copy()
    labels = df["date"].map(pmap.label_of)
    df["period"] = labels.map(month_key).astype(int) if len(df) else pd.Series(dtype=int)
    df["period_day"] = [pmap.day_index(d, lab) for d, lab in zip(df["date"], labels)]
    return df
