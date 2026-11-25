"""Testy modułu okresów (src/periods.py)."""
import pandas as pd
import pytest

from src.periods import PeriodMap, assign_periods, find_paydays


def make_raw(rows):
    """rows: lista (data, kwota, typ, tekst)."""
    return pd.DataFrame(
        [{"date": pd.Timestamp(d), "amount": a, "bank_type": t, "text": x} for d, a, t, x in rows]
    )


@pytest.fixture
def raw():
    return make_raw([
        ("2026-09-03", 5400.0, "Przelew przychodzący", "firma"),
        ("2026-10-05", 5443.0, "Przelew przychodzący", "firma"),
        ("2026-10-20", 3000.0, "Przelew przychodzący", "zwrot od kolegi"),   # mniejszy, ten sam miesiąc
        ("2026-11-03", 5443.96, "Przelew przychodzący", "firma"),
        ("2026-11-09", 87.93, "Transakcja kartą", "zwrot"),                   # za mała kwota
        ("2026-11-12", 2500.0, "Transakcja kartą", "zwrot duży"),             # zły typ
    ])


def test_find_paydays_takes_largest_per_month(raw):
    paydays = find_paydays(raw, {})
    assert paydays == {
        pd.Period("2026-09"): pd.Timestamp("2026-09-03"),
        pd.Period("2026-10"): pd.Timestamp("2026-10-05"),
        pd.Period("2026-11"): pd.Timestamp("2026-11-03"),
    }


def test_find_paydays_respects_settings(raw):
    paydays = find_paydays(raw, {"payday": {"min_amount": 100, "bank_types": []}})
    assert paydays[pd.Period("2026-11")] == pd.Timestamp("2026-11-03")
    assert find_paydays(raw, {"payday": {"min_amount": 99999}}) == {}


def test_find_paydays_keywords(raw):
    paydays = find_paydays(raw, {"payday": {"keywords": ["zwrot od"], "min_amount": 1000}})
    assert list(paydays.values()) == [pd.Timestamp("2026-10-20")]


def test_calendar_mode_bounds_and_labels():
    pmap = PeriodMap("calendar", "2026-10-25", "2026-11-23")
    assert pmap.label_of("2026-11-15") == pd.Period("2026-11")
    assert pmap.bounds(pd.Period("2026-02")) == (pd.Timestamp("2026-02-01"), pd.Timestamp("2026-02-28"))
    assert pmap.length(pd.Period("2026-02")) == 28
    assert pmap.status(pd.Period("2026-10")) == "niepełny"
    assert pmap.status(pd.Period("2026-11")) == "niepełny"
    assert pmap.status(pd.Period("2026-08")) == "brak danych"
    assert pmap.elapsed(pd.Period("2026-11")) == 23
    assert pmap.elapsed(pd.Period("2026-10")) == 31       # dzień okresu liczony od 1., nie od początku danych


def test_calendar_full_month_status():
    pmap = PeriodMap("calendar", "2026-10-01", "2026-12-15")
    assert pmap.status(pd.Period("2026-10")) == "pełny"
    assert pmap.status(pd.Period("2026-11")) == "pełny"
    assert pmap.status(pd.Period("2026-12")) == "niepełny"


def test_payday_mode_bounds(raw):
    paydays = find_paydays(raw, {})
    pmap = PeriodMap("payday", "2026-09-01", "2026-11-23", paydays)
    # okres 2026-10: od 5.10 do dnia przed 3.11
    assert pmap.bounds(pd.Period("2026-10")) == (pd.Timestamp("2026-10-05"), pd.Timestamp("2026-11-02"))
    assert pmap.length(pd.Period("2026-10")) == 29
    assert pmap.status(pd.Period("2026-10")) == "pełny"
    # okres 2026-11 jest otwarty
    assert pmap.bounds(pd.Period("2026-11")) == (pd.Timestamp("2026-11-03"), None)
    assert pmap.status(pd.Period("2026-11")) == "niepełny"
    assert pmap.elapsed(pd.Period("2026-11")) == 21        # 3-23 listopada
    assert pmap.length(pd.Period("2026-11")) == 29         # bierze długość poprzedniego
    assert pmap.describe(pd.Period("2026-11")).endswith("otwarty")


def test_payday_labels_and_before_first_payday(raw):
    pmap = PeriodMap("payday", "2026-08-20", "2026-11-23", find_paydays(raw, {}))
    assert pmap.label_of("2026-11-02") == pd.Period("2026-10")
    assert pmap.label_of("2026-11-03") == pd.Period("2026-11")
    assert pmap.label_of("2026-09-02") == pd.Period("2026-08")   # przed pierwszą wypłatą
    assert pmap.bounds(pd.Period("2026-08")) == (None, pd.Timestamp("2026-09-02"))
    assert pmap.status(pd.Period("2026-08")) == "niepełny"       # nie znamy początku
    assert pmap.status(pd.Period("2025-01")) == "brak danych"


def test_payday_mode_without_paydays_fails_clearly():
    with pytest.raises(ValueError, match="nie znaleziono żadnej wypłaty"):
        PeriodMap("payday", "2026-10-01", "2026-11-01", {})


def test_unknown_mode():
    with pytest.raises(ValueError, match="Nieznany tryb"):
        PeriodMap("weekly", "2026-10-01", "2026-11-01")


def test_assign_periods_calendar():
    df = make_raw([("2026-11-01", -10, "x", "a"), ("2026-11-15", -20, "x", "b"), ("2026-12-01", -5, "x", "c")])
    out = assign_periods(df, PeriodMap("calendar", "2026-11-01", "2026-12-01"))
    assert list(out["period"]) == [202611, 202611, 202612]
    assert list(out["period_day"]) == [1, 15, 1]


def test_assign_periods_payday(raw):
    pmap = PeriodMap("payday", "2026-09-01", "2026-11-23", find_paydays(raw, {}))
    df = make_raw([("2026-11-02", -10, "x", "a"), ("2026-11-03", -20, "x", "b"), ("2026-11-10", -5, "x", "c")])
    out = assign_periods(df, pmap)
    assert list(out["period"]) == [202610, 202611, 202611]
    assert list(out["period_day"]) == [29, 1, 8]


def test_assign_periods_empty():
    empty = pd.DataFrame({"date": pd.to_datetime([]), "amount": [], "bank_type": [], "text": []})
    out = assign_periods(empty, PeriodMap("calendar", "2026-01-01", "2026-01-31"))
    assert len(out) == 0
