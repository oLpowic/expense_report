"""Testy loader.py: wczytywanie, czyszczenie, scalanie i duplikaty."""
import pandas as pd
import pytest

from src.loader import load_expenses, parse_dates, read_one_file, to_number
from tests.conftest import d, make_row, write_export


# ---------------------------------------------------------------- to_number
@pytest.mark.parametrize("value, expected", [
    ("-1 234,50", -1234.5),
    ("-1 234,50", -1234.5),      # twarda spacja
    ("−45,00", -45.0),           # unicode minus
    ("12,5", 12.5),
    (7, 7.0),
    (-3.25, -3.25),
])
def test_to_number_parses_polish_formats(value, expected):
    assert to_number(value) == pytest.approx(expected)


@pytest.mark.parametrize("value", [None, float("nan"), "abc", ""])
def test_to_number_returns_none_for_garbage(value):
    assert to_number(value) is None


# ---------------------------------------------------------------- parse_dates
def test_parse_dates_iso_is_not_mangled():
    """2026-11-02 to 2 listopada (nie 11 lutego)."""
    result = parse_dates(pd.Series(["2026-11-02"]))
    assert result.iloc[0] == pd.Timestamp(2026, 11, 2)


def test_parse_dates_polish_format_and_time_suffix():
    result = parse_dates(pd.Series(["20.11.2026", "2026-11-20 13:45:00"]))
    assert list(result) == [pd.Timestamp(2026, 11, 20)] * 2


def test_parse_dates_keeps_datetime_series():
    series = pd.Series([pd.Timestamp(2026, 1, 5)])
    assert parse_dates(series).iloc[0] == pd.Timestamp(2026, 1, 5)


def test_parse_dates_invalid_becomes_nat():
    assert pd.isna(parse_dates(pd.Series(["nie data"])).iloc[0])


# ---------------------------------------------------------------- read_one_file
def read(tmp_path, rows, settings):
    path = tmp_path / "a.xlsx"
    write_export(path, rows)
    return read_one_file(path, settings)


def test_rejected_and_pending_are_skipped(tmp_path, settings):
    rows = [
        make_row(d("2026-11-01"), -10.0, "ok"),
        make_row(d("2026-11-02"), -20.0, "z datą odrzucenia", rejected=d("2026-11-03")),
        make_row(d("2026-11-03"), -30.0, "odrzucona statusem", status="Odrzucona"),
        make_row(d("2026-11-04"), -40.0, "anulowana", status="Anulowana"),
        make_row(d("2026-11-05"), -50.0, "blokada", status="W trakcie realizacji"),
    ]
    df = read(tmp_path, rows, settings)
    assert list(df["amount"]) == [-10.0]


def test_pending_can_be_included(tmp_path, settings):
    settings = dict(settings, include_pending=True)
    rows = [make_row(d("2026-11-05"), -50.0, "blokada", status="W trakcie realizacji")]
    assert len(read(tmp_path, rows, settings)) == 1


def test_text_is_lowercase_and_merges_recipient_and_description(tmp_path, settings):
    rows = [make_row(d("2026-11-01"), -10.0, "Zakupy  w   SKLEPIE", recipient="Sklep ABC")]
    df = read(tmp_path, rows, settings)
    assert df["text"].iloc[0] == "sklep abc zakupy w sklepie"


def test_account_numbers_and_card_masks_are_removed(tmp_path, settings):
    desc = "Przelew na konto 12345678901234567890123456 oraz PL12345678901234567890123456 karta 516842------9439 sklep"
    df = read(tmp_path, [make_row(d("2026-11-01"), -10.0, desc)], settings)
    text = df["text"].iloc[0]
    assert "1234567890" not in text
    assert "516842" not in text
    assert "sklep" in text


def test_rows_without_valid_date_or_amount_are_dropped(tmp_path, settings, capsys):
    rows = [
        make_row(d("2026-11-01"), -10.0, "dobry"),
        make_row("to nie data", -20.0, "zła data"),
        make_row(d("2026-11-03"), "to nie kwota", "zła kwota"),
    ]
    df = read(tmp_path, rows, settings)
    assert list(df["amount"]) == [-10.0]
    assert "pominięto 2 wierszy" in capsys.readouterr().out


def test_foreign_currency_gives_warning(tmp_path, settings, capsys):
    rows = [make_row(d("2026-11-01"), -10.0, "w euro", currency="EUR")]
    read(tmp_path, rows, settings)
    assert "nie jest w PLN" in capsys.readouterr().out


def test_missing_required_column_gives_helpful_error(tmp_path, settings):
    path = tmp_path / "bad.xlsx"
    pd.DataFrame({"Coś": [1], "Innego": [2]}).to_excel(path, index=False)
    with pytest.raises(ValueError) as err:
        read_one_file(path, settings)
    assert "Data transakcji" in str(err.value)
    assert "Coś" in str(err.value)       # podaje, jakie kolumny faktycznie są w pliku


def test_column_names_with_extra_spaces_are_handled(tmp_path, settings):
    path = tmp_path / "spaces.xlsx"
    df = pd.DataFrame([make_row(d("2026-11-01"), -10.0, "x")])
    df.columns = [f" {c} " for c in df.columns]
    df.to_excel(path, index=False)
    assert len(read_one_file(path, settings)) == 1


# ---------------------------------------------------------------- load_expenses
def test_no_files_raises(tmp_path, settings):
    (tmp_path / "data" / "exports").mkdir(parents=True)
    with pytest.raises(FileNotFoundError):
        load_expenses(tmp_path, settings)


def test_overlapping_exports_are_not_double_counted(tmp_path, settings):
    first = [make_row(d("2026-11-01"), -10.0, "a", balance=100.0),
             make_row(d("2026-11-02"), -20.0, "b", balance=80.0)]
    second = first + [make_row(d("2026-11-03"), -30.0, "c", balance=50.0)]   # nakłada się + nowy wiersz
    write_export(tmp_path / "data" / "exports" / "a.xlsx", first)
    write_export(tmp_path / "data" / "exports" / "b.xlsx", second)
    df = load_expenses(tmp_path, settings)
    assert sorted(df["amount"]) == [-30.0, -20.0, -10.0]


def test_identical_rows_inside_one_file_are_kept(tmp_path, settings):
    """Dwie takie same płatności w jednym pliku to dwie prawdziwe płatności (np. 2 × 89 zł)."""
    rows = [make_row(d("2026-11-09"), -89.0, "subskrypcja", balance=500.0),
            make_row(d("2026-11-09"), -89.0, "subskrypcja", balance=500.0)]
    write_export(tmp_path / "data" / "exports" / "a.xlsx", rows)
    assert len(load_expenses(tmp_path, settings)) == 2


def test_same_amount_different_balance_across_files_is_kept(tmp_path, settings):
    """Dwa zakupy z tą samą kwotą i opisem, ale innym saldem, to różne transakcje."""
    write_export(tmp_path / "data" / "exports" / "a.xlsx",
                 [make_row(d("2026-11-09"), -10.0, "kawa", balance=500.0)])
    write_export(tmp_path / "data" / "exports" / "b.xlsx",
                 [make_row(d("2026-11-09"), -10.0, "kawa", balance=400.0)])
    assert len(load_expenses(tmp_path, settings)) == 2


def test_result_is_sorted_by_date_and_has_source_column(tmp_path, settings):
    rows = [make_row(d("2026-11-05"), -1.0, "późno", balance=1.0),
            make_row(d("2026-11-01"), -2.0, "wcześnie", balance=2.0)]
    write_export(tmp_path / "data" / "exports" / "a.xlsx", rows)
    df = load_expenses(tmp_path, settings)
    assert list(df["date"]) == sorted(df["date"])
    assert set(df["source"]) == {"a.xlsx"}


def test_excel_lock_files_are_ignored(tmp_path, settings):
    write_export(tmp_path / "data" / "exports" / "a.xlsx", [make_row(d("2026-11-01"), -1.0, "x")])
    (tmp_path / "data" / "exports" / "~$a.xlsx").write_bytes(b"to nie jest xlsx")
    assert len(load_expenses(tmp_path, settings)) == 1
