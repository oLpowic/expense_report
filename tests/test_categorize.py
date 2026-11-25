"""Testy categorize.py: reguły, zwroty, typy kategorii i filtry."""
import json

import pandas as pd
import pytest

from src.categorize import apply_filter, categorize, find_category, load_filter


def frame(rows):
    """Minimalna tabela jak z loadera: (amount, text, bank_type)."""
    return pd.DataFrame(
        [{"date": pd.Timestamp(2026, 11, 1), "amount": a, "text": t, "bank_type": b, "balance": 0.0}
         for a, t, b in rows])


# ---------------------------------------------------------------- find_category
def test_first_matching_rule_wins():
    rules = {"rules": [
        {"category": "A", "keywords": ["abc"]},
        {"category": "B", "keywords": ["abc", "xyz"]},
    ]}
    assert find_category("coś abc", rules, "Inne") == "A"
    assert find_category("xyz", rules, "Inne") == "B"


def test_default_category_when_nothing_matches():
    assert find_category("nic", {"rules": [{"category": "A", "keywords": ["abc"]}]}, "Inne") == "Inne"


@pytest.mark.parametrize("text, expected", [
    ("koleo bilety kolejowe", "Podróże"),         # regresja: "bilety" nie może kraść do Rozrywki
    ("biedronka 123 szczecin", "Zakupy spożywcze"),
    ("wolt dostawa", "Jedzenie na mieście"),
    ("uber eats zamówienie", "Jedzenie na mieście"),   # szczegółowa reguła przed ogólnym "uber"
    ("uber trip", "Transport"),
    ("czesne za semestr", "Edukacja"),
    ("bossa wpłata", "Inwestycje"),
    ("netflix", "Subskrypcje i opłaty"),
    ("zupełnie nieznany sklep", "Inne"),
])
def test_template_rules_assign_expected_categories(text, expected, rules):
    assert find_category(text, rules, "Inne") == expected


# ---------------------------------------------------------------- categorize
def test_expenses_get_positive_spent_and_refunds_negative(rules, categories, settings):
    df = categorize(frame([(-100.0, "biedronka", "Transakcja kartą"),
                           (30.0, "reserved zwrot", "Transakcja kartą")]),
                    rules, categories, settings)
    by_kind = df.set_index("kind")["spent"]
    assert by_kind["expense"] == 100.0
    assert by_kind["refund"] == -30.0


def test_refund_reduces_its_category_total(rules, categories, settings):
    df = categorize(frame([(-100.0, "reserved sukienka", "Transakcja kartą"),
                           (40.0, "reserved zwrot", "Transakcja kartą")]),
                    rules, categories, settings)
    assert df.groupby("category")["spent"].sum()["Ubrania i zakupy"] == pytest.approx(60.0)


def test_other_incoming_money_is_ignored(rules, categories, settings):
    df = categorize(frame([(5000.0, "wypłata", "Przelew przychodzący"),
                           (94.0, "zwrot od znajomego", "Przelew na telefon")]),
                    rules, categories, settings)
    assert df.empty


def test_blik_refund_is_netted(rules, categories, settings):
    df = categorize(frame([(25.0, "pyszne zwrot", "Transakcja BLIK")]), rules, categories, settings)
    assert df["spent"].iloc[0] == -25.0


def test_refunds_can_be_switched_off(rules, categories, settings):
    settings = dict(settings, net_refunds=False)
    df = categorize(frame([(-10.0, "biedronka", "Transakcja kartą"),
                           (30.0, "reserved zwrot", "Transakcja kartą")]),
                    rules, categories, settings)
    assert list(df["kind"]) == ["expense"]


def test_category_type_comes_from_categories_json(rules, categories, settings):
    df = categorize(frame([(-10.0, "czynsz", "Przelew wychodzący"),
                           (-10.0, "biedronka", "Transakcja kartą"),
                           (-10.0, "bossa", "Przelew wychodzący")]),
                    rules, categories, settings)
    assert dict(zip(df["category"], df["type"])) == {
        "Mieszkanie": "fixed", "Zakupy spożywcze": "variable", "Inwestycje": "savings"}


def test_unknown_category_warns_and_falls_back_to_variable(categories, settings, capsys):
    rules = {"rules": [{"category": "Nieistniejąca", "keywords": ["x"]}]}
    df = categorize(frame([(-10.0, "x", "Transakcja kartą")]), rules, categories, settings)
    assert df["type"].iloc[0] == "variable"
    assert "Nieistniejąca" in capsys.readouterr().out


def test_categorize_does_not_modify_input(rules, categories, settings):
    raw = frame([(-10.0, "biedronka", "Transakcja kartą")])
    categorize(raw, rules, categories, settings)
    assert "category" not in raw.columns


# ---------------------------------------------------------------- filtry
@pytest.fixture
def categorized(rules, categories, settings):
    return categorize(frame([
        (-2000.0, "czynsz", "Przelew wychodzący"),          # Mieszkanie, fixed, duża kwota
        (-50.0, "biedronka", "Transakcja kartą"),            # Zakupy, variable
        (-80.0, "wolt dostawa", "Transakcja BLIK"),          # Jedzenie, variable
        (-1500.0, "czesne", "Przelew wychodzący"),           # Edukacja
        (-500.0, "bossa", "Przelew wychodzący"),             # Inwestycje, savings
    ]), rules, categories, settings)


def test_filter_by_type(categorized):
    out = apply_filter(categorized, {"type": "fixed"})
    assert set(out["category"]) == {"Mieszkanie"}


def test_filter_by_categories(categorized):
    out = apply_filter(categorized, {"categories": ["Zakupy spożywcze", "Jedzenie na mieście"]})
    assert set(out["category"]) == {"Zakupy spożywcze", "Jedzenie na mieście"}


def test_filter_exclude_categories(categorized):
    out = apply_filter(categorized, {"exclude_categories": ["Edukacja", "Inwestycje"]})
    assert "Edukacja" not in set(out["category"])
    assert "Inwestycje" not in set(out["category"])
    assert len(out) == 3


def test_filter_min_amount_keeps_only_big_expenses(categorized):
    out = apply_filter(categorized, {"min_amount": 200})
    assert sorted(out["spent"]) == [500.0, 1500.0, 2000.0]


def test_filter_conditions_are_combined(categorized):
    out = apply_filter(categorized, {"type": "variable", "min_amount": 60})
    assert list(out["category"]) == ["Jedzenie na mieście", "Edukacja"]


def test_empty_filter_changes_nothing(categorized):
    assert len(apply_filter(categorized, {})) == len(categorized)


def test_load_filter_reads_file(tmp_path):
    folder = tmp_path / "config" / "filters"
    folder.mkdir(parents=True)
    (folder / "moj.json").write_text(json.dumps({"min_amount": 10}), encoding="utf-8")
    assert load_filter(tmp_path, "moj") == {"min_amount": 10}


def test_load_filter_unknown_lists_available(tmp_path):
    folder = tmp_path / "config" / "filters"
    folder.mkdir(parents=True)
    (folder / "moj.json").write_text("{}", encoding="utf-8")
    with pytest.raises(FileNotFoundError) as err:
        load_filter(tmp_path, "nie_ma")
    assert "moj" in str(err.value)
