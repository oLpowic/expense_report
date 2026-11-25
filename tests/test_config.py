"""
Testy plików konfiguracji: szablon config/template/ oraz (jeśli istnieje) Twój config/.

Łapią typowe literówki: słowo kluczowe z wielką literą (nigdy by nie zadziałało, bo tekst
transakcji jest zamieniany na małe litery), kategoria w regułach bez wpisu w categories.json,
filtr wskazujący nieistniejącą kategorię itd.
"""
import json
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
DIRS = [ROOT / "config" / "template"]
if (ROOT / "config").exists():       # Twoja konfiguracja (nie ma jej w repozytorium)
    DIRS.append(ROOT / "config")

VALID_TYPES = {"fixed", "variable", "savings"}


def load(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


@pytest.fixture(params=DIRS, ids=lambda p: p.name)
def cfg(request):
    folder = request.param
    return {
        "folder": folder,
        "settings": load(folder / "settings.json"),
        "categories": load(folder / "categories.json"),
        "rules": load(folder / "rules.json"),
        "filters": {p.stem: load(p) for p in sorted((folder / "filters").glob("*.json"))},
    }


def category_names(cfg):
    return [c["name"] for c in cfg["categories"]["categories"]]


def test_all_json_files_are_valid(cfg):
    files = list(cfg["folder"].glob("*.json")) + list((cfg["folder"] / "filters").glob("*.json"))
    assert files
    for path in files:
        load(path)  # rzuci wyjątek przy błędnej składni


def test_settings_have_all_required_keys(cfg):
    required = {"input_folder", "columns", "text_columns", "skip_status_contains", "include_pending",
                "pending_status_contains", "net_refunds", "refund_types", "default_category",
                "compare_months", "auto_yearly_in_january"}
    assert required <= set(cfg["settings"])
    assert {"date", "rejected_date", "amount", "currency", "status", "type", "balance"} \
        <= set(cfg["settings"]["columns"])


def test_categories_are_unique_and_have_valid_types(cfg):
    names = category_names(cfg)
    assert len(names) == len(set(names))
    assert {c["type"] for c in cfg["categories"]["categories"]} <= VALID_TYPES


def test_default_category_exists(cfg):
    assert cfg["settings"]["default_category"] in category_names(cfg)


def test_every_rule_category_exists_in_categories(cfg):
    unknown = {r["category"] for r in cfg["rules"]["rules"]} - set(category_names(cfg))
    assert not unknown, f"Reguły wskazują nieznane kategorie: {unknown}"


def test_keywords_are_lowercase_and_not_empty(cfg):
    """Tekst transakcji jest zamieniany na małe litery, więc słowa z wielkimi literami nigdy nie pasują."""
    for rule in cfg["rules"]["rules"]:
        assert rule["keywords"], f"Pusta lista słów w regule {rule['category']}"
        for word in rule["keywords"]:
            assert word == word.lower(), f"Słowo '{word}' ({rule['category']}) ma wielkie litery"
            assert word.strip(), "Puste słowo kluczowe dopasowałoby wszystko"


def test_keyword_is_not_listed_in_two_categories(cfg):
    """Gdyby słowo było w dwóch kategoriach, działałaby tylko pierwsza (druga byłaby martwa)."""
    seen = {}
    for rule in cfg["rules"]["rules"]:
        for word in rule["keywords"]:
            assert word not in seen, f"'{word}' jest w '{seen[word]}' i w '{rule['category']}'"
            seen[word] = rule["category"]


def test_filters_reference_existing_categories_and_types(cfg):
    names = set(category_names(cfg))
    allowed_fields = {"name", "description", "type", "categories", "exclude_categories", "min_amount"}
    for fname, flt in cfg["filters"].items():
        assert set(flt) <= allowed_fields, f"Filtr {fname}: nieznane pola {set(flt) - allowed_fields}"
        assert set(flt.get("categories", [])) <= names, f"Filtr {fname}: nieznana kategoria"
        assert set(flt.get("exclude_categories", [])) <= names, f"Filtr {fname}: nieznana kategoria"
        if "type" in flt:
            assert flt["type"] in VALID_TYPES
        if "min_amount" in flt:
            assert flt["min_amount"] >= 0


def test_filter_name_matches_file_name(cfg):
    for fname, flt in cfg["filters"].items():
        assert flt.get("name", fname) == fname
