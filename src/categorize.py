"""
categorize.py - przypisywanie kategorii i filtrowanie.

Co robi ten plik:
1. z całej tabeli zostawia tylko wydatki i zwroty (resztę wpływów pomija)
2. dopasowuje kategorię do każdej transakcji po słowach kluczowych z rules.json
3. dopisuje typ kategorii (fixed / variable / savings) z categories.json
4. pozwala zastosować filtr z config/filters/

Po tym kroku tabela ma dodatkowe kolumny:
    kind     - "expense" (wydatek) albo "refund" (zwrot)
    spent    - kwota dodatnia dla wydatku, ujemna dla zwrotu (dzięki temu zwrot
               automatycznie zmniejsza sumę kategorii)
    category - nazwa kategorii
    type     - fixed / variable / savings
"""
from src.loader import load_json


def find_category(text, rules, default_category):
    """Zwraca kategorię pierwszej pasującej reguły (reguły są ułożone od góry do dołu)."""
    for rule in rules["rules"]:
        for word in rule["keywords"]:
            if word in text:
                return rule["category"]
    return default_category


def categorize(df, rules, categories, settings):
    """Zostawia wydatki i zwroty oraz dopisuje kategorię i typ."""
    df = df.copy()

    # wydatek = kwota ujemna
    is_expense = df["amount"] < 0

    # zwrot = kwota dodatnia kartą lub BLIK-iem (typy z settings.json)
    # inne wpływy (wypłata, przelewy od osób) pomijamy
    is_refund = (df["amount"] > 0) & df["bank_type"].isin(settings["refund_types"])

    if settings["net_refunds"]:
        df = df[is_expense | is_refund].copy()
    else:
        df = df[is_expense].copy()

    df["kind"] = df["amount"].map(lambda a: "expense" if a < 0 else "refund")
    df["spent"] = -df["amount"]  # wydatek: dodatnia, zwrot: ujemna

    # kategoria po słowach kluczowych
    default_category = settings["default_category"]
    df["category"] = df["text"].map(lambda t: find_category(t, rules, default_category))

    # typ kategorii z categories.json (jeśli czegoś brakuje, to "variable")
    types = {c["name"]: c["type"] for c in categories["categories"]}
    unknown = sorted(set(df["category"]) - set(types))
    if unknown:
        print(f"UWAGA: kategorie z rules.json, których nie ma w categories.json: {unknown}")
    df["type"] = df["category"].map(types).fillna("variable")
    return df


def load_filter(base_dir, name):
    """Wczytuje filtr z config/filters/<name>.json."""
    folder = base_dir / "config" / "filters"
    path = folder / f"{name}.json"
    if not path.exists():
        available = sorted(p.stem for p in folder.glob("*.json"))
        raise FileNotFoundError(f"Nie ma filtru '{name}'. Dostępne filtry: {available}")
    return load_json(path)


def apply_filter(df, flt):
    """Stosuje filtr. Wszystkie podane w nim warunki muszą być spełnione."""
    if "type" in flt:
        df = df[df["type"] == flt["type"]]
    if "categories" in flt:
        df = df[df["category"].isin(flt["categories"])]
    if "exclude_categories" in flt:
        df = df[~df["category"].isin(flt["exclude_categories"])]
    if "min_amount" in flt:
        df = df[df["spent"] >= flt["min_amount"]]
    return df