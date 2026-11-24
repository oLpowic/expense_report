"""
loader.py - wczytywanie eksportów z banku.

Co robi ten plik:
1. czyta wszystkie pliki .xlsx z folderu data/exports
2. zamienia daty i kwoty na liczby
3. usuwa odrzucone transakcje i blokady środków
4. scala pliki w jedną tabelę i usuwa duplikaty z nakładających się eksportów

Wynik: tabela z kolumnami
    date      - data transakcji
    amount    - kwota ze znakiem (wydatek jest ujemny, wpływ dodatni)
    text      - Odbiorca + Opis, małymi literami (po tym szukamy słów kluczowych)
    bank_type - "Typ transakcji" z banku (np. Transakcja kartą)
    balance   - saldo po transakcji (służy tylko do wykrywania duplikatów)
    source    - nazwa pliku, z którego pochodzi wiersz
"""
import json
from pathlib import Path

import pandas as pd


def load_json(path):
    """Wczytuje plik JSON i zwraca słownik."""
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def parse_dates(series):
    """Zamienia daty na datetime. Obsługuje 2026-11-20 oraz 20.11.2026 (z godziną lub bez)."""
    if pd.api.types.is_datetime64_any_dtype(series):
        return series
    text = series.astype(str).str.strip().str[:10]  # odetnij ewentualną godzinę
    iso = pd.to_datetime(text, format="%Y-%m-%d", errors="coerce")
    polish = pd.to_datetime(text, format="%d.%m.%Y", errors="coerce")
    return iso.fillna(polish)


def to_number(value):
    """Zamienia kwotę na liczbę. Obsługuje '-1 234,50', zwykły minus i unicode minus."""
    if pd.isna(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).replace(" ", "").replace(" ", "").replace("−", "-")
    text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def read_one_file(path, settings):
    """Wczytuje jeden plik xlsx i zwraca oczyszczoną tabelę."""
    cols = settings["columns"]
    df = pd.read_excel(path)
    df.columns = [str(c).strip() for c in df.columns]  # usuń spacje z nazw kolumn

    # bez daty i kwoty nie da się nic policzyć
    for key in ("date", "amount"):
        if cols[key] not in df.columns:
            raise ValueError(
                f"{path.name}: brak kolumny '{cols[key]}'. Kolumny w pliku: {list(df.columns)}"
            )

    # 1) odrzucone transakcje (mają datę odrzucenia)
    if cols["rejected_date"] in df.columns:
        df = df[df[cols["rejected_date"]].isna()]

    # 2) pomijamy transakcje o statusie z listy (odrzucone, anulowane, a opcjonalnie "w trakcie")
    if cols["status"] in df.columns:
        words = list(settings["skip_status_contains"])
        if not settings["include_pending"]:
            words += settings["pending_status_contains"]
        status = df[cols["status"]].astype(str).str.lower()
        skip = status.map(lambda s: any(w in s for w in words))
        df = df[~skip]

    # 3) ostrzeżenie o innych walutach (kwoty nie są przeliczane)
    if cols["currency"] in df.columns:
        foreign = df[df[cols["currency"]].astype(str).str.upper() != "PLN"]
        if len(foreign):
            print(f"UWAGA ({path.name}): {len(foreign)} transakcji nie jest w PLN - bez przewalutowania.")

    # 4) składamy tekst do szukania słów kluczowych z kolumn Odbiorca i Opis
    text_cols = [c for c in settings["text_columns"] if c in df.columns]
    if not text_cols:
        print(f"UWAGA ({path.name}): brak kolumn z opisem. Kolumny w pliku: {list(df.columns)}")
        text = pd.Series("", index=df.index)
    else:
        text = df[text_cols].fillna("").astype(str).agg(" ".join, axis=1)
        text = text.str.lower()
        # usuwamy numery kont (20-26 cyfr, opcjonalnie z "pl") i maskę karty - tylko śmiecą w opisie
        text = text.str.replace(r"\b(?:pl)?\d{20,26}\b", " ", regex=True)
        text = text.str.replace(r"\b\d{6}[-_]{6}\d{4}\b", " ", regex=True)
        text = text.str.replace(r"\s+", " ", regex=True).str.strip()

    out = pd.DataFrame({
        "date": parse_dates(df[cols["date"]]),
        "amount": df[cols["amount"]].map(to_number),
        "text": text,
        "bank_type": df[cols["type"]].astype(str) if cols["type"] in df.columns else "",
        "balance": df[cols["balance"]].map(to_number) if cols["balance"] in df.columns else None,
    })

    # wiersze bez daty lub kwoty są bezużyteczne
    before = len(out)
    out = out.dropna(subset=["date", "amount"])
    if len(out) < before:
        print(f"UWAGA ({path.name}): pominięto {before - len(out)} wierszy bez poprawnej daty lub kwoty.")
    return out


def make_key(row):
    """Klucz transakcji do wykrywania duplikatów: data + kwota + tekst + saldo."""
    balance = None if pd.isna(row["balance"]) else round(row["balance"], 2)
    return (row["date"].date(), round(row["amount"], 2), row["text"], balance)


def load_expenses(base_dir, settings):
    """Wczytuje wszystkie eksporty z folderu i zwraca jedną tabelę bez duplikatów."""
    folder = Path(base_dir) / settings["input_folder"]
    files = sorted(p for p in folder.glob("*.xlsx") if not p.name.startswith("~$"))
    if not files:
        raise FileNotFoundError(f"Brak plików .xlsx w folderze: {folder}")

    seen = set()      # klucze transakcji z wcześniejszych plików
    parts = []
    skipped = 0
    for file in files:
        df = read_one_file(file, settings)
        keys = [make_key(row) for _, row in df.iterrows()]

        # wiersze, które były już w poprzednim pliku, odrzucamy (nakładające się eksporty)
        # Duplikaty WEWNĄTRZ jednego pliku zostają - to mogą być dwie prawdziwe płatności.
        is_new = [k not in seen for k in keys]
        skipped += len(df) - sum(is_new)
        seen.update(keys)

        df = df[is_new].copy()
        df["source"] = file.name
        parts.append(df)

    result = pd.concat(parts, ignore_index=True).sort_values("date", kind="stable")
    result = result.reset_index(drop=True)
    print(f"Wczytano {len(files)} plik(ów): {len(result)} transakcji, "
          f"pominięto {skipped} duplikatów z nakładających się eksportów.")
    return result
