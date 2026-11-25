# Raport wydatków

🇬🇧 English version: [README.en.md](README.en.md)

Skrypt czyta eksporty z banku (xlsx), przypisuje wydatki do kategorii i tworzy raporty
w Excelu (z wykresami): miesięczne i roczne.

## Wymagania

- Python 3.10+
- biblioteki: `pip install -r requirements.txt` (pandas, openpyxl)
- do testów dodatkowo: `pip install -r requirements-dev.txt` (pytest)

## Szybki start

1. Skopiuj szablon konfiguracji (tylko za pierwszym razem):
   ```
   cp -r config/template/* config/
   ```
   Potem dopisz w `config/rules.json` swoje sklepy i usługi (patrz niżej).
2. Wyeksportuj z banku historię operacji do xlsx i wrzuć plik do `data/exports/`.
   (Kolejne eksporty możesz dokładać obok. Nakładające się dni się nie zdublują.)
3. W folderze projektu uruchom:
   ```
   python main.py
   ```
4. Raport znajdziesz w `output/<rok>/<rrrr-mm>.xlsx`.

## Komendy

| Komenda | Co robi |
|---|---|
| `python main.py` | raport miesięczny za ostatni miesiąc z danych |
| `python main.py 2026-11` | raport miesięczny za wybrany miesiąc |
| `python main.py 2026-11 everyday` | raport miesięczny z filtrem `everyday` |
| `python main.py year` | raport roczny za rok ostatniej transakcji |
| `python main.py year 2026` | raport roczny za wybrany rok |
| `python main.py year 2026 everyday` | raport roczny z filtrem |
| `python main.py help` | krótka pomoc |

Raport roczny powstaje też sam: jeśli uruchomisz `python main.py` w styczniu, a pliku
`output/<poprzedni rok>_roczny.xlsx` jeszcze nie ma, skrypt go utworzy.
Raport roczny możesz wygenerować wcześniej, w dowolnym momencie (rok niepełny jest oznaczony
i nadpisuje się przy kolejnym uruchomieniu).

## Struktura plików

```
expense_report/
├── main.py                 ← jedyny plik, który uruchamiasz
├── src/
│   ├── loader.py           ← wczytanie eksportów, scalanie, usuwanie duplikatów
│   ├── categorize.py       ← kategorie, zwroty, filtry
│   ├── report_monthly.py   ← raport miesięczny
│   └── report_yearly.py    ← raport roczny
├── config/template/         ← SZABLON konfiguracji (ogólny, do wersjonowania w git)
│   ├── settings.json
│   ├── categories.json
│   ├── rules.json
│   └── filters/
├── config/                 ← TWOJA konfiguracja (kopia szablonu, nie trafia do git)
│   ├── settings.json       ← ustawienia (nazwy kolumn, zwroty, blokady...)
│   ├── categories.json     ← lista kategorii i ich typ
│   ├── rules.json          ← słowa kluczowe przypisujące transakcje do kategorii
│   └── filters/            ← filtry (po jednym pliku JSON)
├── tests/                  ← testy automatyczne (pytest, dane syntetyczne)
├── requirements.txt        ← biblioteki potrzebne do działania
├── requirements-dev.txt    ← + pytest (do testów)
├── pytest.ini              ← ustawienia pytest
├── data/exports/           ← tu wrzucasz eksporty z banku (nie trafia do git)
└── output/
    ├── 2026/2026-11.xlsx   ← raporty miesięczne z danego roku (nie trafia do git)
    └── 2026_roczny.xlsx    ← raport roczny
```

## Prywatność i git

Eksporty z banku i raporty zawierają prawdziwe transakcje, a `rules.json` pokazuje Twoje
nawyki (sklepy, uczelnia, inwestycje). Nie wrzucaj ich do repozytorium. Proponowany `.gitignore`:

```
# dane z banku i raporty
data/
output/

# osobista konfiguracja (w git zostaje tylko config/template/)
config/

# Python
__pycache__/
*.pyc
```

W repozytorium trzymaj tylko szablon `config/template/` (ogólne reguły, bez danych osobistych).
Jeśli dodasz do szablonu coś nowego i ogólnego (np. nowy filtr), skopiuj to do `config/template/`.
Jeśli kiedyś zmienisz format konfiguracji, zaktualizuj też szablon, żeby nowa kopia działała.

## Jak działa kategoryzacja

Skrypt składa tekst z kolumn `Odbiorca` i `Opis`, zamienia go na małe litery i przechodzi
po regułach z `config/rules.json` od góry do dołu. **Pierwsza reguła, której słowo
kluczowe występuje w tekście, wygrywa.** Jeśli żadna nie pasuje, transakcja trafia do
kategorii "Inne".

### Dodanie nowego sklepu
W `config/rules.json` dopisz słowo kluczowe (małymi literami) do właściwej kategorii:
```json
{"category": "Zakupy spożywcze", "keywords": ["biedronka", "lidl", "twoj nowy sklep"]}
```
Słowo kluczowe może być fragmentem nazwy. Uważaj na zbyt krótkie słowa (np. "bar"), bo
mogą pasować do czegoś innego. Szczegółowe reguły trzymaj wyżej niż ogólne.

### Dodanie nowej kategorii
1. Dopisz ją w `config/categories.json` z typem: `fixed` (stały wydatek), `variable`
   (zmienny) albo `savings` (inwestycje, liczone osobno, poza sumą wydatków).
2. Dodaj dla niej regułę w `config/rules.json`.

## Co dzieje się z poszczególnymi transakcjami

- **Wydatki:** kwoty ujemne.
- **Zwroty:** dodatnie kwoty kartą lub BLIK-iem odejmują się od kategorii dopasowanej po
  opisie (np. zwrot z Reserved zmniejsza "Ubrania i zakupy").
- **Inne wpływy** (wypłata, przelewy od osób): pomijane.
- **Odrzucone i anulowane:** pomijane.
- **Blokady środków ("W trakcie realizacji"):** domyślnie pomijane, bo po zaksięgowaniu
  mogą pojawić się drugi raz. Zmienisz to w `settings.json` (`include_pending`).
- **Inwestycje** (np. wpłata do XTB): liczone osobno, nie wchodzą do sumy wydatków.

## Filtry

Filtr to plik JSON w `config/filters/`. Wszystkie podane w nim warunki muszą być spełnione.
Pola (wszystkie opcjonalne):

| Pole | Znaczenie |
|---|---|
| `type` | tylko typ `fixed` albo `variable` |
| `categories` | tylko te kategorie |
| `exclude_categories` | wszystkie oprócz tych kategorii |
| `min_amount` | tylko pojedyncze wydatki od tej kwoty |

Dostępne filtry: `fixed_only`, `variable_only`, `big_expenses`, `everyday`, `food`,
`leisure`, `transport_housing`, `education`. Nowy filtr = nowy plik, np.
`config/filters/moj.json`, a potem `python main.py 2026-11 moj`.

## Co jest w raporcie (xlsx)

- **Podsumowanie:** tabela kategorii (suma, udział, zmiana vs poprzedni miesiąc), stałe i
  zmienne, inwestycje osobno, 5 największych wydatków, wykres.
- **Porównanie** (miesięczny) / **Miesiące** (roczny): kategorie × miesiące, wykresy.
- **Transakcje:** wszystkie wiersze z kategorią, z filtrami w nagłówkach.
- **Do przypisania:** transakcje z kategorii "Inne", czyli co warto dopisać do reguł.

Liczby w raporcie liczą się formułami z arkusza "Transakcje". Jeśli zmienisz tam kategorię
transakcji, sumy i wykresy przeliczą się same. W raporcie miesięcznym możesz też zmienić
żółtą komórkę z miesiącem (zapisany jako `rrrrmm`, np. `202610`).

Miesiące, w których eksport nie obejmuje wszystkich dni, są oznaczone **"niepełny"**. Nie
porównuj ich 1:1 z miesiącami pełnymi.

## Ustawienia (`config/settings.json`)

| Klucz | Znaczenie |
|---|---|
| `input_folder` | folder z eksportami |
| `columns` | nazwy kolumn w pliku z banku |
| `text_columns` | kolumny, w których szukamy słów kluczowych |
| `skip_status_contains` | statusy do pominięcia (fragmenty tekstu) |
| `include_pending` | `true` = licz też blokady "w trakcie" |
| `pending_status_contains` | jak rozpoznać blokady w kolumnie Status |
| `net_refunds` | `true` = odejmuj zwroty od wydatków |
| `refund_types` | typy transakcji, w których dodatnia kwota to zwrot |
| `default_category` | kategoria dla nierozpoznanych transakcji |
| `compare_months` | ile poprzednich miesięcy w porównaniu (domyślnie 6) |
| `auto_yearly_in_january` | `true` = w styczniu utwórz raport roczny za poprzedni rok |

## Testy

```
pip install -r requirements-dev.txt
python -m pytest
```

Testy używają wyłącznie danych syntetycznych (zmyślonych), więc nie dotykają Twoich
eksportów ani raportów. Sprawdzają m.in.:

- wczytywanie: formaty kwot i dat, pomijanie odrzuconych i blokad, usuwanie duplikatów
  z nakładających się eksportów (a zostawianie dwóch identycznych płatności z jednego pliku),
- kategoryzację: kolejność reguł, zwroty, wpływy, typy kategorii, filtry,
- raporty xlsx: strukturę, formuły i policzone wartości (sumy, udziały, stałe/zmienne,
  inwestycje osobno, rekordy miesięcy),
- `main.py`: komendy, komunikaty błędów i automatyczny raport roczny w styczniu,
- konfigurację: szablon `config/template/` oraz Twój `config/` (jeśli istnieje). Test wyłapie
  słowo kluczowe z wielką literą, kategorię z reguły bez wpisu w `categories.json`, słowo
  powtórzone w dwóch kategoriach albo filtr wskazujący nieistniejącą kategorię.

**Po każdej zmianie w `rules.json`, `categories.json` lub filtrach uruchom testy.**
Testy sprawdzające wartości w raportach korzystają z LibreOffice (`soffice`). Jeśli go nie
masz, te testy są pomijane (`skipped`), a reszta działa normalnie.

## Najczęstsze problemy

- **"Brak folderu config/"**: skopiuj szablon: `cp -r config/template/* config/`.
- **"Brak kolumny ..."**: bank zmienił nagłówki. Popraw nazwy w `settings.json` → `columns`.
- **Dużo w "Inne"**: sprawdź arkusz "Do przypisania" i dopisz słowa kluczowe do `rules.json`.
- **"Nie ma filtru ..."**: nazwa filtru to nazwa pliku z `config/filters/` bez `.json`.
- **Transakcja w innej walucie**: skrypt ostrzega i liczy kwotę bez przewalutowania.
- **Pusty wykres lub zera po otwarciu pliku**: otwórz plik w Excelu. Formuły przeliczają
  się przy otwarciu (w niektórych podglądach, np. w telefonie, mogą być puste).