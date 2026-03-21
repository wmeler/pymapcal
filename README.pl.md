# pymapcal

MVP aplikacji Qt do kalibracji arkuszy w albumie map.
Album projektu zawiera wiele skanów, a każdy skan zawiera własne arkusze.

## Wymagania
- Python 3.10+
- PySide6 (`pip install -r requirements.txt`)

## Uruchomienie
```bash
python3 main.py
```

Możesz też podać ścieżkę jako pierwszy argument pozycyjny, np. `python3 main.py ./scan.tif`, `python3 main.py ./arkusz.map` albo `python3 main.py ./projekt.json`.

## Build PyInstaller
Linux/macOS:
```bash
python3 -m venv venv
./venv/bin/pip install -r requirements-build.txt
./build_pyinstaller.sh
```

Windows:
```bat
py -m venv venv
venv\Scripts\python.exe -m pip install -r requirements-build.txt
build_pyinstaller.bat
```

Gotowy build pojawi się w `./dist/pymapcal/`.

## Obsługa
1. `Plik -> Dodaj skan mapy` (tif/tiff/bmp/png/jpg/jpeg).
   - lub `Plik -> Importuj MAP...` aby zaimportować kalibrację z plików OziExplorer `.map`.
   - jeśli wiele plików `.map` wskazuje ten sam obraz skanu, skan pojawi się w albumie tylko raz, a każdy plik `.map` doda kolejny arkusz.
   - `Arkusz -> Eksportuj KAP...` generuje pliki `.kap` dla wszystkich arkuszy aktywnego skanu.
   - `Plik -> Eksportuj wszystkie arkusze do KAP...` generuje `.kap` dla wszystkich arkuszy ze wszystkich skanów w albumie.
2. Wybierz `Arkusz -> Nowy arkusz` i dodaj kolejne narożniki klikami na mapie.
3. Dodaj co najmniej 3 punkty obrysu; po trzecim punkcie obrys jest traktowany jako domknięty automatycznie.
4. `Arkusz -> Dodaj punkt kalibracyjny` dodaje punkt kalibracyjny (max 9 na arkusz).
   - `Arkusz -> Dodaj punkt obrysu` dodaje punkt maski/obrysu bez punktu kalibracyjnego.
   - po kliknięciu współrzędne są automatycznie podpowiadane na podstawie bieżącej kalibracji,
   - jeśli klik jest blisko linii siatki wynikającej ze skali arkusza, proponowana pozycja punktu kalibracyjnego jest automatycznie dociągana do tej siatki,
   - możesz je ręcznie edytować.
   - podczas nanoszenia/przesuwania punktu kalibracyjnego widoczne są linie pozycji kursora.
   - `Esc` kończy tryb dodawania punktu kalibracyjnego lub obrysu.
   - każde dodanie punktu kalibracyjnego dopisuje przypadek diagnostyczny do `./geo_position_cases.jsonl`.
5. Punkty można przesuwać przeciągając myszą.
6. W panelu bocznym:
   - drzewo projektu: `Skan -> Arkusze -> (Obrys kadrowania / Punkty kalibracyjne)`,
   - kliknięcie elementu drzewa zaznacza odpowiedni arkusz/punkt na mapie,
   - menu kontekstowe drzewa (prawy przycisk myszy) pozwala zależnie od węzła dodać/usunąć arkusz oraz dodać/usunąć punkty obrysu lub kalibracyjne,
   - `Arkusz -> Dodaj punkt kalibracji do obrysu` pozwala użyć punktu kalibracyjnego jako punktu kadrowania,
   - edytuj nazwę i skalę zaznaczonego arkusza,
   - dla zaznaczonego punktu wpisz `Lat` i `Lon`, a potem kliknij `Zapisz punkt` (z walidacją formatu).
7. Pasek statusu pokazuje pozycję kursora w pikselach i przybliżoną pozycję geo.
8. Pan i zoom:
   - rolka myszy: zoom względem kursora,
   - środkowy przycisk myszy + przeciąganie: pan,
   - `Widok -> Zoom +`, `Widok -> Zoom -`, `Widok -> Zoom 100%`.

## Uwagi
- Pozycja geo kursora i siatka są liczone:
  - z 2 punktów: transformacja podobieństwa (obrót + skala + przesunięcie),
  - z 3+ punktów: dopasowanie afiniczne (least squares).
- Linie obrysu mają stałą grubość ekranową (nie skalują się przy zoom).
- `Plik -> Zapisz` / `Zapisz jako...` / `Wczytaj album` operuje na całym albumie (zbiorze skanów).
- `Arkusz -> Eksportuj KAP...` eksportuje arkusze z bieżącego skanu do wskazanego katalogu.
  - każdy arkusz musi mieć skalę (np. `1:50000`),
  - punkty obrysu muszą mieć współrzędne geo (brakujące są wyliczane z bieżącej kalibracji, jeśli to możliwe).
- `Narzędzia -> Edytuj ustawienia...` pozwala zmienić parametry wyświetlania i język oraz zapisać je do `.pymapcal`.
- Ścieżka do binarki `imgkap` jest konfigurowalna w `Narzędzia -> Edytuj ustawienia...`.
- `Plik -> Zapisz` zapisuje do aktualnie otwartego projektu (bez pytania),
- `Plik -> Zapisz jako...` zapisuje pod nową nazwą/ścieżką.

## Ustawienia `.pymapcal`
Aplikacja wczytuje ustawienia z:
1. `./.pymapcal` (bieżący katalog),
2. `~/.pymapcal` (fallback, jeśli brak lokalnego).

Format pliku: JSON (może być bezpośrednio lub pod kluczem `display`).

Obsługa i18n:
- `language: "pl"` lub `language: "en"`
- `imgkap_path`: ścieżka do binarki `imgkap` (np. `imgkap` lub `/home/user/imgkap/imgkap`)
- `imgkap_work_dir`: opcjonalny katalog debug eksportu KAP; jeśli ustawiony, zapisuje tam log `imgkap_calls.log` i wszystkie pliki tymczasowe
- `kap_sounding_datum`: wartość `SD` zapisywana w nagłówku KAP (np. `UNKNOWN`)

Przykład:
```json
{
  "language": "pl",
  "imgkap_path": "imgkap",
  "imgkap_work_dir": "",
  "kap_sounding_datum": "UNKNOWN",
  "display": {
    "outline_width": 2,
    "outline_selected_width": 3,
    "draft_outline_width": 2,
    "crosshair_arm_corner": 16,
    "crosshair_arm_cal": 14,
    "crosshair_ring_corner": 4,
    "crosshair_ring_cal": 2,
    "crosshair_selected_arm_bonus": 6,
    "crosshair_selected_ring_bonus": 2,
    "cursor_guide_width": 2,
    "cursor_guide_alpha": 200,
    "cursor_guide_dash": 10,
    "cursor_guide_gap": 6,
    "cursor_guide_color": "#FFD84D"
  }
}
```

## Format współrzędnych
Obsługiwane dla `Lon/Lat`:
- `DD` (stopnie dziesiętne), np. `18.654321`, `-54.1234`
- `DD + półkula`, np. `54.1234N`, `18.6543 E`
- `DMM + półkula`, np. `54 12.34 N`, `18° 39.26' E`
- `DMS + półkula`, np. `54 12 20.5 N`, `18°39'15.2"E`

Półkule:
- `N/S` dla szerokości (`Lat`)
- `E/W` dla długości (`Lon`)

## Przykłady zapisu (gotowe do wklejenia)
To samo położenie zapisane różnie:

- `Lat`:
  - `54.205694`
  - `54.205694N`
  - `54 12.3416 N`
  - `54°12.3416'N`
  - `54 12 20.5 N`
  - `54°12'20.5"N`

- `Lon`:
  - `18.652611`
  - `18.652611E`
  - `18 39.1567 E`
  - `18°39.1567'E`
  - `18 39 9.4 E`
  - `18°39'9.4"E`

Przykłady dla półkuli zachodniej/południowej:
- `Lat`: `33.9249S`, `33 55.494 S`, `33 55 29.6 S`
- `Lon`: `151.2093W`, `151 12.558 W`, `151 12 33.5 W`

## Start z plikiem projektu
Można od razu otworzyć projekt przy uruchomieniu:
```bash
python3 main.py /ścieżka/do/projektu.json
```
