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

## Obsługa
1. `Plik -> Dodaj skan mapy` (tif/tiff/bmp/png/jpg/jpeg).
   - lub `Plik -> Importuj MAP...` aby zaimportować kalibrację z plików OziExplorer `.map`.
   - jeśli wiele plików `.map` wskazuje ten sam obraz skanu, skan pojawi się w albumie tylko raz, a każdy plik `.map` doda kolejny arkusz.
2. Kliknij `Nowy arkusz` i dodaj kolejne narożniki klikami na mapie.
3. Kliknij `Zamknij arkusz` aby utworzyć wielokąt.
4. `Dodaj punkt kalibracyjny` dodaje punkt kalibracyjny (max 9 na arkusz).
   - `Dodaj punkt obrysu` dodaje punkt maski/obrysu bez punktu kalibracyjnego.
   - po kliknięciu współrzędne są automatycznie podpowiadane na podstawie bieżącej kalibracji,
   - możesz je ręcznie edytować.
   - podczas nanoszenia/przesuwania punktu kalibracyjnego widoczne są linie pozycji kursora.
5. Punkty można przesuwać przeciągając myszą.
6. W panelu bocznym:
   - drzewo projektu: `Skan -> Arkusze -> (Obrys kadrowania / Punkty kalibracyjne)`,
   - kliknięcie elementu drzewa zaznacza odpowiedni arkusz/punkt na mapie,
   - `Usuń zaznaczony punkt` usuwa aktualny punkt kalibracyjny lub punkt obrysu,
   - przycisk `Dodaj punkt kalibracji do obrysu` pozwala użyć punktu kalibracyjnego jako punktu kadrowania,
   - edytuj nazwę i skalę zaznaczonego arkusza,
   - dla zaznaczonego punktu wpisz `Lat` i `Lon`, a potem kliknij `Zapisz punkt` (z walidacją formatu).
7. Pasek statusu pokazuje pozycję kursora w pikselach i przybliżoną pozycję geo.
8. Pan i zoom:
   - rolka myszy: zoom względem kursora,
   - środkowy przycisk myszy + przeciąganie: pan,
   - przyciski `Zoom +`, `Zoom -`, `Zoom 100%` w panelu bocznym.

## Uwagi
- Pozycja geo kursora i siatka są liczone:
  - z 2 punktów: transformacja podobieństwa (obrót + skala + przesunięcie),
  - z 3+ punktów: dopasowanie afiniczne (least squares).
- Linie obrysu mają stałą grubość ekranową (nie skalują się przy zoom).
- `Plik -> Zapisz` / `Zapisz jako...` / `Wczytaj album` operuje na całym albumie (zbiorze skanów).
- `Ustawienia -> Edytuj ustawienia...` pozwala zmienić parametry wyświetlania i język oraz zapisać je do `.pymapcal`.
- `Plik -> Zapisz` zapisuje do aktualnie otwartego projektu (bez pytania),
- `Plik -> Zapisz jako...` zapisuje pod nową nazwą/ścieżką.

## Ustawienia `.pymapcal`
Aplikacja wczytuje ustawienia z:
1. `./.pymapcal` (bieżący katalog),
2. `~/.pymapcal` (fallback, jeśli brak lokalnego).

Format pliku: JSON (może być bezpośrednio lub pod kluczem `display`).

Obsługa i18n:
- `language: "pl"` lub `language: "en"`

Przykład:
```json
{
  "language": "pl",
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
