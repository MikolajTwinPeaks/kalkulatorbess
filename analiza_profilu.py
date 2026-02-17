"""
Analiza profilu mocy godzinowej — parser plików SKADEN (XLS/XLSX) + funkcje analityczne.

Plik referencyjny: eksport SKADEN z OSD, 2 bloki (godz. 1-12 i 13-24),
format kolumn: [2,5,7,10,12,13,15,16,17,19,20,21].
"""

import io
import re
from dataclasses import dataclass

import pandas as pd
import numpy as np


# ============================================================
# DATACLASS
# ============================================================

@dataclass
class ProfilMocy:
    """Wynik parsowania pliku z mocami godzinowymi."""
    dane: pd.DataFrame  # kolumny: datetime, moc_kw (8760/8784 wierszy)
    rok: int
    taryfa: str
    firma: str
    nr_transformatora: str
    p_max_kw: float
    p_min_kw: float
    p_srednia_kw: float
    zuzycie_roczne_kwh: float  # suma(moc_kw) — dane godzinowe


# ============================================================
# PARSER
# ============================================================

# Kolumny danych w formacie SKADEN — 12 godzin per blok
_KOLUMNY_DANYCH = [2, 5, 7, 10, 12, 13, 15, 16, 17, 19, 20, 21]


def _parsuj_date_label(label: str) -> tuple[int, int, bool]:
    """Parsuje etykietę daty SKADEN, np. ' 01-01 <Pn>' → (1, 1, True).

    Zwraca (dzien, miesiac, czy_roboczy).
    Dni w <> to weekendy/święta.
    """
    label = str(label).strip()
    is_working = '<' not in label
    # Wyciągnij DD-MM
    m = re.search(r'(\d{1,2})-(\d{1,2})', label)
    if not m:
        raise ValueError(f'Nie można sparsować daty: {label!r}')
    dd, mm = int(m.group(1)), int(m.group(2))
    return dd, mm, is_working


def parsuj_profil_mocy(file_bytes: bytes, filename: str) -> ProfilMocy:
    """Parsuje plik SKADEN (.xls/.xlsx) z mocami godzinowymi.

    Args:
        file_bytes: zawartość pliku
        filename: nazwa pliku (do rozpoznania formatu)

    Returns:
        ProfilMocy z DataFrame 8760/8784 wierszy
    """
    engine = 'xlrd' if filename.lower().endswith('.xls') else 'openpyxl'
    df_raw = pd.read_excel(io.BytesIO(file_bytes), header=None, engine=engine)

    # --- Metadane z nagłówka ---
    taryfa = ''
    firma = ''
    nr_transformatora = ''
    rok = 0

    # Szukaj metadanych w pierwszych 6 wierszach
    for i in range(min(6, len(df_raw))):
        row_text = ' '.join(str(v) for v in df_raw.iloc[i] if pd.notna(v))

        # Taryfa (np. "B23", "C22a")
        m = re.search(r'\b([ABC]\d{2}[ab]?)\b', row_text, re.IGNORECASE)
        if m and not taryfa:
            taryfa = m.group(1).upper()

        # Rok z zakresu dat okresu (np. "za okres: 01-01-2024 - 31-12-2024")
        # Szukamy daty po "okres" lub bierzemy pierwszą datę z zakresu (DD-MM-YYYY - DD-MM-YYYY)
        m_okres = re.search(r'(\d{2}-\d{2}-\d{4})\s*-\s*(\d{2}-\d{2}-(\d{4}))', row_text)
        if m_okres and not rok:
            rok = int(m_okres.group(3))

    # Firma i transformator z footer (ostatnie wiersze)
    for i in range(max(0, len(df_raw) - 10), len(df_raw)):
        row_text = ' '.join(str(v) for v in df_raw.iloc[i] if pd.notna(v))

        if 'Transformator' in row_text or 'trafo' in row_text.lower():
            nr_transformatora = row_text.strip()

        # Firma — wiersz z "sp. z o.o." lub "S.A." itp.
        if re.search(r'(sp\.\s*z\s*o\.o\.|S\.A\.|sp\.j\.|sp\.k\.)', row_text, re.IGNORECASE):
            # Wyciągnij tylko nazwę firmy (do końca formy prawnej)
            m_firma = re.search(
                r'(.+?(?:sp\.\s*z\s*o\.o\.|S\.A\.|sp\.j\.|sp\.k\.))',
                row_text, re.IGNORECASE,
            )
            firma = m_firma.group(1).strip() if m_firma else row_text.strip()

        # SKADEN version line often has company name nearby
        if 'SKADEN' in row_text:
            continue

    # --- Znajdź bloki danych ---
    # Format SKADEN: "Godziny" na jednym wierszu, wartości godzin (np. "1:00") na następnym.
    # Szukamy wierszy z wartościami godzin w kolumnie 2.
    blok1_start = None
    blok2_start = None

    for i in range(len(df_raw)):
        cell2 = str(df_raw.iloc[i, 2]).strip() if pd.notna(df_raw.iloc[i, 2]) else ''

        if cell2 == '1:00' and blok1_start is None:
            blok1_start = i + 1  # dane zaczynają się w następnym wierszu
        elif cell2 == '13:00' and blok2_start is None:
            blok2_start = i + 1

    if blok1_start is None or blok2_start is None:
        raise ValueError(
            'Nie znaleziono nagłówków bloków godzinowych. '
            'Sprawdź czy plik jest w formacie SKADEN.'
        )

    # --- Parsuj dane z obu bloków ---
    records = []  # lista (datetime, moc_kw)

    def _parsuj_blok(start_row: int, godziny_offset: int):
        """Parsuje blok danych (12 godzin).

        godziny_offset: 0 dla bloku 1 (godz. 1-12), 12 dla bloku 2 (godz. 13-24)
        """
        i = start_row
        while i < len(df_raw):
            cell0 = df_raw.iloc[i, 0]
            if pd.isna(cell0):
                i += 1
                continue

            cell0_str = str(cell0).strip()

            # Koniec bloku — wiersz PMax/PMin/S
            if cell0_str.startswith('PMax') or cell0_str.startswith('PMin') or cell0_str.startswith('S '):
                break

            # Spróbuj sparsować datę
            try:
                dd, mm, _ = _parsuj_date_label(cell0_str)
            except ValueError:
                i += 1
                continue

            # Odczytaj 12 wartości mocy
            for h_idx, col_idx in enumerate(_KOLUMNY_DANYCH):
                godzina = godziny_offset + h_idx + 1  # 1-24
                if col_idx < len(df_raw.columns):
                    val = df_raw.iloc[i, col_idx]
                    moc = float(val) if pd.notna(val) else 0.0
                else:
                    moc = 0.0

                # Godzina 24 → następny dzień godzina 0 (ale traktujemy jako 0:00)
                if godzina == 24:
                    dt = pd.Timestamp(year=rok, month=mm, day=dd, hour=0)
                    # Przesuń o 1 dzień
                    dt = dt + pd.Timedelta(days=1)
                else:
                    dt = pd.Timestamp(year=rok, month=mm, day=dd, hour=godzina)

                records.append((dt, moc))

            i += 1

    _parsuj_blok(blok1_start, 0)   # godziny 1-12
    _parsuj_blok(blok2_start, 12)  # godziny 13-24

    if not records:
        raise ValueError('Nie znaleziono żadnych danych mocy w pliku.')

    # --- Buduj DataFrame ---
    df = pd.DataFrame(records, columns=['datetime', 'moc_kw'])
    df = df.sort_values('datetime').reset_index(drop=True)

    # Usuń duplikaty (np. godz. 24 = następny dzień 0:00)
    df = df.drop_duplicates(subset='datetime', keep='first').reset_index(drop=True)

    # Statystyki
    p_max = df['moc_kw'].max()
    p_min = df['moc_kw'].min()
    p_srednia = df['moc_kw'].mean()
    zuzycie = df['moc_kw'].sum()  # 1h × kW = kWh

    return ProfilMocy(
        dane=df,
        rok=rok,
        taryfa=taryfa,
        firma=firma,
        nr_transformatora=nr_transformatora,
        p_max_kw=p_max,
        p_min_kw=p_min,
        p_srednia_kw=p_srednia,
        zuzycie_roczne_kwh=zuzycie,
    )


# ============================================================
# ANALIZA
# ============================================================

def analiza_profilu(profil: ProfilMocy) -> dict:
    """Pełna analiza profilu mocy godzinowej.

    Returns:
        dict z kluczami: statystyki, rozklad_strefowy, profil_dobowy,
        profil_miesieczny, top_szczyty, heatmapa, rekomendacja_moc_umowna,
        kategoria_mocowa
    """
    df = profil.dane.copy()
    df['godzina'] = df['datetime'].dt.hour
    df['miesiac'] = df['datetime'].dt.month
    df['dzien_tygodnia'] = df['datetime'].dt.dayofweek  # 0=Pn, 6=Nd
    df['dzien_roku'] = df['datetime'].dt.dayofyear
    df['nazwa_dnia'] = df['datetime'].dt.day_name()

    wynik = {}

    # 1. Statystyki podstawowe
    load_factor = profil.p_srednia_kw / profil.p_max_kw if profil.p_max_kw > 0 else 0
    wynik['statystyki'] = {
        'p_max_kw': profil.p_max_kw,
        'p_min_kw': profil.p_min_kw,
        'p_srednia_kw': profil.p_srednia_kw,
        'zuzycie_roczne_kwh': profil.zuzycie_roczne_kwh,
        'load_factor': load_factor,
        'liczba_godzin': len(df),
    }

    # 2. Rozkład strefowy (definicje taryfowe)
    # Szczyt: 7-13, 16-21 w dni robocze
    # Pozaszczyt: reszta w dni robocze
    # Noc: 22-6 (wszystkie dni)
    def _strefa(row):
        h = row['godzina']
        wd = row['dzien_tygodnia']
        if 22 <= h or h < 6:
            return 'noc'
        if wd < 5:  # Pn-Pt
            if (7 <= h < 13) or (16 <= h < 21):
                return 'szczyt'
        return 'pozaszczyt'

    df['strefa'] = df.apply(_strefa, axis=1)
    zuzycie_per_strefa = df.groupby('strefa')['moc_kw'].sum()
    total = zuzycie_per_strefa.sum()
    wynik['rozklad_strefowy'] = {
        'szczyt_kwh': zuzycie_per_strefa.get('szczyt', 0),
        'pozaszczyt_kwh': zuzycie_per_strefa.get('pozaszczyt', 0),
        'noc_kwh': zuzycie_per_strefa.get('noc', 0),
        'szczyt_pct': zuzycie_per_strefa.get('szczyt', 0) / total * 100 if total > 0 else 0,
        'pozaszczyt_pct': zuzycie_per_strefa.get('pozaszczyt', 0) / total * 100 if total > 0 else 0,
        'noc_pct': zuzycie_per_strefa.get('noc', 0) / total * 100 if total > 0 else 0,
    }

    # 3. Profil dobowy — średnia moc per godzina
    profil_dobowy = df.groupby('godzina')['moc_kw'].mean()
    wynik['profil_dobowy'] = profil_dobowy.to_dict()

    # 4. Profil miesięczny — zużycie per miesiąc
    profil_miesieczny = df.groupby('miesiac')['moc_kw'].sum()
    wynik['profil_miesieczny'] = profil_miesieczny.to_dict()

    # 5. Top 10 szczytów
    top10 = df.nlargest(10, 'moc_kw')[['datetime', 'moc_kw']].copy()
    top10['datetime'] = top10['datetime'].dt.strftime('%Y-%m-%d %H:%M')
    wynik['top_szczyty'] = top10.reset_index(drop=True)

    # 6. Heatmapa — pivot: dzień tygodnia × godzina → średnia moc
    nazwy_dni = ['Pn', 'Wt', 'Sr', 'Cz', 'Pt', 'Sb', 'Nd']
    heatmapa = df.pivot_table(
        values='moc_kw',
        index='dzien_tygodnia',
        columns='godzina',
        aggfunc='mean',
    )
    heatmapa.index = [nazwy_dni[i] for i in heatmapa.index]
    wynik['heatmapa'] = heatmapa

    # 7. Rekomendacja mocy umownej — percentyl 99.5
    p995 = np.percentile(df['moc_kw'], 99.5)
    wynik['rekomendacja_moc_umowna'] = {
        'p_max_kw': profil.p_max_kw,
        'percentyl_995_kw': p995,
        'rekomendacja_kw': round(p995 / 10) * 10,  # zaokrąglenie do 10 kW
    }

    # 8. Kategoria mocowa — zużycie w szczycie systemowym
    # Szczyt systemowy: 17-19 w dni robocze (Pn-Pt), XI-III
    mask_szczyt_sys = (
        (df['godzina'] >= 17) & (df['godzina'] < 19) &
        (df['dzien_tygodnia'] < 5) &
        (df['miesiac'].isin([11, 12, 1, 2, 3]))
    )
    zuzycie_szczyt_sys = df.loc[mask_szczyt_sys, 'moc_kw'].sum()
    # Kategoria wg progów (uproszczone)
    if zuzycie_szczyt_sys <= 100_000:
        kategoria = 'K1'
    elif zuzycie_szczyt_sys <= 500_000:
        kategoria = 'K2'
    elif zuzycie_szczyt_sys <= 2_000_000:
        kategoria = 'K3'
    else:
        kategoria = 'K4'

    wynik['kategoria_mocowa'] = {
        'zuzycie_szczyt_systemowy_kwh': zuzycie_szczyt_sys,
        'kategoria': kategoria,
    }

    return wynik


# ============================================================
# MAPPING NA FORMULARZ
# ============================================================

def mapuj_profil_na_dane(profil: ProfilMocy, analiza: dict) -> dict:
    """Mapuje dane z profilu na klucze session_state formularza."""
    mapped = {
        'roczne_zuzycie_ee_kwh': round(profil.zuzycie_roczne_kwh, 0),
        'moc_umowna_kw': analiza['rekomendacja_moc_umowna']['rekomendacja_kw'],
    }

    if profil.taryfa:
        mapped['grupa_taryfowa'] = profil.taryfa

    if analiza.get('kategoria_mocowa'):
        mapped['kategoria_mocowa'] = analiza['kategoria_mocowa']['kategoria']

    return mapped
