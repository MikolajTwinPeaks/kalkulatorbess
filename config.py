#!/usr/bin/env python3
"""
ConfigManager — zarządzanie parametrami kalkulatora w bazie SQLite.

Tabela `config` w ceny_tge.db przechowuje key-value z typowaniem i kategoriami.
Domyślne wartości = obecne hardcoded z kalkulator_oferta.py.
"""

import json
import os
import sqlite3
from typing import Any


DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ceny_tge.db')

# (klucz, wartosc_domyslna, opis, kategoria, typ)
_DEFAULTS = [
    # ── Ceny energii ──
    ('cena_fix', 0.58, 'Cena FIX rynkowa 2026 (PLN/kWh)', 'Ceny energii', 'float'),
    ('cena_rdn_srednia', 0.50, 'Średnia cena RDN (PLN/kWh)', 'Ceny energii', 'float'),
    ('cena_mix', 0.54, 'Cena MIX 50% FIX + 50% RDN (PLN/kWh)', 'Ceny energii', 'float'),
    ('cena_net_billing_mnoznik', 0.5, 'Mnożnik ceny ee dla net-billingu', 'Ceny energii', 'float'),

    # ── Koszty inwestycyjne PV ──
    ('pv_capex_maly', 3800.0, 'CAPEX PV < 50 kWp (PLN/kWp)', 'Koszty inwestycyjne', 'float'),
    ('pv_capex_sredni', 3200.0, 'CAPEX PV 50-200 kWp (PLN/kWp)', 'Koszty inwestycyjne', 'float'),
    ('pv_capex_duzy', 2800.0, 'CAPEX PV >= 200 kWp (PLN/kWp)', 'Koszty inwestycyjne', 'float'),
    ('bess_koszt_kwh', 2000.0, 'Koszt BESS (PLN/kWh)', 'Koszty inwestycyjne', 'float'),
    ('bess_ems_koszt', 30000.0, 'Koszt systemu EMS (PLN)', 'Koszty inwestycyjne', 'float'),
    ('bess_instalacja_procent', 0.10, 'Koszt instalacji BESS (% CAPEX)', 'Koszty inwestycyjne', 'float'),
    ('kmb_capex_kvar', 120.0, 'Koszt KMB (PLN/kvar)', 'Koszty inwestycyjne', 'float'),
    ('dsr_koszt_bazowy', 15000.0, 'Koszt bazowy wdrożenia DSR (PLN)', 'Koszty inwestycyjne', 'float'),
    ('dsr_koszt_kw', 50.0, 'Koszt DSR za kW (PLN/kW)', 'Koszty inwestycyjne', 'float'),

    # ── Parametry techniczne PV ──
    ('pv_m2_per_kwp', 5.5, 'Powierzchnia dachu na 1 kWp (m²)', 'Parametry techniczne', 'float'),
    ('pv_pokrycie_zuzycia', 0.70, 'Optymalne pokrycie zużycia ee przez PV', 'Parametry techniczne', 'float'),
    ('pv_produkcja_kwh_per_kwp', 1050.0, 'Roczna produkcja PV (kWh/kWp)', 'Parametry techniczne', 'float'),
    ('pv_autokonsumpcja_24h', 0.50, 'Autokonsumpcja PV — praca 24h/3 zmiany', 'Parametry techniczne', 'float'),
    ('pv_autokonsumpcja_6dni', 0.40, 'Autokonsumpcja PV — praca 6 dni/tyg', 'Parametry techniczne', 'float'),
    ('pv_autokonsumpcja_5dni', 0.35, 'Autokonsumpcja PV — praca 5 dni/tyg', 'Parametry techniczne', 'float'),
    ('bess_rte', 0.90, 'Sprawność round-trip BESS', 'Parametry techniczne', 'float'),
    ('bess_min_pojemnosc', 50.0, 'Minimalna pojemność BESS (kWh)', 'Parametry techniczne', 'float'),
    ('bess_degradacja_bufor', 1.25, 'Bufor na degradację BESS (mnożnik)', 'Parametry techniczne', 'float'),
    ('bess_spread', 0.30, 'Spread cenowy arbitrażu (PLN/kWh)', 'Parametry techniczne', 'float'),
    ('bess_dni_efektywne', 300, 'Liczba efektywnych dni arbitrażu/rok', 'Parametry techniczne', 'int'),
    ('kmb_cos_phi_docelowy', 0.95, 'Docelowy cos(φ) po kompensacji', 'Parametry techniczne', 'float'),
    ('kmb_oszczednosc_dystr_procent', 0.10, 'Oszczędność na dystrybucji z KMB', 'Parametry techniczne', 'float'),

    # ── Peak shaving ──
    ('peak_shaving_redukcja', 0.30, 'Redukcja mocy szczytowej (%)', 'Peak shaving', 'float'),
    ('peak_shaving_godziny', 3.0, 'Godziny peak shavingu', 'Peak shaving', 'float'),
    ('peak_zuzycie_szczytu', 0.65, 'Udział zużycia w szczycie', 'Peak shaving', 'float'),
    ('peak_arbitraz_procent', 0.20, 'Udział dziennego zużycia na arbitraż', 'Peak shaving', 'float'),
    ('peak_kat_mn', '{"K1": 0.17, "K2": 0.40, "K3": 0.70, "K4": 1.00}', 'Mnożniki kategorii mocowych', 'Peak shaving', 'json'),
    ('peak_nowa_kat', '{"K4": "K2", "K3": "K1", "K2": "K1", "K1": "K1"}', 'Mapowanie nowej kategorii po peak shavingu', 'Peak shaving', 'json'),

    # ── DSR ──
    ('dsr_procent_mocy', 0.15, 'Potencjał DSR (% mocy umownej)', 'DSR', 'float'),
    ('dsr_min_kw', 50.0, 'Minimalny potencjał DSR (kW)', 'DSR', 'float'),
    ('dsr_przychod_kw_rok', 300.0, 'Przychód DSR (PLN/kW/rok)', 'DSR', 'float'),

    # ── Finansowanie ──
    ('fin_amortyzacja_procent', 0.10, 'Roczna stawka amortyzacji', 'Finansowanie', 'float'),
    ('fin_cit_procent', 0.19, 'Stawka CIT', 'Finansowanie', 'float'),
    ('fin_leasing_okres', 84, 'Okres leasingu operacyjnego (mies.)', 'Finansowanie', 'int'),
    ('fin_leasing_oprocentowanie', 0.065, 'RRSO leasingu', 'Finansowanie', 'float'),
    ('fin_leasing_wykup', 0.01, 'Wykup leasingu (%)', 'Finansowanie', 'float'),
    ('fin_leasing_fin_okres', 120, 'Okres leasingu finansowego (mies.)', 'Finansowanie', 'int'),
    ('fin_leasing_fin_wklad', 5.0, 'Wkład własny leasing finansowy (%)', 'Finansowanie', 'float'),
    ('fin_bgk_premia', 0.50, 'Premia ekologiczna BGK (%)', 'Finansowanie', 'float'),
    ('fin_bgk_oprocentowanie', 0.07, 'Oprocentowanie kredytu BGK', 'Finansowanie', 'float'),
    ('fin_bgk_okres', 120, 'Okres kredytu BGK (mies.)', 'Finansowanie', 'int'),
    ('fin_esco_mnoznik', 1.50, 'Mnożnik kosztu ESCO vs CAPEX', 'Finansowanie', 'float'),
    ('fin_esco_okres', 180, 'Okres ESCO (mies.)', 'Finansowanie', 'int'),
    ('fin_ppa_okres', 180, 'Okres PPA (mies.)', 'Finansowanie', 'int'),
]


class ConfigManager:
    """Zarządzanie parametrami kalkulatora w SQLite."""

    def __init__(self, db_path: str = DB_PATH):
        self._db = db_path
        self._init_db()

    def _conn(self) -> sqlite3.Connection:
        return sqlite3.connect(self._db)

    def _init_db(self):
        with self._conn() as conn:
            conn.execute('''
                CREATE TABLE IF NOT EXISTS config (
                    klucz TEXT PRIMARY KEY,
                    wartosc TEXT NOT NULL,
                    opis TEXT DEFAULT '',
                    kategoria TEXT DEFAULT '',
                    typ TEXT DEFAULT 'float'
                )
            ''')
            # Seed defaults only if table is empty
            count = conn.execute('SELECT COUNT(*) FROM config').fetchone()[0]
            if count == 0:
                conn.executemany(
                    'INSERT INTO config (klucz, wartosc, opis, kategoria, typ) VALUES (?, ?, ?, ?, ?)',
                    [(k, str(v), o, kat, t) for k, v, o, kat, t in _DEFAULTS],
                )

    def _cast(self, value_str: str, typ: str) -> Any:
        if typ == 'float':
            return float(value_str)
        elif typ == 'int':
            return int(float(value_str))
        elif typ == 'json':
            return json.loads(value_str)
        return value_str

    def get(self, klucz: str, default: Any = None) -> Any:
        with self._conn() as conn:
            row = conn.execute(
                'SELECT wartosc, typ FROM config WHERE klucz = ?', (klucz,)
            ).fetchone()
        if row is None:
            return default
        return self._cast(row[0], row[1])

    def get_all(self) -> dict[str, Any]:
        with self._conn() as conn:
            rows = conn.execute('SELECT klucz, wartosc, typ FROM config ORDER BY klucz').fetchall()
        return {k: self._cast(v, t) for k, v, t in rows}

    def get_by_category(self, kategoria: str) -> list[dict]:
        with self._conn() as conn:
            rows = conn.execute(
                'SELECT klucz, wartosc, opis, kategoria, typ FROM config WHERE kategoria = ? ORDER BY klucz',
                (kategoria,),
            ).fetchall()
        return [
            {'klucz': k, 'wartosc': self._cast(v, t), 'wartosc_raw': v, 'opis': o, 'kategoria': kat, 'typ': t}
            for k, v, o, kat, t in rows
        ]

    def set(self, klucz: str, wartosc: Any):
        with self._conn() as conn:
            conn.execute(
                'UPDATE config SET wartosc = ? WHERE klucz = ?',
                (str(wartosc) if not isinstance(wartosc, str) else wartosc, klucz),
            )

    def set_many(self, updates: dict[str, Any]):
        with self._conn() as conn:
            for klucz, wartosc in updates.items():
                conn.execute(
                    'UPDATE config SET wartosc = ? WHERE klucz = ?',
                    (str(wartosc) if not isinstance(wartosc, str) else wartosc, klucz),
                )

    def categories(self) -> list[str]:
        with self._conn() as conn:
            rows = conn.execute(
                'SELECT DISTINCT kategoria FROM config ORDER BY kategoria'
            ).fetchall()
        return [r[0] for r in rows]
