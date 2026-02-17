#!/usr/bin/env python3
"""
Manager bazy danych SQLite z cenami 15-minutowymi z TGE RDB.

Tabele:
- ceny_15min: ceny energii z interwałem 15-minutowym
- scraper_log: historia uruchomień scrapera

Użycie:
    from baza_cen import BazaCen
    db = BazaCen()
    db.zapisz_ceny([...])
    df = db.pobierz_ceny('2026-02-01', '2026-02-15')
"""

import sqlite3
import os
from datetime import datetime, timedelta
from typing import Optional

import pandas as pd


# Domyślna ścieżka bazy danych — obok tego pliku
_DEFAULT_DB_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), 'ceny_tge.db'
)


class BazaCen:
    """Manager bazy SQLite z cenami 15-minutowymi TGE."""

    def __init__(self, db_path: str = _DEFAULT_DB_PATH):
        self.db_path = db_path
        self._init_db()

    # ------------------------------------------------------------------
    # Inicjalizacja
    # ------------------------------------------------------------------

    def _conn(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self.db_path)
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA foreign_keys=ON")
        return conn

    def _init_db(self):
        with self._conn() as conn:
            conn.executescript("""
                CREATE TABLE IF NOT EXISTS ceny_15min (
                    timestamp_start TEXT NOT NULL,
                    timestamp_end   TEXT NOT NULL,
                    rynek           TEXT NOT NULL DEFAULT 'RDB',
                    cena_pln_mwh    REAL NOT NULL,
                    wolumen         REAL,
                    waluta          TEXT NOT NULL DEFAULT 'PLN',
                    zrodlo          TEXT NOT NULL DEFAULT 'TGE',
                    PRIMARY KEY (timestamp_start, rynek)
                );

                CREATE INDEX IF NOT EXISTS idx_ceny_ts
                    ON ceny_15min(timestamp_start);
                CREATE INDEX IF NOT EXISTS idx_ceny_rynek
                    ON ceny_15min(rynek);

                CREATE TABLE IF NOT EXISTS scraper_log (
                    id              INTEGER PRIMARY KEY AUTOINCREMENT,
                    timestamp       TEXT NOT NULL,
                    rynek           TEXT NOT NULL DEFAULT 'RDB',
                    data_sesji      TEXT,
                    liczba_rekordow INTEGER DEFAULT 0,
                    status          TEXT NOT NULL,
                    komunikat       TEXT
                );
            """)

    # ------------------------------------------------------------------
    # Zapis
    # ------------------------------------------------------------------

    def zapisz_ceny(self, rekordy: list[dict]) -> int:
        """Zapisuje listę rekordów cenowych (upsert).

        Każdy rekord: {timestamp_start, timestamp_end, cena_pln_mwh,
                       wolumen?, rynek?, waluta?, zrodlo?}
        Zwraca liczbę zapisanych rekordów.
        """
        if not rekordy:
            return 0
        with self._conn() as conn:
            conn.executemany("""
                INSERT OR REPLACE INTO ceny_15min
                    (timestamp_start, timestamp_end, rynek, cena_pln_mwh,
                     wolumen, waluta, zrodlo)
                VALUES
                    (:timestamp_start, :timestamp_end,
                     :rynek, :cena_pln_mwh,
                     :wolumen, :waluta, :zrodlo)
            """, [
                {
                    'timestamp_start': r['timestamp_start'],
                    'timestamp_end': r['timestamp_end'],
                    'rynek': r.get('rynek', 'RDB'),
                    'cena_pln_mwh': r['cena_pln_mwh'],
                    'wolumen': r.get('wolumen'),
                    'waluta': r.get('waluta', 'PLN'),
                    'zrodlo': r.get('zrodlo', 'TGE'),
                }
                for r in rekordy
            ])
        return len(rekordy)

    def zapisz_log(self, rynek: str, data_sesji: str, liczba_rekordow: int,
                   status: str, komunikat: str = ''):
        with self._conn() as conn:
            conn.execute("""
                INSERT INTO scraper_log
                    (timestamp, rynek, data_sesji, liczba_rekordow, status, komunikat)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (
                datetime.now().isoformat(),
                rynek, data_sesji, liczba_rekordow, status, komunikat,
            ))

    # ------------------------------------------------------------------
    # Odczyt
    # ------------------------------------------------------------------

    def pobierz_ceny(self, data_od: str, data_do: str,
                     rynek: str = 'RDB') -> pd.DataFrame:
        """Pobiera ceny z zakresu dat (format YYYY-MM-DD lub ISO)."""
        with self._conn() as conn:
            df = pd.read_sql_query("""
                SELECT timestamp_start, timestamp_end, rynek,
                       cena_pln_mwh, wolumen, waluta, zrodlo
                FROM ceny_15min
                WHERE timestamp_start >= ? AND timestamp_start < ?
                  AND rynek = ?
                ORDER BY timestamp_start
            """, conn, params=(data_od, data_do, rynek))
        if not df.empty:
            df['timestamp_start'] = pd.to_datetime(df['timestamp_start'])
            df['timestamp_end'] = pd.to_datetime(df['timestamp_end'])
            df['cena_pln_kwh'] = df['cena_pln_mwh'] / 1000.0
        return df

    def pobierz_ostatnie(self, n: int = 96, rynek: str = 'RDB') -> pd.DataFrame:
        """Pobiera ostatnich n rekordów (domyślnie 96 = 24h)."""
        with self._conn() as conn:
            df = pd.read_sql_query("""
                SELECT timestamp_start, timestamp_end, rynek,
                       cena_pln_mwh, wolumen, waluta, zrodlo
                FROM ceny_15min
                WHERE rynek = ?
                ORDER BY timestamp_start DESC
                LIMIT ?
            """, conn, params=(rynek, n))
        if not df.empty:
            df = df.sort_values('timestamp_start').reset_index(drop=True)
            df['timestamp_start'] = pd.to_datetime(df['timestamp_start'])
            df['timestamp_end'] = pd.to_datetime(df['timestamp_end'])
            df['cena_pln_kwh'] = df['cena_pln_mwh'] / 1000.0
        return df

    def statystyki_dzienne(self, data: str, rynek: str = 'RDB') -> dict:
        """Statystyki cenowe dla danego dnia."""
        data_do = (datetime.strptime(data, '%Y-%m-%d') + timedelta(days=1)).strftime('%Y-%m-%d')
        df = self.pobierz_ceny(data, data_do, rynek)
        if df.empty:
            return {}
        return {
            'data': data,
            'srednia_pln_mwh': round(df['cena_pln_mwh'].mean(), 2),
            'min_pln_mwh': round(df['cena_pln_mwh'].min(), 2),
            'max_pln_mwh': round(df['cena_pln_mwh'].max(), 2),
            'mediana_pln_mwh': round(df['cena_pln_mwh'].median(), 2),
            'liczba_rekordow': len(df),
            'srednia_pln_kwh': round(df['cena_pln_mwh'].mean() / 1000.0, 4),
        }

    def profil_godzinowy(self, data_od: str, data_do: str,
                         rynek: str = 'RDB') -> pd.DataFrame:
        """Średnia cena w podziale na godziny (0-23)."""
        df = self.pobierz_ceny(data_od, data_do, rynek)
        if df.empty:
            return pd.DataFrame()
        df['godzina'] = df['timestamp_start'].dt.hour
        profil = df.groupby('godzina')['cena_pln_mwh'].mean().reset_index()
        profil.columns = ['godzina', 'srednia_cena_pln_mwh']
        profil['srednia_cena_pln_kwh'] = profil['srednia_cena_pln_mwh'] / 1000.0
        return profil

    def spread_dzienny(self, data: str, rynek: str = 'RDB') -> Optional[float]:
        """Spread max-min dla danego dnia (PLN/MWh)."""
        stats = self.statystyki_dzienne(data, rynek)
        if not stats:
            return None
        return round(stats['max_pln_mwh'] - stats['min_pln_mwh'], 2)

    def srednia_rdb(self, dni: int = 30, rynek: str = 'RDB') -> Optional[float]:
        """Średnia cena RDB z ostatnich N dni (PLN/MWh)."""
        data_do = datetime.now().strftime('%Y-%m-%d')
        data_od = (datetime.now() - timedelta(days=dni)).strftime('%Y-%m-%d')
        df = self.pobierz_ceny(data_od, data_do, rynek)
        if df.empty:
            return None
        return round(df['cena_pln_mwh'].mean(), 2)

    def srednia_rdb_kwh(self, dni: int = 30, rynek: str = 'RDB') -> Optional[float]:
        """Średnia cena RDB z ostatnich N dni (PLN/kWh)."""
        avg = self.srednia_rdb(dni, rynek)
        if avg is None:
            return None
        return round(avg / 1000.0, 4)

    def spread_sredni_kwh(self, dni: int = 30, rynek: str = 'RDB') -> Optional[float]:
        """Średni dzienny spread z ostatnich N dni (PLN/kWh)."""
        data_do = datetime.now()
        spreads = []
        for i in range(dni):
            d = (data_do - timedelta(days=i)).strftime('%Y-%m-%d')
            s = self.spread_dzienny(d, rynek)
            if s is not None:
                spreads.append(s)
        if not spreads:
            return None
        return round(sum(spreads) / len(spreads) / 1000.0, 4)

    def liczba_rekordow(self, rynek: str = 'RDB') -> int:
        """Łączna liczba rekordów cenowych."""
        with self._conn() as conn:
            row = conn.execute(
                "SELECT COUNT(*) FROM ceny_15min WHERE rynek = ?", (rynek,)
            ).fetchone()
        return row[0] if row else 0

    def pobierz_logi(self, limit: int = 20) -> pd.DataFrame:
        """Pobiera ostatnie logi scrapera."""
        with self._conn() as conn:
            df = pd.read_sql_query("""
                SELECT timestamp, rynek, data_sesji, liczba_rekordow, status, komunikat
                FROM scraper_log
                ORDER BY id DESC
                LIMIT ?
            """, conn, params=(limit,))
        return df

    # ------------------------------------------------------------------
    # Import z pliku CSV / XLSX
    # ------------------------------------------------------------------

    def importuj_csv(self, sciezka: str, rynek: str = 'RDB') -> int:
        """Importuje ceny z pliku CSV.

        Oczekiwane kolumny: timestamp_start, timestamp_end, cena_pln_mwh
        Opcjonalne: wolumen, waluta, zrodlo
        """
        df = pd.read_csv(sciezka)
        return self._importuj_df(df, rynek)

    def importuj_xlsx(self, sciezka_lub_bytes, rynek: str = 'RDB') -> int:
        """Importuje ceny z pliku XLSX.

        Oczekiwane kolumny: timestamp_start, timestamp_end, cena_pln_mwh
        Opcjonalne: wolumen, waluta, zrodlo
        """
        df = pd.read_excel(sciezka_lub_bytes)
        return self._importuj_df(df, rynek)

    def _importuj_df(self, df: pd.DataFrame, rynek: str) -> int:
        """Importuje DataFrame do bazy."""
        wymagane = {'timestamp_start', 'timestamp_end', 'cena_pln_mwh'}
        if not wymagane.issubset(set(df.columns)):
            # Próba automatycznego mapowania popularnych nazw kolumn
            mapping = {
                'Data': 'timestamp_start',
                'Czas': 'timestamp_start',
                'Cena': 'cena_pln_mwh',
                'Kurs': 'cena_pln_mwh',
                'Wolumen': 'wolumen',
            }
            df = df.rename(columns={
                k: v for k, v in mapping.items() if k in df.columns
            })
            if not wymagane.issubset(set(df.columns)):
                raise ValueError(
                    f"Brak wymaganych kolumn: {wymagane - set(df.columns)}. "
                    f"Dostępne: {list(df.columns)}"
                )

        rekordy = []
        for _, row in df.iterrows():
            rekordy.append({
                'timestamp_start': str(row['timestamp_start']),
                'timestamp_end': str(row['timestamp_end']),
                'rynek': rynek,
                'cena_pln_mwh': float(row['cena_pln_mwh']),
                'wolumen': float(row['wolumen']) if 'wolumen' in row and pd.notna(row.get('wolumen')) else None,
                'waluta': row.get('waluta', 'PLN'),
                'zrodlo': row.get('zrodlo', 'import'),
            })
        return self.zapisz_ceny(rekordy)


if __name__ == '__main__':
    db = BazaCen()
    print(f"Baza: {db.db_path}")
    print(f"Rekordów RDB: {db.liczba_rekordow()}")
