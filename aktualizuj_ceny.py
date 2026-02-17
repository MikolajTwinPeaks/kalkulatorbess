#!/usr/bin/env python3
"""
Skrypt automatyzacji pobierania cen z TGE RDB.

Łączy scraper (scraper_tge.py) z bazą danych (baza_cen.py).
Gotowy do crontaba: np. codziennie o 23:30.

Użycie:
    python3 aktualizuj_ceny.py                  # pobierz ceny za dziś
    python3 aktualizuj_ceny.py --date 2026-02-15  # konkretna data
    python3 aktualizuj_ceny.py --backfill 7       # ostatnie 7 dni
    python3 aktualizuj_ceny.py --verbose          # szczegółowe logowanie

Crontab:
    30 23 * * * cd /path/to/BESS_PV_Kalkulator && python3 aktualizuj_ceny.py >> cron.log 2>&1
"""

import argparse
import logging
import sys
from datetime import datetime, timedelta

from baza_cen import BazaCen
from scraper_tge import ScraperTGE

logger = logging.getLogger(__name__)


def pobierz_i_zapisz(scraper: ScraperTGE, db: BazaCen, data: str,
                      verbose: bool = False) -> bool:
    """Pobiera ceny za daną datę i zapisuje do bazy.

    Returns:
        True jeśli sukces, False w razie błędu.
    """
    logger.info("=" * 60)
    logger.info("Pobieranie cen RDB za: %s", data)
    logger.info("=" * 60)

    try:
        ceny = scraper.pobierz_ceny_rdb(data)
    except Exception as e:
        logger.error("Błąd scrapera: %s", e)
        db.zapisz_log('RDB', data, 0, 'ERROR', str(e))
        return False

    if not ceny:
        msg = f"Brak danych cenowych za {data}"
        logger.warning(msg)
        db.zapisz_log('RDB', data, 0, 'EMPTY', msg)
        return False

    # Konwersja CenaRDB → dict do zapisu
    rekordy = []
    for c in ceny:
        rekordy.append({
            'timestamp_start': c.timestamp_start,
            'timestamp_end': c.timestamp_end,
            'cena_pln_mwh': c.cena_pln_mwh,
            'wolumen': c.wolumen_mwh,
            'rynek': 'RDB',
            'waluta': 'PLN',
            'zrodlo': 'TGE_scraper',
        })

    n = db.zapisz_ceny(rekordy)
    msg = f"Zapisano {n} rekordów za {data}"
    logger.info(msg)
    db.zapisz_log('RDB', data, n, 'OK', msg)

    if verbose:
        logger.info("Przykładowe rekordy:")
        for c in ceny[:5]:
            logger.info("  %s — %s: %.2f PLN/MWh",
                         c.timestamp_start, c.timestamp_end, c.cena_pln_mwh)

    return True


def _eksportuj_seed_csv(db: BazaCen):
    """Eksportuje całą bazę cen do dane/ceny_seed.csv (dla Cloud)."""
    import csv
    import os as _os
    base_dir = _os.path.dirname(_os.path.abspath(__file__))
    dane_dir = _os.path.join(base_dir, 'dane')
    _os.makedirs(dane_dir, exist_ok=True)
    seed_path = _os.path.join(dane_dir, 'ceny_seed.csv')

    import sqlite3
    conn = sqlite3.connect(db.db_path)
    cur = conn.execute(
        'SELECT timestamp_start, timestamp_end, cena_pln_mwh, '
        'wolumen, rynek, waluta, zrodlo FROM ceny_15min ORDER BY timestamp_start'
    )
    rows = cur.fetchall()
    conn.close()

    with open(seed_path, 'w', newline='') as f:
        w = csv.writer(f)
        w.writerow(['timestamp_start', 'timestamp_end', 'cena_pln_mwh',
                     'wolumen', 'rynek', 'waluta', 'zrodlo'])
        w.writerows(rows)
    logger.info("Wyeksportowano %d rekordów do %s", len(rows), seed_path)


def _git_push_seed():
    """Commituje i pushuje zaktualizowany seed CSV do GitHub."""
    import subprocess
    import os as _os
    base_dir = _os.path.dirname(_os.path.abspath(__file__))
    try:
        subprocess.run(
            ['git', 'add', 'dane/ceny_seed.csv'],
            cwd=base_dir, check=True, capture_output=True,
        )
        result = subprocess.run(
            ['git', 'diff', '--cached', '--quiet'],
            cwd=base_dir, capture_output=True,
        )
        if result.returncode == 0:
            logger.info("Seed CSV bez zmian — pomijam push")
            return
        subprocess.run(
            ['git', 'commit', '-m', 'Auto-update ceny_seed.csv'],
            cwd=base_dir, check=True, capture_output=True,
        )
        subprocess.run(
            ['git', 'push'],
            cwd=base_dir, check=True, capture_output=True, timeout=30,
        )
        logger.info("Seed CSV zpushowany do GitHub")
    except Exception as e:
        logger.warning("Nie udało się zpushować seed CSV: %s", e)


def main():
    parser = argparse.ArgumentParser(
        description='Pobieranie cen 15-minutowych z TGE RDB'
    )
    parser.add_argument('--date', type=str, default=None,
                        help='Data sesji (YYYY-MM-DD). Domyślnie: dziś.')
    parser.add_argument('--backfill', type=int, default=0,
                        help='Pobierz ceny za ostatnie N dni')
    parser.add_argument('--headless', action='store_true', default=True,
                        help='Tryb headless (domyślnie włączony)')
    parser.add_argument('--no-headless', action='store_true',
                        help='Wyłącz tryb headless (otwiera przeglądarkę)')
    parser.add_argument('--verbose', '-v', action='store_true',
                        help='Szczegółowe logowanie')
    parser.add_argument('--db', type=str, default=None,
                        help='Ścieżka do bazy SQLite')
    args = parser.parse_args()

    # Logowanie
    log_level = logging.DEBUG if args.verbose else logging.INFO
    logging.basicConfig(
        level=log_level,
        format='%(asctime)s [%(levelname)s] %(name)s: %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S',
    )

    headless = not args.no_headless

    # Baza danych
    db = BazaCen(args.db) if args.db else BazaCen()
    logger.info("Baza danych: %s (rekordów: %d)", db.db_path, db.liczba_rekordow())

    # Lista dat do pobrania
    daty = []
    if args.backfill > 0:
        for i in range(args.backfill):
            d = (datetime.now() - timedelta(days=i)).strftime('%Y-%m-%d')
            daty.append(d)
        daty.reverse()  # od najstarszej
    elif args.date:
        daty = [args.date]
    else:
        daty = [datetime.now().strftime('%Y-%m-%d')]

    logger.info("Daty do pobrania: %s", daty)

    # Scraping
    sukcesy = 0
    bledy = 0

    with ScraperTGE(headless=headless, verbose=args.verbose) as scraper:
        for data in daty:
            ok = pobierz_i_zapisz(scraper, db, data, verbose=args.verbose)
            if ok:
                sukcesy += 1
            else:
                bledy += 1

    # Eksport seed CSV + push do GitHub (sync Cloud)
    if sukcesy > 0:
        _eksportuj_seed_csv(db)
        _git_push_seed()

    # Podsumowanie
    logger.info("")
    logger.info("=" * 60)
    logger.info("PODSUMOWANIE")
    logger.info("  Dat przetworzonych: %d", len(daty))
    logger.info("  Sukcesy: %d", sukcesy)
    logger.info("  Błędy: %d", bledy)
    logger.info("  Łączna liczba rekordów w bazie: %d", db.liczba_rekordow())
    logger.info("=" * 60)

    if bledy > 0:
        sys.exit(1)


if __name__ == '__main__':
    main()
