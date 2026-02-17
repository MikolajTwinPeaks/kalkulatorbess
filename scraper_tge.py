#!/usr/bin/env python3
"""
Scraper cen 15-minutowych z TGE (Rynek Dnia Bieżącego).

Używa Selenium + Chrome (headless) z automatycznym pobieraniem ChromeDriver
przez webdriver-manager.

Strategia:
1. Nawigacja na stronę TGE RDB
2. Tryb discovery — logowanie struktury DOM
3. Scraping tabeli HTML z cenami 15-minutowymi
4. Fallback: pobranie pliku XLSX z linku na stronie

Użycie:
    from scraper_tge import ScraperTGE
    with ScraperTGE() as scraper:
        ceny = scraper.pobierz_ceny_rdb()

    # Tryb discovery (nie-headless):
    python3 scraper_tge.py --discovery
"""

import argparse
import logging
import re
import time
from dataclasses import dataclass
from datetime import datetime, timedelta
from typing import Optional

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, NoSuchElementException, WebDriverException,
)

try:
    from webdriver_manager.chrome import ChromeDriverManager
    _HAS_WDM = True
except ImportError:
    _HAS_WDM = False

logger = logging.getLogger(__name__)

# URL-e TGE
URL_RDB = 'https://tge.pl/energia-elektryczna-rdb'
URL_RDB_VARIANTS = [
    'https://tge.pl/energia-elektryczna-rdb',
    'https://www.tge.pl/energia-elektryczna-rdb',
]


@dataclass
class CenaRDB:
    """Pojedynczy rekord ceny 15-minutowej z TGE RDB."""
    timestamp_start: str   # ISO format
    timestamp_end: str     # ISO format
    cena_pln_mwh: float
    wolumen_mwh: Optional[float] = None


class ScraperTGE:
    """Scraper cen z TGE z obsługą Selenium Chrome."""

    MAX_RETRIES = 3
    RETRY_DELAY = 5  # sekundy
    PAGE_LOAD_TIMEOUT = 30  # sekundy

    def __init__(self, headless: bool = True, verbose: bool = False):
        self.headless = headless
        self.verbose = verbose
        self.driver = None

        if verbose:
            logging.basicConfig(level=logging.DEBUG)
        else:
            logging.basicConfig(level=logging.INFO)

    def __enter__(self):
        self._init_driver()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.zamknij()
        return False

    def _init_driver(self):
        """Inicjalizacja Chrome WebDriver."""
        options = Options()
        if self.headless:
            options.add_argument('--headless=new')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--disable-gpu')
        options.add_argument('--window-size=1920,1080')
        options.add_argument('--lang=pl-PL')
        options.add_argument(
            '--user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) '
            'AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
        )

        if _HAS_WDM:
            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=options)
        else:
            # Fallback — zakłada że chromedriver jest w PATH
            self.driver = webdriver.Chrome(options=options)

        self.driver.set_page_load_timeout(self.PAGE_LOAD_TIMEOUT)
        logger.info("Chrome WebDriver zainicjalizowany (headless=%s)", self.headless)

    def zamknij(self):
        """Zamyka przeglądarkę."""
        if self.driver:
            try:
                self.driver.quit()
            except Exception:
                pass
            self.driver = None

    # ------------------------------------------------------------------
    # Cookie consent
    # ------------------------------------------------------------------

    def _akceptuj_cookies(self):
        """Próba zamknięcia okienka cookies."""
        selektory = [
            "button[id*='cookie']",
            "button[class*='cookie']",
            "button[class*='accept']",
            "a[class*='cookie']",
            ".cookies-accept",
            "#cookies-accept",
            "button.btn-accept",
            "[data-action='accept-cookies']",
        ]
        for sel in selektory:
            try:
                btn = self.driver.find_element(By.CSS_SELECTOR, sel)
                if btn.is_displayed():
                    btn.click()
                    logger.info("Cookie consent zamknięty: %s", sel)
                    time.sleep(1)
                    return True
            except (NoSuchElementException, WebDriverException):
                continue

        # Próba po tekście
        for tekst in ['Akceptuję', 'Zgadzam się', 'Accept', 'OK', 'Rozumiem']:
            try:
                btn = self.driver.find_element(
                    By.XPATH, f"//button[contains(text(), '{tekst}')]"
                )
                if btn.is_displayed():
                    btn.click()
                    logger.info("Cookie consent zamknięty (tekst: %s)", tekst)
                    time.sleep(1)
                    return True
            except (NoSuchElementException, WebDriverException):
                continue

        logger.debug("Brak okienka cookies do zamknięcia")
        return False

    # ------------------------------------------------------------------
    # Parsowanie polskich liczb
    # ------------------------------------------------------------------

    @staticmethod
    def _parse_pl_number(text: str) -> Optional[float]:
        """Parsuje polską liczbę: '1 234,56' → 1234.56"""
        if not text or text.strip() in ('-', '', 'n/a', 'b.d.'):
            return None
        cleaned = text.strip()
        cleaned = cleaned.replace('\xa0', '').replace(' ', '')
        cleaned = cleaned.replace(',', '.')
        try:
            return float(cleaned)
        except ValueError:
            return None

    # ------------------------------------------------------------------
    # Discovery mode
    # ------------------------------------------------------------------

    def discovery(self, url: str = URL_RDB) -> dict:
        """Tryb discovery — loguje strukturę strony TGE."""
        if not self.driver:
            self._init_driver()

        logger.info("Discovery mode: %s", url)
        self.driver.get(url)
        time.sleep(3)
        self._akceptuj_cookies()
        time.sleep(2)

        info = {
            'url': url,
            'title': self.driver.title,
            'tables': [],
            'links_xlsx': [],
            'links_csv': [],
            'iframes': [],
        }

        # Tabele
        tables = self.driver.find_elements(By.TAG_NAME, 'table')
        for i, t in enumerate(tables):
            rows = t.find_elements(By.TAG_NAME, 'tr')
            headers = []
            if rows:
                ths = rows[0].find_elements(By.TAG_NAME, 'th')
                headers = [th.text.strip() for th in ths]
            info['tables'].append({
                'index': i,
                'rows': len(rows),
                'headers': headers,
                'class': t.get_attribute('class'),
                'id': t.get_attribute('id'),
            })

        # Linki do plików
        links = self.driver.find_elements(By.TAG_NAME, 'a')
        for link in links:
            href = link.get_attribute('href') or ''
            text = link.text.strip()
            if '.xlsx' in href.lower() or '.xls' in href.lower():
                info['links_xlsx'].append({'href': href, 'text': text})
            if '.csv' in href.lower():
                info['links_csv'].append({'href': href, 'text': text})

        # Iframes
        iframes = self.driver.find_elements(By.TAG_NAME, 'iframe')
        for iframe in iframes:
            info['iframes'].append({
                'src': iframe.get_attribute('src'),
                'id': iframe.get_attribute('id'),
            })

        logger.info("Discovery zakończony:")
        logger.info("  Tytuł: %s", info['title'])
        logger.info("  Tabele: %d", len(info['tables']))
        for t in info['tables']:
            logger.info("    Tabela %d: %d wierszy, headers=%s, class=%s",
                         t['index'], t['rows'], t['headers'], t['class'])
        logger.info("  Linki XLSX: %d", len(info['links_xlsx']))
        for lnk in info['links_xlsx']:
            logger.info("    %s → %s", lnk['text'], lnk['href'])
        logger.info("  Linki CSV: %d", len(info['links_csv']))
        logger.info("  Iframes: %d", len(info['iframes']))

        return info

    # ------------------------------------------------------------------
    # Pobieranie cen
    # ------------------------------------------------------------------

    def pobierz_ceny_rdb(self, data: Optional[str] = None) -> list[CenaRDB]:
        """Pobiera ceny 15-minutowe RDB z TGE.

        Args:
            data: Data sesji (YYYY-MM-DD). Domyślnie dzisiaj.

        Returns:
            Lista obiektów CenaRDB.
        """
        if not self.driver:
            self._init_driver()

        if data is None:
            data = datetime.now().strftime('%Y-%m-%d')

        for attempt in range(1, self.MAX_RETRIES + 1):
            try:
                logger.info("Próba %d/%d pobierania cen RDB za %s",
                            attempt, self.MAX_RETRIES, data)
                ceny = self._scrape_tabela_rdb(data)
                if ceny:
                    logger.info("Pobrano %d rekordów cenowych", len(ceny))
                    return ceny

                # Fallback: próba pobrania XLSX
                logger.info("Tabela pusta/niedostępna, próba fallback XLSX")
                ceny = self._scrape_xlsx_fallback(data)
                if ceny:
                    logger.info("Fallback XLSX: pobrano %d rekordów", len(ceny))
                    return ceny

                logger.warning("Próba %d: brak danych", attempt)
            except Exception as e:
                logger.error("Próba %d: błąd — %s", attempt, e)

            if attempt < self.MAX_RETRIES:
                logger.info("Czekam %ds przed ponowną próbą...", self.RETRY_DELAY)
                time.sleep(self.RETRY_DELAY)

        logger.error("Nie udało się pobrać cen po %d próbach", self.MAX_RETRIES)
        return []

    def _scrape_tabela_rdb(self, data: str) -> list[CenaRDB]:
        """Scraping tabeli HTML z cenami RDB."""
        url = URL_RDB
        logger.debug("Ładuję stronę: %s", url)
        self.driver.get(url)

        # Czekaj na załadowanie DataTables
        try:
            WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'table.table-rdb'))
            )
        except TimeoutException:
            logger.warning("Timeout oczekiwania na tabelę")
            return []

        time.sleep(3)
        self._akceptuj_cookies()

        # Szukaj największej tabeli z klasą table-rdb (główna tabela danych)
        tables = self.driver.find_elements(By.CSS_SELECTOR, 'table.table-rdb')
        ceny = []

        for table in tables:
            # Zbierz wszystkie wiersze tbody (pomijamy thead — na TGE jest pusty/dynamiczny)
            tbody_rows = table.find_elements(By.CSS_SELECTOR, 'tbody tr')
            if not tbody_rows:
                tbody_rows = table.find_elements(By.TAG_NAME, 'tr')
            if len(tbody_rows) < 10:
                continue

            logger.debug("Tabela table-rdb: %d wierszy w tbody", len(tbody_rows))

            # Loguj pierwsze 3 wiersze dla diagnostyki
            for i, row in enumerate(tbody_rows[:3]):
                cells = row.find_elements(By.TAG_NAME, 'td')
                cell_texts = [c.text.strip()[:30] for c in cells]
                logger.debug("  Wiersz %d (%d komórek): %s", i, len(cells), cell_texts)

            # TGE RDB structure (29 cols):
            #   [0] Period: "2026-02-17_H01" (hourly) or "2026-02-17_Q00:15" (15-min)
            #   [1] Duration: 60 or 15
            #   [2-7]  CT session (various prices/volumes, often "-")
            #   [8-17] IDA1 session prices/volumes
            #   [18-22] IDA2/IDA3 sessions
            #   [23-28] "Łącznie" (total): kurs min, kurs max, kurs śr. ważony, wolumen...
            # We want: 15-min rows (_Q) with the weighted average price from "Łącznie"
            sample_cells = tbody_rows[0].find_elements(By.TAG_NAME, 'td')
            n_cols = len(sample_cells)
            logger.info("Tabela: %d wierszy, %d kolumn", len(tbody_rows), n_cols)

            # Find the weighted avg price column: scan from the right side
            # looking for the last group of price-like values (100-2000 PLN/MWh)
            # In "Łącznie" section: [kurs_min, kurs_max, kurs_sr_wazony, vol1, vol2, vol3]
            col_cena = None
            col_wolumen = None

            # Analyze a 15-min row (skip _H rows which may have all dashes)
            for test_idx in range(min(10, len(tbody_rows))):
                test_cells = tbody_rows[test_idx].find_elements(By.TAG_NAME, 'td')
                period = test_cells[0].text.strip() if test_cells else ''
                if '_Q' not in period:
                    continue  # Skip hourly rows

                # Scan from the right for the "Łącznie" section
                # Find rightmost group of price-range values (100-2000)
                price_cols = []
                for ci in range(n_cols - 1, 1, -1):
                    val = self._parse_pl_number(test_cells[ci].text)
                    if val is not None and 100 < val < 2000:
                        price_cols.append(ci)
                    elif price_cols:
                        break  # End of price group

                if len(price_cols) >= 2:
                    # price_cols is reversed (right to left), so:
                    # rightmost group = [kurs_min, kurs_max, kurs_sr_wazony]
                    # The middle/last one in the group is the weighted average
                    price_cols.sort()
                    # Pick the last (rightmost) price column = weighted average
                    col_cena = price_cols[-1]
                    # Look for volume column right after prices
                    if col_cena + 1 < n_cols:
                        vol_val = self._parse_pl_number(test_cells[col_cena + 1].text)
                        if vol_val is not None and 0 < vol_val < 10000:
                            col_wolumen = col_cena + 1
                    break

            if col_cena is None:
                # Fallback: scan any column > col 2 for price-like value
                for test_idx in range(min(10, len(tbody_rows))):
                    test_cells = tbody_rows[test_idx].find_elements(By.TAG_NAME, 'td')
                    for ci in range(2, len(test_cells)):
                        val = self._parse_pl_number(test_cells[ci].text)
                        if val is not None and 100 < val < 2000:
                            col_cena = ci
                            break
                    if col_cena is not None:
                        break

            if col_cena is None:
                logger.debug("Nie znaleziono kolumny z ceną w tabeli")
                continue

            logger.info("Wykryto kolumny: okres=0, cena=%d, wolumen=%s (z %d)",
                        col_cena, col_wolumen, n_cols)

            # Parse all 15-minute rows
            for row in tbody_rows:
                cells = row.find_elements(By.TAG_NAME, 'td')
                if len(cells) <= col_cena:
                    continue

                period_text = cells[0].text.strip()
                # Only 15-min quarters (_Q), skip hourly aggregates (_H)
                if '_Q' not in period_text and '_q' not in period_text:
                    continue

                cena_val = self._parse_pl_number(cells[col_cena].text)
                if cena_val is None or cena_val < 1:
                    continue

                ts_start, ts_end = self._parse_time_range(data, period_text)
                if ts_start is None:
                    continue

                wolumen = None
                if col_wolumen is not None and col_wolumen < len(cells):
                    wolumen = self._parse_pl_number(cells[col_wolumen].text)

                ceny.append(CenaRDB(
                    timestamp_start=ts_start,
                    timestamp_end=ts_end,
                    cena_pln_mwh=cena_val,
                    wolumen_mwh=wolumen,
                ))

            if ceny:
                break  # Found data in this table

        return ceny

    def _scrape_xlsx_fallback(self, data: str) -> list[CenaRDB]:
        """Fallback: pobieranie XLSX z linku na stronie TGE."""
        import tempfile
        import os

        links = self.driver.find_elements(By.TAG_NAME, 'a')
        xlsx_links = []
        for link in links:
            href = link.get_attribute('href') or ''
            if '.xlsx' in href.lower() or '.xls' in href.lower():
                xlsx_links.append(href)

        if not xlsx_links:
            logger.debug("Brak linków XLSX na stronie")
            return []

        # Pobierz plik XLSX (requests z obsługą SSL)
        try:
            import requests as _requests
        except ImportError:
            import urllib.request
            import ssl
            _requests = None

        for xlsx_url in xlsx_links:
            try:
                logger.info("Pobieram XLSX: %s", xlsx_url)
                tmp = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)

                if _requests is not None:
                    resp = _requests.get(xlsx_url, timeout=30)
                    resp.raise_for_status()
                    tmp.write(resp.content)
                else:
                    ctx = ssl.create_default_context()
                    ctx.check_hostname = False
                    ctx.verify_mode = ssl.CERT_NONE
                    with urllib.request.urlopen(xlsx_url, context=ctx) as r:
                        tmp.write(r.read())

                tmp.close()
                ceny = self._parse_xlsx_rdb(tmp.name, data)
                os.unlink(tmp.name)
                if ceny:
                    return ceny
            except Exception as e:
                logger.warning("Błąd pobierania XLSX: %s", e)
                try:
                    os.unlink(tmp.name)
                except OSError:
                    pass
                continue

        return []

    def _parse_xlsx_rdb(self, path: str, data: str) -> list[CenaRDB]:
        """Parsuje plik XLSX z cenami RDB."""
        import pandas as pd

        try:
            df = pd.read_excel(path)
        except Exception as e:
            logger.error("Błąd odczytu XLSX: %s", e)
            return []

        logger.debug("XLSX kolumny: %s", list(df.columns))
        logger.debug("XLSX wierszy: %d", len(df))

        ceny = []
        # Próba identyfikacji kolumn
        col_names = [str(c).lower() for c in df.columns]

        col_cena = None
        col_czas = None
        col_wolumen = None

        for i, c in enumerate(col_names):
            if any(w in c for w in ('cena', 'kurs', 'price', 'fixing')):
                col_cena = i
            elif any(w in c for w in ('godz', 'czas', 'time', 'hour', 'okres')):
                col_czas = i
            elif any(w in c for w in ('wolumen', 'volume')):
                col_wolumen = i

        if col_cena is None:
            logger.warning("Nie znaleziono kolumny z ceną w XLSX")
            return []

        for _, row in df.iterrows():
            cena_val = self._parse_pl_number(str(row.iloc[col_cena]))
            if cena_val is None:
                continue

            if col_czas is not None:
                time_text = str(row.iloc[col_czas])
            else:
                time_text = str(row.iloc[0])

            ts_start, ts_end = self._parse_time_range(data, time_text)
            if ts_start is None:
                continue

            wolumen = None
            if col_wolumen is not None:
                wolumen = self._parse_pl_number(str(row.iloc[col_wolumen]))

            ceny.append(CenaRDB(
                timestamp_start=ts_start,
                timestamp_end=ts_end,
                cena_pln_mwh=cena_val,
                wolumen_mwh=wolumen,
            ))

        return ceny

    # ------------------------------------------------------------------
    # Parsowanie czasu
    # ------------------------------------------------------------------

    @staticmethod
    def _parse_time_range(data: str, text: str) -> tuple[Optional[str], Optional[str]]:
        """Parsuje tekst z godziną/zakresem na (timestamp_start, timestamp_end).

        Obsługuje formaty TGE i standardowe:
        - '2026-02-17_Q00:15' (TGE 15-min quarter)
        - '2026-02-17_H01' (TGE hourly)
        - '08:00 - 08:15', '08:00-08:15'
        - '08:00', '8:00'
        - '8', '08'
        - '1' (numer okresu 15-min: 1=00:00-00:15, 96=23:45-24:00)
        """
        text = text.strip()

        # TGE format: "YYYY-MM-DD_QHH:MM" (15-minute quarter)
        m = re.match(r'(\d{4}-\d{2}-\d{2})_[Qq](\d{2}):(\d{2})', text)
        if m:
            d, h, mi = m.group(1), int(m.group(2)), int(m.group(3))
            # Handle 24:00 = midnight next day
            if h == 24:
                start = datetime.strptime(d, '%Y-%m-%d') + timedelta(days=1, minutes=mi)
            else:
                start = datetime.strptime(f"{d} {h:02d}:{mi:02d}", '%Y-%m-%d %H:%M')
            end = start + timedelta(minutes=15)
            return start.strftime('%Y-%m-%dT%H:%M:%S'), end.strftime('%Y-%m-%dT%H:%M:%S')

        # TGE format: "YYYY-MM-DD_HNN" (hourly, NN=01..24)
        m = re.match(r'(\d{4}-\d{2}-\d{2})_[Hh](\d{2})', text)
        if m:
            d, h = m.group(1), int(m.group(2))
            # TGE hours are 1-24 where H01=00:00-01:00, H24=23:00-24:00
            start = datetime.strptime(d, '%Y-%m-%d') + timedelta(hours=h - 1)
            end = start + timedelta(hours=1)
            return start.strftime('%Y-%m-%dT%H:%M:%S'), end.strftime('%Y-%m-%dT%H:%M:%S')

        # Format zakres: "HH:MM - HH:MM" lub "HH:MM-HH:MM"
        m = re.match(r'(\d{1,2}):(\d{2})\s*[-–]\s*(\d{1,2}):(\d{2})', text)
        if m:
            h1, m1, h2, m2 = int(m.group(1)), int(m.group(2)), int(m.group(3)), int(m.group(4))
            ts_start = f"{data}T{h1:02d}:{m1:02d}:00"
            ts_end = f"{data}T{h2:02d}:{m2:02d}:00"
            return ts_start, ts_end

        # Format godzina: "HH:MM"
        m = re.match(r'^(\d{1,2}):(\d{2})$', text)
        if m:
            h, mi = int(m.group(1)), int(m.group(2))
            start = datetime.strptime(f"{data} {h:02d}:{mi:02d}", '%Y-%m-%d %H:%M')
            end = start + timedelta(minutes=15)
            return start.strftime('%Y-%m-%dT%H:%M:%S'), end.strftime('%Y-%m-%dT%H:%M:%S')

        # Format numer okresu (1-96)
        m = re.match(r'^(\d{1,2})$', text)
        if m:
            n = int(m.group(1))
            if 0 <= n <= 23:
                start = datetime.strptime(f"{data} {n:02d}:00", '%Y-%m-%d %H:%M')
                end = start + timedelta(hours=1)
                return start.strftime('%Y-%m-%dT%H:%M:%S'), end.strftime('%Y-%m-%dT%H:%M:%S')
            elif 1 <= n <= 96:
                minutes = (n - 1) * 15
                start = datetime.strptime(data, '%Y-%m-%d') + timedelta(minutes=minutes)
                end = start + timedelta(minutes=15)
                return start.strftime('%Y-%m-%dT%H:%M:%S'), end.strftime('%Y-%m-%dT%H:%M:%S')

        logger.debug("Nie rozpoznano formatu czasu: '%s'", text)
        return None, None


# ======================================================================
# CLI
# ======================================================================

def main():
    parser = argparse.ArgumentParser(description='Scraper TGE RDB')
    parser.add_argument('--discovery', action='store_true',
                        help='Tryb discovery — logowanie struktury DOM')
    parser.add_argument('--headless', action='store_true', default=False,
                        help='Uruchom w trybie headless')
    parser.add_argument('--date', type=str, default=None,
                        help='Data sesji (YYYY-MM-DD)')
    parser.add_argument('--verbose', '-v', action='store_true')
    args = parser.parse_args()

    if args.discovery:
        with ScraperTGE(headless=args.headless, verbose=True) as scraper:
            info = scraper.discovery()
            print("\n=== DISCOVERY RESULTS ===")
            print(f"Tytuł: {info['title']}")
            print(f"Tabele: {len(info['tables'])}")
            for t in info['tables']:
                print(f"  [{t['index']}] {t['rows']} wierszy, "
                      f"headers={t['headers']}, class={t['class']}")
            print(f"Linki XLSX: {len(info['links_xlsx'])}")
            for lnk in info['links_xlsx']:
                print(f"  {lnk['text']} → {lnk['href']}")
            print(f"Linki CSV: {len(info['links_csv'])}")
            print(f"Iframes: {len(info['iframes'])}")
    else:
        headless = args.headless if not args.discovery else False
        with ScraperTGE(headless=headless, verbose=args.verbose) as scraper:
            ceny = scraper.pobierz_ceny_rdb(args.date)
            print(f"\nPobrano {len(ceny)} rekordów cenowych:")
            for c in ceny[:10]:
                print(f"  {c.timestamp_start} — {c.timestamp_end}: "
                      f"{c.cena_pln_mwh:.2f} PLN/MWh"
                      f"{f', vol={c.wolumen_mwh:.1f}' if c.wolumen_mwh else ''}")
            if len(ceny) > 10:
                print(f"  ... i {len(ceny) - 10} więcej")


if __name__ == '__main__':
    main()
