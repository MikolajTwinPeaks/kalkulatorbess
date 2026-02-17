"""
Analiza faktury za energię elektryczną — parser PDF + rekomendacje bezkosztowe.

Bezpieczeństwo: plik przetwarzany wyłącznie w RAM (BytesIO), nigdy zapisywany na dysk.
"""

import io
import re
import math
from dataclasses import dataclass, field
from typing import Optional


# ============================================================
# DATACLASSES
# ============================================================

@dataclass
class DaneFaktury:
    """Dane odczytane z faktury za energię elektryczną."""
    # Identyfikacja
    ppe: str = ''
    nr_faktury: str = ''
    okres_od: str = ''
    okres_do: str = ''
    # Taryfa i OSD
    taryfa: str = ''
    osd: str = ''
    # Moc
    moc_umowna_kw: float = 0.0
    # Zużycie per strefa (kWh)
    zuzycie_szczyt_kwh: float = 0.0
    zuzycie_pozaszczyt_kwh: float = 0.0
    zuzycie_noc_kwh: float = 0.0
    zuzycie_calkowite_kwh: float = 0.0
    # Ceny i opłaty
    cena_energii_pln_kwh: float = 0.0
    oplata_dystrybucyjna_pln_kwh: float = 0.0
    oplata_mocowa_pln_mwh: float = 0.0
    oplata_oze_pln_mwh: float = 0.0
    oplata_kogeneracja_pln_mwh: float = 0.0
    oplata_jakosciowa_pln_mwh: float = 0.0
    oplata_przejsciowa_pln_mwh: float = 0.0
    oplata_abonamentowa_pln: float = 0.0
    oplata_sieciowa_stala_pln: float = 0.0
    # Moc bierna
    tg_phi: float = 0.0
    moc_bierna_kvarh: float = 0.0
    oplata_moc_bierna_pln: float = 0.0
    # Kwoty
    kwota_netto_pln: float = 0.0
    kwota_brutto_pln: float = 0.0
    kwota_energia_pln: float = 0.0
    kwota_dystrybucja_pln: float = 0.0
    # Kategoria mocowa
    kategoria_mocowa: str = ''
    # Typ faktury
    typ_faktury: str = ''  # 'rozliczeniowa', 'prognozowa', 'korygująca'
    # Pewność parsowania (0-1)
    pewnosc: float = 0.0


@dataclass
class Rekomendacja:
    """Pojedyncza rekomendacja optymalizacyjna."""
    tytul: str = ''
    opis: str = ''
    oszczednosc_roczna_pln: float = 0.0
    priorytet: str = 'sredni'  # 'wysoki', 'sredni', 'niski'


@dataclass
class AnalizaOptymalizacji:
    """Wyniki analizy optymalizacyjnej faktury."""
    analiza_taryfy: str = ''
    analiza_moc_umowna: str = ''
    analiza_moc_bierna: str = ''
    analiza_oplata_mocowa: str = ''
    rekomendacje: list = field(default_factory=list)
    laczna_oszczednosc_roczna_pln: float = 0.0


# ============================================================
# HELPERS
# ============================================================

def _parse_polish_float(s: str) -> Optional[float]:
    """Parsuje polską notację liczbową: '1 234,56' → 1234.56"""
    if not s:
        return None
    s = s.strip()
    # Usuń spacje (separator tysięcy)
    s = s.replace(' ', '').replace('\u00a0', '')
    # Zamień przecinek na kropkę
    s = s.replace(',', '.')
    # Usuń PLN, zł, kWh itp.
    s = re.sub(r'[a-zA-ZłŁ/]+$', '', s).strip()
    try:
        return float(s)
    except ValueError:
        return None


def _first_match(pattern: str, text: str, group: int = 1) -> str:
    """Zwraca pierwszy match regex lub pusty string."""
    m = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
    return m.group(group).strip() if m else ''


def _first_float(pattern: str, text: str, group: int = 1) -> float:
    """Zwraca pierwszy match regex sparsowany jako float, lub 0."""
    raw = _first_match(pattern, text, group)
    val = _parse_polish_float(raw)
    return val if val is not None else 0.0


# ============================================================
# PDF TEXT EXTRACTION
# ============================================================

def _ocr_pdf(pdf_bytes: bytes) -> str:
    """OCR fallback dla skanowanych PDF — 100% in-memory, offline."""
    from pdf2image import convert_from_bytes
    import pytesseract

    images = convert_from_bytes(pdf_bytes)
    ocr_text = ''
    for img in images:
        ocr_text += pytesseract.image_to_string(img, lang='pol') + '\n'
        img.close()
    del images
    return ocr_text


# Minimalny próg znaków z pdfplumber, poniżej którego uruchamiamy OCR
_MIN_TEXT_LEN = 50


def _extract_text(pdf_bytes: bytes) -> tuple[str, list]:
    """Wyciąga tekst i tabele z PDF (100% in-memory).

    Próbuje pdfplumber (natywny tekst). Jeśli wynik jest zbyt krótki
    (skan), automatycznie uruchamia OCR (pytesseract) jako fallback.
    """
    import pdfplumber

    full_text = ''
    all_tables = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                full_text += page_text + '\n'
            tables = page.extract_tables()
            if tables:
                all_tables.extend(tables)

    # Fallback OCR dla skanów
    if len(full_text.strip()) < _MIN_TEXT_LEN:
        try:
            full_text = _ocr_pdf(pdf_bytes)
        except Exception:
            pass  # jeśli OCR niedostępny, zwróć co mamy

    return full_text, all_tables


# ============================================================
# REGEX EXTRACTORS
# ============================================================

def _extract_ppe(text: str) -> str:
    """Numer PPE — format: PL + cyfry lub PLXXXX...."""
    return _first_match(r'(?:PPE|punkt\s+poboru)[:\s]*([A-Z0-9]{16,18})', text)


def _extract_nr_faktury(text: str) -> str:
    return _first_match(
        r'(?:nr\s+faktury|faktura\s+(?:nr|VAT)|numer\s+faktury)[:\s]*([A-Z0-9/\-]+)',
        text,
    )


def _extract_okres(text: str) -> tuple[str, str]:
    """Okres rozliczeniowy — np. '01.01.2026 - 31.01.2026'."""
    m = re.search(
        r'(?:okres|za\s+okres)[:\s]*(\d{2}[.\-/]\d{2}[.\-/]\d{4})\s*[-–]\s*(\d{2}[.\-/]\d{2}[.\-/]\d{4})',
        text, re.IGNORECASE,
    )
    if m:
        return m.group(1).strip(), m.group(2).strip()
    return '', ''


def _extract_taryfa(text: str) -> str:
    """Grupa taryfowa — C11, C12a, C12b, C21, C22a, C22b, B21, B23, itp."""
    return _first_match(
        r'(?:grupa\s+taryfowa|taryfa)[:\s]*([A-Z]\d{2}[a-z]?)',
        text,
    )


def _extract_osd(text: str) -> str:
    """Operator Systemu Dystrybucyjnego."""
    osd_names = ['Tauron', 'Enea', 'Energa', 'PGE', 'innogy', 'Stoen']
    text_lower = text.lower()
    for name in osd_names:
        if name.lower() in text_lower:
            return name
    return _first_match(r'(?:OSD|operator)[:\s]*(\S+)', text)


def _extract_moc_umowna(text: str) -> float:
    return _first_float(
        r'(?:moc\s+umowna|moc\s+przy)[:\s]*([\d\s,\.]+)\s*kW',
        text,
    )


def _extract_zuzycie(text: str) -> dict:
    """Zużycie energii per strefa."""
    result = {
        'szczyt': 0.0,
        'pozaszczyt': 0.0,
        'noc': 0.0,
        'calkowite': 0.0,
    }

    # Zużycie całkowite
    result['calkowite'] = _first_float(
        r'(?:zu[żz]ycie|energia\s+czynna|ilo[śs][ćc])\s*(?:ca[łl]kowit[ea]|razem|og[óo][łl]em)?[:\s]*([\d\s,\.]+)\s*kWh',
        text,
    )

    # Strefy
    result['szczyt'] = _first_float(
        r'(?:strefa\s+szczytowa|szczyt|strefa\s+I|dzień)[:\s]*([\d\s,\.]+)\s*kWh',
        text,
    )
    result['pozaszczyt'] = _first_float(
        r'(?:strefa\s+pozaszczytowa|pozaszczyt|strefa\s+II)[:\s]*([\d\s,\.]+)\s*kWh',
        text,
    )
    result['noc'] = _first_float(
        r'(?:strefa\s+nocna|noc|strefa\s+III)[:\s]*([\d\s,\.]+)\s*kWh',
        text,
    )

    # Fallback: jeśli brak stref, całkowite = suma
    if result['calkowite'] == 0 and (result['szczyt'] > 0 or result['pozaszczyt'] > 0):
        result['calkowite'] = result['szczyt'] + result['pozaszczyt'] + result['noc']

    return result


def _extract_cena_energii(text: str) -> float:
    return _first_float(
        r'(?:cena\s+energii|cena\s+ee|stawka\s+za\s+energi[eę])[:\s]*([\d\s,\.]+)\s*(?:PLN|z[łl])?(?:/kWh|/1\s*kWh)',
        text,
    )


def _extract_oplata_dystrybucyjna(text: str) -> float:
    return _first_float(
        r'(?:op[łl]ata\s+dystrybucyjna|sk[łl]adnik\s+zmienny|stawka\s+dystrybucyjna)[:\s]*([\d\s,\.]+)\s*(?:PLN|z[łl])?(?:/kWh|/1\s*kWh)',
        text,
    )


def _extract_oplata_mocowa(text: str) -> float:
    return _first_float(
        r'(?:op[łl]ata\s+mocowa|stawka\s+mocowa)[:\s]*([\d\s,\.]+)\s*(?:PLN|z[łl])?/MWh',
        text,
    )


def _extract_tg_phi(text: str) -> float:
    return _first_float(
        r'(?:tg\s*[(\[]?\s*[φfi]\s*[)\]]?|tangent|wsp[óo][łl]czynnik\s+mocy\s+biernej)[:\s=]*([\d\s,\.]+)',
        text,
    )


def _extract_moc_bierna_kvarh(text: str) -> float:
    return _first_float(
        r'(?:moc\s+bierna|energia\s+bierna)[:\s]*([\d\s,\.]+)\s*kvarh',
        text,
    )


def _extract_oplata_moc_bierna(text: str) -> float:
    return _first_float(
        r'(?:op[łl]ata\s+(?:za\s+)?(?:moc[ą]?\s+)?biern[aą]|nadwy[żz]ka\s+mocy\s+biernej)[:\s]*([\d\s,\.]+)\s*(?:PLN|z[łl])',
        text,
    )


def _extract_kwota_netto(text: str) -> float:
    return _first_float(
        r'(?:razem\s+netto|kwota\s+netto|warto[śs][ćc]\s+netto|do\s+zap[łl]aty\s+netto)[:\s]*([\d\s,\.]+)\s*(?:PLN|z[łl])',
        text,
    )


def _extract_kwota_brutto(text: str) -> float:
    return _first_float(
        r'(?:razem\s+brutto|kwota\s+brutto|do\s+zap[łl]aty\s+brutto|do\s+zap[łl]aty)[:\s]*([\d\s,\.]+)\s*(?:PLN|z[łl])',
        text,
    )


def _extract_kategoria_mocowa(text: str) -> str:
    return _first_match(r'(?:kategoria\s+mocowa|kat\.\s*mocowa)[:\s]*(K[1-4])', text)


def _extract_typ_faktury(text: str) -> str:
    text_lower = text.lower()
    if 'koryguj' in text_lower or 'korekta' in text_lower:
        return 'korygująca'
    if 'prognoz' in text_lower:
        return 'prognozowa'
    return 'rozliczeniowa'


def _extract_oplata_sieciowa_stala(text: str) -> float:
    return _first_float(
        r'(?:sk[łl]adnik\s+sta[łl]y|op[łl]ata\s+sieciowa\s+sta[łl]a|stawka\s+sieciowa\s+sta[łl]a)[:\s]*([\d\s,\.]+)\s*(?:PLN|z[łl])',
        text,
    )


# ============================================================
# CONFIDENCE
# ============================================================

def _calculate_confidence(dane: DaneFaktury) -> float:
    """Oblicza pewność parsowania (0-1) na podstawie odczytanych pól krytycznych."""
    critical_fields = [
        dane.taryfa != '',
        dane.moc_umowna_kw > 0,
        dane.zuzycie_calkowite_kwh > 0,
        dane.cena_energii_pln_kwh > 0,
        dane.kwota_netto_pln > 0 or dane.kwota_brutto_pln > 0,
    ]
    important_fields = [
        dane.ppe != '',
        dane.osd != '',
        dane.oplata_dystrybucyjna_pln_kwh > 0,
        dane.oplata_mocowa_pln_mwh > 0,
        dane.nr_faktury != '',
        dane.okres_od != '',
    ]
    critical_score = sum(critical_fields) / len(critical_fields) * 0.7
    important_score = sum(important_fields) / len(important_fields) * 0.3
    return round(critical_score + important_score, 2)


# ============================================================
# MAIN PARSER
# ============================================================

def parsuj_fakture(pdf_bytes: bytes) -> DaneFaktury:
    """Parsuje fakturę PDF i zwraca odczytane dane."""
    text, tables = _extract_text(pdf_bytes)

    # Dodatkowo szukaj w tabelach
    table_text = ''
    for table in tables:
        for row in table:
            if row:
                table_text += ' '.join(str(cell) for cell in row if cell) + '\n'

    combined = text + '\n' + table_text

    okres_od, okres_do = _extract_okres(combined)
    zuzycie = _extract_zuzycie(combined)

    dane = DaneFaktury(
        ppe=_extract_ppe(combined),
        nr_faktury=_extract_nr_faktury(combined),
        okres_od=okres_od,
        okres_do=okres_do,
        taryfa=_extract_taryfa(combined),
        osd=_extract_osd(combined),
        moc_umowna_kw=_extract_moc_umowna(combined),
        zuzycie_szczyt_kwh=zuzycie['szczyt'],
        zuzycie_pozaszczyt_kwh=zuzycie['pozaszczyt'],
        zuzycie_noc_kwh=zuzycie['noc'],
        zuzycie_calkowite_kwh=zuzycie['calkowite'],
        cena_energii_pln_kwh=_extract_cena_energii(combined),
        oplata_dystrybucyjna_pln_kwh=_extract_oplata_dystrybucyjna(combined),
        oplata_mocowa_pln_mwh=_extract_oplata_mocowa(combined),
        oplata_sieciowa_stala_pln=_extract_oplata_sieciowa_stala(combined),
        tg_phi=_extract_tg_phi(combined),
        moc_bierna_kvarh=_extract_moc_bierna_kvarh(combined),
        oplata_moc_bierna_pln=_extract_oplata_moc_bierna(combined),
        kwota_netto_pln=_extract_kwota_netto(combined),
        kwota_brutto_pln=_extract_kwota_brutto(combined),
        kategoria_mocowa=_extract_kategoria_mocowa(combined),
        typ_faktury=_extract_typ_faktury(combined),
    )
    dane.pewnosc = _calculate_confidence(dane)
    return dane


# ============================================================
# ANALIZA OPTYMALIZACYJNA
# ============================================================

# Stawki opłaty mocowej 2026 (PLN/MWh)
STAWKI_MOCOWE_2026 = {
    'K1': 37.28,
    'K2': 87.76,
    'K3': 153.58,
    'K4': 219.40,
}


def _analiza_taryfy(dane: DaneFaktury) -> tuple[str, list[Rekomendacja]]:
    """Analiza optymalności obecnej grupy taryfowej."""
    rekomendacje = []
    info_parts = []

    taryfa = dane.taryfa.upper()
    total = dane.zuzycie_calkowite_kwh
    if total <= 0:
        return 'Brak danych o zużyciu — nie można przeanalizować taryfy.', []

    pct_pozaszczyt = dane.zuzycie_pozaszczyt_kwh / total if total > 0 else 0
    pct_szczyt = dane.zuzycie_szczyt_kwh / total if total > 0 else 0
    pct_noc = dane.zuzycie_noc_kwh / total if total > 0 else 0

    info_parts.append(
        f'Obecna taryfa: {taryfa}. '
        f'Rozkład zużycia: szczyt {pct_szczyt:.0%}, pozaszczyt {pct_pozaszczyt:.0%}, noc {pct_noc:.0%}.'
    )

    # C22a → C21: jeśli <50% zużycia pozaszczytowego
    if taryfa in ('C22A', 'C22B') and pct_pozaszczyt < 0.50:
        szacunek = total * 12 * dane.cena_energii_pln_kwh * 0.03  # ~3% oszczędności
        rekomendacje.append(Rekomendacja(
            tytul='Zmiana taryfy na C21 (jednostrefowa)',
            opis=(
                f'Przy {pct_pozaszczyt:.0%} zużycia pozaszczytowego, taryfa dwustrefowa nie jest optymalna. '
                f'C21 eliminuje podział na strefy i może obniżyć koszty.'
            ),
            oszczednosc_roczna_pln=round(szacunek, 0),
            priorytet='sredni',
        ))

    # C11 → C12a: jeśli >35% zużycia pozaszczytowego
    if taryfa == 'C11' and pct_pozaszczyt > 0.35:
        szacunek = total * 12 * dane.cena_energii_pln_kwh * 0.05
        rekomendacje.append(Rekomendacja(
            tytul='Zmiana taryfy na C12a (dwustrefowa)',
            opis=(
                f'Przy {pct_pozaszczyt:.0%} zużycia pozaszczytowego, taryfa dwustrefowa może dać oszczędności '
                f'dzięki niższej stawce pozaszczytowej.'
            ),
            oszczednosc_roczna_pln=round(szacunek, 0),
            priorytet='sredni',
        ))

    # C12a → C22a: jeśli >25% zużycia nocnego
    if taryfa == 'C12A' and pct_noc > 0.25:
        szacunek = total * 12 * dane.cena_energii_pln_kwh * 0.04
        rekomendacje.append(Rekomendacja(
            tytul='Zmiana taryfy na C22a (trzystrefowa)',
            opis=(
                f'Przy {pct_noc:.0%} zużycia nocnego, taryfa trzystrefowa z osobną strefą nocną '
                f'może obniżyć koszty.'
            ),
            oszczednosc_roczna_pln=round(szacunek, 0),
            priorytet='sredni',
        ))

    if not rekomendacje:
        info_parts.append('Obecna taryfa wydaje się optymalna dla profilu zużycia.')

    return ' '.join(info_parts), rekomendacje


def _analiza_moc_umowna(dane: DaneFaktury) -> tuple[str, list[Rekomendacja]]:
    """Analiza optymalności mocy umownej."""
    rekomendacje = []

    moc = dane.moc_umowna_kw
    zuzycie_mies = dane.zuzycie_calkowite_kwh

    if moc <= 0 or zuzycie_mies <= 0:
        return 'Brak danych o mocy umownej lub zużyciu.', []

    # Load factor = zużycie / (moc × 720h) — 720h ≈ średnia miesiąca
    load_factor = zuzycie_mies / (moc * 720)

    info = (
        f'Moc umowna: {moc:.0f} kW. '
        f'Zużycie miesięczne: {zuzycie_mies:,.0f} kWh. '
        f'Współczynnik obciążenia (load factor): {load_factor:.2f}.'
    )

    if load_factor < 0.30:
        # Sugerowana moc — target load factor ~0.5
        sugerowana_moc = zuzycie_mies / (720 * 0.50)
        sugerowana_moc = max(sugerowana_moc, 10)  # minimum
        redukcja_kw = moc - sugerowana_moc

        if redukcja_kw > 0:
            # Oszczędność: opłata sieciowa stała + składnik stały dystrybucji
            # Szacunek: ~12 PLN/kW/mies
            oszczednosc = redukcja_kw * 12 * 12
            rekomendacje.append(Rekomendacja(
                tytul='Redukcja mocy umownej',
                opis=(
                    f'Load factor {load_factor:.2f} wskazuje na przewymiarowaną moc umowną. '
                    f'Sugerowana redukcja z {moc:.0f} kW do ~{sugerowana_moc:.0f} kW '
                    f'(redukcja o {redukcja_kw:.0f} kW).'
                ),
                oszczednosc_roczna_pln=round(oszczednosc, 0),
                priorytet='wysoki',
            ))
    else:
        info += ' Moc umowna wydaje się adekwatna do zużycia.'

    return info, rekomendacje


def _analiza_moc_bierna(dane: DaneFaktury) -> tuple[str, list[Rekomendacja]]:
    """Analiza kompensacji mocy biernej."""
    rekomendacje = []

    tg = dane.tg_phi
    # Jeśli nie odczytano tg(φ), ale jest cos(φ) — przelicz
    if tg <= 0 and dane.moc_bierna_kvarh > 0 and dane.zuzycie_calkowite_kwh > 0:
        tg = dane.moc_bierna_kvarh / dane.zuzycie_calkowite_kwh

    if tg <= 0:
        return 'Brak danych o mocy biernej (tg φ).', []

    info = f'tg(φ) = {tg:.2f} (norma: ≤ 0.4).'

    if tg > 0.4:
        # Potrzebne kvar do kompensacji
        if dane.zuzycie_calkowite_kwh > 0:
            # P = zużycie / h_pracy (szacunek: 720h/mies)
            p_kw = dane.zuzycie_calkowite_kwh / 720
            q_obecne = p_kw * tg
            q_docelowe = p_kw * 0.4
            potrzebne_kvar = q_obecne - q_docelowe
        else:
            potrzebne_kvar = dane.moc_bierna_kvarh * (1 - 0.4 / tg)

        info += f' Przekroczenie normy! Potrzebna kompensacja: ~{potrzebne_kvar:.0f} kvar.'

        # Oszczędność = eliminacja opłaty za moc bierną
        if dane.oplata_moc_bierna_pln > 0:
            oszczednosc = dane.oplata_moc_bierna_pln * 12
        else:
            # Szacunek: ~5% rachunku
            oszczednosc = dane.kwota_netto_pln * 0.05 * 12 if dane.kwota_netto_pln > 0 else 0

        rekomendacje.append(Rekomendacja(
            tytul='Instalacja kompensacji mocy biernej (KMB)',
            opis=(
                f'tg(φ) = {tg:.2f} przekracza normę 0.4. '
                f'Instalacja baterii kondensatorów ~{potrzebne_kvar:.0f} kvar '
                f'wyeliminuje opłaty za moc bierną.'
            ),
            oszczednosc_roczna_pln=round(oszczednosc, 0),
            priorytet='wysoki',
        ))
    else:
        info += ' Wartość w normie — kompensacja nie jest potrzebna.'

    return info, rekomendacje


def _analiza_oplata_mocowa(dane: DaneFaktury) -> tuple[str, list[Rekomendacja]]:
    """Analiza kategorii opłaty mocowej."""
    rekomendacje = []

    kat = dane.kategoria_mocowa.upper() if dane.kategoria_mocowa else ''
    stawka = dane.oplata_mocowa_pln_mwh

    # Próba ustalenia kategorii z obecnej stawki
    if not kat and stawka > 0:
        # Znajdź najbliższą stawkę
        min_diff = float('inf')
        for k, s in STAWKI_MOCOWE_2026.items():
            diff = abs(stawka - s)
            if diff < min_diff:
                min_diff = diff
                kat = k

    if not kat:
        return 'Brak danych o kategorii mocowej.', []

    stawka_obecna = STAWKI_MOCOWE_2026.get(kat, stawka)
    info = f'Kategoria mocowa: {kat} (stawka: {stawka_obecna:.2f} PLN/MWh).'

    zuzycie_roczne_mwh = dane.zuzycie_calkowite_kwh * 12 / 1000

    if kat in ('K3', 'K4'):
        # Sugestia przejścia na niższą kategorię
        if kat == 'K4':
            kat_docelowa = 'K3'
        else:
            kat_docelowa = 'K2'

        stawka_docelowa = STAWKI_MOCOWE_2026[kat_docelowa]
        roznica_stawki = stawka_obecna - stawka_docelowa

        if zuzycie_roczne_mwh > 0:
            oszczednosc = roznica_stawki * zuzycie_roczne_mwh
        else:
            oszczednosc = 0

        rekomendacje.append(Rekomendacja(
            tytul=f'Obniżenie kategorii mocowej {kat} → {kat_docelowa}',
            opis=(
                f'Przejście z {kat} ({stawka_obecna:.2f} PLN/MWh) na {kat_docelowa} ({stawka_docelowa:.2f} PLN/MWh) '
                f'przez peak shaving / przesunięcie zużycia. '
                f'Różnica stawki: {roznica_stawki:.2f} PLN/MWh.'
            ),
            oszczednosc_roczna_pln=round(oszczednosc, 0),
            priorytet='wysoki' if oszczednosc > 5000 else 'sredni',
        ))
    else:
        info += ' Kategoria mocowa jest już korzystna.'

    return info, rekomendacje


def analizuj_fakture(dane: DaneFaktury) -> AnalizaOptymalizacji:
    """Przeprowadza pełną analizę optymalizacyjną na podstawie danych z faktury."""
    analiza = AnalizaOptymalizacji()

    info_taryfa, rek_taryfa = _analiza_taryfy(dane)
    info_moc, rek_moc = _analiza_moc_umowna(dane)
    info_bierna, rek_bierna = _analiza_moc_bierna(dane)
    info_mocowa, rek_mocowa = _analiza_oplata_mocowa(dane)

    analiza.analiza_taryfy = info_taryfa
    analiza.analiza_moc_umowna = info_moc
    analiza.analiza_moc_bierna = info_bierna
    analiza.analiza_oplata_mocowa = info_mocowa

    analiza.rekomendacje = rek_taryfa + rek_moc + rek_bierna + rek_mocowa
    analiza.laczna_oszczednosc_roczna_pln = sum(
        r.oszczednosc_roczna_pln for r in analiza.rekomendacje
    )

    return analiza


# ============================================================
# MAPPING NA FORMULARZ
# ============================================================

def mapuj_na_dane_klienta(dane: DaneFaktury) -> dict:
    """Mapuje dane z faktury na klucze session_state formularza."""
    mapping = {}

    if dane.taryfa:
        # Upewnij się, że format pasuje do selectboxa
        t = dane.taryfa
        if t.upper() in ('C11', 'C12A', 'C12B', 'C21', 'C22A', 'C22B', 'B21', 'B23'):
            mapping['grupa_taryfowa'] = t[0].upper() + t[1:]

    if dane.osd:
        mapping['osd'] = dane.osd

    if dane.moc_umowna_kw > 0:
        mapping['moc_umowna_kw'] = dane.moc_umowna_kw

    if dane.zuzycie_calkowite_kwh > 0:
        # Annualizacja: zużycie z 1 faktury × 12
        mapping['roczne_zuzycie_ee_kwh'] = dane.zuzycie_calkowite_kwh * 12

    if dane.cena_energii_pln_kwh > 0:
        mapping['cena_ee_pln_kwh'] = dane.cena_energii_pln_kwh

    if dane.oplata_dystrybucyjna_pln_kwh > 0:
        mapping['oplata_dystr_pln_kwh'] = dane.oplata_dystrybucyjna_pln_kwh

    if dane.oplata_mocowa_pln_mwh > 0:
        mapping['oplata_mocowa_pln_mwh'] = dane.oplata_mocowa_pln_mwh

    if dane.kategoria_mocowa:
        mapping['kategoria_mocowa'] = dane.kategoria_mocowa.upper()

    if dane.kwota_netto_pln > 0:
        mapping['sredni_rachunek_ee_mies_pln'] = dane.kwota_netto_pln

    # cos(φ) z tg(φ)
    if dane.tg_phi > 0:
        cos_phi = 1 / math.sqrt(1 + dane.tg_phi ** 2)
        mapping['wspolczynnik_cos_phi'] = round(cos_phi, 2)

    return mapping
