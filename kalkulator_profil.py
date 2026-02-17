#!/usr/bin/env python3
"""
Kalkulacje BESS/PV oparte na rzeczywistym profilu mocy godzinowej.

Zastępuje flat scalars z kalkulator_oferta.py symulacjami godzinowymi 8760h
gdy dostępny jest profil z pliku SKADEN (XLS/XLSX).

Zwraca te same dataclassy (RekomendacjaBESS, RekomendacjaPV) co kalkulator_oferta.py.
"""

import numpy as np
import pandas as pd

from config import ConfigManager
from kalkulator_oferta import DaneKlienta, RekomendacjaBESS, RekomendacjaPV
from analiza_profilu import ProfilMocy


def _cfg() -> ConfigManager:
    return ConfigManager()


# ============================================================
# PROFIL PV — syntetyczny
# ============================================================

def generuj_profil_pv_godzinowy(moc_kwp: float, rok: int = 2024) -> pd.Series:
    """Syntetyczny profil generacji PV (8760h).

    Krzywa sinusoidalna z sezonowością. Kalibracja: roczna suma = moc_kwp × pv_produkcja_kwh_per_kwp.
    """
    cfg = _cfg()
    produkcja_per_kwp = cfg.get('pv_produkcja_kwh_per_kwp')

    # Generuj indeks godzinowy
    hours = pd.date_range(start=f'{rok}-01-01 01:00', periods=8760, freq='h')

    doy = hours.dayofyear.values.astype(float)
    hour = hours.hour.values.astype(float)

    # Sezonowość — dłuższe i mocniejsze słońce latem
    # Maksimum ok. dnia 172 (21 czerwca)
    sezon = 0.5 + 0.5 * np.sin(2 * np.pi * (doy - 80) / 365)

    # Krzywa dobowa — generacja 5:00-20:00, szczyt 12:00
    # Przesunięcie wschodu/zachodu wg sezonu
    sunrise = 6.0 - 2.0 * sezon  # lato: 4, zima: 6
    sunset = 18.0 + 2.0 * sezon   # lato: 20, zima: 18

    # Sinusoida między sunrise a sunset
    day_fraction = (hour - sunrise) / (sunset - sunrise + 1e-9)
    solar = np.sin(np.pi * day_fraction)
    solar = np.clip(solar, 0, None)

    # Zeruj poza godzinami słonecznymi
    solar[(hour < sunrise) | (hour > sunset)] = 0.0

    # Łączna generacja = sezon × solar
    profil_raw = sezon * solar

    # Kalibracja do docelowej produkcji rocznej
    roczna_docelowa = moc_kwp * produkcja_per_kwp
    suma_raw = profil_raw.sum()
    if suma_raw > 0:
        profil = profil_raw * (roczna_docelowa / suma_raw)
    else:
        profil = profil_raw

    return pd.Series(profil, index=hours, name='pv_kwh')


# ============================================================
# AUTOKONSUMPCJA PV — godzinowa
# ============================================================

def oblicz_autokonsumpcje_pv(
    profil_zuzycia: pd.DataFrame,
    profil_pv: pd.Series,
) -> dict:
    """Oblicza autokonsumpcję PV godzina po godzinie.

    Args:
        profil_zuzycia: DataFrame z kolumnami 'datetime', 'moc_kw' (zużycie w kW = kWh/h)
        profil_pv: Series z generacją PV (kWh/h), index = datetime

    Returns:
        dict z kluczami: autokonsumpcja_kwh, nadwyzka_kwh, autokonsumpcja_procent
    """
    # Wyrównaj indeksy — dopasuj po godzinie
    df = profil_zuzycia[['datetime', 'moc_kw']].copy()
    df = df.set_index('datetime')
    df = df.sort_index()

    # Resample profil PV do tych samych godzin
    pv = profil_pv.copy()
    pv.index = pv.index.round('h')

    # Merge
    merged = df.join(pv, how='left').fillna(0)

    autokonsumpcja = np.minimum(merged['moc_kw'].values, merged['pv_kwh'].values)
    nadwyzka = np.maximum(merged['pv_kwh'].values - merged['moc_kw'].values, 0)

    autokonsumpcja_kwh = float(autokonsumpcja.sum())
    nadwyzka_kwh = float(nadwyzka.sum())
    produkcja_total = float(merged['pv_kwh'].sum())

    return {
        'autokonsumpcja_kwh': autokonsumpcja_kwh,
        'nadwyzka_kwh': nadwyzka_kwh,
        'autokonsumpcja_procent': (autokonsumpcja_kwh / produkcja_total * 100) if produkcja_total > 0 else 0,
        'produkcja_total_kwh': produkcja_total,
    }


# ============================================================
# SYMULACJA BESS — dispatch 8760h
# ============================================================

def symuluj_bess(
    profil_zuzycia: pd.DataFrame,
    profil_pv: pd.Series | None,
    pojemnosc_kwh: float,
    moc_kw: float,
) -> dict:
    """Symulacja dispatchu BESS na 8760h.

    Strategia priorytetowa:
    1. Magazynuj nadwyżkę PV
    2. Arbitraż: ładuj w tanich godzinach (22-6, 10-15), rozładuj w drogich (17-21)
    3. Peak shaving: rozładuj gdy moc > percentyl 95

    Ograniczenia: SOC 10-90%, moc C-rate, RTE 90%.

    Returns:
        dict z kluczami: oszcz_autokonsumpcja_kwh, oszcz_arbitraz_kwh,
        oszcz_peak_shaving_kwh, cykle_roczne, profil_soc
    """
    cfg = _cfg()
    rte = cfg.get('bess_rte')
    soc_min = 0.10 * pojemnosc_kwh
    soc_max = 0.90 * pojemnosc_kwh

    # Przygotuj dane
    df = profil_zuzycia[['datetime', 'moc_kw']].copy()
    df = df.set_index('datetime').sort_index()

    if profil_pv is not None:
        pv = profil_pv.copy()
        pv.index = pv.index.round('h')
        df = df.join(pv, how='left').fillna(0)
    else:
        df['pv_kwh'] = 0.0

    zuzycie = df['moc_kw'].values
    generacja_pv = df['pv_kwh'].values
    hours = df.index
    n = len(zuzycie)

    # Peak shaving threshold — percentyl 95 zużycia
    p95 = np.percentile(zuzycie, 95)

    # SOC tracking
    soc = np.zeros(n + 1)
    soc[0] = pojemnosc_kwh * 0.5  # start at 50%

    # Wyniki per godzina
    autokons_kwh = np.zeros(n)
    arbitraz_ladowanie_kwh = np.zeros(n)
    arbitraz_rozladowanie_kwh = np.zeros(n)
    peak_shaving_kwh = np.zeros(n)

    for i in range(n):
        h = hours[i].hour if hasattr(hours[i], 'hour') else pd.Timestamp(hours[i]).hour
        current_soc = soc[i]

        # Dostępna pojemność do ładowania/rozładowania
        space_to_charge = soc_max - current_soc
        space_to_discharge = current_soc - soc_min

        energy_delta = 0.0  # dodatnie = ładowanie, ujemne = rozładowanie

        # --- 1. Autokonsumpcja PV: magazynuj nadwyżkę ---
        nadwyzka_pv = max(0, generacja_pv[i] - zuzycie[i])
        if nadwyzka_pv > 0 and space_to_charge > 0:
            # Ładuj nadwyżką PV (z uwzględnieniem RTE i mocy)
            laduj = min(nadwyzka_pv, space_to_charge / rte, moc_kw)
            autokons_kwh[i] = laduj * rte  # energia odzyskana później
            energy_delta += laduj * rte  # do SOC trafia mniej (RTE)
            space_to_charge -= laduj * rte
            space_to_discharge += laduj * rte

        # --- 2. Arbitraż ---
        # Tanie godziny: 22-6, 10-15 → ładuj
        # Drogie godziny: 17-21 → rozładuj
        if (22 <= h or h < 6 or 10 <= h < 15):
            # Ładuj
            if space_to_charge > 0:
                laduj = min(space_to_charge / rte, moc_kw - abs(energy_delta / rte))
                laduj = max(0, laduj)
                arbitraz_ladowanie_kwh[i] = laduj
                energy_delta += laduj * rte
        elif 17 <= h < 21:
            # Rozładuj
            discharge_available = min(space_to_discharge, moc_kw)
            if discharge_available > 0:
                arbitraz_rozladowanie_kwh[i] = discharge_available
                energy_delta -= discharge_available

        # --- 3. Peak shaving ---
        if zuzycie[i] > p95 and energy_delta >= 0:
            # Rozładuj aby zredukować szczyt
            nadmiar = zuzycie[i] - p95
            rozladuj = min(nadmiar, current_soc + energy_delta - soc_min, moc_kw)
            rozladuj = max(0, rozladuj)
            peak_shaving_kwh[i] = rozladuj
            energy_delta -= rozladuj

        # Aktualizuj SOC
        new_soc = current_soc + energy_delta
        soc[i + 1] = np.clip(new_soc, soc_min, soc_max)

    # Oblicz cykle roczne
    total_discharge = arbitraz_rozladowanie_kwh.sum() + peak_shaving_kwh.sum()
    # Autokonsumpcja jest ładowaniem — rozładowanie następuje w kolejnych godzinach
    # Przybliżenie: autokonsumpcja_kwh to energia która przeszła przez baterię
    total_throughput = autokons_kwh.sum() + total_discharge
    cykle = total_throughput / pojemnosc_kwh if pojemnosc_kwh > 0 else 0

    return {
        'oszcz_autokonsumpcja_kwh': float(autokons_kwh.sum()),
        'oszcz_arbitraz_kwh': float(arbitraz_rozladowanie_kwh.sum()),
        'oszcz_peak_shaving_kwh': float(peak_shaving_kwh.sum()),
        'arbitraz_ladowanie_kwh': float(arbitraz_ladowanie_kwh.sum()),
        'cykle_roczne': float(cykle),
        'profil_soc': soc[:-1],  # 8760 values
        'profil_soc_index': hours,
    }


# ============================================================
# REKOMENDACJA BESS — z profilem
# ============================================================

def oblicz_rekomendacje_bess_z_profilem(
    dane: DaneKlienta,
    profil: ProfilMocy,
) -> RekomendacjaBESS:
    """Rekomendacja BESS oparta na symulacji godzinowej.

    Iteruje pojemności i wybiera optymalną wg najkrótszego okresu zwrotu.
    """
    cfg = _cfg()

    # Profil PV (istniejące + ewentualnie nowe)
    moc_pv = dane.moc_pv_kwp if dane.ma_pv else 0
    profil_pv = generuj_profil_pv_godzinowy(moc_pv, profil.rok) if moc_pv > 0 else None

    # Parametry kosztowe
    koszt_kwh = cfg.get('bess_koszt_kwh')
    capex_ems = cfg.get('bess_ems_koszt')
    capex_inst_pct = cfg.get('bess_instalacja_procent')
    rte = cfg.get('bess_rte')

    # Ceny do przeliczenia kWh → PLN
    cena_autokons = dane.cena_ee_pln_kwh + dane.oplata_dystr_pln_kwh
    cena_net_bill = dane.cena_ee_pln_kwh * cfg.get('cena_net_billing_mnoznik')
    spread = cfg.get('bess_spread')

    # Peak shaving — oblicz oszczędność z kategorii mocowej
    zuzycie_szczytu = dane.roczne_zuzycie_ee_kwh * cfg.get('peak_zuzycie_szczytu')
    kat_mn = cfg.get('peak_kat_mn')
    mn_obecny = kat_mn.get(dane.kategoria_mocowa, 1.0)
    oplata_obecna = zuzycie_szczytu / 1000 * dane.oplata_mocowa_pln_mwh * mn_obecny
    nowa_kat = cfg.get('peak_nowa_kat')
    mn_nowy = kat_mn[nowa_kat[dane.kategoria_mocowa]]
    oplata_nowa = zuzycie_szczytu / 1000 * dane.oplata_mocowa_pln_mwh * mn_nowy
    oszcz_peak_pln_max = oplata_obecna - oplata_nowa

    # Iteruj pojemności
    bess_min = cfg.get('bess_min_pojemnosc')
    max_pojemnosc = max(500, dane.roczne_zuzycie_ee_kwh / 365 * 0.5)  # max ~ 50% dziennego zużycia
    pojemnosci = list(range(int(bess_min), int(max_pojemnosc) + 1, 50))
    if not pojemnosci:
        pojemnosci = [int(bess_min)]

    best = None
    best_okres = 999.0
    best_sim = None

    for poj in pojemnosci:
        moc = poj / 2  # C-rate 0.5

        sim = symuluj_bess(profil.dane, profil_pv, poj, moc)

        # Przelicz kWh → PLN
        oszcz_autokons = sim['oszcz_autokonsumpcja_kwh'] * (cena_autokons - cena_net_bill)
        oszcz_arbitraz = sim['oszcz_arbitraz_kwh'] * spread
        # Peak shaving — proporcjonalnie do redukcji szczytu
        if sim['oszcz_peak_shaving_kwh'] > 0:
            oszcz_peak = oszcz_peak_pln_max  # pełna oszczędność z kategorii mocowej
        else:
            oszcz_peak = 0.0

        oszcz_total = oszcz_autokons + oszcz_arbitraz + oszcz_peak

        capex_bess = poj * koszt_kwh
        capex_total = capex_bess + capex_ems + capex_bess * capex_inst_pct

        okres = capex_total / oszcz_total if oszcz_total > 0 else 999

        if okres < best_okres:
            best_okres = okres
            best_sim = sim
            best = {
                'pojemnosc': poj,
                'moc': moc,
                'capex': capex_total,
                'oszcz_autokons': oszcz_autokons,
                'oszcz_arbitraz': oszcz_arbitraz,
                'oszcz_peak': oszcz_peak,
                'oszcz_total': oszcz_total,
                'okres': okres,
            }

    if best is None:
        # Fallback — minimalna pojemność
        poj = int(bess_min)
        moc = poj / 2
        capex_bess = poj * koszt_kwh
        capex_total = capex_bess + capex_ems + capex_bess * capex_inst_pct
        return RekomendacjaBESS(
            pojemnosc_kwh=poj, moc_kw=moc, capex_pln=capex_total,
            oszczednosc_autokonsumpcja_pln=0, oszczednosc_arbitraz_pln=0,
            oszczednosc_peak_shaving_pln=0, oszczednosc_calkowita_pln=0,
            okres_zwrotu_lat=99,
        )

    return RekomendacjaBESS(
        pojemnosc_kwh=best['pojemnosc'],
        moc_kw=best['moc'],
        capex_pln=best['capex'],
        oszczednosc_autokonsumpcja_pln=best['oszcz_autokons'],
        oszczednosc_arbitraz_pln=best['oszcz_arbitraz'],
        oszczednosc_peak_shaving_pln=best['oszcz_peak'],
        oszczednosc_calkowita_pln=best['oszcz_total'],
        okres_zwrotu_lat=best['okres'],
    )


# ============================================================
# REKOMENDACJA PV — z profilem
# ============================================================

def oblicz_rekomendacje_pv_z_profilem(
    dane: DaneKlienta,
    profil: ProfilMocy,
) -> RekomendacjaPV | None:
    """Rekomendacja PV oparta na godzinowej autokonsumpcji.

    Iteruje moce PV (krok 10 kWp), liczy autokonsumpcję godzinowo,
    wybiera optymalną moc wg ROI.
    """
    cfg = _cfg()

    if dane.powierzchnia_dachu_m2 < 20:
        return None

    m2_per_kwp = cfg.get('pv_m2_per_kwp')
    max_moc_z_dachu = dane.powierzchnia_dachu_m2 / m2_per_kwp
    produkcja_per_kwp = cfg.get('pv_produkcja_kwh_per_kwp')
    istniejaca_moc = dane.moc_pv_kwp if dane.ma_pv else 0

    # Potrzebna dodatkowa moc
    istniejaca_produkcja = dane.roczna_produkcja_pv_kwh if dane.ma_pv else 0
    potrzebna_produkcja = dane.roczne_zuzycie_ee_kwh * cfg.get('pv_pokrycie_zuzycia') - istniejaca_produkcja
    if potrzebna_produkcja <= 0:
        return None

    # Ceny
    cena_pelna = dane.cena_ee_pln_kwh + dane.oplata_dystr_pln_kwh
    cena_net_billing = dane.cena_ee_pln_kwh * cfg.get('cena_net_billing_mnoznik')

    # Iteruj moce PV
    max_moc = min(potrzebna_produkcja / produkcja_per_kwp, max_moc_z_dachu)
    max_moc = max(10, max_moc)
    moce = list(range(10, int(max_moc) + 10, 10))
    if not moce:
        moce = [10]

    best = None
    best_okres = 999.0

    for moc_kwp in moce:
        if moc_kwp > max_moc_z_dachu:
            break

        # Generuj profil PV (istniejące + nowe)
        total_moc = istniejaca_moc + moc_kwp
        profil_pv = generuj_profil_pv_godzinowy(total_moc, profil.rok)

        # Oblicz autokonsumpcję godzinowo
        wynik = oblicz_autokonsumpcje_pv(profil.dane, profil_pv)

        # Oszczędności z nowej mocy PV
        roczna_produkcja_nowa = moc_kwp * produkcja_per_kwp
        autokons_pct = wynik['autokonsumpcja_procent'] / 100

        oszczednosc = roczna_produkcja_nowa * (
            autokons_pct * cena_pelna + (1 - autokons_pct) * cena_net_billing
        )

        # CAPEX
        capex_pln_kwp = (
            cfg.get('pv_capex_maly') if moc_kwp < 50
            else cfg.get('pv_capex_sredni') if moc_kwp < 200
            else cfg.get('pv_capex_duzy')
        )
        capex = moc_kwp * capex_pln_kwp

        okres = capex / oszczednosc if oszczednosc > 0 else 999

        if okres < best_okres:
            best_okres = okres
            best = {
                'moc_kwp': moc_kwp,
                'produkcja': roczna_produkcja_nowa,
                'capex': capex,
                'oszczednosc': oszczednosc,
                'okres': okres,
                'autokons_pct': autokons_pct * 100,
            }

    if best is None:
        return None

    return RekomendacjaPV(
        nowa_moc_kwp=best['moc_kwp'],
        roczna_produkcja_kwh=best['produkcja'],
        capex_pln=best['capex'],
        oszczednosc_roczna_pln=best['oszczednosc'],
        okres_zwrotu_lat=best['okres'],
        autokonsumpcja_procent=best['autokons_pct'],
    )


# ============================================================
# HELPER: dane do wykresów SOC
# ============================================================

def get_bess_soc_data(
    dane: DaneKlienta,
    profil: ProfilMocy,
    pojemnosc_kwh: float,
    moc_kw: float,
) -> dict | None:
    """Uruchamia symulację BESS i zwraca dane SOC do wykresów."""
    moc_pv = dane.moc_pv_kwp if dane.ma_pv else 0
    profil_pv = generuj_profil_pv_godzinowy(moc_pv, profil.rok) if moc_pv > 0 else None

    sim = symuluj_bess(profil.dane, profil_pv, pojemnosc_kwh, moc_kw)
    return sim
