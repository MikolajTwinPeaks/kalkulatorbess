#!/usr/bin/env python3
"""
Kalkulator Ofertowy Energii – aplikacja Streamlit.

Uruchomienie:
    streamlit run app.py
"""

import os
import base64
import streamlit as st
import pandas as pd

from kalkulator_oferta import (
    DaneKlienta,
    oblicz_rekomendacje_ee,
    oblicz_rekomendacje_pv,
    oblicz_rekomendacje_bess,
    oblicz_rekomendacje_dsr,
    oblicz_rekomendacje_kmb,
    oblicz_opcje_finansowania,
    generuj_oferte_bytes,
)
from generuj_raport import create_report_bytes
from generuj_formularz_klienta import create_intake_form_bytes
from baza_cen import BazaCen

# ============================================================
# CONFIG
# ============================================================
st.set_page_config(
    page_title='Kalkulator Ofertowy Energii',
    page_icon='⚡',
    layout='centered',
)

# ============================================================
# RESPONSIVE CSS
# ============================================================
st.markdown("""
<style>
/* Ciemny sidebar */
[data-testid="stSidebar"] {
    background-color: #003366;
}
[data-testid="stSidebar"] * {
    color: white !important;
}
[data-testid="stSidebar"] .stRadio label span {
    color: white !important;
}

/* Mobile: lepsze metryki */
@media (max-width: 768px) {
    /* Metryki w jednej kolumnie */
    [data-testid="stMetric"] {
        background-color: #f8f9fa;
        border-radius: 8px;
        padding: 12px;
        margin-bottom: 8px;
        border-left: 4px solid #003366;
    }
    [data-testid="stMetric"] label {
        font-size: 0.85rem !important;
    }
    [data-testid="stMetric"] [data-testid="stMetricValue"] {
        font-size: 1.2rem !important;
    }
    /* Ekspandery - większy dotyk */
    .streamlit-expanderHeader {
        padding: 14px 8px !important;
        font-size: 1rem !important;
    }
    /* Przyciski - pełna szerokość na mobile */
    .stButton > button {
        width: 100% !important;
        padding: 12px !important;
        font-size: 1rem !important;
    }
    .stDownloadButton > button {
        width: 100% !important;
        padding: 12px !important;
    }
    /* Inputy - większy tekst */
    .stTextInput input, .stNumberInput input {
        font-size: 16px !important;  /* zapobiega zoom na iOS */
    }
    /* Tabela - scroll horyzontalny */
    .stDataFrame {
        overflow-x: auto !important;
    }
}
</style>
""", unsafe_allow_html=True)

# ============================================================
# NAVIGATION
# ============================================================
PAGES = [
    'Dane klienta',
    'Analiza & Rekomendacje',
    'Finansowanie',
    'Generuj ofertę',
    'Baza cen',
]

# Logo w sidebarze
_logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'logo.svg')
with open(_logo_path, 'r') as _f:
    _svg = _f.read()
_b64 = base64.b64encode(_svg.encode()).decode()
st.sidebar.markdown(
    f'<div style="text-align:center;padding:10px 0 20px 0;">'
    f'<img src="data:image/svg+xml;base64,{_b64}" width="180">'
    f'</div>',
    unsafe_allow_html=True,
)

st.sidebar.title('Kalkulator Ofertowy')
page = st.sidebar.radio('Nawigacja', PAGES)

# Demo button
if st.sidebar.button('Załaduj dane DEMO'):
    st.session_state['demo'] = True
    st.rerun()


def _load_demo():
    """Ładuje dane demo do session_state."""
    d = {
        'nazwa_firmy': 'Przykładowy Zakład Produkcyjny Sp. z o.o.',
        'nip': '1234567890',
        'branza': 'Produkcja metalowa',
        'dni_pracy': 'Pn-Pt',
        'godziny_pracy': '6:00-22:00 (2 zmiany)',
        'roczne_zuzycie_ee_kwh': 800_000.0,
        'moc_umowna_kw': 350.0,
        'moc_przylaczeniowa_kw': 400.0,
        'grupa_taryfowa': 'C22a',
        'osd': 'Tauron',
        'sredni_rachunek_ee_mies_pln': 55_000.0,
        'cena_ee_pln_kwh': 0.68,
        'oplata_dystr_pln_kwh': 0.27,
        'oplata_mocowa_pln_mwh': 219.40,
        'kategoria_mocowa': 'K3',
        'data_konca_umowy_ee': '2026-09-30',
        'typ_umowy_ee': 'FIX',
        'roczne_zuzycie_gaz_kwh': 200_000.0,
        'sredni_rachunek_gaz_mies_pln': 8_000.0,
        'cena_gaz_pln_kwh': 0.28,
        'data_konca_umowy_gaz': '2026-12-31',
        'ma_pv': True,
        'moc_pv_kwp': 200.0,
        'roczna_produkcja_pv_kwh': 210_000.0,
        'autokonsumpcja_pv_procent': 35.0,
        'ma_kmb': False,
        'moc_bierna_kvar': 0.0,
        'ma_agregat': False,
        'potrzebuje_go': True,
        'powierzchnia_dachu_m2': 800.0,
        'wspolczynnik_cos_phi': 0.85,
    }
    for k, v in d.items():
        st.session_state[k] = v


if st.session_state.pop('demo', False):
    _load_demo()


def _get_dane() -> DaneKlienta:
    """Buduje DaneKlienta z session_state."""
    s = st.session_state
    return DaneKlienta(
        nazwa_firmy=s.get('nazwa_firmy', ''),
        nip=s.get('nip', ''),
        branza=s.get('branza', ''),
        dni_pracy=s.get('dni_pracy', 'Pn-Pt'),
        godziny_pracy=s.get('godziny_pracy', ''),
        roczne_zuzycie_ee_kwh=s.get('roczne_zuzycie_ee_kwh', 0.0),
        moc_umowna_kw=s.get('moc_umowna_kw', 0.0),
        moc_przylaczeniowa_kw=s.get('moc_przylaczeniowa_kw', 0.0),
        grupa_taryfowa=s.get('grupa_taryfowa', 'C22a'),
        osd=s.get('osd', 'Tauron'),
        sredni_rachunek_ee_mies_pln=s.get('sredni_rachunek_ee_mies_pln', 0.0),
        cena_ee_pln_kwh=s.get('cena_ee_pln_kwh', 0.65),
        oplata_dystr_pln_kwh=s.get('oplata_dystr_pln_kwh', 0.25),
        oplata_mocowa_pln_mwh=s.get('oplata_mocowa_pln_mwh', 219.40),
        kategoria_mocowa=s.get('kategoria_mocowa', 'K3'),
        data_konca_umowy_ee=s.get('data_konca_umowy_ee', ''),
        typ_umowy_ee=s.get('typ_umowy_ee', 'FIX'),
        roczne_zuzycie_gaz_kwh=s.get('roczne_zuzycie_gaz_kwh', 0.0),
        sredni_rachunek_gaz_mies_pln=s.get('sredni_rachunek_gaz_mies_pln', 0.0),
        cena_gaz_pln_kwh=s.get('cena_gaz_pln_kwh', 0.25),
        data_konca_umowy_gaz=s.get('data_konca_umowy_gaz', ''),
        ma_pv=s.get('ma_pv', False),
        moc_pv_kwp=s.get('moc_pv_kwp', 0.0),
        roczna_produkcja_pv_kwh=s.get('roczna_produkcja_pv_kwh', 0.0),
        autokonsumpcja_pv_procent=s.get('autokonsumpcja_pv_procent', 0.0),
        ma_kmb=s.get('ma_kmb', False),
        moc_bierna_kvar=s.get('moc_bierna_kvar', 0.0),
        ma_agregat=s.get('ma_agregat', False),
        potrzebuje_go=s.get('potrzebuje_go', False),
        powierzchnia_dachu_m2=s.get('powierzchnia_dachu_m2', 0.0),
        wspolczynnik_cos_phi=s.get('wspolczynnik_cos_phi', 0.85),
    )


def _dane_ready() -> bool:
    """Sprawdza czy dane klienta zostały uzupełnione (minimum)."""
    s = st.session_state
    return bool(s.get('nazwa_firmy')) and s.get('roczne_zuzycie_ee_kwh', 0) > 0


# ============================================================
# PAGE 1: DANE KLIENTA
# ============================================================
def page_dane_klienta():
    st.header('Dane klienta')

    with st.expander('Firma', expanded=True):
        st.text_input('Nazwa firmy', key='nazwa_firmy')
        st.text_input('NIP', key='nip')
        st.text_input('Branża', key='branza')
        col1, col2 = st.columns(2)
        col1.selectbox('Dni pracy', ['Pn-Pt', 'Pn-Sob', '7 dni'], key='dni_pracy')
        col2.text_input('Godziny pracy', key='godziny_pracy',
                        placeholder='np. 6:00-22:00 (2 zmiany)')

    with st.expander('Energia elektryczna', expanded=True):
        st.number_input('Roczne zużycie ee (kWh/rok)', min_value=0.0,
                        step=10000.0, key='roczne_zuzycie_ee_kwh')
        col1, col2 = st.columns(2)
        col1.number_input('Moc umowna (kW)', min_value=0.0,
                          step=10.0, key='moc_umowna_kw')
        col2.number_input('Moc przyłączeniowa (kW)', min_value=0.0,
                          step=10.0, key='moc_przylaczeniowa_kw')

        col1, col2 = st.columns(2)
        col1.selectbox('Grupa taryfowa', ['C11', 'C12a', 'C12b', 'C21', 'C22a', 'C22b', 'B21', 'B23'],
                       key='grupa_taryfowa', index=4)
        col2.selectbox('OSD', ['Tauron', 'Enea', 'Energa', 'PGE', 'innogy'], key='osd')

        st.number_input('Średni rachunek ee (PLN/mies.)', min_value=0.0,
                        step=1000.0, key='sredni_rachunek_ee_mies_pln')

        col1, col2 = st.columns(2)
        col1.number_input('Cena ee z umowy (PLN/kWh)', min_value=0.0,
                          step=0.01, format='%.2f', key='cena_ee_pln_kwh')
        col2.number_input('Opłata dystrybucyjna (PLN/kWh)', min_value=0.0,
                          step=0.01, format='%.2f', key='oplata_dystr_pln_kwh')

        st.number_input('Stawka opłaty mocowej (PLN/MWh)', min_value=0.0,
                        step=1.0, key='oplata_mocowa_pln_mwh')

        col1, col2 = st.columns(2)
        col1.selectbox('Kategoria mocowa', ['K1', 'K2', 'K3', 'K4'],
                       key='kategoria_mocowa', index=2)
        col2.selectbox('Typ obecnej umowy', ['FIX', 'RDN', 'MIX'], key='typ_umowy_ee')

        st.text_input('Data końca umowy ee', key='data_konca_umowy_ee',
                      placeholder='RRRR-MM-DD')

    with st.expander('Gaz'):
        col1, col2 = st.columns(2)
        col1.number_input('Roczne zużycie gazu (kWh/rok)', min_value=0.0,
                          step=10000.0, key='roczne_zuzycie_gaz_kwh')
        col2.number_input('Średni rachunek gaz (PLN/mies.)', min_value=0.0,
                          step=100.0, key='sredni_rachunek_gaz_mies_pln')
        col1.number_input('Cena gazu (PLN/kWh)', min_value=0.0,
                          step=0.01, format='%.2f', key='cena_gaz_pln_kwh')
        col2.text_input('Data końca umowy gaz', key='data_konca_umowy_gaz',
                        placeholder='RRRR-MM-DD')

    with st.expander('Instalacja PV'):
        ma_pv = st.checkbox('Czy jest instalacja PV?', key='ma_pv')
        if ma_pv:
            st.number_input('Moc PV (kWp)', min_value=0.0,
                            step=10.0, key='moc_pv_kwp')
            st.number_input('Roczna produkcja PV (kWh)', min_value=0.0,
                            step=1000.0, key='roczna_produkcja_pv_kwh')
            st.number_input('Autokonsumpcja PV (%)', min_value=0.0,
                            max_value=100.0, step=5.0, key='autokonsumpcja_pv_procent')

    with st.expander('Infrastruktura'):
        st.checkbox('Kompensacja mocy biernej (KMB)?', key='ma_kmb')
        if st.session_state.get('ma_kmb'):
            st.number_input('Moc kompensacji (kvar)', min_value=0.0,
                            step=10.0, key='moc_bierna_kvar')
        col1, col2 = st.columns(2)
        col1.checkbox('Agregat prądotwórczy?', key='ma_agregat')
        col2.checkbox('Gwarancje pochodzenia?', key='potrzebuje_go')
        st.number_input('Wolna powierzchnia na PV (m²)', min_value=0.0,
                        step=50.0, key='powierzchnia_dachu_m2')
        st.number_input('Współczynnik cos(φ)', min_value=0.5,
                        max_value=1.0, step=0.01, format='%.2f',
                        key='wspolczynnik_cos_phi')

    st.divider()
    if _dane_ready():
        st.success('Dane uzupełnione. Przejdź do "Analiza & Rekomendacje" w menu bocznym.')
    else:
        st.warning('Uzupełnij co najmniej nazwę firmy i roczne zużycie ee.')


# ============================================================
# PAGE 2: ANALIZA & REKOMENDACJE
# ============================================================
def page_analiza():
    st.header('Analiza & Rekomendacje')

    if not _dane_ready():
        st.warning('Najpierw uzupełnij dane klienta.')
        return

    dane = _get_dane()

    # Obliczenia
    rek_ee = oblicz_rekomendacje_ee(dane)
    rek_pv = oblicz_rekomendacje_pv(dane)
    rek_bess = oblicz_rekomendacje_bess(dane)
    rek_dsr = oblicz_rekomendacje_dsr(dane)
    rek_kmb = oblicz_rekomendacje_kmb(dane)

    # Łączny CAPEX / oszczędność
    capex_items = [('BESS', rek_bess.capex_pln)]
    if rek_pv:
        capex_items.append(('PV', rek_pv.capex_pln))
    if rek_dsr and rek_dsr.potencjal_redukcji_kw > 0:
        capex_items.append(('DSR', rek_dsr.koszt_wdrozenia_pln))
    if rek_kmb:
        capex_items.append(('KMB', rek_kmb.capex_pln))
    capex_total = sum(c for _, c in capex_items)

    oszcz_items = [
        ('Kontrakt ee', rek_ee.oszczednosc_roczna_pln),
        ('BESS – autokonsumpcja', rek_bess.oszczednosc_autokonsumpcja_pln),
        ('BESS – arbitraż', rek_bess.oszczednosc_arbitraz_pln),
        ('BESS – peak shaving', rek_bess.oszczednosc_peak_shaving_pln),
    ]
    if rek_pv:
        oszcz_items.append(('PV', rek_pv.oszczednosc_roczna_pln))
    if rek_dsr and rek_dsr.przychod_roczny_pln > 0:
        oszcz_items.append(('DSR', rek_dsr.przychod_roczny_pln))
    if rek_kmb:
        oszcz_items.append(('KMB', rek_kmb.oszczednosc_roczna_pln))
    oszcz_total = sum(o for _, o in oszcz_items)

    okres_zw = capex_total / oszcz_total if oszcz_total > 0 else 99
    roczny_rach = dane.sredni_rachunek_ee_mies_pln * 12 + dane.sredni_rachunek_gaz_mies_pln * 12
    redukcja = oszcz_total / roczny_rach * 100 if roczny_rach > 0 else 0

    # Metryki – 2x2 grid (lepiej na mobile niż 4 w rzędzie)
    col1, col2 = st.columns(2)
    col1.metric('Łączny CAPEX', f'{capex_total:,.0f} PLN')
    col2.metric('Oszczędność roczna', f'{oszcz_total:,.0f} PLN')
    col1, col2 = st.columns(2)
    col1.metric('Okres zwrotu', f'{okres_zw:.1f} lat')
    col2.metric('Redukcja kosztów', f'{redukcja:.0f}%')

    # Wykres oszczędności
    st.subheader('Struktura oszczędności rocznych')
    df_oszcz = pd.DataFrame(oszcz_items, columns=['Źródło', 'PLN/rok'])
    df_oszcz = df_oszcz[df_oszcz['PLN/rok'] > 0]
    st.bar_chart(df_oszcz.set_index('Źródło'))

    # Szczegóły produktów
    with st.expander('Kontrakt na energię elektryczną', expanded=True):
        st.markdown(f'**Rekomendowany produkt:** {rek_ee.produkt}')
        col1, col2, col3 = st.columns(3)
        col1.metric('Cena FIX', f'{rek_ee.cena_fix_pln_kwh:.2f} PLN/kWh')
        col2.metric('Cena RDN', f'{rek_ee.cena_rdn_srednia_pln_kwh:.2f} PLN/kWh')
        col3.metric('Cena MIX', f'{rek_ee.cena_mix_pln_kwh:.2f} PLN/kWh')
        st.metric('Oszczędność roczna (zmiana taryfy)', f'{rek_ee.oszczednosc_roczna_pln:,.0f} PLN')
        st.info(rek_ee.uzasadnienie)

    with st.expander('Magazyn energii (BESS)', expanded=True):
        col1, col2 = st.columns(2)
        col1.metric('Pojemność', f'{rek_bess.pojemnosc_kwh:.0f} kWh')
        col2.metric('Moc', f'{rek_bess.moc_kw:.0f} kW')
        st.metric('CAPEX', f'{rek_bess.capex_pln:,.0f} PLN')

        st.markdown('**Oszczędności roczne (3 strumienie):**')
        st.metric('Autokonsumpcja PV', f'{rek_bess.oszczednosc_autokonsumpcja_pln:,.0f} PLN')
        st.metric('Arbitraż cenowy', f'{rek_bess.oszczednosc_arbitraz_pln:,.0f} PLN')
        st.metric('Peak shaving', f'{rek_bess.oszczednosc_peak_shaving_pln:,.0f} PLN')
        st.divider()
        col1, col2 = st.columns(2)
        col1.metric('Razem BESS', f'{rek_bess.oszczednosc_calkowita_pln:,.0f} PLN/rok')
        col2.metric('Okres zwrotu BESS', f'{rek_bess.okres_zwrotu_lat:.1f} lat')

    if rek_pv:
        with st.expander('Fotowoltaika (PV) – nowa instalacja'):
            col1, col2 = st.columns(2)
            col1.metric('Nowa moc', f'{rek_pv.nowa_moc_kwp:.0f} kWp')
            col2.metric('Roczna produkcja', f'{rek_pv.roczna_produkcja_kwh:,.0f} kWh')
            st.metric('CAPEX', f'{rek_pv.capex_pln:,.0f} PLN')
            col1, col2 = st.columns(2)
            col1.metric('Oszczędność roczna', f'{rek_pv.oszczednosc_roczna_pln:,.0f} PLN')
            col2.metric('Okres zwrotu', f'{rek_pv.okres_zwrotu_lat:.1f} lat')
            st.metric('Autokonsumpcja (bez BESS)', f'{rek_pv.autokonsumpcja_procent:.0f}%')

    with st.expander('DSR (Demand Side Response)'):
        if rek_dsr and rek_dsr.potencjal_redukcji_kw > 0:
            col1, col2 = st.columns(2)
            col1.metric('Potencjał redukcji', f'{rek_dsr.potencjal_redukcji_kw:.0f} kW')
            col2.metric('Przychód roczny', f'{rek_dsr.przychod_roczny_pln:,.0f} PLN')
            st.metric('Koszt wdrożenia', f'{rek_dsr.koszt_wdrozenia_pln:,.0f} PLN')
            st.info(rek_dsr.uzasadnienie)
        else:
            st.info(rek_dsr.uzasadnienie if rek_dsr else 'Brak potencjału DSR.')

    with st.expander('KMB (Kompensacja Mocy Biernej)'):
        if rek_kmb:
            col1, col2 = st.columns(2)
            col1.metric('Moc kompensacji', f'{rek_kmb.moc_potrzebna_kvar:.0f} kvar')
            col2.metric('CAPEX', f'{rek_kmb.capex_pln:,.0f} PLN')
            col1, col2 = st.columns(2)
            col1.metric('Oszczędność roczna', f'{rek_kmb.oszczednosc_roczna_pln:,.0f} PLN')
            col2.metric('Okres zwrotu', f'{rek_kmb.okres_zwrotu_lat:.1f} lat')
        elif dane.ma_kmb:
            st.info('Kompensacja mocy biernej już zainstalowana.')
        else:
            st.info('cos(φ) >= 0.95 – kompensacja nie jest potrzebna.')

    # Zapisz wyniki do session_state
    st.session_state['obliczenia_done'] = True
    st.session_state['capex_total'] = capex_total
    st.session_state['oszcz_total'] = oszcz_total


# ============================================================
# PAGE 3: FINANSOWANIE
# ============================================================
def page_finansowanie():
    st.header('Opcje finansowania')

    if not st.session_state.get('obliczenia_done'):
        st.warning('Najpierw przejdź do "Analiza & Rekomendacje" aby obliczyć CAPEX.')
        return

    capex_total = st.session_state.get('capex_total', 0)
    oszcz_total = st.session_state.get('oszcz_total', 0)

    st.metric('Łączny CAPEX do sfinansowania', f'{capex_total:,.0f} PLN netto')
    st.divider()

    opcje = oblicz_opcje_finansowania(capex_total)

    # Tabela porównawcza
    rows = []
    for o in opcje:
        rows.append({
            'Model': o.nazwa,
            'Wkład własny': f'{o.wklad_wlasny_procent:.0f}%',
            'Rata/mies.': f'{o.rata_miesieczna_pln:,.0f}' if o.rata_miesieczna_pln > 0 else '-',
            'Koszt całk.': f'{o.koszt_calkowity_pln:,.0f}' if o.koszt_calkowity_pln > 0 else '-',
        })
    df = pd.DataFrame(rows)
    st.dataframe(df, use_container_width=True, hide_index=True)

    # Wykres kosztów
    st.subheader('Porównanie kosztów całkowitych')
    chart_data = pd.DataFrame({
        'Model': [o.nazwa for o in opcje],
        'Koszt całkowity (PLN)': [o.koszt_calkowity_pln for o in opcje],
    })
    chart_data = chart_data[chart_data['Koszt całkowity (PLN)'] > 0]
    st.bar_chart(chart_data.set_index('Model'))

    # Szczegóły każdej opcji
    st.subheader('Szczegóły')
    for o in opcje:
        with st.expander(o.nazwa):
            st.markdown(f'**{o.opis}**')
            col1, col2 = st.columns(2)
            col1.metric('Wkład własny', f'{o.wklad_wlasny_procent:.0f}%')
            col2.metric('Rata miesięczna', f'{o.rata_miesieczna_pln:,.0f} PLN' if o.rata_miesieczna_pln > 0 else 'Brak')
            col1, col2 = st.columns(2)
            col1.metric('Koszt całkowity', f'{o.koszt_calkowity_pln:,.0f} PLN' if o.koszt_calkowity_pln > 0 else '-')
            col2.metric('Korzyść podatkowa/rok', f'{o.korzysc_podatkowa_roczna_pln:,.0f} PLN' if o.korzysc_podatkowa_roczna_pln > 0 else '-')
            st.caption(o.uwagi)


# ============================================================
# PAGE 4: GENERUJ OFERTĘ
# ============================================================
def page_generuj():
    st.header('Generuj ofertę')

    if not _dane_ready():
        st.warning('Najpierw uzupełnij dane klienta.')
        return

    dane = _get_dane()
    nazwa = dane.nazwa_firmy.replace(' ', '_')[:30]

    st.subheader('Pobierz dokumenty')

    # Oferta XLSX
    st.markdown('#### Oferta XLSX')
    st.caption('Pełna oferta z 7 arkuszami: podsumowanie, ee, BESS, PV/DSR/KMB, finansowanie, analiza 10-letnia, korzyści.')
    if st.button('Generuj ofertę XLSX', use_container_width=True):
        with st.spinner('Generowanie oferty...'):
            data = generuj_oferte_bytes(dane)
        st.download_button(
            label='Pobierz ofertę XLSX',
            data=data,
            file_name=f'Oferta_{nazwa}.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            use_container_width=True,
        )

    st.divider()

    # Raport DOCX
    st.markdown('#### Raport DOCX')
    st.caption('Raport analityczny: PV + BESS dla zakładów produkcyjnych. Regulacje, technologia, rynek.')
    if st.button('Generuj raport DOCX', use_container_width=True):
        with st.spinner('Generowanie raportu...'):
            data = create_report_bytes()
        st.download_button(
            label='Pobierz raport DOCX',
            data=data,
            file_name='Raport_PV_BESS_Zaklady_Produkcyjne.docx',
            mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            use_container_width=True,
        )

    st.divider()

    # Formularz klienta XLSX
    st.markdown('#### Formularz klienta XLSX')
    st.caption('Formularz zbierania danych od klienta z checklistą dokumentów i workflow.')
    if st.button('Generuj formularz', use_container_width=True):
        with st.spinner('Generowanie formularza...'):
            data = create_intake_form_bytes()
        st.download_button(
            label='Pobierz formularz XLSX',
            data=data,
            file_name='Formularz_Dane_Klienta.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            use_container_width=True,
        )


# ============================================================
# PAGE 5: BAZA CEN
# ============================================================
def page_baza_cen():
    st.header('Baza cen TGE RDB')

    db = BazaCen()

    tab1, tab2, tab3, tab4 = st.tabs([
        'Ostatnie ceny', 'Wykres historyczny', 'Statystyki', 'Import / Scraping'
    ])

    # --- Tab 1: Ostatnie ceny ---
    with tab1:
        st.subheader('Ostatnie ceny 15-minutowe')
        n_rec = st.selectbox('Liczba rekordów', [96, 192, 288, 672],
                             format_func=lambda x: f'{x} ({x // 4}h)')
        df = db.pobierz_ostatnie(n_rec)
        if df.empty:
            st.info('Brak danych w bazie. Użyj zakładki "Import / Scraping" aby dodać ceny.')
        else:
            col1, col2, col3 = st.columns(3)
            col1.metric('Rekordów', len(df))
            col2.metric('Śr. cena', f'{df["cena_pln_mwh"].mean():.2f} PLN/MWh')
            col3.metric('Śr. cena', f'{df["cena_pln_kwh"].mean():.4f} PLN/kWh')

            st.dataframe(
                df[['timestamp_start', 'timestamp_end', 'cena_pln_mwh', 'cena_pln_kwh', 'wolumen']].rename(
                    columns={
                        'timestamp_start': 'Od',
                        'timestamp_end': 'Do',
                        'cena_pln_mwh': 'PLN/MWh',
                        'cena_pln_kwh': 'PLN/kWh',
                        'wolumen': 'Wolumen',
                    }
                ),
                use_container_width=True,
                hide_index=True,
            )

    # --- Tab 2: Wykres historyczny ---
    with tab2:
        st.subheader('Wykres cen historycznych')
        col1, col2 = st.columns(2)
        from datetime import datetime, timedelta
        default_od = (datetime.now() - timedelta(days=7)).date()
        default_do = datetime.now().date()
        data_od = col1.date_input('Od', value=default_od)
        data_do = col2.date_input('Do', value=default_do)

        data_do_query = (data_do + timedelta(days=1)).strftime('%Y-%m-%d')
        df_hist = db.pobierz_ceny(
            data_od.strftime('%Y-%m-%d'), data_do_query
        )

        if df_hist.empty:
            st.info('Brak danych dla wybranego zakresu dat.')
        else:
            # Wykres liniowy cen
            st.markdown('#### Ceny 15-minutowe (PLN/MWh)')
            chart_df = df_hist.set_index('timestamp_start')[['cena_pln_mwh']]
            chart_df.columns = ['PLN/MWh']
            st.line_chart(chart_df)

            # Profil godzinowy
            st.markdown('#### Średni profil godzinowy')
            profil = db.profil_godzinowy(
                data_od.strftime('%Y-%m-%d'), data_do_query
            )
            if not profil.empty:
                profil_chart = profil.set_index('godzina')[['srednia_cena_pln_mwh']]
                profil_chart.columns = ['Śr. PLN/MWh']
                st.bar_chart(profil_chart)

    # --- Tab 3: Statystyki ---
    with tab3:
        st.subheader('Statystyki cenowe')
        total = db.liczba_rekordow()
        st.metric('Łączna liczba rekordów w bazie', total)

        if total > 0:
            st.markdown('---')
            col1, col2 = st.columns(2)

            with col1:
                st.markdown('**Ostatnie 30 dni**')
                avg30 = db.srednia_rdb(30)
                avg30k = db.srednia_rdb_kwh(30)
                spr30 = db.spread_sredni_kwh(30)
                if avg30 is not None:
                    st.metric('Średnia RDB', f'{avg30:.2f} PLN/MWh')
                    st.metric('Średnia RDB', f'{avg30k:.4f} PLN/kWh')
                if spr30 is not None:
                    st.metric('Średni spread dzienny', f'{spr30:.4f} PLN/kWh')

            with col2:
                st.markdown('**Ostatnie 7 dni**')
                avg7 = db.srednia_rdb(7)
                avg7k = db.srednia_rdb_kwh(7)
                spr7 = db.spread_sredni_kwh(7)
                if avg7 is not None:
                    st.metric('Średnia RDB', f'{avg7:.2f} PLN/MWh')
                    st.metric('Średnia RDB', f'{avg7k:.4f} PLN/kWh')
                if spr7 is not None:
                    st.metric('Średni spread dzienny', f'{spr7:.4f} PLN/kWh')

            # Logi scrapera
            st.markdown('---')
            st.markdown('**Ostatnie uruchomienia scrapera**')
            logi = db.pobierz_logi(10)
            if not logi.empty:
                st.dataframe(logi, use_container_width=True, hide_index=True)
            else:
                st.caption('Brak logów.')

    # --- Tab 4: Import / Scraping ---
    with tab4:
        st.subheader('Import danych')

        # Ręczny scraping
        st.markdown('#### Scraping z TGE')
        st.caption(
            'Uruchom scraper, aby pobrać najnowsze ceny z TGE RDB. '
            'Wymaga zainstalowanego Chrome i selenium.'
        )
        scrape_date = st.date_input('Data sesji', value=datetime.now().date(),
                                    key='scrape_date')
        if st.button('Uruchom scraper', use_container_width=True):
            with st.spinner('Pobieram ceny z TGE...'):
                try:
                    from scraper_tge import ScraperTGE
                    with ScraperTGE(headless=True) as scraper:
                        ceny = scraper.pobierz_ceny_rdb(scrape_date.strftime('%Y-%m-%d'))

                    if ceny:
                        rekordy = [{
                            'timestamp_start': c.timestamp_start,
                            'timestamp_end': c.timestamp_end,
                            'cena_pln_mwh': c.cena_pln_mwh,
                            'wolumen': c.wolumen_mwh,
                            'rynek': 'RDB',
                            'waluta': 'PLN',
                            'zrodlo': 'TGE_scraper',
                        } for c in ceny]
                        n = db.zapisz_ceny(rekordy)
                        db.zapisz_log('RDB', scrape_date.strftime('%Y-%m-%d'),
                                      n, 'OK', f'Scraping z UI: {n} rekordów')
                        st.success(f'Pobrano i zapisano {n} rekordów cenowych.')
                    else:
                        db.zapisz_log('RDB', scrape_date.strftime('%Y-%m-%d'),
                                      0, 'EMPTY', 'Scraping z UI: brak danych')
                        st.warning('Scraper nie znalazł danych cenowych na stronie TGE.')
                except ImportError:
                    st.error('Brak modułu selenium. Zainstaluj: pip install selenium webdriver-manager')
                except Exception as e:
                    db.zapisz_log('RDB', scrape_date.strftime('%Y-%m-%d'),
                                  0, 'ERROR', str(e))
                    st.error(f'Błąd scrapera: {e}')

        st.divider()

        # Upload pliku
        st.markdown('#### Import z pliku CSV / XLSX')
        st.caption(
            'Wymagane kolumny: timestamp_start, timestamp_end, cena_pln_mwh. '
            'Opcjonalne: wolumen, waluta, zrodlo.'
        )
        uploaded = st.file_uploader('Wybierz plik', type=['csv', 'xlsx', 'xls'])
        if uploaded is not None:
            if st.button('Importuj plik', use_container_width=True):
                with st.spinner('Importuję...'):
                    try:
                        if uploaded.name.endswith('.csv'):
                            import io
                            tmp_df = pd.read_csv(io.BytesIO(uploaded.read()))
                            uploaded.seek(0)
                            # Zapis tymczasowy
                            import tempfile
                            with tempfile.NamedTemporaryFile(suffix='.csv', delete=False) as f:
                                f.write(uploaded.read())
                                tmp_path = f.name
                            n = db.importuj_csv(tmp_path)
                            os.unlink(tmp_path)
                        else:
                            import tempfile
                            with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
                                f.write(uploaded.read())
                                tmp_path = f.name
                            n = db.importuj_xlsx(tmp_path)
                            os.unlink(tmp_path)

                        db.zapisz_log('RDB', '-', n, 'OK',
                                      f'Import z pliku {uploaded.name}: {n} rekordów')
                        st.success(f'Zaimportowano {n} rekordów z pliku {uploaded.name}.')
                    except Exception as e:
                        st.error(f'Błąd importu: {e}')


# ============================================================
# ROUTER
# ============================================================
if page == PAGES[0]:
    page_dane_klienta()
elif page == PAGES[1]:
    page_analiza()
elif page == PAGES[2]:
    page_finansowanie()
elif page == PAGES[3]:
    page_generuj()
elif page == PAGES[4]:
    page_baza_cen()
