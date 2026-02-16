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

# ============================================================
# CONFIG
# ============================================================
st.set_page_config(
    page_title='Kalkulator Ofertowy Energii',
    page_icon='⚡',
    layout='wide',
)

# ============================================================
# NAVIGATION
# ============================================================
PAGES = [
    'Dane klienta',
    'Analiza & Rekomendacje',
    'Finansowanie',
    'Generuj ofertę',
]

# Ciemny sidebar dla białego logo
st.markdown("""
<style>
[data-testid="stSidebar"] {
    background-color: #003366;
}
[data-testid="stSidebar"] * {
    color: white !important;
}
[data-testid="stSidebar"] .stRadio label span {
    color: white !important;
}
</style>
""", unsafe_allow_html=True)

# Logo w sidebarze
_logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'logo.svg')
with open(_logo_path, 'r') as _f:
    _svg = _f.read()
_b64 = base64.b64encode(_svg.encode()).decode()
st.sidebar.markdown(
    f'<div style="text-align:center;padding:10px 0 20px 0;">'
    f'<img src="data:image/svg+xml;base64,{_b64}" width="200">'
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
        col1, col2 = st.columns(2)
        col1.text_input('Nazwa firmy', key='nazwa_firmy')
        col2.text_input('NIP', key='nip')
        col1.text_input('Branża', key='branza')
        col2.selectbox('Dni pracy', ['Pn-Pt', 'Pn-Sob', '7 dni'], key='dni_pracy')
        col1.text_input('Godziny pracy', key='godziny_pracy',
                        placeholder='np. 6:00-22:00 (2 zmiany)')

    with st.expander('Energia elektryczna', expanded=True):
        col1, col2, col3 = st.columns(3)
        col1.number_input('Roczne zużycie ee (kWh/rok)', min_value=0.0,
                          step=10000.0, key='roczne_zuzycie_ee_kwh')
        col2.number_input('Moc umowna (kW)', min_value=0.0,
                          step=10.0, key='moc_umowna_kw')
        col3.number_input('Moc przyłączeniowa (kW)', min_value=0.0,
                          step=10.0, key='moc_przylaczeniowa_kw')

        col1, col2, col3 = st.columns(3)
        col1.selectbox('Grupa taryfowa', ['C11', 'C12a', 'C12b', 'C21', 'C22a', 'C22b', 'B21', 'B23'],
                       key='grupa_taryfowa', index=4)
        col2.selectbox('OSD', ['Tauron', 'Enea', 'Energa', 'PGE', 'innogy'], key='osd')
        col3.number_input('Średni rachunek ee (PLN/mies.)', min_value=0.0,
                          step=1000.0, key='sredni_rachunek_ee_mies_pln')

        col1, col2, col3 = st.columns(3)
        col1.number_input('Cena ee z umowy (PLN/kWh)', min_value=0.0,
                          step=0.01, format='%.2f', key='cena_ee_pln_kwh')
        col2.number_input('Opłata dystrybucyjna (PLN/kWh)', min_value=0.0,
                          step=0.01, format='%.2f', key='oplata_dystr_pln_kwh')
        col3.number_input('Stawka opłaty mocowej (PLN/MWh)', min_value=0.0,
                          step=1.0, key='oplata_mocowa_pln_mwh')

        col1, col2, col3 = st.columns(3)
        col1.selectbox('Kategoria mocowa', ['K1', 'K2', 'K3', 'K4'],
                       key='kategoria_mocowa', index=2)
        col2.text_input('Data końca umowy ee', key='data_konca_umowy_ee',
                        placeholder='RRRR-MM-DD')
        col3.selectbox('Typ obecnej umowy', ['FIX', 'RDN', 'MIX'], key='typ_umowy_ee')

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
            col1, col2, col3 = st.columns(3)
            col1.number_input('Moc PV (kWp)', min_value=0.0,
                              step=10.0, key='moc_pv_kwp')
            col2.number_input('Roczna produkcja PV (kWh)', min_value=0.0,
                              step=1000.0, key='roczna_produkcja_pv_kwh')
            col3.number_input('Autokonsumpcja PV (%)', min_value=0.0,
                              max_value=100.0, step=5.0, key='autokonsumpcja_pv_procent')

    with st.expander('Infrastruktura'):
        col1, col2 = st.columns(2)
        col1.checkbox('Kompensacja mocy biernej (KMB)?', key='ma_kmb')
        if st.session_state.get('ma_kmb'):
            col2.number_input('Moc kompensacji (kvar)', min_value=0.0,
                              step=10.0, key='moc_bierna_kvar')
        col1.checkbox('Agregat prądotwórczy?', key='ma_agregat')
        col2.checkbox('Gwarancje pochodzenia?', key='potrzebuje_go')
        col1.number_input('Wolna powierzchnia na PV (m²)', min_value=0.0,
                          step=50.0, key='powierzchnia_dachu_m2')
        col2.number_input('Współczynnik cos(φ)', min_value=0.5,
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

    # Metryki
    col1, col2, col3, col4 = st.columns(4)
    col1.metric('Łączny CAPEX', f'{capex_total:,.0f} PLN')
    col2.metric('Oszczędność roczna', f'{oszcz_total:,.0f} PLN')
    col3.metric('Okres zwrotu', f'{okres_zw:.1f} lat')
    col4.metric('Redukcja kosztów', f'{redukcja:.0f}%')

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
        col2.metric('Cena RDN (średnia)', f'{rek_ee.cena_rdn_srednia_pln_kwh:.2f} PLN/kWh')
        col3.metric('Cena MIX', f'{rek_ee.cena_mix_pln_kwh:.2f} PLN/kWh')
        st.metric('Oszczędność roczna (zmiana taryfy)', f'{rek_ee.oszczednosc_roczna_pln:,.0f} PLN')
        st.info(rek_ee.uzasadnienie)

    with st.expander('Magazyn energii (BESS)', expanded=True):
        col1, col2, col3 = st.columns(3)
        col1.metric('Pojemność', f'{rek_bess.pojemnosc_kwh:.0f} kWh')
        col2.metric('Moc', f'{rek_bess.moc_kw:.0f} kW')
        col3.metric('CAPEX', f'{rek_bess.capex_pln:,.0f} PLN')

        st.markdown('**Oszczędności roczne (3 strumienie):**')
        col1, col2, col3 = st.columns(3)
        col1.metric('Autokonsumpcja PV', f'{rek_bess.oszczednosc_autokonsumpcja_pln:,.0f} PLN')
        col2.metric('Arbitraż cenowy', f'{rek_bess.oszczednosc_arbitraz_pln:,.0f} PLN')
        col3.metric('Peak shaving', f'{rek_bess.oszczednosc_peak_shaving_pln:,.0f} PLN')
        st.metric('Razem BESS', f'{rek_bess.oszczednosc_calkowita_pln:,.0f} PLN/rok')
        st.metric('Okres zwrotu BESS', f'{rek_bess.okres_zwrotu_lat:.1f} lat')

    if rek_pv:
        with st.expander('Fotowoltaika (PV) – nowa instalacja'):
            col1, col2, col3 = st.columns(3)
            col1.metric('Nowa moc', f'{rek_pv.nowa_moc_kwp:.0f} kWp')
            col2.metric('Roczna produkcja', f'{rek_pv.roczna_produkcja_kwh:,.0f} kWh')
            col3.metric('CAPEX', f'{rek_pv.capex_pln:,.0f} PLN')
            col1.metric('Oszczędność roczna', f'{rek_pv.oszczednosc_roczna_pln:,.0f} PLN')
            col2.metric('Okres zwrotu', f'{rek_pv.okres_zwrotu_lat:.1f} lat')
            col3.metric('Autokonsumpcja (bez BESS)', f'{rek_pv.autokonsumpcja_procent:.0f}%')

    with st.expander('DSR (Demand Side Response)'):
        if rek_dsr and rek_dsr.potencjal_redukcji_kw > 0:
            col1, col2, col3 = st.columns(3)
            col1.metric('Potencjał redukcji', f'{rek_dsr.potencjal_redukcji_kw:.0f} kW')
            col2.metric('Przychód roczny', f'{rek_dsr.przychod_roczny_pln:,.0f} PLN')
            col3.metric('Koszt wdrożenia', f'{rek_dsr.koszt_wdrozenia_pln:,.0f} PLN')
            st.info(rek_dsr.uzasadnienie)
        else:
            st.info(rek_dsr.uzasadnienie if rek_dsr else 'Brak potencjału DSR.')

    with st.expander('KMB (Kompensacja Mocy Biernej)'):
        if rek_kmb:
            col1, col2, col3 = st.columns(3)
            col1.metric('Moc kompensacji', f'{rek_kmb.moc_potrzebna_kvar:.0f} kvar')
            col2.metric('CAPEX', f'{rek_kmb.capex_pln:,.0f} PLN')
            col3.metric('Oszczędność roczna', f'{rek_kmb.oszczednosc_roczna_pln:,.0f} PLN')
            st.metric('Okres zwrotu', f'{rek_kmb.okres_zwrotu_lat:.1f} lat')
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
            'Rata/mies. (PLN)': f'{o.rata_miesieczna_pln:,.0f}' if o.rata_miesieczna_pln > 0 else '-',
            'Okres (mies.)': o.okres_mies if o.okres_mies > 0 else '-',
            'Koszt całkowity (PLN)': f'{o.koszt_calkowity_pln:,.0f}' if o.koszt_calkowity_pln > 0 else '-',
            'Korzyść podat./rok (PLN)': f'{o.korzysc_podatkowa_roczna_pln:,.0f}' if o.korzysc_podatkowa_roczna_pln > 0 else '-',
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

    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown('#### Oferta XLSX')
        st.caption('Pełna oferta z 7 arkuszami: podsumowanie, ee, BESS, PV/DSR/KMB, finansowanie, analiza 10-letnia, korzyści.')
        if st.button('Generuj ofertę XLSX'):
            with st.spinner('Generowanie oferty...'):
                data = generuj_oferte_bytes(dane)
            st.download_button(
                label='Pobierz ofertę XLSX',
                data=data,
                file_name=f'Oferta_{nazwa}.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            )

    with col2:
        st.markdown('#### Raport DOCX')
        st.caption('Raport analityczny: PV + BESS dla zakładów produkcyjnych. Regulacje, technologia, rynek.')
        if st.button('Generuj raport DOCX'):
            with st.spinner('Generowanie raportu...'):
                data = create_report_bytes()
            st.download_button(
                label='Pobierz raport DOCX',
                data=data,
                file_name='Raport_PV_BESS_Zaklady_Produkcyjne.docx',
                mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            )

    with col3:
        st.markdown('#### Formularz klienta XLSX')
        st.caption('Formularz zbierania danych od klienta z checklistą dokumentów i workflow.')
        if st.button('Generuj formularz'):
            with st.spinner('Generowanie formularza...'):
                data = create_intake_form_bytes()
            st.download_button(
                label='Pobierz formularz XLSX',
                data=data,
                file_name='Formularz_Dane_Klienta.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            )


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
