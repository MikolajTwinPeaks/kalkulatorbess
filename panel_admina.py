#!/usr/bin/env python3
"""
Panel administracyjny — strona Streamlit.

Dwa taby:
1. Zarządzanie użytkownikami (CRUD)
2. Konfiguracja parametrów kalkulatora
"""

import json
import streamlit as st
from auth import AuthManager
from config import ConfigManager


def page_panel_admina():
    st.markdown('<h1>Panel administracyjny</h1>', unsafe_allow_html=True)

    tab_users, tab_config = st.tabs(['Zarządzanie użytkownikami', 'Konfiguracja parametrów'])

    # ── Tab 1: Użytkownicy ──
    with tab_users:
        auth = AuthManager()
        users = auth.list_users()

        st.subheader('Użytkownicy')
        if users:
            for u in users:
                status = 'aktywny' if u['aktywny'] else 'nieaktywny'
                with st.expander(f"{u['username']} — {u['rola']} ({status})"):
                    col1, col2 = st.columns(2)
                    with col1:
                        nowa_rola = st.selectbox(
                            'Rola', ['admin', 'handlowiec', 'guest'],
                            index=['admin', 'handlowiec', 'guest'].index(u['rola']),
                            key=f"rola_{u['id']}",
                        )
                    with col2:
                        nowy_status = st.selectbox(
                            'Status', ['aktywny', 'nieaktywny'],
                            index=0 if u['aktywny'] else 1,
                            key=f"status_{u['id']}",
                        )
                    if st.button('Zapisz zmiany', key=f"save_{u['id']}"):
                        auth.update_user(u['id'], rola=nowa_rola, aktywny=1 if nowy_status == 'aktywny' else 0)
                        st.success(f"Zaktualizowano użytkownika {u['username']}.")
                        st.rerun()

                    st.divider()
                    st.markdown('**Zmiana hasła**')
                    nowe_haslo = st.text_input('Nowe hasło', type='password', key=f"pass_{u['id']}")
                    if st.button('Zmień hasło', key=f"chpass_{u['id']}"):
                        if nowe_haslo:
                            auth.change_password(u['id'], nowe_haslo)
                            st.success('Hasło zmienione.')
                        else:
                            st.warning('Podaj nowe hasło.')

                    st.divider()
                    if st.button('Dezaktywuj użytkownika', key=f"del_{u['id']}"):
                        auth.delete_user(u['id'])
                        st.success(f"Użytkownik {u['username']} dezaktywowany.")
                        st.rerun()

                    st.caption(f"Utworzony: {u['utworzony'] or '-'} | Ostatnie logowanie: {u['ostatnie_logowanie'] or '-'}")

        st.divider()
        st.subheader('Dodaj nowego użytkownika')
        with st.form('add_user_form'):
            new_username = st.text_input('Login')
            new_password = st.text_input('Hasło', type='password')
            new_rola = st.selectbox('Rola', ['handlowiec', 'admin', 'guest'])
            submitted = st.form_submit_button('Dodaj użytkownika')
            if submitted:
                if new_username and new_password:
                    ok = auth.create_user(new_username, new_password, new_rola)
                    if ok:
                        st.success(f'Dodano użytkownika: {new_username}')
                        st.rerun()
                    else:
                        st.error(f'Użytkownik "{new_username}" już istnieje.')
                else:
                    st.warning('Podaj login i hasło.')

    # ── Tab 2: Konfiguracja parametrów ──
    with tab_config:
        cfg = ConfigManager()
        categories = cfg.categories()

        for kategoria in categories:
            with st.expander(kategoria, expanded=False):
                params = cfg.get_by_category(kategoria)
                form_key = f"cfg_form_{kategoria}"
                with st.form(form_key):
                    new_values = {}
                    for p in params:
                        klucz = p['klucz']
                        typ = p['typ']
                        label = f"{p['opis']} (`{klucz}`)"

                        if typ == 'float':
                            new_values[klucz] = st.number_input(
                                label, value=float(p['wartosc']),
                                format='%.4f', step=0.01,
                                key=f"cfg_{klucz}",
                            )
                        elif typ == 'int':
                            new_values[klucz] = st.number_input(
                                label, value=int(p['wartosc']),
                                step=1,
                                key=f"cfg_{klucz}",
                            )
                        elif typ == 'json':
                            raw = p['wartosc_raw']
                            try:
                                formatted = json.dumps(json.loads(raw), indent=2, ensure_ascii=False)
                            except (json.JSONDecodeError, TypeError):
                                formatted = raw
                            new_values[klucz] = st.text_area(
                                label, value=formatted, height=100,
                                key=f"cfg_{klucz}",
                            )
                        else:
                            new_values[klucz] = st.text_input(
                                label, value=str(p['wartosc']),
                                key=f"cfg_{klucz}",
                            )

                    if st.form_submit_button(f'Zapisz — {kategoria}'):
                        # Validate JSON fields
                        ok = True
                        for p in params:
                            if p['typ'] == 'json':
                                try:
                                    json.loads(new_values[p['klucz']])
                                except json.JSONDecodeError:
                                    st.error(f"Nieprawidłowy JSON w polu {p['klucz']}")
                                    ok = False
                        if ok:
                            cfg.set_many(new_values)
                            st.success(f'Zapisano parametry: {kategoria}')
