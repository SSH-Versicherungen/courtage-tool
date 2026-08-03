"""
Courtage-Extraktor - Web-Oberflaeche (Streamlit)
==================================================

Lokale Browser-Oberflaeche fuer courtage_extraktor.py: PDFs hochladen,
verarbeiten lassen, Ergebnis direkt im Browser ansehen und als Excel
herunterladen. Enthaelt keine eigene Extraktionslogik - ruft nur
process_files()/write_excel() aus courtage_extraktor.py auf.

Start (einmal pro Nutzung, oeffnet automatisch den Browser):
    streamlit run app.py
"""

import io
import os
import re
import shutil
import tempfile

import pandas as pd
import streamlit as st

import courtage_extraktor as ce

st.set_page_config(page_title="Courtage-Extraktor", page_icon="📄", layout="wide")


def require_password():
    """Passwort-Sperre - nur aktiv, wenn in den Streamlit-Secrets ein
    'app_password' hinterlegt ist (Streamlit Cloud: ueber das Dashboard
    unter 'Settings -> Secrets' gesetzt). Bei rein lokaler Nutzung
    (streamlit run app.py auf dem eigenen PC) ist kein Secret gesetzt -
    dann laeuft die App ohne Login, da das lokale Netz/der eigene Rechner
    bereits die Zugriffskontrolle ist."""
    try:
        required = st.secrets.get("app_password")
    except Exception:
        # Keine secrets.toml vorhanden (z.B. bei rein lokaler Nutzung) -
        # dann ist kein Passwort verlangt.
        required = None
    if not required:
        return True
    if st.session_state.get("password_ok"):
        return True

    st.title("📄 Courtage-Extraktor")
    pw = st.text_input("Passwort", type="password")
    if st.button("Anmelden"):
        if pw == required:
            st.session_state["password_ok"] = True
            st.rerun()
        else:
            st.error("Falsches Passwort.")
    return False


if not require_password():
    st.stop()

st.title("📄 Courtage-Extraktor")
st.caption(
    "Courtageabrechnungs-PDFs eines Monats hochladen -> Kunde + Provision je "
    "Buchung automatisch auslesen -> Ergebnis als Excel herunterladen."
)

with st.expander("So funktioniert's", expanded=False):
    st.markdown(
        "1. **Monat** eintragen (nur fuer die Erkennung von Datumsangaben "
        "in den PDFs wichtig, z.B. bei Abrechnungen, die einen Abschnitt "
        "fuer den Folgemonat enthalten).\n"
        "2. Alle PDFs des Monats hochladen (Mehrfachauswahl/Drag & Drop "
        "moeglich).\n"
        "3. Optional: den Kontoauszug (VR Bank Rhein-Neckar) desselben "
        "Monats hochladen - dann wird zusaetzlich geprueft, ob es "
        "Zahlungseingaenge ohne passende Abrechnungs-PDF gibt.\n"
        "4. Auf **Verarbeiten** klicken.\n"
        "5. Ergebnis unten pruefen und Excel-Datei herunterladen."
    )

MONTHS_DE = ce.MONTHS_DE

col1, col2 = st.columns([1, 1])
with col1:
    month_name = st.selectbox("Monat", MONTHS_DE, index=5)
with col2:
    year = st.number_input("Jahr", min_value=2020, max_value=2100, value=2026, step=1)

month_label = f"{month_name}-{year}"

uploaded_files = st.file_uploader(
    "Courtageabrechnungs-PDFs hierher ziehen oder auswaehlen",
    type=["pdf"],
    accept_multiple_files=True,
)

if uploaded_files:
    st.write(f"{len(uploaded_files)} Datei(en) ausgewaehlt.")

uploaded_bank_files = st.file_uploader(
    "Kontoauszug (optional) - prueft auf Zahlungseingaenge ohne passende Abrechnung",
    type=["pdf"],
    accept_multiple_files=True,
)

process_clicked = st.button("Verarbeiten", type="primary", disabled=not uploaded_files)

if process_clicked and uploaded_files:
    tmp_dir = tempfile.mkdtemp(prefix="courtage_")
    try:
        saved_paths = []
        for uf in uploaded_files:
            # Originaldateiname beibehalten - insurer_name_from_filename()
            # erkennt den Versicherer am Dateinamensmuster
            # "Abrechnung-<N>-<Versicherer>-<Monat>-<Jahr>.pdf".
            safe_name = re.sub(r"[\\/]", "_", uf.name)
            dest = os.path.join(tmp_dir, safe_name)
            with open(dest, "wb") as fh:
                fh.write(uf.getbuffer())
            saved_paths.append(dest)

        progress_bar = st.progress(0.0, text="Starte Verarbeitung ...")
        status_text = st.empty()

        def report_progress(i, total, filename):
            status_text.text(f"Verarbeite {i + 1}/{total}: {filename}")
            progress_bar.progress((i + 1) / total)

        df_rows, df_control, df_agg, df_problem = ce.process_files(
            saved_paths, month_label, report_progress
        )
        progress_bar.empty()
        status_text.empty()

        df_bank_unmatched = None
        if uploaded_bank_files:
            insurer_keys = pd.concat([
                df_control["Versicherer"], df_agg["Versicherer"], df_problem["Versicherer"],
            ]).unique().tolist() if not df_control.empty or not df_agg.empty or not df_problem.empty else []
            all_credits = []
            for bf in uploaded_bank_files:
                bank_dest = os.path.join(tmp_dir, re.sub(r"[\\/]", "_", bf.name))
                with open(bank_dest, "wb") as fh:
                    fh.write(bf.getbuffer())
                all_credits.extend(ce.extract_bank_credits(bank_dest))
            unmatched = ce.reconcile_bank_credits(all_credits, insurer_keys)
            df_bank_unmatched = pd.DataFrame(
                unmatched, columns=["datum", "betrag", "sender", "verwendungszweck"]
            )

        out_buffer = io.BytesIO()
        ce.write_excel(df_rows, df_control, df_agg, df_problem, out_buffer, df_bank_unmatched)
        out_buffer.seek(0)

        st.success(f"Fertig! {len(df_rows)} Buchungszeilen aus {len(saved_paths)} Dateien extrahiert.")

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Buchungszeilen", len(df_rows))
        m2.metric("Summe Provision", f"{df_rows['Provision'].sum():,.2f} €" if not df_rows.empty else "0,00 €")
        m3.metric("Sammelbelege ohne Details", len(df_agg))
        m4.metric("Manuelle Pruefung noetig", len(df_problem))

        st.download_button(
            "⬇️ Excel-Ergebnis herunterladen",
            data=out_buffer,
            file_name=f"Kunde_Provision_{month_label}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

        st.subheader("Kontrolle je Datei")
        st.dataframe(df_control, use_container_width=True)

        if not df_agg.empty:
            st.subheader("Sammelbelege ohne Kundendetail")
            st.dataframe(df_agg, use_container_width=True)

        if not df_problem.empty:
            st.subheader("Manuelle Pruefung noetig")
            st.dataframe(df_problem, use_container_width=True)

        if df_bank_unmatched is not None:
            st.subheader("Zahlungseingaenge ohne passende Abrechnung")
            if df_bank_unmatched.empty:
                st.success("Alle Zahlungseingaenge im Kontoauszug konnten einer verarbeiteten Abrechnung zugeordnet werden.")
            else:
                st.warning(
                    f"{len(df_bank_unmatched)} Zahlungseingang/-eingaenge "
                    f"({df_bank_unmatched['betrag'].sum():,.2f} €) ohne erkennbare "
                    "zugehoerige Abrechnungs-PDF - ggf. beim Versicherer nachfragen "
                    "oder pruefen, ob die PDF-Datei einfach nicht hochgeladen wurde. "
                    "Heuristischer Abgleich per Absendername/Verwendungszweck - "
                    "bitte gegenpruefen, bevor Sie einen Versicherer kontaktieren."
                )
                st.dataframe(df_bank_unmatched, use_container_width=True)

        st.subheader("Alle Buchungszeilen (Kunde + Provision)")
        st.dataframe(df_rows, use_container_width=True)
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)
