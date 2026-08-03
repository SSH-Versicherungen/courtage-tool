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

ASSETS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")
LOGO_PATH = os.path.join(ASSETS_DIR, "ssh_logo.png")
FAVICON_PATH = os.path.join(ASSETS_DIR, "favicon.png")

st.set_page_config(
    page_title="Courtage-Extraktor",
    page_icon=FAVICON_PATH if os.path.exists(FAVICON_PATH) else "📄",
    layout="wide",
)

# Dezente Anpassungen, die ueber .streamlit/config.toml (Theme-Farben,
# Serifenschrift) hinausgehen: zentriertes Logo, etwas Luft darunter.
st.markdown(
    """
    <style>
    div[data-testid="stImage"] { display: flex; justify-content: center; }
    </style>
    """,
    unsafe_allow_html=True,
)


def show_header():
    if os.path.exists(LOGO_PATH):
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.image(LOGO_PATH, use_container_width=True)
    else:
        st.title("Courtage-Extraktor")


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

    show_header()
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

show_header()
st.markdown(
    "<h3 style='text-align:center; font-weight:400; color:#242A4E;'>Courtage-Extraktor</h3>",
    unsafe_allow_html=True,
)
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
        "moeglich) - die VEMA-Pool-CSV-Datei(en) (`VEMA-Poolabrechnung-"
        "N.csv`) koennen einfach mit hochgeladen werden, werden automatisch "
        "erkannt.\n"
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
    "Courtageabrechnungs-PDFs (+ ggf. VEMA-Pool-CSV) hierher ziehen oder auswaehlen",
    type=["pdf", "csv"],
    accept_multiple_files=True,
)

if uploaded_files:
    st.write(f"{len(uploaded_files)} Datei(en) ausgewaehlt.")

uploaded_bank_files = st.file_uploader(
    "Kontoauszug (optional) - prueft auf Zahlungseingaenge ohne passende Abrechnung",
    type=["pdf"],
    accept_multiple_files=True,
)

uploaded_betreuer_file = st.file_uploader(
    "Betreuer.xlsx (optional) - ordnet jeden Umsatz Robin Heckmann/Tim Selle/Andreas Selle "
    "zu und markiert die Kundenzeilen farbig (Rot/Gelb/Blau)",
    type=["xlsx"],
)

make_pdf_summary = st.checkbox(
    "PDF-Uebersicht zusaetzlich erstellen (Versicherer nach Umsatz sortiert, "
    "mit Kunden/Vertraegen je Versicherer)"
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

        if uploaded_betreuer_file is not None:
            betreuer_dest = os.path.join(tmp_dir, re.sub(r"[\\/]", "_", uploaded_betreuer_file.name))
            with open(betreuer_dest, "wb") as fh:
                fh.write(uploaded_betreuer_file.getbuffer())
            betreuer_lookup = ce.load_betreuer_lookup(betreuer_dest)
            df_rows = ce.apply_betreuer(df_rows, betreuer_lookup)

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

        if make_pdf_summary:
            pdf_buffer = io.BytesIO()
            ce.build_summary_pdf(df_rows, pdf_buffer, month_label)
            pdf_buffer.seek(0)
            st.download_button(
                "⬇️ PDF-Uebersicht herunterladen",
                data=pdf_buffer,
                file_name=f"Courtage-Uebersicht_{month_label}.pdf",
                mime="application/pdf",
            )

        st.subheader("Kontrolle je Datei")
        st.dataframe(df_control, use_container_width=True)

        if not df_agg.empty:
            st.subheader("Sammelbelege ohne Kundendetail")
            st.dataframe(df_agg, use_container_width=True)

        if not df_problem.empty:
            st.subheader("Manuelle Pruefung noetig")
            st.dataframe(df_problem, use_container_width=True)

        if "Betreuer" in df_rows.columns:
            df_unmatched_betreuer = df_rows[df_rows["Betreuer"].isna()][
                ["Versicherer", "Kunde", "Provision", "Datei"]
            ].drop_duplicates()
            if not df_unmatched_betreuer.empty:
                st.subheader("Kunden ohne Betreuer-Zuordnung")
                st.warning(
                    f"{len(df_unmatched_betreuer)} Kunde(n) konnten keinem Betreuer eindeutig "
                    "zugeordnet werden - bitte manuell pruefen (auch im Excel-Blatt "
                    "'Kunde_ohne_Betreuer')."
                )
                st.dataframe(df_unmatched_betreuer, use_container_width=True)

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
