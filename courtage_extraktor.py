"""
Courtage-Extraktor fuer SSH Versicherungsmakler GmbH & Co. KG
================================================================

Liest alle Courtageabrechnungs-PDFs eines Monatsordners (z.B. "Juni-2026")
und extrahiert je Buchungszeile: Versicherer, Kunde, Provision.

Aufruf:
    python courtage_extraktor.py Juni-2026

Ergebnis:
    Courtage-Tool/output/Juni-2026/Kunde_Provision_Juni-2026.xlsx

Die Zuordnung Kunde -> Betreuer (Courtage-Verteilung) ist bewusst NICHT
Teil dieses Skripts (das ist der naechste Ausbauschritt). Dieses Skript
liefert nur die Rohdaten je Buchung: Kunde + Provision + Herkunft.

Aufbau, kurz:
- Manche Versicherer liefern eine echte Positionstabelle (Kunde + Betrag
  pro Zeile) -> generische, koordinatenbasierte Tabellen-Engine.
- Manche Versicherer (aktuell: Fondsfinanz) liefern ein Block-Format
  (ein Kundenblock ueber mehrere Zeilen) -> eigener Parser.
- Manche PDFs enthalten keinen Text (Scan) -> OCR-Fallback (Tesseract).
- Manche PDFs sind reine Sammelabrechnungen ohne Kundenbezug (z.B. ARAG,
  VEMA-Pool-Uebersicht) -> werden als "Sammelbeleg" gekennzeichnet, kein
  Kunde wird erfunden.
- Manche PDFs (Allianz, AXA, Gothaer) haben einen Sonderfall in der
  PDF-Erzeugung (Text pro Seite komplett gespiegelt/verdreht) -> eigene,
  koordinatenbasierte Parser (siehe extract_allianz/extract_axa/
  extract_gothaer).
- Dialog und Continentale liefern echte Kundenpositionstabellen, aber im
  Scan zu klein/dicht/verschachtelt fuer zuverlaessige OCR -> bleiben
  manuelle Pruefung statt falsche Zahlen zu riskieren (bei Dialog zusaetzlich
  ein OCR-Kontrollbetrag, siehe find_dialog_total_hint).
- VEMA-Pool liefert Kundenpositionen nicht im PDF, sondern separat als CSV
  (VEMA-Poolabrechnung-N.csv) -> wird automatisch mitverarbeitet, wenn sie
  im Monatsordner liegt bzw. mit hochgeladen wird (siehe extract_vema_csv).
"""

import csv
import glob
import os
import re
import sys
import unicodedata
from collections import Counter
from decimal import Decimal, ROUND_HALF_UP

import pdfplumber
import pandas as pd
import numpy as np

try:
    import pytesseract
    from PIL import ImageOps, ImageFilter
    TESSERACT_EXE = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    if os.path.exists(TESSERACT_EXE):
        pytesseract.pytesseract.tesseract_cmd = TESSERACT_EXE
    TOOL_DIR = os.path.dirname(os.path.abspath(__file__))
    LOCAL_TESSDATA = os.path.join(TOOL_DIR, "tessdata")
    if os.path.isdir(LOCAL_TESSDATA) and os.listdir(LOCAL_TESSDATA):
        # Lokales Sprachpaket vorhanden (Windows-Installation ohne Schreib-
        # zugriff auf das Tesseract-Programmverzeichnis) - andernfalls nutzt
        # Tesseract sein eigenes System-Tessdata-Verzeichnis, z.B. auf
        # Streamlit Cloud (Linux) ueber das apt-Paket tesseract-ocr-deu
        # (siehe packages.txt).
        os.environ["TESSDATA_PREFIX"] = LOCAL_TESSDATA
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # .../Umsatz
TOOL_DIR = os.path.dirname(os.path.abspath(__file__))                   # .../Umsatz/Courtage-Tool

# ---------------------------------------------------------------------------
# Konfiguration: Schluesselwoerter fuer die generische Tabellen-Engine
# ---------------------------------------------------------------------------

# Tier 1: eindeutige, volle Schluesselwoerter - werden immer zuerst versucht.
NAME_KEYWORDS_STRONG = [
    "versicherungsnehmer", "kunde", "name, vorname", "vertragsinhaber",
    "versicherter", "name vn", "mandant",
    # "Vers.-nehmer" (z.B. Ergo) - abgekuerzte Schreibweise von
    # "Versicherungsnehmer", die von "versicherungsnehmer" (voll ausgeschrieben)
    # und "versicherungs-" (WEAK, kollidiert mit "Versicherungs-Nr.") nicht
    # erfasst wird.
    "vers.-nehmer",
]
# Tier 2: nur als Fallback, wenn Tier 1 nichts findet (z.B. Kopfzeile ueber
# zwei Zeilen gebrochen: "Versicherungs-" / "nehmer"). "versicherungs-"
# kollidiert sonst mit einer Spalte "Versicherungs-Nr.". "vn/tu" ist die in
# manchen Abrechnungen (z.B. HDI-Leben: "VN/TU") uebliche Abkuerzung fuer
# Versicherungsnehmer - bewusst NICHT das kuerzere "vn" allein, das z.B. in
# Kontaktangaben wie "Prov./Court.(VN):" faelschlich matchen wuerde.
# "vertragsinformationen" (Mannheimer) ist keine Namens-Spaltenueberschrift
# im engeren Sinn, aber die Spalte, in der der Name/die Firma tatsaechlich
# steht.
NAME_KEYWORDS_WEAK = ["versicherungs-", "name", "vn/tu", "vertragsinformationen"]

# Woerter, die in manchen Abrechnungen (z.B. Mannheimer) an der Position der
# Namens-Spalte auftauchen, aber Buchungsart-Label sind, kein Kundenname
# (z.B. eine eigene Zeile "Bestandspflege" oder "Folgebeitrag" zwischen der
# echten Namenszeile und der Betragszeile). Wird zusaetzlich zur
# Alphalauf-Pruefung angewendet, da diese Woerter lang genug sind, um die
# Alphalauf-Pruefung fuer sich allein zu bestehen.
NON_NAME_WORDS = {
    "bestandspflege", "folgebeitrag", "erstbeitrag", "nachtrag",
    "vers.schein", "praemienregul.haftpflicht", "beitrag/gebuehr/zahlung",
    "fremdversicherungsnummer",
    # ohne fuehrenden Buchstaben, da manche PDFs "ae" als kaputtes "�"
    # rendern (siehe UMLAUT_MAP-Hinweis) - Teilstring-Vergleich statt exakt.
    "lter90tage", "provisioni",
}

AMOUNT_KEYWORDS = [
    # Hinweis: "verguet" statt "vergüt", weil normalize() Umlaute VOR dem
    # Vergleich in ue/oe/ae/ss uebersetzt (siehe UMLAUT_MAP/normalize()).
    "courtage", "provision", "verguet", "betrag", "gesamtverg",
    # "buch-" statt "buchwert", weil das Wort manchmal ueber zwei Zeilen
    # getrennt ist ("Buch-" / "wert"). Mit Bindestrich, damit es nicht mit
    # "Buchungs-" (Datum/Typ-Spalte, keine Betragsspalte) kollidiert.
    "buch-",
    # "prov." (mit Punkt, z.B. Itzehoer: "Prov.(€)") statt "provision" -
    # manche Abrechnungen kuerzen die Spaltenueberschrift ab. "%" wird ueber
    # AMOUNT_EXCLUDE separat ausgeschlossen (z.B. "Prov.(%)").
    "prov.",
    # "gutschrift" (z.B. Alte Leipziger: "Gutschrift/Belastung").
    "gutschrift",
]
AMOUNT_EXCLUDE = ["%", "satz", "-satz", "erwart"]

BLACKLIST_LINE_KEYWORDS = [
    "summe", "saldo", "gesamt", "endsaldo", "kontostand", "uebertrag",
    "vortrag", "zwischensumme", "kontenstand",
    "auszahlungsbetrag", "abrechnungsbetrag",
    "seite", "kontoauszug", "davon", "total", "bezeichnung",
    "jahreswerte", "zahlungsausgang",
    # "Auszahlung" pauschal (nicht nur "auszahlungsbetrag"): faengt u.a. die
    # Amex-Pool-Zeile "Sonstige Positionen" / "<Datum> Auszahlung freie
    # Stornoreserve <Betrag>" ab, die extract_generic_table sonst mit den
    # zuletzt erkannten (nicht mehr passenden) Spaltenbreiten der
    # vorangegangenen Tabelle als Muell-"Kundenzeile" einliest - der
    # eigentliche Betrag wird stattdessen sauber ueber
    # extract_amex_stornoreserve_auszahlung() erfasst.
    "auszahlung",
    # Fusszeile "Abrechnung vom <Datum>" (z.B. Amex-Pool) - kein Kunde.
    "abrechnung vom",
]
# "stornoreserve" bewusst NICHT (mehr) auf der Sperrliste: bei Fondsfinanz
# (eigener Parser, siehe extract_fondsfinanz) spielt das ohnehin keine
# Rolle, aber bei anderen Versicherern (z.B. Blaudirekt) ist eine
# "Stornoreserve"-Zeile ein echter, vertragsbezogener Abzug/Zuschlag, der
# dem jeweiligen Kunden gutzuschreiben ist (Nutzer-Bestaetigung), nicht
# eine rein informative Sammel-Zwischensumme.
# Nur DIESE (Teilmenge der Blacklist) setzen den "aktuellen Kunden" auch
# zurueck, statt die Zeile nur zu ueberspringen: sie markieren ein echtes
# Abschnittsende (z.B. eine Spaltenkopf-Wiederholung wie "VGART Bezeichnung
# ..." bei Deurag, nach der die naechste Zeile sonst faelschlich dem letzten
# Kunden zugerechnet wuerde). Andere Blacklist-Woerter wie "uebertrag" oder
# "kontostand" sind dagegen oft nur eine harmlose Salden-Notiz MITTEN in
# einem mehrzeiligen Kundenblock (z.B. HDI: "Uebertrag 6.941,03 738,23"
# gefolgt von einer weiteren Buchungszeile DESSELBEN Kunden) - dort wuerde
# ein Reset eine legitime Fortsetzungszeile verwerfen.
RESET_CUSTOMER_KEYWORDS = ["bezeichnung", "davon", "jahreswerte"]
# Ausnahmen: diese Zeilen ENTHALTEN "summe", zaehlen aber trotzdem als
# gueltige Kundenzeile (Fondsfinanz-Blockformat: "Vertragssumme EUR X")
BLACKLIST_EXCEPTIONS = ["vertragssumme"]

# Ein einzelner Buchungsbetrag ueber dieser Schwelle ist bei diesem Makler-
# volumen unplausibel (typische Courtage-Betraege liegen im ein- bis
# niedrigen vierstelligen Bereich) - eher ein Fehltreffer (z.B. eine
# Handelsregister- oder Kontonummer, die zufaellig in der Betrags-Spalte
# liegt). Solche Zeilen werden verworfen statt falsche Zahlen zu liefern.
MAX_PLAUSIBLE_AMOUNT = 20000

TOTAL_LINE_PATTERNS = [
    r"Endsumme[n]?.{0,40}?([\d.,]+)\s*$",
    r"Gesamtbetrag[:]?\s*([\d.,]+)",
    r"Auszahlungsbetrag[:]?\s*([\d.,]+)",
    r"Abrechnungsbetrag[:]?\s*([\d.,]+)",
    r"Gesamt-Auszahlungsbetrag[:]?\s*([\d.,]+)",
    r"Summe Vermittler.*?([\-\d.,]+)\s*$",
    # Erste Zahl nach "Auszahlung" (nicht letzte!): manche Abrechnungen
    # zeigen "aktueller Zeitraum | kumuliert" nebeneinander - die kumulierte
    # Spalte (meist die letzte Zahl der Zeile) waere sonst falsch.
    r"Auszahlung\s+(?:\d\d\.\d\d\.\d{4}\s+)?([\-\d.,]+)",
    r"Wir überweisen[^0-9]*([\d.,]+)",
    r"Saldo (?:neu|zum \d\d\.\d\d\.\d{4})[:]?\s*([\-\d.,]+)",
]

REVERSED_MARKERS = [
    "gnunhcerbA", "ztasmU", "gnurehcisreV", "esiewnihztasuZ",
    "remhensgnurehcisreV", "gnunneK",
]

MONTHS_DE = [
    "Januar", "Februar", "März", "April", "Mai", "Juni",
    "Juli", "August", "September", "Oktober", "November", "Dezember",
]


# ---------------------------------------------------------------------------
# Hilfsfunktionen
# ---------------------------------------------------------------------------

def parse_amount(text):
    """Parst deutsche (1.234,56 / 12,34-) und englische (1,234.56 / -12.34)
    Zahlenformate zu float. Gibt None zurueck, wenn kein plausibler Betrag."""
    if not text:
        return None
    s = text.strip()
    s = s.replace("€", "").replace("EUR", "").replace("�", "").strip()
    if not s:
        return None

    neg = False
    if s.startswith("-"):
        neg = True
        s = s[1:]
    if s.endswith("-"):
        neg = True
        s = s[:-1]
    s = s.strip()
    if not re.fullmatch(r"[\d.,]+", s):
        return None

    last_comma = s.rfind(",")
    last_dot = s.rfind(".")
    if last_comma == -1 and last_dot == -1:
        if not s.isdigit():
            return None
        val = float(s)
    elif last_comma > last_dot:
        # deutsches Format: Punkt = Tausender, Komma = Dezimal
        s2 = s.replace(".", "").replace(",", ".")
        try:
            val = float(s2)
        except ValueError:
            return None
    elif last_dot > last_comma:
        # englisches Format: Komma = Tausender, Punkt = Dezimal
        s2 = s.replace(",", "")
        try:
            val = float(s2)
        except ValueError:
            return None
    else:
        try:
            val = float(s)
        except ValueError:
            return None

    return -val if neg else val


UMLAUT_MAP = str.maketrans({
    "ü": "ue", "ö": "oe", "ä": "ae", "ß": "ss",
    "Ü": "ue", "Ö": "oe", "Ä": "ae",
})


def normalize(word):
    return word.lower().strip().translate(UMLAUT_MAP)


def matches_any(word_lower, keywords):
    return any(k in word_lower for k in keywords)


def _blacklist_check(text):
    if any(exc in text for exc in BLACKLIST_EXCEPTIONS):
        return False
    return any(k in text for k in BLACKLIST_LINE_KEYWORDS)


def is_blacklisted_line(line_text_lower):
    """Prueft eine Zeile gegen die Blacklist. Prueft zusaetzlich eine um
    Ziffern bereinigte Variante, da manche PDFs Kopf-/Fusszeilen-Texte
    (z.B. eine Positionsnummer) mit Wort-Ueberlagerung zeichenweise
    ineinander schieben, z.B. "3S0u1m2m11e" statt "Summe" - sowie eine
    Leerzeichen-bereinigte Variante, da manche PDFs Woerter mit
    Buchstaben-Spatium rendern, z.B. "S U M M E" statt "Summe"."""
    normalized = line_text_lower.translate(UMLAUT_MAP)
    if _blacklist_check(normalized):
        return True
    if _blacklist_check(re.sub(r"\d", "", normalized)):
        return True
    return _blacklist_check(normalized.replace(" ", ""))


def cluster_lines(words, tol=2.5):
    """Gruppiert Woerter nach vertikaler Position (top) zu Zeilen."""
    lines = []
    for w in sorted(words, key=lambda w: (w["top"], w["x0"])):
        placed = False
        for line in lines:
            if abs(line["top"] - w["top"]) <= tol:
                line["words"].append(w)
                line["top"] = (line["top"] * len(line["words"]) + w["top"]) / (len(line["words"]) + 1)
                placed = True
                break
        if not placed:
            lines.append({"top": w["top"], "words": [w]})
    for line in lines:
        line["words"].sort(key=lambda w: w["x0"])
    lines.sort(key=lambda l: l["top"])
    return lines


def get_words_for_page(page, resolution=300):
    """Liefert Woerter (text,x0,x1,top,bottom) einer Seite - aus der Text-
    ebene, oder per OCR falls die Seite keinen Text enthaelt (Scan)."""
    if len(page.chars) > 0:
        words = page.extract_words(x_tolerance=1.5, keep_blank_chars=False)
        return words, "text"

    if not OCR_AVAILABLE:
        return [], "kein-ocr"

    scale = 72.0 / resolution
    im = page.to_image(resolution=resolution).original
    data = pytesseract.image_to_data(im, lang="deu", output_type=pytesseract.Output.DICT)
    words = []
    n = len(data["text"])
    for i in range(n):
        txt = data["text"][i].strip()
        if not txt:
            continue
        conf = data.get("conf", ["-1"] * n)[i]
        try:
            if float(conf) < 25:
                continue
        except (ValueError, TypeError):
            pass
        x, y, w_, h_ = data["left"][i], data["top"][i], data["width"][i], data["height"][i]
        words.append({
            "text": txt,
            "x0": x * scale,
            "x1": (x + w_) * scale,
            "top": y * scale,
            "bottom": (y + h_) * scale,
        })
    return words, "ocr"


def _merge_touching_words(words, max_gap=1.0):
    """Fuehrt Woerter derselben Zeile zusammen, die sich beruehren (Abstand
    < max_gap pt). Itzehoer (September-2026): jedes Wort/jede Zahl wurde
    nach dem ERSTEN Buchstaben aufgetrennt ('B' + 'ozkurt', '0' +
    '1.09.2026'), vermutlich ein Schriftart-Wechsel im Quell-PDF - ohne
    Zusammenfuehren erkennt die Tabellen-Engine weder Namen noch Daten."""
    ordered = sorted(words, key=lambda w: (round(w["top"] / 3.0), w["x0"]))
    merged = []
    for w in ordered:
        if merged:
            prev = merged[-1]
            same_line = abs(prev["top"] - w["top"]) < 2.0
            if same_line and -0.5 <= w["x0"] - prev["x1"] < max_gap:
                merged[-1] = {
                    "text": prev["text"] + w["text"],
                    "x0": prev["x0"], "x1": max(prev["x1"], w["x1"]),
                    "top": min(prev["top"], w["top"]),
                    "bottom": max(prev["bottom"], w["bottom"]),
                }
                continue
        merged.append(dict(w))
    return merged


def _find_name_and_amount(combined_words):
    """Liefert (name_word, amount_word) oder None, falls in den gegebenen
    Woertern nicht beides eindeutig vorkommt. Nutzt zuerst die eindeutigen
    (Tier-1) Namens-Schluesselwoerter, erst danach die mehrdeutigeren
    Tier-2-Woerter (die z.B. mit 'Versicherungs-Nr.' kollidieren koennen)."""
    def next_word_is_satz(word):
        # "Prov." + "Satz" als ZWEI getrennte Woerter (z.B. Domcura:
        # "Prov. in % ... Prov. Satz" ueber zwei Zeilen gebrochener
        # Tabellenkopf) ist eine Provisionssatz-/Prozent-Spalte, keine
        # Betragsspalte - als EIN Wort ("Provisionssatz") faengt bereits
        # AMOUNT_EXCLUDE ("satz" als Teilstring) das ab, hier zusaetzlich
        # fuer den Fall zweier getrennter Woerter in derselben Zeile.
        same_line = [w for w in combined_words if w["_line"] == word["_line"] and w is not word]
        after = [w for w in same_line if w["x0"] > word["x0"]]
        if not after:
            return False
        nxt = min(after, key=lambda w: w["x0"])
        return normalize(nxt["text"]).startswith("satz")

    amount_candidates = [
        w for w in combined_words
        if matches_any(normalize(w["text"]), AMOUNT_KEYWORDS)
        and not matches_any(normalize(w["text"]), AMOUNT_EXCLUDE)
        and not next_word_is_satz(w)
    ]
    if not amount_candidates:
        return None

    for name_keywords in (NAME_KEYWORDS_STRONG, NAME_KEYWORDS_WEAK):
        name_candidates = [w for w in combined_words if matches_any(normalize(w["text"]), name_keywords)]
        if name_candidates:
            name_word = min(name_candidates, key=lambda w: w["x0"])
            amount_word = max(amount_candidates, key=lambda w: w["x0"])
            return name_word, amount_word
    return None


def _header_from_span(lines, start, span):
    end = min(start + span, len(lines))
    combined = []
    for k in range(start, end):
        for w in lines[k]["words"]:
            tagged = dict(w)
            tagged["_line"] = k
            combined.append(tagged)
    if not combined:
        return None
    found = _find_name_and_amount(combined)
    if not found:
        return None
    name_word, amount_word = found

    def right_boundary(word, all_words):
        # nur Woerter aus derselben Zeile wie das gefundene Schluesselwort
        # heranziehen, damit fachfremde Woerter aus einer bei groesseren
        # Fenstern zufaellig mit erfassten Nachbarzeile die Spaltenbreite
        # nicht verfaelschen.
        same_line = [w for w in all_words if w["_line"] == word["_line"]]
        candidates = [w["x0"] for w in same_line if w["x0"] > word["x0"] + 3]
        # kleiner Sicherheitsabstand: Kopfzeilen-Wort und Datenzeilen-Wert
        # sind selten pixelgenau gleich ausgerichtet - ohne Puffer wuerde ein
        # Datenwert, der zufaellig 0,1pt vor der naechsten Spaltenueberschrift
        # beginnt, faelschlich noch in diese Spalte hineingezaehlt.
        return (min(candidates) - 2) if candidates else word["x1"] + 250

    def left_boundary(word, all_words):
        # Zahlen-Spalten sind meist rechtsbuendig: ein groesserer Betrag
        # (mehr Ziffern) beginnt weiter links als die Kopfzeilen-Ueberschrift
        # selbst. Als linke Grenze daher das rechte Ende der VORHERGEHENDEN
        # Spalte in derselben Zeile nehmen (mit etwas Puffer), statt der
        # eigenen x0 der Ueberschrift - sonst werden breite Zahlen
        # abgeschnitten.
        same_line = [w for w in all_words if w["_line"] == word["_line"]]
        candidates = [w["x1"] for w in same_line if w["x1"] < word["x0"] - 3]
        return max(candidates) + 2 if candidates else max(0, word["x0"] - 5)

    # Nur fuer die Betragsspalte auf die vorherige Spalte stuetzen (rechts-
    # buendige, unterschiedlich breite Zahlen wie "1.134,16" wuerden sonst
    # abgeschnitten). Fuer die Namensspalte NICHT, da die Namens-/Produkt-
    # Spalte oft eine kurze Kopfzeilen-Ueberschrift (z.B. "VSNR") aber viel
    # laengere Dateninhalte (z.B. lange Produktnamen) hat - dort wuerde die
    # vorige-Spalte-Grenze faelschlich in die Nachbarspalte hineinragen
    # (siehe R+V: "VSNR" ist kurz, "Produktname" in derselben Spalte lang).
    name_left = max(0, name_word["x0"] - 5)
    name_right = right_boundary(name_word, combined)
    amount_left = left_boundary(amount_word, combined)
    amount_right = right_boundary(amount_word, combined)
    return {
        "name_range": (name_left, name_right),
        "amount_range": (amount_left, amount_right),
    }, end


def extract_generic_table(pages_words):
    """Generische koordinatenbasierte Extraktion fuer 'normale' Tabellen-PDFs.
    pages_words: Liste von (page_index, words, source) Tupeln.

    Kopfzeilen werden nicht nur einmal pro Seite gesucht, sondern bei jeder
    Zeile neu geprueft: manche Versicherer (z.B. Amex-Pool) haben mehrere
    Teiltabellen mit unterschiedlichen Spaltenbreiten auf derselben Seite
    (z.B. 'Abschlussprovisionen' und 'Folgeprovisionen')."""
    rows = []
    columns = None
    current_customer = None

    for page_idx, words, source in pages_words:
        if not words:
            continue
        lines = cluster_lines(words)
        li = 0
        n = len(lines)

        while li < n:
            header_found = None
            for span in (1, 2, 3):
                result = _header_from_span(lines, li, span)
                if result:
                    header_found = result
                    break
            if header_found:
                columns, next_li = header_found
                li = next_li
                continue

            if columns is None:
                li += 1
                continue

            line = lines[li]
            line_text = " ".join(w["text"] for w in line["words"])
            line_lower = line_text.lower()
            if is_blacklisted_line(line_lower):
                if any(k in line_lower for k in RESET_CUSTOMER_KEYWORDS):
                    # Echtes Abschnittsende (siehe RESET_CUSTOMER_KEYWORDS):
                    # aktuellen Kunden zuruecksetzen, sonst wuerde eine
                    # nachfolgende, selbst nicht geblacklistete Rekap-Zeile
                    # (z.B. eine Aufschluesselung nach Vertragsart/Jahres-
                    # werten, die denselben Betrag nochmal zusammenfasst)
                    # faelschlich per Fortschreibung dem letzten echten
                    # Kunden zugerechnet und so doppelt gezaehlt.
                    current_customer = None
                li += 1
                continue

            name_l, name_r = columns["name_range"]
            amt_l, amt_r = columns["amount_range"]
            name_words = [w for w in line["words"] if name_l <= w["x0"] < name_r]
            amount_words = [w for w in line["words"] if amt_l <= w["x0"] < amt_r]
            other_words = [
                w for w in line["words"]
                if w not in name_words and w not in amount_words
            ]

            name_text = " ".join(w["text"] for w in name_words).strip(" ,")
            if name_text:
                letter_runs = re.findall(r"[A-Za-zÀ-ÿ]+", name_text)
                longest_run = max((len(r) for r in letter_runs), default=0)
                name_joined = normalize(name_text).replace(" ", "")
                if longest_run < 3:
                    # kein echter Kundenname: z.B. eine verirrte Saldozahl,
                    # oder ein Fragment einer Folgezeile (Datum/Kennzeichen
                    # wie "4 01.07.26 184,76 FB"), das zufaellig in die
                    # Namensspalte faellt.
                    name_text = ""
                elif any(nn in name_joined for nn in NON_NAME_WORDS):
                    # Buchungsart-/Spaltenkopf-Label (z.B. "Bestandspflege",
                    # "Erstbeitrag Vers.Schein"), kein Kundenname - siehe
                    # NON_NAME_WORDS. Teilstring-Vergleich (nicht exakte
                    # Gleichheit), da diese Woerter manchmal an einen
                    # zusaetzlichen echten Namensteil angrenzen.
                    name_text = ""
                elif re.match(r"^[A-ZÀ-Ü]{2,4}-[A-ZÀ-Ü]{1,4}(\s+\d+)?$", name_text):
                    # Kurzes Risikoort-/Produkt-Kuerzel mit Bindestrich (z.B.
                    # "HAL-SL 321", "LU-SW 312"), kein Kundenname - manche
                    # Abrechnungen (z.B. HDI) verwenden dieselbe Spalten-
                    # position auf Folgezeilen fuer ein anderes Feld.
                    name_text = ""

            # Von rechts nach links das erste als Zahl parsebare Wort in der
            # Betragsspalte nehmen - nicht nur die letzten ein/zwei Woerter
            # pruefen: manche Abrechnungen (z.B. Blaudirekt) haengen nach dem
            # Betrag noch einen Hinweis wie "Festbetrag" an, der sonst die
            # Zeile faelschlich verwerfen wuerde (Betrag stuende dann drei
            # statt zwei Woerter vor Zeilenende).
            amount_val = None
            for w in reversed(amount_words):
                amount_val = parse_amount(w["text"])
                if amount_val is not None:
                    break

            if name_text:
                current_customer = name_text

            li += 1
            if amount_val is None:
                continue
            if abs(amount_val) > MAX_PLAUSIBLE_AMOUNT:
                continue

            # Erlaeuterungstext einer Fortsetzungszeile ohne eigenen Kunden-
            # namen (z.B. "Stornoreserve 10%" oder "Vertrauensschaden-
            # haftpflicht" bei einem vertragsbezogenen Abzug/Zuschlag) landet
            # manchmal in derselben Spalte wie der Betrag selbst, statt in
            # other_words - z.B. "-4,86 EUR Stornoreserve" mit x0 noch
            # innerhalb der Betragsspalte. Nur "EUR"/"€" und das erkannte
            # Betragswort selbst gelten als reine Fuellwoerter ohne
            # Aussagekraft.
            extra_amount_words = [
                w for w in amount_words
                if parse_amount(w["text"]) is None and normalize(w["text"]) not in ("eur", "€")
            ]
            # other_words muss ECHTEN Erlaeuterungstext enthalten, um eine
            # Fortsetzungszeile zu erkennen - eine reine Tabellen-Zwischen-
            # summe ohne Textlabel (z.B. Amex-Pool: "1.846,76 115,48" als
            # letzte Zeile einer Kundentabelle, N-Praemie- und Courtage-
            # Spaltensumme ohne "Summe"-Wort) faellt sonst in other_words
            # (die N-Praemie-Zahl liegt ausserhalb von Name-/Betragsspalte)
            # und wuerde faelschlich als weitere Buchungszeile des zuletzt
            # gesehenen Kunden gezaehlt (doppelt gezaehlte Spaltensumme).
            meaningful_other_words = [
                w for w in other_words
                if parse_amount(w["text"]) is None and normalize(w["text"]) not in ("eur", "€")
            ]
            has_context = bool(meaningful_other_words) or bool(extra_amount_words)

            if name_text:
                rows.append((page_idx, current_customer, amount_val, line_text, source))
            elif has_context:
                if current_customer:
                    rows.append((page_idx, current_customer, amount_val, line_text, source))
            # reine Zahlen-Zeile ohne sonstigen Inhalt -> Zwischensumme, skip

    return rows


def extract_fondsfinanz(pages_words):
    """Bloc-Format von Fondsfinanz: 'VN Nachname, Vorname ...' gefolgt von
    einer Provisionsuebersicht und einer Zeile 'Vertragssumme EUR X'.

    Fondsfinanz behaelt zusaetzlich einen pauschalen, nicht kundenbezogenen
    'Stornoreserve'-Betrag als Sicherheit ein (siehe Deckblatt). Nach
    Rueckmeldung des Nutzers zaehlt fuer die Courtage-Erfassung die volle
    (versteuerte) Courtage inkl. dieses Einbehalts, nicht der um die Reserve
    gekuerzte Auszahlungsbetrag - der Einbehalt wird daher als eigene,
    ausdruecklich nicht-kundenspezifische Zeile wieder hinzugerechnet."""
    rows = []
    current_name = None
    storno_total = 0.0
    storno_line = None
    storno_page = None
    for page_idx, words, source in pages_words:
        if not words:
            continue
        lines = cluster_lines(words)
        for line in lines:
            toks = [w["text"] for w in line["words"]]
            if not toks:
                continue
            line_text = " ".join(toks)
            if toks[0] == "VN":
                if len(toks) >= 3 and toks[1].endswith(","):
                    current_name = toks[1] + " " + toks[2]
                elif len(toks) >= 2:
                    current_name = toks[1]
            elif toks[0] == "VP":
                continue
            elif "Vertragssumme" in toks:
                idx = toks.index("Vertragssumme")
                amt = None
                for t in toks[idx + 1:]:
                    amt = parse_amount(t)
                    if amt is not None:
                        break
                if amt is not None and current_name:
                    rows.append((page_idx, current_name, amt, line_text, source))
            elif toks[0] == "Stornoreserve":
                amt = parse_amount(toks[1]) if len(toks) > 1 else None
                if amt is not None and amt != 0:
                    storno_total += abs(amt)
                    storno_line = line_text
                    storno_page = page_idx

    if storno_total:
        rows.append((
            storno_page,
            "Stornoreserve-Einbehalt (nicht kundenspezifisch, siehe Deckblatt)",
            round(storno_total, 2),
            storno_line,
            "text",
        ))
    return rows


def extract_vn_summevertrag_block(pages_words):
    """Format wie Swiss Life 'Anhang Einzelaufstellung': Name und Betrag
    stehen auf derselben Zeile, z.B.
    'Vers.Nr. 9317437 VN/VP Knaus, Maximilian Referenznummer 74
    Summe Vertrag 82,03'."""
    rows = []
    for page_idx, words, source in pages_words:
        if not words:
            continue
        lines = cluster_lines(words)
        for line in lines:
            toks = [w["text"] for w in line["words"]]
            if "VN/VP" in toks:
                vn_idx = toks.index("VN/VP")
            elif "VN" in toks:
                vn_idx = toks.index("VN")
            else:
                continue
            if vn_idx + 1 >= len(toks):
                continue
            if toks[vn_idx + 1].endswith(",") and vn_idx + 2 < len(toks):
                name = toks[vn_idx + 1] + " " + toks[vn_idx + 2]
            else:
                name = toks[vn_idx + 1]

            amt = None
            for i in range(len(toks) - 2):
                if toks[i] == "Summe" and toks[i + 1] == "Vertrag":
                    amt = parse_amount(toks[i + 2])
                    break
            if amt is not None:
                line_text = " ".join(toks)
                rows.append((page_idx, name, amt, line_text, source))
    return rows


SWISS_LIFE_VSV_RE = re.compile(
    r"BeiträgezurVertrauensschadenversicherung\s+[\d.,]+\s+(-?[\d.,]+)\s+(-?[\d.,]+)"
)


def extract_swiss_life_vsv_deduction(pages_full_text):
    """Swiss Life behaelt pro Abrechnung einen pauschalen (nicht kunden-
    bezogenen) Beitrag zur Vertrauensschadenversicherung ein - anders als
    urspruenglich angenommen wird dieser NICHT separat von RH persoenlich
    getragen, sondern ist direkt von der Abschlussprovision abzuziehen
    (Nutzer-Rueckmeldung). Er steht explizit auf der "Abrechnungsuebersicht
    zum ..."-Seite (nicht auf der "Kumulierte Jahreswerte"-Seite, die
    dieselbe Zeile fuer das gesamte Jahr zeigt) als eigene Zeile
    "BeitraegezurVertrauensschadenversicherung <Gutschrift> <Belastung>
    <Saldo>", z.B. "... 0,00 -3,58 -3,58" - der Saldo-Wert wird als eigene,
    ausdruecklich nicht-kundenspezifische Abzugszeile zurueckgegeben."""
    for pidx, text in pages_full_text:
        if "SummenausaktuellerAbrechnung" not in text:
            continue
        m = SWISS_LIFE_VSV_RE.search(text)
        if m:
            saldo = parse_amount(m.group(2))
            if saldo:
                return pidx, saldo, m.group(0)
    return None


AMEX_STORNORESERVE_AUSZAHLUNG_RE = re.compile(
    r"\d{2}\.\d{2}\.\d{4}\s+Auszahlung\s+freie\s+Stornoreserve\s+([\d.,]+)"
)


def extract_amex_stornoreserve_auszahlung(pages_full_text):
    """Amex-Pool: manuell angestossene Auszahlung einer 'freien Stornoreserve'
    steht als eigene Zeile im Abschnitt 'Sonstige Positionen', z.B.
    '17.08.2026 Auszahlung freie Stornoreserve 1898,39'. Lebensversicherungs-
    Provisionen werden bei SSH bereits vollstaendig (inkl. des einbehaltenen
    Stornoreserve-Anteils) beim urspruenglichen Buchungsdatum verbucht und
    einem Kunden/Partner zugeordnet - die spaetere Auszahlung der Reserve
    betrifft rueckwirkend viele verschiedene, nicht mehr einzeln nachvoll-
    ziehbare Kunden (Nutzer-Bestaetigung) und wird daher als eigene, mit
    NICHT_ZUORDENBAR_MARKER gekennzeichnete Zeile ohne Kunden-/Betreuer-
    Zuordnung zurueckgegeben (siehe process_file(), apply_betreuer())."""
    for pidx, text in pages_full_text:
        m = AMEX_STORNORESERVE_AUSZAHLUNG_RE.search(text)
        if m:
            amt = parse_amount(m.group(1))
            if amt:
                return pidx, amt, m.group(0)
    return None


VHV_ADDR_RE = re.compile(r"^(.+?),\s*.+?,\s*\d{5}\s+\S")

# Links von der Name/Adresse-Spalte (beginnt konstant bei x0 ~165) steht bei
# VHV eine eigene, schmale Spalte (x0 ~77-95) mit einem internen Kuerzel -
# mal eine Kundengruppen-/Sparten-Nummer (z.B. "3374"), mal ein
# Vermittler-Kuerzel ("TS"/"RH"/"AS"). cluster_lines() gruppiert beides in
# dieselbe Zeile wie den Namen, da rein visuell (gleiche Zeilenhoehe)
# nicht unterscheidbar - ohne Filterung landet das Kuerzel als Praefix im
# extrahierten Kundennamen (z.B. "TS Boris Eichhorn" statt "Boris Eichhorn")
# und die Namenszuordnung zu Betreuer.xlsx schlaegt fehl. Schwellwert 120
# liegt sicher zwischen beiden Spalten (Kuerzel-Spalte endet bei ca. 95,
# Name-Spalte beginnt bei ca. 165).
VHV_NAME_COLUMN_X0_MIN = 120

# Am Ende jeder Abrechnungsgruppe druckt VHV eine Sparten-Aufschluesselung
# ("10 Haftpflicht 51,32 12,84" / "20 KFZ VHV Allgemeine 8.853,06 622,78" /
# "60 Sachversicherungen 49,57 12,40") gefolgt von "Gesamtsumme ...". Die
# Gesamtsumme-Zeile selbst ist bereits blacklisted (siehe
# BLACKLIST_LINE_KEYWORDS), die "20 KFZ ..."-Zeile davor aber nicht - sie hat
# zufaellig >=6 Tokens und endet auf einen gueltigen Betrag, wuerde also als
# Buchungszeile in "pending" landen. Bisher beobachtet nur am Ende der
# jeweils letzten Seite einer Gruppe (kein nachfolgender Adresstreffer mehr,
# daher folgenlos), aber sicherheitshalber hier explizit ausgeschlossen -
# nur fuer VHV, nicht global (die Woerter koennten bei anderen Versicherern
# legitim in einer Kundenzeile vorkommen).
VHV_SPARTE_RECAP_WORDS = ("haftpflicht", "sachversicherungen", "kfz")


def extract_vhv(pages_words):
    """VHV: 1-3 Buchungszeilen (Betrag als letztes Feld) stehen VOR der
    zugehoerigen Zeile 'Firma/Name, Strasse , PLZ Ort' - nicht danach wie
    bei den meisten anderen Versicherern. Die gepufferten Betraege werden
    rueckwirkend der folgenden Name+Adresse-Zeile zugeordnet. 'pending' wird
    bewusst NICHT pro Seite zurueckgesetzt, sondern nur pro Datei: manche
    Kundenbloecke (Storno + Neubuchung, oft mehrere Buchungszeilen) enden
    exakt am Seitenende, die zugehoerige Name+Adresse-Zeile folgt dann erst
    als allererste Zeile der naechsten Seite - ein Reset pro Seite wuerde
    diese Betraege sonst stillschweigend verwerfen (beobachtet bei
    August-2026: 'Nazim Turan', 12 am Ende von Seite 8 gepufferte Zeilen,
    Summe -13,82 EUR, gingen so komplett verloren)."""
    rows = []
    pending = []
    for page_idx, words, source in pages_words:
        if not words:
            continue
        lines = cluster_lines(words)
        for line in lines:
            toks = [w["text"] for w in line["words"]]
            if not toks:
                continue
            line_text = " ".join(toks)
            m = VHV_ADDR_RE.match(line_text)
            if m:
                name_toks = [w["text"] for w in line["words"] if w["x0"] >= VHV_NAME_COLUMN_X0_MIN]
                name = (" ".join(name_toks) if name_toks else line_text).split(",")[0].strip()
                for amt, ltxt in pending:
                    rows.append((page_idx, name, amt, ltxt, source))
                pending = []
                continue
            if is_blacklisted_line(line_text.lower()):
                # Nur bei einer ECHTEN Abschnittsende-Zeile (Gesamtsumme/
                # Zwischensumme) den Puffer verwerfen - eine reine Seiten-
                # fusszeile wie "Seite 3 von 6" (blacklisted wegen "seite")
                # darf laufende Buchungen NICHT loeschen, sonst gehen genau
                # die Faelle verloren, die schon den Seitenumbruch-Fix oben
                # noetig gemacht haben (die Fusszeile steht nach den letzten
                # Buchungszeilen, aber vor der Adresszeile auf der naechsten
                # Seite).
                if any(k in line_text.lower() for k in ("summe", "gesamt")):
                    pending = []
                continue
            if any(w in line_text.lower() for w in VHV_SPARTE_RECAP_WORDS):
                continue
            if len(toks) >= 6:
                amt = parse_amount(toks[-1])
                if amt is not None and abs(amt) <= MAX_PLAUSIBLE_AMOUNT:
                    pending.append((amt, line_text))
    return rows


DOMCURA_POLICY_RE = re.compile(r"^[A-Z]{2,6}-\d+-\d+$")


def extract_domcura(pages_words):
    """Domcura-Einzelaufstellung: jede Buchungszeile beginnt mit einem
    Policennummer-Kuerzel (z.B. 'IDCPT-03-1012980', 'DCDO-94-2078570'),
    gefolgt vom Kundennamen bis zum naechsten festen Schluesselwort
    'DOMCURA' (Produktbezeichnung), und endet mit dem Provisionsbetrag
    direkt vor einem kurzen Vergluetungsart-Kuerzel (z.B. 'FP', 'AP') -
    NICHT dem vorherigen 'Nettobetrag', der ebenfalls im selben Format
    dasteht. Die zugehoerige Adress-/VP-Folgezeile (z.B. 'Vinh Duc D-67240
    ... 0,00 %') beginnt nicht mit einem Policennummer-Kuerzel und wird
    dadurch automatisch uebersprungen - ebenso die abschliessenden
    'Zwischensumme'/'Zusammenfassung'/'Gesamtbetrag'-Rekapitulationszeilen,
    die den generischen Tabellen-Parser hier faelschlich als zusaetzliche
    Kundenzeile einlesen wuerden (kein Blacklist-Wort wie 'summe'/'gesamt'
    in der Zeile 'Sachversicherung 190,35 €' der Zusammenfassung)."""
    rows = []
    for page_idx, words, source in pages_words:
        if not words:
            continue
        lines = cluster_lines(words)
        for line in lines:
            toks = [w["text"] for w in line["words"]]
            if not toks or not DOMCURA_POLICY_RE.match(toks[0]):
                continue
            if "DOMCURA" not in toks:
                continue
            name_idx = toks.index("DOMCURA")
            name = " ".join(toks[1:name_idx]).strip(" ,")
            amt = None
            for t in reversed(toks):
                amt = parse_amount(t)
                if amt is not None:
                    break
            if amt is not None and name:
                rows.append((page_idx, name, amt, " ".join(toks), source))
    return rows


ERGO_ART_CODES = ("AP", "BP")


def extract_ergo(pages_words):
    """Ergo 'Verguetungsnachweis': jede Buchungszeile beginnt mit einem
    Verguetungsart-Kuerzel ('AP' Abschlussprovision oder 'BP'
    Bestandsprovision), gefolgt vom Kundennamen - mal als EIN Token
    ('Nachname,Vorname'), mal als ZWEI Token ('Nachname Vorname') ohne
    Komma, uneinheitlich sogar fuer denselben Kunden innerhalb derselben
    Datei (z.B. 'Gerszewski Christine' und 'Gerszewski,Christine'). Der
    Provisionsbetrag steht immer als letztes Token der Zeile. Der
    zweizeilige Tabellenkopf ('Verg.-art Vers.-nehmer ... / VNR Grund ...')
    sowie die 'AP/BP Provisionsvertrags-Nr: <N>'-Zwischenzeilen werden
    automatisch ausgeschlossen (zu wenige Tokens bzw. keine gueltige
    Betragszahl am Ende)."""
    rows = []
    for page_idx, words, source in pages_words:
        if not words:
            continue
        lines = cluster_lines(words)
        for line in lines:
            toks = [w["text"] for w in line["words"]]
            if len(toks) < 5 or toks[0] not in ERGO_ART_CODES:
                continue
            if "," in toks[1]:
                name = toks[1]
            elif re.fullmatch(r"[A-Za-zÀ-ÿ.\-]+", toks[2]):
                name = toks[1] + " " + toks[2]
            else:
                name = toks[1]
            name = name.rstrip(",").strip()
            amt = parse_amount(toks[-1])
            if amt is None or abs(amt) > MAX_PLAUSIBLE_AMOUNT:
                continue
            rows.append((page_idx, name, amt, " ".join(toks), source))
    return rows


def extract_sv_sparkasse(pages_words):
    """SV Sparkassenversicherung (OCR, kein Tabellenkopf): jede Buchungszeile
    beginnt mit dem Kundennamen, gefolgt von Codes/Daten, und endet mit einem
    Betrag gefolgt von einem 1-2-stelligen Flag-Zeichen (Ziffer oder durch
    OCR verfaelschter Buchstabe), z.B.
    'Hu, Hsiao-Lei 10 00 2 50090787898 BP ... 29,06 25,000 7,27 2'.
    Seite 0 (Anschreiben) und 1 (Aenderungsgrund-Legende) enthalten keine
    Buchungen und werden uebersprungen."""
    rows = []
    for page_idx, words, source in pages_words:
        if not words or page_idx < 2:
            continue
        lines = cluster_lines(words)
        for line in lines:
            toks = [w["text"] for w in line["words"] if "#" not in w["text"]]
            if len(toks) < 8:
                continue
            line_text = " ".join(toks)
            if is_blacklisted_line(line_text.lower()):
                continue
            name_toks = []
            i = 0
            while i < len(toks) and not re.fullmatch(r"[\d.,]+", toks[i]):
                name_toks.append(toks[i])
                i += 1
            if not name_toks or i >= len(toks) - 3:
                continue
            name = " ".join(name_toks).strip(" ,")
            if len(re.sub(r"[^A-Za-zÀ-ÿ]", "", name)) < 3:
                continue
            last = toks[-1]
            # der Betrag hat ein Komma (Dezimaltrenner), das nachgestellte
            # Flag-Zeichen nicht.
            amt_tok = last if "," in last else (toks[-2] if len(toks) >= 2 else last)
            amt = parse_amount(amt_tok)
            if amt is None or abs(amt) > MAX_PLAUSIBLE_AMOUNT:
                continue
            rows.append((page_idx, name, amt, line_text, source))
    return rows


ARAG_NAME_RE = re.compile(r"07(?!\d)\s*[:.\|]?\s*(.+?)\s+70\b")
ARAG_NUM_RE = re.compile(r"-?\d{1,3}(?:\.\d{3})*,\d{2}")


def extract_arag(pdf):
    """ARAG-Einzelaufstellung (z.B. 'PROVISIONS-PRODUKTIONSABRECHNUNG'):
    gescanntes PDF, dessen Positionstabelle mit dem Standard-OCR-Pfad
    (get_words_for_page, 300dpi) zu schlecht erkannt wird. Hier wird pro
    Seite mit hoeherer Aufloesung (400dpi) und tabellenorientiertem OCR-Modus
    (--psm 6) als Fliesstext gelesen und zeilenweise per Regex geparst, statt
    ueber Wort-Positionen: 'AT-Nr 07 Nachname, Vorname 70 Beitrag Satz
    Provision [Provision wiederholt] Datum'.

    Storno-Zeilen (Beitrag mit Minus) tauchen manchmal als zweite Zeile mit
    demselben Namen und demselben (positiven) Provisionsbetrag auf, ohne
    dass die OCR das Minuszeichen zuverlaessig erkennt (schwankt je nach
    Aufloesung). Statt uns auf dieses Vorzeichen zu verlassen, wird daher
    jede Zeile, deren (normalisierter) Name UND Betrag mit der direkt
    vorherigen Zeile uebereinstimmen, als Storno der vorherigen Zeile
    behandelt und negativ gezaehlt - das entspricht der Beobachtung, dass
    beide Zeilen sich in der echten Abrechnung zu 0 aufheben."""
    rows = []
    prev_norm_key = None
    for page_idx, page in enumerate(pdf.pages):
        if len(page.chars) > 0:
            continue
        im = page.to_image(resolution=400).original
        text = pytesseract.image_to_string(im, lang="deu", config="--psm 6")

        for line in text.splitlines():
            if not re.search(r"07(?!\d)", line):
                continue
            m = ARAG_NAME_RE.search(line)
            if not m:
                continue
            name = re.sub(r"^[^A-Za-zÀ-ÿ]+", "", m.group(1)).strip(" .|_-;:")
            name_norm = re.sub(r"[^A-Za-zÀ-ÿ]", "", name).lower()
            if len(name_norm) < 3:
                continue
            nums = ARAG_NUM_RE.findall(line)
            if not nums:
                continue
            amt = parse_amount(nums[-1])
            if amt is None or abs(amt) > MAX_PLAUSIBLE_AMOUNT:
                continue

            norm_key = (name_norm, round(abs(amt), 2))
            if norm_key == prev_norm_key:
                # Storno-Gegenbuchung derselben Person/desselben Betrags.
                amt = -abs(amt)
                prev_norm_key = None
            else:
                prev_norm_key = norm_key
            rows.append((page_idx, name, amt, line.strip(), "ocr"))
    return rows


BHV_HAUSBESITZER_ROW_RE = re.compile(
    r"^\d{6,}\s+.+?\s(\d+[.,]\d{2})\s+(\d+[.,]\d{2})\s+(Herrn?|Frau|Firma)\s+(.+)$"
)
BHV_HAUSBESITZER_TOTAL_RE = re.compile(
    r"Provisionsanspruch von insgesamt\s+([\d.,]+)\s*EUR", re.IGNORECASE
)


def extract_bayerische_hausbesitzer(pdf):
    """Bayerische Hausbesitzer-Versicherung 'Provisionsabrechnung': gescanntes
    Einzelblatt pro Vertrag mit Sparten-Unterzeilen (z.B. 'ELEMENTAR V 218.67
    25.00 54.67', 'FEUER V ... ', ohne Namen) und einer abschliessenden
    Gesamtzeile 'Vertragsnr Tarif Beitrag Provision Herr/Frau/Firma Name PLZ
    Ort'. Der generische Tabellen-Parser fing vorher zusaetzlich Zahlen aus
    dem Anschreiben ein (Agenturnummer, IBAN-Ziffern) und jede Sparten-
    Unterzeile als eigene 'Kunden' - real ist pro Datei nur die Gesamtzeile
    eine echte Kundenposition (beobachtet: Juli-2026, komplette Datei mit
    298,17 EUR echtem Inhalt wurde so faelschlich zu 8.378,34 EUR aufgeblaeht).
    Das im Anschreiben aufgedruckte 'Provisionsanspruch von insgesamt X EUR'
    dient als Kontrollsumme; weicht die extrahierte Summe davon ab, faellt
    die Datei auf manuelle Pruefung zurueck statt falsche Zahlen zu liefern."""
    rows = []
    total_hint = None
    for page_idx, page in enumerate(pdf.pages):
        if len(page.chars) > 0:
            continue
        im = page.to_image(resolution=400).original
        text = pytesseract.image_to_string(im, lang="deu", config="--psm 6")

        m_total = BHV_HAUSBESITZER_TOTAL_RE.search(text)
        if m_total:
            total_hint = parse_amount(m_total.group(1))

        for line in text.splitlines():
            m = BHV_HAUSBESITZER_ROW_RE.match(line.strip())
            if not m:
                continue
            amt = parse_amount(m.group(2))
            if amt is None or abs(amt) > MAX_PLAUSIBLE_AMOUNT:
                continue
            name = re.sub(r"\s+\d{5}\b.*$", "", m.group(4)).strip()
            if not name:
                continue
            rows.append((page_idx, name, amt, line.strip(), "ocr"))
    return rows, total_hint


ALLIANZ_NUM_RE = re.compile(r"^-?\d{1,3}(?:\.\d{3})*,\d{2}$")
ALLIANZ_CONTRACT_RE = re.compile(r"^AS-\d+$")
ALLIANZ_DIGITS_RE = re.compile(r"^\d+$")


def extract_allianz(pdf):
    """Allianz-Buchungsnote: Text pro Seite ist gespiegelt (siehe
    REVERSED_MARKERS/is_mirrored_text), UND die Tabelle hat zwei getrennte
    Betragsspalten "Belastung" (Storno/Ruckbuchung, negativ) und
    "Gutschrift" (Neuprovision, positiv). Beide Spalten koennen fuer
    dieselbe Vertragsnummer im selben Dokument vorkommen (z.B. Storno einer
    Rate UND eine neue Gutschrift fuer denselben Vertrag) - das ist KEIN
    Duplikat, siehe Nutzer-Rueckmeldung. Fruehere Versuche, die beide
    Vorkommen einfach zu addieren, ueberzaehlten deshalb systematisch.

    Wegen der Text-Spiegelung liefert pdfplumber verdrehte Koordinaten: die
    'top'-Achse verhaelt sich wie die horizontale (Spalten-)Achse, 'x0' wie
    die vertikale (Zeilen-)Achse. Die Spaltenzugehoerigkeit eines Betrags
    wird daher ueber seine 'top'-Position relativ zu den Kopfzeilen-Woertern
    'Belastung'/'Gutschrift' bestimmt. Jede Seite traegt zusaetzlich eine
    laufende Summenzeile ("Summe SACH-PROVISIONEN") ohne Vertragsnummer -
    wird ueber das Fehlen einer "AS-..."-Vertragsnummer ausgeschlossen.

    Verifiziert gegen den in jedem Dokument selbst aufgedruckten
    "Sach-Provisionen"-Endbetrag (Deckblatt) - exakte Uebereinstimmung bei
    allen 5 getesteten Juni-2026-Dateien."""
    rows = []
    for page_idx, page in enumerate(pdf.pages):
        words = page.extract_words()
        if not words:
            continue
        rwords = [{"text": w["text"][::-1], "top": w["top"], "x0": w["x0"]} for w in words]
        gut_top = next((w["top"] for w in rwords if w["text"] == "Gutschrift"), None)
        bel_top = next((w["top"] for w in rwords if w["text"] == "Belastung"), None)
        if gut_top is None or bel_top is None:
            continue  # Deckblatt oder sonstige Seite ohne SACH-PROVISIONEN-Tabelle
        boundary = (gut_top + bel_top) / 2

        groups = {}
        for w in rwords:
            key = round(w["x0"], 1)
            groups.setdefault(key, []).append(w)
        lines = [sorted(groups[k], key=lambda w: -w["top"]) for k in sorted(groups.keys())]

        for gi, line in enumerate(lines):
            texts = [w["text"] for w in line]
            contract = next((t for t in texts if ALLIANZ_CONTRACT_RE.match(t)), None)
            if not contract:
                continue
            amount_word = line[-1]
            if not ALLIANZ_NUM_RE.match(amount_word["text"]):
                continue
            sign = 1 if amount_word["top"] < boundary else -1
            amt = sign * parse_amount(amount_word["text"])

            # Kundenname steht in der jeweils naechsten Zeilengruppe, vor
            # dem ersten rein numerischen Token (Datev-Kennziffer, z.B.
            # "8072") - siehe Modul-Docstring.
            name = contract
            if gi + 1 < len(lines):
                name_tokens = []
                for t in (w["text"] for w in lines[gi + 1]):
                    if ALLIANZ_DIGITS_RE.match(t):
                        break
                    name_tokens.append(t)
                if name_tokens:
                    name = " ".join(name_tokens).strip()

            line_text = " ".join(texts)
            rows.append((page_idx, name, round(amt, 2), line_text, "text"))
    return rows


AXA_VN_MARKER = "VN:"


def extract_axa(pdf):
    """AXA-Einzelnachweis zum Kontoauszug: Text pro Seite ist gespiegelt
    (siehe REVERSED_MARKERS/is_mirrored_text), Layout wie bei Allianz mit
    vertauschten Koordinatenachsen (siehe extract_allianz-Docstring).

    Jeder Kundenblock beginnt mit einer Zeile 'VN: Nachname, Vorname ...'
    und kann mehrere Sparten-/Vertragszeilen enthalten; die letzte Zeile
    des Blocks ist eine mit Sternchen maskierte Summenzeile, die auf 'VD'
    endet und den Netto-Provisionsbetrag des ganzen Blocks traegt (Format:
    '<Betrag> <H|S> <Betrag> <H|S> VD' - Provisions- und Abrechnungsbetrag
    sind identisch). Nur diese Summenzeile wird verwendet (nicht die
    einzelnen Sparten-Teilbetraege), sonst wuerde doppelt gezaehlt.

    Auf der letzten Seite folgt zusaetzlich ein 'ALLGEMEINE UMSAETZE
    VD-KONTO'-Abschnitt mit kontobezogenen Buchungen. Darunter sind zwei
    Faelle zu unterscheiden (Nutzer-Bestaetigung anhand eines echten
    Beispiels): eine reine 'Kontoausgleich'-Buchung (z.B. eine Auszahlung/
    Ueberweisung durch AXA) ist tatsaechlich nicht kundenbezogen und wird
    ausgeschlossen - eine 'Umbuchung ... Prov.-Ausgleich'-Buchung dagegen
    IST einem Kunden zuzurechnen, auch wenn ihr (anders als bei einem
    echten Kundenblock) keine 'VN:'-Zeile vorausgeht, sondern eine mit dem
    Namen verschmolzene Referenzzeile ('VN:Name') NACH der Betragszeile
    folgt. Ein solcher noch nicht zugeordneter Betrag wird daher als
    'pending' zwischengespeichert und, falls direkt danach eine
    'VN:Name'-Referenzzeile folgt, diesem Kunden zugerechnet - andernfalls
    (z.B. beim reinen Kontoausgleich ohne jede Namensreferenz) verworfen."""
    rows = []
    for page_idx, page in enumerate(pdf.pages):
        words = page.extract_words()
        if not words:
            continue
        rwords = [{"text": w["text"][::-1], "top": w["top"], "x0": w["x0"]} for w in words]
        groups = {}
        for w in rwords:
            key = round(w["x0"], 1)
            groups.setdefault(key, []).append(w)
        lines = [sorted(groups[k], key=lambda w: -w["top"]) for k in sorted(groups.keys())]

        current_name = None
        pending = None  # (amt, line_text) ohne current_name - evtl. gleich per "VN:Name"-Referenzzeile zuzuordnen
        for line in lines:
            texts = [w["text"] for w in line]
            # Nur der EXAKTE Token "VN:" zaehlt als Start eines echten
            # Kundenblocks.
            if AXA_VN_MARKER in texts:
                vn_idx = texts.index(AXA_VN_MARKER)
                name_toks = texts[vn_idx + 1:vn_idx + 3]
                current_name = " ".join(name_toks).replace(",", "").strip()
                pending = None
                continue
            if len(texts) == 1 and texts[0].startswith(AXA_VN_MARKER):
                # Verschmolzene Referenzzeile "VN:Name" (siehe Docstring) -
                # loest einen zwischengespeicherten "pending"-Betrag auf.
                if pending:
                    ref_name = texts[0][len(AXA_VN_MARKER):].strip()
                    amt, line_text = pending
                    if ref_name:
                        rows.append((page_idx, ref_name, amt, line_text, "text"))
                    pending = None
                continue
            if texts and texts[-1] == "VD" and len(texts) >= 4:
                amt_tok, flag_tok = texts[-3], texts[-2]
                if ALLIANZ_NUM_RE.match(amt_tok) and flag_tok in ("H", "S"):
                    sign = 1 if flag_tok == "H" else -1
                    amt = round(sign * parse_amount(amt_tok), 2)
                    line_text = " ".join(texts)
                    if current_name:
                        rows.append((page_idx, current_name, amt, line_text, "text"))
                        pending = None
                    else:
                        pending = (amt, line_text)
    return rows


GOTHAER_CONTRACT_RE = re.compile(r"^\d[\d.\-/]{5,}$")
GOTHAER_NAME_RE = re.compile(r"^([A-ZÀ-Ü][a-zà-ÿß]+),([A-ZÀ-Ü][a-zà-ÿß]+)")


def extract_gothaer(pdf):
    """Gothaer (Allgemeine + Leben): gespiegelter Text wie bei Allianz/AXA
    (siehe extract_allianz-Docstring fuer die vertauschten Koordinaten).

    Jede Buchungszeile beginnt mit einer Vertrags-/Vers.-Nr. (z.B.
    '38.038.906337' oder '64-818940-60/0000') und endet mit dem C/P-Betrag
    (Courtage/Provision); der zugehoerige Kundenname ('Nachname,Vorname',
    ohne Leerzeichen zusammengeschrieben) steht in der jeweils naechsten
    Zeilengruppe. Verifiziert gegen die auf dem Deckblatt aufgedruckte
    "Endsumme" (Gunsten-Spalte) - exakte Uebereinstimmung bei beiden
    getesteten Juni-2026-Dateien (14,21 EUR / 1,50 EUR)."""
    rows = []
    for page_idx, page in enumerate(pdf.pages):
        words = page.extract_words()
        if not words:
            continue
        rwords = [{"text": w["text"][::-1], "top": w["top"], "x0": w["x0"]} for w in words]
        groups = {}
        for w in rwords:
            key = round(w["x0"], 1)
            groups.setdefault(key, []).append(w)
        keys = sorted(groups.keys())
        lines = [sorted(groups[k], key=lambda w: -w["top"]) for k in keys]

        for gi, line in enumerate(lines):
            texts = [w["text"] for w in line]
            if not texts:
                continue
            if GOTHAER_CONTRACT_RE.match(texts[0]) and ALLIANZ_NUM_RE.match(texts[-1]):
                amt = parse_amount(texts[-1])
                name = None
                if gi + 1 < len(lines):
                    ntexts = [w["text"] for w in lines[gi + 1]]
                    if ntexts:
                        m = GOTHAER_NAME_RE.match(ntexts[0])
                        if m:
                            name = f"{m.group(1)} {m.group(2)}"
                if name:
                    line_text = " ".join(texts)
                    rows.append((page_idx, name, round(amt, 2), line_text, "text"))
    return rows


CONTINENTALE_NAME_RE = re.compile(r"([A-ZÀ-Ü][a-zà-ÿß]+-[A-ZÀ-Ü][a-zà-ÿß]+)")
CONTINENTALE_CLEAN_AMT_RE = re.compile(r"\+(\d{1,3}(?:\.\d{3})*,\d{2})")
CONTINENTALE_FALLBACK_AMT_RE = re.compile(r"[+H]\s*(\d{2,4})\b")
CONTINENTALE_SALDO_RE = re.compile(r"Neuer Saldo:?\s*(-?[\d.,]+)")


def _continentale_candidates(lines):
    """Liefert je Namenszeile (name, sauberer_Betrag_oder_None,
    Rueckfall_Betrag_oder_None, Zeile) - siehe extract_continentale()."""
    out = []
    for line in lines:
        if "P__" not in line and "P___" not in line:
            continue
        m = CONTINENTALE_NAME_RE.search(line)
        if not m:
            continue
        rest = line[m.end():]
        clean = None
        cm = CONTINENTALE_CLEAN_AMT_RE.search(rest)
        if cm:
            clean = parse_amount(cm.group(1))
        fallback = None
        fm = CONTINENTALE_FALLBACK_AMT_RE.search(rest)
        if fm:
            digits = fm.group(1)
            # Bei manchen Zeilen faellt der OCR das Komma weg (z.B. "021"
            # statt "0,21") - die letzten beiden Ziffern sind dann die
            # Nachkommastellen (empirisch verifiziert: ergibt zusammen mit
            # den sauber gelesenen Zeilen exakt den aufgedruckten "Neuer
            # Saldo"-Betrag).
            fallback = parse_amount(f"{digits[:-2] or '0'},{digits[-2:]}")
        out.append((m.group(1), clean, fallback, line.strip()))
    return out


CONTINENTALE_AMT_TOKEN_RE = re.compile(r"^-?\d{1,3}(?:\.\d{3})*,\d{2}$")
CONTINENTALE_PROTOKOLL_RE = re.compile(
    r"^\W*\d{4}\W+\d{5}\s+(-?\d{1,3}(?:\.\d{3})*,\d{2})\W*\s+(.+)$"
)


def _continentale_kontoauszug_gutschriften(page):
    """Summe der Gutschriften der Kontoauszug-Seite ('Nr. BuZi Buchungstext
    Lastschrift Gutschrift'): je Zeile 4-stellige Nr., 5-stellige BuZi,
    Text, Betrag. Die Zeile 'UEBERWEISUNG' (Auszahlung des Vorsaldos) zaehlt
    nicht mit; Betraege in der linken Spalte (Lastschrift) zaehlen negativ.
    Die Spalte wird ueber die x-Position des Betrags bestimmt (Gutschrift
    steht rechts von ca. 84 % der Seitenbreite). Liefert None, wenn keine
    Buchungszeile erkannt wurde."""
    im = page.to_image(resolution=300).original
    data = pytesseract.image_to_data(
        im, lang="deu", config="--psm 6", output_type=pytesseract.Output.DICT)
    lines = {}
    for i, txt in enumerate(data["text"]):
        txt = txt.strip()
        if not txt:
            continue
        key = (data["block_num"][i], data["par_num"][i], data["line_num"][i])
        lines.setdefault(key, []).append((data["left"][i], data["width"][i], txt))
    total, found = 0.0, False
    for words in lines.values():
        words.sort()
        texts = [w[2] for w in words]
        # BuZi-Spalte wird vom OCR oft mit Rahmenzeichen verunreinigt
        # ("1|13301", "/13123") - nur die 4-stellige laufende Nr. am
        # Zeilenanfang und der Betrag am Zeilenende sind verlaesslich.
        if len(texts) < 3 or not re.match(r"^(?:\S{0,4}\s+)?\W*\d{4}\b", " ".join(texts)):
            continue
        if "UEBERWEISUNG" in " ".join(texts).upper().replace("Ü", "UE"):
            continue
        left, width, token = words[-1]
        if not CONTINENTALE_AMT_TOKEN_RE.match(token):
            continue
        value = parse_amount(token)
        if value is None:
            continue
        credit = (left + width / 2.0) / im.width > 0.84
        total += value if credit else -value
        found = True
    return round(total, 2) if found else None


def extract_continentale(pdf):
    """Continentale 'Provisionsnote Einzelergebnisse' (gescannt): sehr
    dichte, mehrspaltige Tabelle mit viel Rahmen-/Trennzeichen-Bildrauschen,
    bei der einzelne Betragswerte je nach OCR-Aufloesung mal sauber
    ("+19,45"), mal mit verlorenem Komma ("+502" statt "+5,02") oder mit
    Buchstaben-Ersatzzeichen ("H021" statt "+0,21", vermutlich ein
    OCR-Fehllesen eines "+"-Symbols) gelesen werden - und das nicht
    konsistent bei derselben Aufloesung fuer alle Zeilen.

    Deshalb: OCR bei mehreren Aufloesungen (400/500/600dpi) durchfuehren,
    je Namenszeile ueber die Reihenfolge (nicht den exakten Text, der pro
    Aufloesung variiert) zuordnen und den saubersten verfuegbaren Treffer
    nehmen (zuerst ein Betrag mit Komma, sonst die Ziffern-Rueckfalllogik).

    Da diese Rueckfalllogik unvermeidbar heuristisch ist, wird das Ergebnis
    IMMER gegen den auf der Kontoauszug-Seite aufgedruckten "Neuer Saldo"-
    Betrag geprueft (siehe process_file()): stimmt die Summe nicht exakt
    ueberein, werden die Zeilen verworfen und die Datei faellt automatisch
    auf manuelle Pruefung zurueck, statt falsche Kundenzuordnungen zu
    riskieren - passend zum Grundsatz "Summe der Einzelpositionen muss immer
    dem Gesamtsaldo entsprechen, im Zweifel Rueckmeldung statt Raten"."""
    rows = []
    row_candidates = []
    all_majority = True
    saldo_total = None
    gutschrift_total = None
    protokoll_rows = []
    for pidx, page in enumerate(pdf.pages):
        if len(page.chars) > 0:
            continue
        quick_text = pytesseract.image_to_string(page.to_image(resolution=200).original, lang="deu")
        if saldo_total is None and "Neuer Saldo" in quick_text:
            m = CONTINENTALE_SALDO_RE.search(quick_text)
            if m:
                saldo_total = parse_amount(m.group(1))
        if gutschrift_total is None and "kontoauszug" in quick_text.lower() and "buzi" in quick_text.lower():
            gutschrift_total = _continentale_kontoauszug_gutschriften(page)
        if "protokoll" in quick_text.lower() and "buzi" in quick_text.lower():
            # Zusatzprotokolle (z.B. "Protokoll EUROPA-Reisekrankenversicherung"):
            # Einzelbuchungen mit Buchungstext "RV-<Nr> <Kunde>", die nicht in
            # der Einzelergebnis-Tabelle stehen, aber im Kontoauszug-Saldo
            # enthalten sind (September-2026: 2,40 EUR Sassmann-H).
            ptext = pytesseract.image_to_string(
                page.to_image(resolution=300).original, lang="deu", config="--psm 6")
            for pline in ptext.splitlines():
                pm = CONTINENTALE_PROTOKOLL_RE.match(pline)
                if not pm:
                    continue
                pval = parse_amount(pm.group(1))
                pname = re.sub(r"^\W*RV-?\d+\s*", "", pm.group(2).strip()).strip(" .|:-")
                if pval is not None and pname:
                    protokoll_rows.append((pidx, pname, round(pval, 2), pline.strip(), "ocr"))
        if "Einzelergebnisse" not in quick_text:
            continue

        per_resolution = []
        for res in (400, 500, 600):
            im = page.to_image(resolution=res).original
            text = pytesseract.image_to_string(im, lang="deu", config="--psm 6")
            per_resolution.append(_continentale_candidates(text.splitlines()))

        n = max((len(c) for c in per_resolution), default=0)
        for i in range(n):
            entries = [c[i] for c in per_resolution if i < len(c)]
            if not entries:
                continue
            name = entries[0][0]
            amt, chosen_line = None, entries[0][3]
            # Mehrheitsentscheid ueber die sauber gelesenen Werte der
            # verschiedenen Aufloesungen (einzelne Aufloesungen verlesen
            # gelegentlich eine Ziffer, z.B. 6,02 statt 5,02 oder 0,24 statt
            # 0,21); bei Gleichstand gewinnt der erste (hoechste Aufloesung).
            clean_votes = [round(c, 2) for _, c, _f, _l in entries if c is not None]
            if clean_votes:
                amt, n_votes = Counter(clean_votes).most_common(1)[0]
                chosen_line = next(l for _, c, _f, l in entries if c is not None and round(c, 2) == amt)
                if n_votes < 2:
                    all_majority = False
            else:
                all_majority = False
            if amt is None:
                for _, _clean, fallback, line in entries:
                    if fallback is not None:
                        amt, chosen_line = fallback, line
                        break
            if amt is not None:
                rows.append((pidx, name, round(amt, 2), chosen_line, "ocr"))
                seen = []
                for _, clean, fallback, _line in entries:
                    for v in (clean, fallback):
                        if v is not None and round(v, 2) not in seen:
                            seen.append(round(v, 2))
                row_candidates.append(seen or [round(amt, 2)])

    # Kontrollsumme: Summe der Gutschriften des Kontoauszugs (je BuZi eine
    # klar gedruckte Zeile) - verlaesslicher als der "Neuer Saldo", der
    # Vorsalden/Ueberweisungen enthaelt. Davon gehen die Protokoll-Buchungen
    # ab, die nicht in der Einzelergebnis-Tabelle stehen. Beide Kontrollen
    # werden akzeptiert (eine einzelne OCR-Fehllesung der Gutschrift-Summe
    # soll nicht zu einer falschen Ablehnung fuehren).
    protokoll_total = round(sum(r[2] for r in protokoll_rows), 2)
    target_total = None
    for control in (gutschrift_total, saldo_total):
        if control is None or not rows:
            continue
        positions_target = round(control - protokoll_total, 2)
        if abs(round(sum(r[2] for r in rows), 2) - positions_target) <= 0.01:
            target_total = control
            break
        if len(rows) <= 10:
            from itertools import product
            hits = [
                combo for combo in product(*row_candidates)
                if abs(round(sum(combo), 2) - positions_target) <= 0.005
            ]
            if len(hits) == 1:
                rows = [(r[0], r[1], hits[0][i], r[3], r[4]) for i, r in enumerate(rows)]
                target_total = control
                break
    if target_total is None:
        target_total = gutschrift_total if gutschrift_total is not None else saldo_total
    if rows:
        rows = rows + protokoll_rows
    return rows, target_total, all_majority


AL_AMOUNT_TOKEN_RE = re.compile(r"-?\d{1,3}(?:\.\d{3})*,\d{2}-?")
AL_CALC_RE = re.compile(
    r"(?P<bsign>[-=]?)(?P<base>\d{1,3}(?:\.\d{3})*,\d{2})\s+(?:[A-Z]{1,3}\s+)?"
    r"(?P<rate>\d{1,2},\d{2})\s+\d{2}\.\d{2}\.\d{4}\s+\d{2}\.\d{2}\.\d{4}"
)
AL_FOOTER_MARKERS = (
    "direktion", "postfach", "rechtsform", "aufsichtsrats", "postbank frankfurt",
    "etwaige zahlungen", "wir nehmen den schutz", "alle positionen dieses",
    "wir bitten um",
)


def _al_commission(line_text):
    """Alte Leipziger: Provision = Bezugsgroesse x Satz (%), kaufmaennisch
    auf 2 Stellen gerundet. Liefert None, wenn die Zeile kein Muster
    'Bezugsgroesse Satz Faelligkeit Buchungsdatum' enthaelt."""
    m = AL_CALC_RE.search(line_text)
    if not m:
        return None
    base = Decimal(m.group("base").replace(".", "").replace(",", "."))
    if m.group("bsign") in ("-", "="):
        base = -base
    rate = Decimal(m.group("rate").replace(",", "."))
    if not (0 < rate <= 100):
        return None
    return float((base * rate / 100).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def extract_alte_leipziger(pages_words):
    """Alte Leipziger (gescannt, OCR): sauber gescannte, saubere Tabelle
    ('Abrechnungskonto (Faellige Verguetungen)'), aber mit zwei Eigenheiten
    gegenueber der generischen Tabellen-Engine:

    1) Der Betrag ('Gutschrift/Belastung') steht auf der ERSTEN Zeile eines
       mehrzeiligen Kundenblocks (zusammen mit der Vertrags-/Versicherungs-
       nummer), der lesbare Kundenname aber oft erst auf der 2./3.
       Folgezeile (Firmenname mit Zeilenumbruch) - umgekehrte Reihenfolge
       zur generischen Engine. Betraege werden daher pro Block gesammelt
       und erst bei 'Zwischensumme' (Blockende) dem inzwischen bekannten
       Namen zugeordnet.
    2) Die Kopfzeile 'Name VN' ist zweiteilig; die generische
       Spaltenbreiten-Erkennung haelt das zweite Wort 'VN' faelschlich fuer
       den Beginn der naechsten Spalte und schneidet lange Firmennamen nach
       dem ersten Wort ab - die Namensspalte wird daher manuell bis zum
       bekannten Beginn der naechsten Spalte ('Ausloeser', x0~128) erweitert.

    Nur Seiten mit "Abrechnungskonto" im Text werden verarbeitet (die
    anderen Seiten sind Deckblatt/Inhaltsverzeichnis/Rechtshinweise ohne
    Buchungsdaten). Verifiziert: Summe je Teilabrechnung entspricht exakt
    dem Betrag der zugehoerigen "Zahlungsausgang"-Zeile (Kontrollsumme im
    PDF selbst) - 7.870,01 EUR fuer die 3 Juni-2026-Dateien zusammen."""
    rows = []
    columns = None
    block_name_parts = []
    block_amounts = []

    def flush(page_idx, source, control=None):
        """block_amounts-Eintraege sind (final, zeilentext, alternative):
        'final' ist der bevorzugte Betrag, 'alternative' der abweichend
        gelesene aufgedruckte Betrag (None, wenn beide uebereinstimmen).
        Mit 'control' (aufgedruckte Zwischensumme des Blocks) wird bei
        Abweichung geprueft, ob ein Umschalten auf die gelesenen Betraege
        die Summe trifft - sonst bleibt es bei den berechneten."""
        nonlocal block_name_parts, block_amounts
        name_full = ""
        for frag in block_name_parts:
            if name_full.endswith("-"):
                name_full += frag
            else:
                name_full = (name_full + " " + frag).strip()
        entries = list(block_amounts)
        if control is not None and entries:
            current = round(sum(e[0] for e in entries), 2)
            conflicts = [i for i, e in enumerate(entries) if e[2] is not None]
            if abs(current - control) > 0.02 and conflicts:
                all_alt = round(sum((e[2] if e[2] is not None else e[0]) for e in entries), 2)
                if abs(all_alt - control) <= 0.02:
                    entries = [(e[2] if e[2] is not None else e[0], e[1], None) for e in entries]
                else:
                    for i in conflicts:
                        trial = [(e[2] if (j == i and e[2] is not None) else e[0]) for j, e in enumerate(entries)]
                        if abs(round(sum(trial), 2) - control) <= 0.02:
                            entries[i] = (entries[i][2], entries[i][1], None)
                            break
        if name_full:
            for amt, ltxt, _alt in entries:
                rows.append((page_idx, name_full, amt, ltxt, source))
        block_name_parts = []
        block_amounts = []

    for page_idx, words, source in pages_words:
        flush(page_idx, source)
        columns = None
        page_text = " ".join(w["text"] for w in words).lower()
        # Folgeseiten eines Auszugs tragen nicht immer die Ueberschrift
        # "Abrechnungskonto" - die Tabellenkopfzeile reicht (beobachtet:
        # September-2026, Seite 2 von 2 mit weiteren Buchungen wurde sonst
        # komplett uebersprungen).
        has_table_header = "herkunftsvermitt" in page_text and "gutschrift" in page_text
        if not words or ("abrechnungskonto" not in page_text and not has_table_header):
            continue
        lines = cluster_lines(words)
        li, n = 0, len(lines)
        while li < n:
            # Kopfzeile ist immer genau 4 Zeilen lang (die 4. Zeile
            # "Herkunftsvermittler (VVNR)" traegt sonst faelschlich zur
            # erkannten Spaltenbreite bei).
            header_found = _header_from_span(lines, li, 4)
            if header_found:
                flush(page_idx, source)
                columns, next_li = header_found
                columns = dict(columns)
                columns["name_range"] = (columns["name_range"][0], 120)
                li = next_li
                continue
            if columns is None:
                li += 1
                continue

            line = lines[li]
            line_text = " ".join(w["text"] for w in line["words"])
            line_lower = line_text.lower()
            if "vvnr" in line_lower or "herkunftsvermittler" in line_lower:
                # Rest-Echo der 4. Kopfzeilen-Zeile, das trotz Kopfzeilen-
                # Erkennung gelegentlich noch als Datenzeile ankommt.
                li += 1
                continue
            if any(marker in line_lower for marker in AL_FOOTER_MARKERS):
                # Briefkopf-/Fusszeilen-Boilerplate (Postfach, Sitz,
                # Aufsichtsrat, Postbank-IBAN ...) - auf Seiten ohne
                # abschliessende "Kontenstand"-Zeile wurden daraus sonst
                # Phantom-Buchungen (z.B. "Postfach 1660" als 1.660,00 EUR).
                flush(page_idx, source)
                break
            if is_blacklisted_line(line_lower):
                if "zwischensumme" in line_lower:
                    control = None
                    for w in reversed(line["words"]):
                        control = parse_amount(w["text"].lstrip("="))
                        if control is not None:
                            break
                    flush(page_idx, source, control)
                elif "kontenstand" in line_lower:
                    # Tabellenende: alles danach ist Fusszeilen-Boilerplate
                    # (IBAN/Postbank/Direktion/...), keine Buchungsdaten mehr.
                    flush(page_idx, source)
                    break
                li += 1
                continue

            # Eine neue Vertrags-/Policen-Nummer (langer Ziffernblock am
            # Zeilenanfang) ODER ein Buchungsart-Schlagwort wie
            # "Folgeinkasso"/"Neuinkasso"/... markiert den Beginn eines NEUEN
            # Kundenblocks - reine Sparten-Folgezeilen DESSELBEN Kunden (z.B.
            # mehrere Elementar-/Feuer-/Leitungswasser-Anteile unter einer
            # Police) tragen weder die Nummer noch dieses Schlagwort erneut,
            # sondern beginnen direkt mit dem Sparten-Namen. Die Policen-
            # nummer allein reicht NICHT als Signal, da sie bei der
            # Worterkennung gelegentlich ganz verloren geht, waehrend das
            # Buchungsart-Wort selbst erhalten bleibt (beobachtet:
            # September-2026 - bei "000031101525 Folgeinkasso Gewerbliche..."
            # fehlte die Policennummer in der erkannten Zeile komplett, nur
            # "Folgeinkasso ..." blieb uebrig). Ohne diese Erkennung haengt
            # sich der Name des naechsten Kunden einfach an den vorherigen
            # an, sobald beide Kundenbloecke vor derselben "Zwischensumme"
            # liegen (beobachtet: "Back, Heinrich" (1 Position) gefolgt von
            # "VIP Baudienstleistungen & Handels GmbH" (8 Sparten-
            # Positionen) wurden faelschlich zu einem Kunden mit beiden
            # Namen zusammengezogen UND alle 9 Betraege diesem einen
            # kombinierten Namen zugerechnet).
            is_new_block_head = bool(
                re.match(r"^\d{6,}\b", line_text)
                or re.match(r"^(Folge|Neu|Erst|Storno|Nach)inkasso\b", line_text, re.IGNORECASE)
            )
            if block_name_parts and is_new_block_head:
                flush(page_idx, source)

            name_l, name_r = columns["name_range"]
            amt_l, amt_r = columns["amount_range"]
            name_words = [w for w in line["words"] if name_l <= w["x0"] < name_r]
            amount_words = [w for w in line["words"] if amt_l <= w["x0"] < amt_r]
            name_text = " ".join(w["text"] for w in name_words).strip(" ,")
            letter_runs = re.findall(r"[A-Za-zÀ-ÿ]+", name_text)
            longest_run = max((len(r) for r in letter_runs), default=0)
            # Aufgedruckter Betrag: nur echte Dezimalzahlen mit Komma
            # (ganze Zahlen wie "1660" aus Briefkopf/Fusszeile sind nie ein
            # Betrag). Ein fuehrendes "=" ist ein OCR-Fehllesen von "-".
            printed = None
            for w in amount_words:
                token = w["text"].lstrip("=")
                if AL_AMOUNT_TOKEN_RE.fullmatch(token):
                    printed = parse_amount(token)
                    if printed is not None:
                        break
            computed = _al_commission(line_text)

            # Die Provision ist immer exakt Bezugsgroesse x Satz: damit
            # lassen sich OCR-Fehllesungen der Betragsspalte erkennen und
            # fehlende Betraege rekonstruieren (beobachtet: "=18,43"/"=-10,18"
            # als Vorzeichen-Fehllesung, ein ganz fehlender Betrag, und
            # "147,73" statt "117,73" bei Bezugsgroesse 470,90 x 25 %).
            alt = None
            if computed is None:
                amount_val = printed
            elif printed is None:
                amount_val = computed
            elif abs(printed - computed) <= 0.02:
                amount_val = printed
            else:
                amount_val, alt = computed, printed

            li += 1
            if longest_run >= 3:
                block_name_parts.append(name_text)
            if amount_val is not None and abs(amount_val) <= MAX_PLAUSIBLE_AMOUNT:
                block_amounts.append((amount_val, line_text, alt))
    return rows


BARMENIA_NAME_RE = re.compile(r"^([A-ZÀ-Ü][A-Za-zà-ÿ\-]+),\s*([A-ZÀ-Ü][A-Za-zà-ÿ\-]+)")
BARMENIA_NUM_RE = re.compile(r"-?\d{1,3}(?:\.\d{3})*,\d{2}")


def extract_barmenia(pdf):
    """Barmenia/Gothaer 'Vergluetungsnachweis Abschlussverguetungen'
    (gescannt): der Standard-OCR-Pfad (300dpi) verliert auf den
    Detailseiten regelmaessig die eigentlichen Betragsspalten (zu dicht/
    klein gedruckte Tabelle) - hoehere Aufloesung (400dpi) mit
    tabellenorientiertem OCR-Modus (--psm 6) als Fliesstext liest sie
    zuverlaessig. Pro Kundenzeile (Muster 'Nachname, Vorname ...') steht
    der Verguetungsbetrag doppelt hintereinander (Verguetungsbetrag =
    Abgerechneter Betrag, sofern kein Storno/keine Proration) - dieses
    Zahlenpaar wird als der gesuchte Betrag genommen. Nur Seiten mit
    'Verguetungsnachweis' im Text enthalten Kundenpositionen (Deckblatt/
    Kontoauszug/Gesamtuebersicht-Seiten werden uebersprungen)."""
    rows = []
    for pidx, page in enumerate(pdf.pages):
        if len(page.chars) > 0:
            continue
        quick_text = pytesseract.image_to_string(page.to_image(resolution=200).original, lang="deu")
        if "vergütungsnachweis" not in quick_text.lower():
            continue
        if "fälliger" in quick_text.lower():
            # 'Gothaer Leben (GL)'-Vorlage - eigene Extraktion, siehe
            # extract_barmenia_abschluss_gl().
            continue
        im = page.to_image(resolution=400).original
        text = pytesseract.image_to_string(im, lang="deu", config="--psm 6")
        for line in text.splitlines():
            m = BARMENIA_NAME_RE.match(line.strip())
            if not m:
                continue
            name = f"{m.group(1)} {m.group(2)}"
            amounts = [parse_amount(a) for a in BARMENIA_NUM_RE.findall(line)]
            amt = None
            for i in range(len(amounts) - 1):
                if amounts[i] == amounts[i + 1] and amounts[i]:
                    amt = amounts[i]
                    break
            if amt is not None:
                rows.append((pidx, name, amt, line.strip(), "ocr"))
    return rows


# Spalten-x-Bereiche (PDF-Punkte) fuer extract_barmenia_laufende() - relativ
# stabil ueber alle "Laufende & Folgeprovisionen"-Seiten desselben Vorlagen-
# Formats: die Kopfzeilen-Woerter "Versicherungs-/nehmer" sitzen konstant bei
# x0~28-46, "Vergütungs-/betrag ... EUR" konstant bei x0~496-617 (siehe
# extract_barmenia_laufende()-Docstring fuer die Herleitung). Nicht dynamisch
# aus der Kopfzeile abgeleitet (waere robuster gegen Layout-Aenderungen in
# kuenftigen Monaten), sondern fest verdrahtet - wie bei VHV_NAME_COLUMN_X0_MIN
# ein bewusster Kompromiss fuer diese eine Vorlage.
BARMENIA_LAUFENDE_NAME_X = (10, 133)
BARMENIA_LAUFENDE_AMOUNT_X = (596, 674)
BARMENIA_LAUFENDE_AMT_ATTEMPTS = (
    (8, 6), (3, 6), (0, 6), (3, 4), (0, 4), (3, 7), (0, 7), (12, 6), (5, 3),
    # Zuletzt: beschnittene Zuschnitte (Zellrahmen stoeren sonst gelegentlich
    # die Ziffernerkennung, siehe BARMENIA_GL_AMT_ATTEMPTS) - nur fuer Zeilen,
    # bei denen alle uebrigen Varianten keinen Betrag lieferten.
    (-10, 11), (-20, 11), (-10, 6),
)
BARMENIA_LAUFENDE_NAME_ATTEMPTS = ((3, 4), (0, 4), (0, 6), (3, 6), (0, 7), (3, 3), (8, 4))
BARMENIA_LAUFENDE_NAME_RE = re.compile(
    r"[A-Za-zÀ-ÿ]{2,}[A-Za-zÀ-ÿ.\-]*,\s*[A-Za-zÀ-ÿ]{2,}[A-Za-zÀ-ÿ.\-]*"
)
BARMENIA_LAUFENDE_AMT_RE = re.compile(r"-?\d{1,3}(?:\.\d{3})*,\d{2}")


def _barmenia_laufende_row_bands(im, resolution, amount_x=None):
    """Erkennt Tabellenzeilen-Grenzen ueber horizontale schwarze Linien im
    Pixelraster der Betragsspalte (siehe extract_barmenia_laufende())."""
    scale = resolution / 72.0
    amount_x = amount_x or BARMENIA_LAUFENDE_AMOUNT_X
    x0, x1 = int(amount_x[0] * scale), int(amount_x[1] * scale)
    arr = np.array(im.convert("L"))
    strip = arr[:, x0:x1]
    dark_frac = (strip < 128).mean(axis=1)
    line_rows = np.where(dark_frac > 0.8)[0]
    lines = []
    if len(line_rows):
        start = prev = line_rows[0]
        for r in line_rows[1:]:
            if r - prev > 3:
                lines.append((start + prev) // 2)
                start = r
            prev = r
        lines.append((start + prev) // 2)
    return [(lines[i], lines[i + 1]) for i in range(len(lines) - 1) if 60 <= lines[i + 1] - lines[i] <= 250]


# 'Gothaer Leben (GL)'-Vorlage der Abschlussverguetungen (Seiten mit Spalte
# "Faelliger Betrag in EUR"): pro Zeile steht der gesuchte Betrag NUR in dieser
# einen Spalte; Verguetungsanspruch (oft 0,00 bei Aenderungen/Storno) und
# Faelliger Betrag unterscheiden sich, der "doppelte Betrag"-Trick aus
# extract_barmenia() greift hier daher nicht.
BARMENIA_GL_NAME_X = (10, 100)
BARMENIA_GL_AMOUNT_X = (641, 696)
# Negative 'pad'-Werte beschneiden den Zuschnitt: die Zellrahmen oben/unten
# stoeren die Ziffernerkennung (beobachtet: mit Rahmen leer, 10-20px
# beschnitten mit --psm 11 sauber '13,50').
BARMENIA_GL_AMT_ATTEMPTS = ((-10, 11), (-20, 11), (-10, 6), (-20, 6), (-30, 11), (0, 11))
BARMENIA_OVERVIEW_ABSCHLUSS_RE = re.compile(
    r"ergebnis\s+abschlussverg\S*\s+(-?\d{1,3}(?:\.\d{3})*,\d{2}-?)", re.IGNORECASE)
BARMENIA_OVERVIEW_LAUFENDE_RE = re.compile(
    r"ergebnis\s+laufende\s+verg\S*\s+(-?\d{1,3}(?:\.\d{3})*,\d{2}-?)", re.IGNORECASE)


def extract_barmenia_abschluss_gl(pdf):
    """Barmenia/Gothaer 'Vergutungsnachweis Abschlussvergutungen' im
    'Gothaer Leben (GL)'-Layout (Spalte 'Faelliger Betrag'): gesamte Zeile
    als Fliesstext liefert unzuverlaessig die Betragsspalten (einzelne Werte
    fehlen), daher wie bei extract_barmenia_laufende() zeilenweise
    Zuschnitt von Namens- und Betragsspalte. Mehrheitsentscheid ueber die
    Zuschnitt-/PSM-Varianten. Wird von process_file() gegen das 'Gesamt-
    ergebnis Abschlussverguetungen' der Gesamtuebersichtsseite geprueft."""
    rows = []
    for pidx, page in enumerate(pdf.pages):
        if len(page.chars) > 0 or page.width <= page.height:
            continue
        quick = pytesseract.image_to_string(page.to_image(resolution=200).original, lang="deu").lower()
        if "abschlussverg" not in quick or "fälliger" not in quick:
            continue
        resolution = 400
        im = page.to_image(resolution=resolution).original
        scale = resolution / 72.0
        name_x0, name_x1 = (int(v * scale) for v in BARMENIA_GL_NAME_X)
        amt_x0, amt_x1 = (int(v * scale) for v in BARMENIA_GL_AMOUNT_X)
        for top, bottom in _barmenia_laufende_row_bands(im, resolution, BARMENIA_GL_AMOUNT_X):
            votes = []
            for pad, psm in BARMENIA_GL_AMT_ATTEMPTS:
                crop = im.crop((amt_x0, max(0, top - pad), amt_x1, bottom + pad))
                text = pytesseract.image_to_string(crop, lang="deu", config=f"--psm {psm}")
                text = re.sub(r"(\d)\s+,\s*(\d{2})", r"\1,\2", text)
                found = BARMENIA_LAUFENDE_AMT_RE.findall(text)
                if found:
                    v = parse_amount(found[-1])
                    if v is not None and abs(v) <= MAX_PLAUSIBLE_AMOUNT:
                        votes.append(round(v, 2))
            if not votes:
                continue
            amt = Counter(votes).most_common(1)[0][0]

            name, raw = None, ""
            for pad, psm in BARMENIA_LAUFENDE_NAME_ATTEMPTS:
                crop = im.crop((name_x0, max(0, top - pad), name_x1, bottom + pad))
                text = pytesseract.image_to_string(crop, lang="deu", config=f"--psm {psm}").strip()
                m = BARMENIA_LAUFENDE_NAME_RE.search(text)
                if m:
                    name, raw = m.group(0).strip(), text
                    break
            if not name:
                continue
            rows.append((pidx, name, amt, raw, "ocr"))
    return rows


def find_barmenia_overview_totals(pdf):
    """Liest 'Gesamtergebnis Abschlussverguetungen' und 'Gesamtergebnis
    Laufende Verguetungen' von der (sauber gescannten) Gesamtuebersicht-
    Seite. Liefert (abschluss, laufende) - je None, wenn nicht gefunden."""
    abschluss = laufende = None
    for page in pdf.pages:
        if len(page.chars) > 0 or page.width > page.height:
            continue
        text = pytesseract.image_to_string(page.to_image(resolution=250).original, lang="deu", config="--psm 6")
        text = re.sub(r"\s+", " ", text)
        m = BARMENIA_OVERVIEW_ABSCHLUSS_RE.search(text)
        if m and abschluss is None:
            abschluss = parse_amount(m.group(1))
        m = BARMENIA_OVERVIEW_LAUFENDE_RE.search(text)
        if m and laufende is None:
            laufende = parse_amount(m.group(1))
    return abschluss, laufende


def extract_barmenia_laufende(pdf):
    """Barmenia 'Verguetungsnachweis Laufende & Folgeprovisionen' (gescannt,
    andere Tabellenvorlage als die Abschlussverguetungen-Seiten in
    extract_barmenia()): eine dicht gedruckte mehrspaltige Tabelle mit
    sichtbaren Zellrahmen. Der Standard-OCR-Pfad (ganze Seite als Fliesstext,
    egal ob --psm 6 oder automatische Wort-Erkennung) liefert dafuer nur
    Bildrauschen - NICHT weil der Scan schlecht waere (eine einzelne isolierte
    Tabellenzelle liest sich einwandfrei), sondern weil Tesseracts Seiten-
    /Spaltenerkennung an der dichten Tabelle scheitert.

    Workaround: die Tabellenzeilen-Grenzen werden direkt aus den Pixeln
    erkannt (durchgehende dunkle horizontale Linien in der Betragsspalte,
    siehe _barmenia_laufende_row_bands()), dann wird JEDE Zeile fuer sich
    zugeschnitten und Name-/Betragsspalte getrennt per OCR gelesen (mehrere
    Zuschnitt-/PSM-Kombinationen werden probiert, bis eine einen gueltigen
    Betrag bzw. ein Name-Komma-Name-Muster liefert). Das behebt das
    Segmentierungsproblem, da jede einzelne Zelle klein/isoliert genug ist.

    Nicht 100% zuverlaessig: manche Zeilen liefern trotzdem keinen gueltigen
    Betrag (werden dann uebersprungen - fehlendes Geld, kein Risiko), und
    einzelne erkannte Namen koennen leicht verlesen sein (z.B. ein Buchstabe
    falsch). Das ist bewusst in Kauf genommen, WEIL match_betreuer() ein
    Sicherheitsnetz bietet: ein verlesener Name matcht in aller Regel keinen
    (falschen) Betreuer.xlsx-Eintrag und landet stattdessen im Blatt
    'Kunde_ohne_Betreuer' zur manuellen Pruefung, statt still einem falschen
    Kunden/Partner zugerechnet zu werden."""
    rows = []
    for pidx, page in enumerate(pdf.pages):
        if len(page.chars) > 0:
            continue
        quick_text = pytesseract.image_to_string(page.to_image(resolution=200).original, lang="deu").lower()
        if "laufende" not in quick_text or "folgeprovision" not in quick_text:
            continue
        resolution = 400
        im = page.to_image(resolution=resolution).original
        scale = resolution / 72.0
        name_x0, name_x1 = (int(v * scale) for v in BARMENIA_LAUFENDE_NAME_X)
        amt_x0, amt_x1 = (int(v * scale) for v in BARMENIA_LAUFENDE_AMOUNT_X)

        for top, bottom in _barmenia_laufende_row_bands(im, resolution):
            amt = None
            for pad, psm in BARMENIA_LAUFENDE_AMT_ATTEMPTS:
                crop = im.crop((amt_x0, max(0, top - pad), amt_x1, bottom + pad))
                text = pytesseract.image_to_string(crop, lang="deu", config=f"--psm {psm}").strip()
                m = BARMENIA_LAUFENDE_AMT_RE.search(text)
                if m:
                    amt = parse_amount(m.group(0))
                    break
            if amt is None or abs(amt) > MAX_PLAUSIBLE_AMOUNT:
                continue

            name = None
            raw_name_text = ""
            for pad, psm in BARMENIA_LAUFENDE_NAME_ATTEMPTS:
                crop = im.crop((name_x0, max(0, top - pad), name_x1, bottom + pad))
                text = pytesseract.image_to_string(crop, lang="deu", config=f"--psm {psm}").strip()
                m = BARMENIA_LAUFENDE_NAME_RE.search(text)
                if m:
                    name = m.group(0).strip()
                    raw_name_text = text
                    break
            if not name:
                continue

            rows.append((pidx, name, amt, raw_name_text, "ocr"))
    return rows


# Wie BARMENIA_LAUFENDE_NAME_X/AMOUNT_X: PDF-Punkte-Bereiche fuer die
# 'Kunde'- und 'Prov./Court'-Spalte des Dialog-'Provisionseinzelnachweis'
# (siehe extract_dialog()), per Sichtpruefung ermittelt.
DIALOG_NAME_X = (133, 227)
DIALOG_AMOUNT_X = (693, 774)
DIALOG_AMT_RE = re.compile(r"-?\d{1,3}(?:\.\d{3})*,\d{2}-?")
DIALOG_NAME_RE = re.compile(r"[A-Za-zÀ-ÿ][A-Za-zÀ-ÿ.&\- ]{2,}")
DIALOG_AMT_ATTEMPTS = ((0, 6), (5, 6), (0, 11), (5, 11), (0, 4), (5, 4), (0, 7), (10, 6))
DIALOG_NAME_ATTEMPTS = ((0, 6), (5, 6), (0, 4), (5, 4), (0, 7), (10, 6))


DIALOG_ORIENT_KEYWORDS = (
    "vertragsnummer", "kunde", "zahlweise", "fälligkeit", "falligkeit",
    "provisionseinzel", "produkt", "geschäftspartner", "geschaftspartner",
)


def _dialog_page_is_upside_down(page):
    """Manche Dialog-Detailseiten sind um 180 Grad gedreht eingescannt
    (beobachtet: September-2026, Seite 2 von 2 stand kopf) - die Zeilen-
    erkennung selbst funktioniert gedreht genauso, aber die OCR liefert nur
    Rauschen. Entscheidung ueber die Zahl erkannter Kopfzeilen-Schluessel-
    woerter in beiden Ausrichtungen (niedrige Aufloesung genuegt)."""
    scores = []
    for rotate in (False, True):
        im = page.to_image(resolution=150).original.convert("L")
        if rotate:
            im = im.rotate(180)
        im = im.filter(ImageFilter.MedianFilter(3))
        text = pytesseract.image_to_string(im, lang="deu", config="--psm 6").lower()
        scores.append(sum(text.count(k) for k in DIALOG_ORIENT_KEYWORDS))
    return scores[1] > scores[0]


def extract_dialog(pages):
    """Dialog 'Provisionseinzelnachweis' (gescannt): pro Kunde ein umrahmter
    Tabellenblock (Name ueber 1-2 Zeilen; je Vertragszeile stehen Prov.Satz
    (Prozent) und Prov.-Betrag (Euro) uebereinander in derselben Zelle; hat
    ein Kunde mehrere Vertragszeilen, schliesst eine 'Summe:'-Zeile den
    Block ab) mit gepunktetem Wasserzeichen-Hintergrund im gesamten
    Tabellenbereich - dieser Hintergrund verwirrt sowohl Tesseracts Standard-
    Volltext-OCR (--psm 6) als auch die wortpositionsbasierte OCR fast
    vollstaendig, obwohl der Scan selbst gestochen scharf ist.

    Workaround wie bei extract_barmenia_laufende(): Kundenbloecke ueber die
    horizontalen schwarzen Rahmenlinien im Pixelraster der Betragsspalte
    finden, dann Name-/Betragsspalte je Block einzeln zuschneiden. Zusaetzlich
    noetig (im Unterschied zu Barmenia): ein Median-Filter auf dem Graustufen-
    bild VOR der OCR entfernt das feine Punktraster, ohne die fetten Text-
    Konturen zu zerstoeren.

    WICHTIG: ein Block = ein Kunde = ein Betrag, naemlich der LETZTE erkannte
    Dezimalwert im Block (bei mehreren Vertragszeilen die 'Summe:'-Zeile am
    Blockende, bei einer Zeile der Prov.-Betrag unter dem Prozentsatz).
    Negative Betraege stehen mit nachgestelltem Minus ("102,72-"). Ein
    frueherer Versuch, hohe Bloecke in gleich grosse Teilzeilen zu zerlegen,
    hat die einzelnen Vertragszeilen UND die Summenzeile desselben Kunden
    gelesen und damit doppelt gezaehlt - verworfen.

    Die Detailtabellen-Seite(n) werden NICHT per Text-/OCR-Inhalt erkannt,
    sondern am Seitenformat: die Tabelle steht quer (Breite > Hoehe),
    waehrend Deckblatt/Uebersichtsseiten hochkant sind. Kopfstehende Seiten
    werden vorab erkannt und gedreht (siehe _dialog_page_is_upside_down())."""
    rows = []
    for pidx, page in enumerate(pages):
        if len(page.chars) > 0 or page.width <= page.height:
            continue
        resolution = 400
        im = page.to_image(resolution=resolution).original
        if _dialog_page_is_upside_down(page):
            im = im.rotate(180)
        scale = resolution / 72.0
        name_x0, name_x1 = (int(v * scale) for v in DIALOG_NAME_X)
        amt_x0, amt_x1 = (int(v * scale) for v in DIALOG_AMOUNT_X)

        arr = np.array(im.convert("L"))
        strip = arr[:, amt_x0:amt_x1]
        dark_frac = (strip < 128).mean(axis=1)
        line_rows = np.where(dark_frac > 0.7)[0]
        lines = []
        if len(line_rows):
            start = prev = line_rows[0]
            for r in line_rows[1:]:
                if r - prev > 3:
                    lines.append((start + prev) // 2)
                    start = r
                prev = r
            lines.append((start + prev) // 2)

        bands = [
            (lines[i], lines[i + 1])
            for i in range(len(lines) - 1)
            if lines[i + 1] - lines[i] >= 150
        ]

        for b_top, b_bottom in bands:
            # Mehrheitsentscheid ueber alle Zuschnitt-/PSM-Varianten statt
            # "erster Treffer gewinnt": einzelne Varianten verlesen gelegentlich
            # eine Ziffer (beobachtet: 38,94 als 38,99), die uebrigen Varianten
            # liefern meist den richtigen Wert.
            candidates = []
            for pad, psm in DIALOG_AMT_ATTEMPTS:
                amt_crop = im.crop((amt_x0, max(0, b_top - pad), amt_x1, b_bottom + pad)) \
                    .convert("L").filter(ImageFilter.MedianFilter(7))
                amt_text = pytesseract.image_to_string(amt_crop, lang="deu", config=f"--psm {psm}")
                # OCR setzt gelegentlich ein Leerzeichen vor das Komma ("99 ,95").
                amt_text = re.sub(r"(\d)\s+,\s*(\d{2})", r"\1,\2", amt_text)
                amt_matches = DIALOG_AMT_RE.findall(amt_text)
                if amt_matches:
                    v = parse_amount(amt_matches[-1])
                    if v is not None and abs(v) <= MAX_PLAUSIBLE_AMOUNT:
                        candidates.append(round(v, 2))
            if not candidates:
                continue
            amt = Counter(candidates).most_common(1)[0][0]

            name = None
            raw_name_text = ""
            for pad, psm in DIALOG_NAME_ATTEMPTS:
                name_crop = im.crop((name_x0, max(0, b_top - pad), name_x1, b_bottom + pad)) \
                    .convert("L").filter(ImageFilter.MedianFilter(7))
                name_text = pytesseract.image_to_string(name_crop, lang="deu", config=f"--psm {psm}")
                m = DIALOG_NAME_RE.search(name_text)
                if m:
                    name = m.group(0).strip()
                    raw_name_text = name_text.strip()
                    break
            if not name:
                continue

            rows.append((pidx, name, amt, raw_name_text, "ocr"))

    # Der unterste Rahmenblock der letzten Seite enthaelt die Gesamtsumme
    # ("Summe DIV"/"Gesamt") und ist kein Kunde - er faellt dadurch auf, dass
    # sein Betrag genau der Summe aller vorherigen Bloecke entspricht.
    while len(rows) >= 2 and abs(rows[-1][2] - round(sum(r[2] for r in rows[:-1]), 2)) <= 0.005:
        rows = rows[:-1]
    return rows


def repair_signs_to_control(rows, control, max_rows=16):
    """Wenn die Summe der gelesenen Betraege nicht der aufgedruckten
    Kontrollsumme entspricht, aber sich durch Umdrehen des Vorzeichens
    einzelner Zeilen EXAKT treffen laesst (typisch: OCR verliert das kleine
    nachgestellte Minus einer Storno-Summe), wird das kleinste eindeutige
    Umdrehen angewendet. Mehrdeutig (mehrere gleich kleine Loesungen) oder
    nicht lösbar -> rows unveraendert zurueck. Der Aufrufer prueft danach
    ohnehin erneut gegen die Kontrollsumme."""
    if control is None or not rows or len(rows) > max_rows:
        return rows
    total = round(sum(r[2] for r in rows), 2)
    if abs(total - control) <= 0.01:
        return rows
    # Umdrehen von Teilmenge S aendert die Summe um -2*sum(S).
    target = round((total - control) / 2.0, 2)
    amounts = [r[2] for r in rows]
    n = len(amounts)
    best_size, solutions = None, []
    for mask in range(1, 1 << n):
        s = sum(amounts[i] for i in range(n) if mask >> i & 1)
        if abs(round(s, 2) - target) <= 0.005:
            size = bin(mask).count("1")
            if best_size is None or size < best_size:
                best_size, solutions = size, [mask]
            elif size == best_size:
                solutions.append(mask)
    if best_size is None or len(solutions) != 1:
        return rows
    mask = solutions[0]
    return [
        (r[0], r[1], -r[2] if (mask >> i & 1) else r[2], r[3], r[4])
        for i, r in enumerate(rows)
    ]


VEMA_CSV_AMOUNT_COL = "Betrag"


def extract_vema_csv(path):
    """VEMA-Pool liefert die Kundenpositionen NICHT im PDF (das ist ein
    reiner Sammelbeleg, siehe README), sondern separat als CSV-Export
    ('VEMA-Poolabrechnung-N.csv'). Spalten (Semikolon-getrennt):
    VEMAintern;VN Anrede;VN Vorname;VN Nachname;Vertragsnummer;
    Gesellschaft;Sparte;Faelligkeit;Provisionsbasis;Courtage;
    Provisionsart;Einbehalt;Betrag;...

    'Courtage' ist der Brutto-Courtagebetrag, 'Betrag' der Betrag NACH
    Abzug des VEMA-Poolbeitrags (siehe Spalte 'Einbehalt', typischerweise
    10%). Anders als z.B. bei Fondsfinanz's Stornoreserve (temporaere
    Sicherheit, siehe extract_fondsfinanz) ist der VEMA-Einbehalt ein
    dauerhafter Pool-Verwaltungsbeitrag, der SSH nie zufliesst - massgeblich
    ist daher 'Betrag' (netto), nicht 'Courtage' (brutto)."""
    rows = []
    with open(path, encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f, delimiter=";")
        if VEMA_CSV_AMOUNT_COL not in (reader.fieldnames or []):
            return rows
        for row in reader:
            anrede = (row.get("VN Anrede") or "").strip()
            vorname = (row.get("VN Vorname") or "").strip()
            nachname = (row.get("VN Nachname") or "").strip()
            if anrede == "Firma" or not vorname:
                name = nachname
            else:
                name = f"{nachname}, {vorname}"
            betrag_s = (row.get(VEMA_CSV_AMOUNT_COL) or "").strip()
            if not betrag_s or not name:
                continue
            betrag = parse_amount(betrag_s)
            if betrag is None:
                continue
            raw_line = ";".join(f"{k}={v}" for k, v in row.items() if v)
            rows.append((0, name, round(betrag, 2), raw_line, "csv"))
    return rows


COVOMO_CSV_AMOUNT_COL = "Provision"


def extract_covomo_csv(path):
    """Covomo (Vergleichsplattform, hier beobachtet fuer eine Reise-
    Krankenversicherung ueber HanseMerkur, Juli-2026): eigenes CSV-Format,
    Semikolon-getrennt, u.a. Spalten 'Provision', 'KundenVorname',
    'KundenNachname'. Wurde vorher stillschweigend als vermeintliche
    VEMA-Pool-CSV mit falschem Spaltenformat verworfen (process_csv_file()
    kannte nur extract_vema_csv), wodurch die Datei komplett unter den
    Tisch fiel statt in 'Sammelbelege_ohne_Details' auf einen Sonderfall
    hinzuweisen. Nur eine Zeile bisher beobachtet - Format kann sich mit
    mehr Beispielen noch aendern."""
    rows = []
    with open(path, encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f, delimiter=";")
        if COVOMO_CSV_AMOUNT_COL not in (reader.fieldnames or []):
            return rows
        for row in reader:
            vorname = (row.get("KundenVorname") or "").strip()
            nachname = (row.get("KundenNachname") or "").strip()
            name = f"{nachname}, {vorname}" if vorname else nachname
            betrag = parse_amount((row.get(COVOMO_CSV_AMOUNT_COL) or "").strip())
            if betrag is None or not name:
                continue
            raw_line = ";".join(f"{k}={v}" for k, v in row.items() if v)
            rows.append((0, name, round(betrag, 2), raw_line, "csv"))
    return rows


BALOISE_DATE_RE = re.compile(r"\d{2}\.\d{2}\.\d{4}")
BALOISE_AMT_RE = re.compile(r"-?\d{1,3}(?:\.\d{3})*,\d{2}[-+]?")
BALOISE_CONTRACT_RE = re.compile(
    r"(?:^|\W)[A-Z]{1,3}\W{0,4}((?=[A-Z0-9]*\d)[A-Z0-9]{8,12})\s+(.+?)\s+\d{2}\.\d{2}\.\d{4}"
)
BALOISE_BELEG_RE = re.compile(r"^\W*\d{2}\.\d{2}\W+\d{8,10}\b")
BALOISE_OCR_VARIANTS = ((200, 5), (300, 5), (200, 3), (300, 3))


def _baloise_signed(token):
    """'34,50-' -> -34.5, '0,00+' -> 0.0, '4,89' -> 4.89."""
    neg = token.endswith("-") or token.startswith("-")
    val = parse_amount(token.strip("+-"))
    if val is None:
        return None
    return -abs(val) if neg else abs(val)


def _baloise_parse_text(text):
    """Zerlegt den OCR-Text einer Baloise-'Kontoauszug'-Seite in
    (Detailzeilen, Auszahlung, Saldo-Vortrag, End-Saldo). Detailzeilen sind
    alle Zeilen mit Faelligkeitsdatum; der Kundenname steht nur in der ERSTEN
    Zeile einer Gruppe (Vertragsnummer + Name), die folgenden Zeilen
    (nur Sparten-Kuerzel wie F/ED/LW/ST) gehoeren zum zuletzt genannten Kunden."""
    rows, payout, vortrag, endsaldo = [], None, 0.0, 0.0
    current = None
    for raw in text.splitlines():
        line = re.sub(r"(\d)\s+,(\d{2})", r"\1,\2", raw.strip())
        if not line:
            continue
        low = line.lower()
        amounts = BALOISE_AMT_RE.findall(line)
        if "saldo-vortrag" in low or "saldo vortrag" in low:
            if amounts:
                vortrag = _baloise_signed(amounts[-1]) or 0.0
            continue
        if "end-saldo" in low or "endsaldo" in low:
            if amounts:
                endsaldo = _baloise_signed(amounts[-1]) or 0.0
            continue
        if not BALOISE_DATE_RE.search(line):
            if ("prov" in low and "zahl" in low) and amounts:
                v = _baloise_signed(amounts[-1])
                if v is not None:
                    payout = abs(v)
            continue
        if not amounts:
            continue
        m = BALOISE_CONTRACT_RE.search(line)
        if m:
            name = re.sub(r"[|_]+", " ", m.group(2))
            name = re.sub(r"(?:^|\s)[.:,;-]+(?=\s|$)", " ", name)
            name = re.sub(r"\s+", " ", name).strip(" .:,;|-")
            if len(re.findall(r"[A-Za-zÀ-ÿ]", name)) >= 3:
                current = name
            else:
                current = f"Baloise Vertrag {m.group(1)} (Name nicht lesbar)"
        elif BALOISE_BELEG_RE.match(line):
            # Neue Buchungsgruppe (Belegnummer am Zeilenanfang), deren
            # Vertrags-/Namensteil nicht erkannt wurde: NICHT dem vorherigen
            # Kunden zuschlagen, sondern als eigene, markierte Gruppe fuehren.
            current = "Baloise (Name nicht lesbar)"
        if current is None:
            continue
        value = _baloise_signed(amounts[-1])
        if value is None or abs(value) > MAX_PLAUSIBLE_AMOUNT:
            continue
        rows.append((current, round(value, 2), line))
    return rows, payout, vortrag, endsaldo


def extract_baloise(pdf):
    """Baloise (frueher Basler) 'Kontoauszug' (gescannt, Querformat, mit
    Wasserzeichen-Punktraster auf dem Tabellenbereich): der Standard-OCR-Pfad
    (300dpi, ohne Filter) liefert dort nur Bruchstuecke und die Datei wurde
    faelschlich als 'Sammelbeleg ohne Details' eingestuft. Mit einem
    Median-Filter (wie bei Dialog) und --psm 6 wird die Tabelle fast
    fehlerfrei gelesen. Je Seite werden mehrere Aufloesungs-/Filter-
    Varianten probiert; gewaehlt wird die, deren Detailsumme zur ausgezahlten
    Summe ('Prov-Zahlung' Buchungsart 420) passt, sonst die mit den meisten
    Zeilen. Rueckgabe: (rows, payout_abs, vortrag, endsaldo)."""
    all_rows = []
    payout_total, vortrag_total, endsaldo_total = None, 0.0, 0.0
    for pidx, page in enumerate(pdf.pages):
        if len(page.chars) > 0 or page.width <= page.height:
            continue
        best = None
        for dpi, med in BALOISE_OCR_VARIANTS:
            im = page.to_image(resolution=dpi).original.convert("L").filter(ImageFilter.MedianFilter(med))
            text = pytesseract.image_to_string(im, lang="deu", config="--psm 6")
            rows, payout, vortrag, endsaldo = _baloise_parse_text(text)
            if not rows:
                continue
            score = (0 if payout is None else (1 if abs(round(sum(r[1] for r in rows), 2) - payout) <= 0.01 else 0), len(rows))
            if best is None or score > best[0]:
                best = (score, rows, payout, vortrag, endsaldo)
            if score[0] == 1:
                break
        if best is None:
            continue
        _, rows, payout, vortrag, endsaldo = best
        for name, value, line in rows:
            all_rows.append((pidx, name, value, line, "ocr"))
        if payout is not None:
            payout_total = (payout_total or 0.0) + payout
        vortrag_total += vortrag
        endsaldo_total += endsaldo
    return all_rows, payout_total, vortrag_total, endsaldo_total


HELVETIA_ROW_RE = re.compile(
    r"^\W*(\d{4})\W+(\d{3}[.,]\d{3}[.,]\d{7}[.,]\d)\W+(.+?)\s+\d{2}\.\d{2}\.\d{4}"
)
HELVETIA_AMT_RE = re.compile(r"-?\d{1,3}(?:\.\d{3})*,\d{2}-?")
HELVETIA_OCR_VARIANTS = ((250, 0), (300, 0), (300, 3), (200, 3))


def _helvetia_parse_text(text):
    """Zeilen 'lfd.Nr | Versicherungsschein-Nr | Name von bis ... Vergütung':
    der gesuchte Betrag ist die LETZTE Dezimalzahl der Zeile (Verguetung,
    bei Storno negativ). Kontrollsumme je Seite: 'Vermittlerinkasso -
    Verguetung = Saldo EUR - Summe: X' bzw. 'EUR - Summe: <Praemie> <...>
    <Verguetung>' (letzte Zahl)."""
    rows, control = [], None
    for raw in text.splitlines():
        line = re.sub(r"(\d)\s+,(\d{2})", r"\1,\2", raw.strip())
        if not line:
            continue
        low = line.lower()
        if "summe" in low and ("saldo" in low or low.startswith("eur") or " eur" in low):
            nums = HELVETIA_AMT_RE.findall(line)
            if nums and "vermittlerinkasso" in low and "saldo" in low:
                control = parse_amount(nums[-1])
            elif nums and control is None and "teilsumme" not in low and "vermittlerinkasso" not in low:
                control = parse_amount(nums[-1])
            continue
        m = HELVETIA_ROW_RE.match(line)
        if not m:
            continue
        policy = m.group(2).replace(",", ".")
        name = re.sub(r"[|{}\[\]_]+", " ", m.group(3))
        name = re.sub(r"\s+", " ", name).strip(" .:,;-")
        nums = HELVETIA_AMT_RE.findall(line[m.end():])
        if not nums or len(re.findall(r"[A-Za-zÀ-ÿ]", name)) < 3:
            continue
        value = parse_amount(nums[-1])
        if value is None or abs(value) > MAX_PLAUSIBLE_AMOUNT:
            continue
        rows.append((f"{policy} | {name}", round(value, 2), line))
    return rows, control


def extract_helvetia(pdf):
    """Helvetia 'Vermittlerliste Komposit' (gescannt, Querformat): je Seite
    eine Liste (Listennummer 000NN/JJJJ) mit Positionen. Der generische
    Parser hat bei neu gescannten Seiten die Tabellenkopfzeile als
    'Kunde' gelesen (Name 'schein Nr.', Betrag 13,36 statt 9,85) - hier wird
    jede Zeile ueber das Muster 'Nr | Versicherungsscheinnummer | Name | von
    | bis' erkannt und der Betrag als letzte Zahl der Zeile gelesen. Pro
    Seite wird die gedruckte 'Summe' (Vermittlerinkasso - Verguetung =
    Saldo) als Kontrolle benutzt; mehrere OCR-Varianten werden probiert,
    bis die Summe der Zeilen trifft. Rueckgabe: (rows, sum_of_page_controls,
    all_pages_ok)."""
    all_rows, control_total, all_ok = [], 0.0, True
    for pidx, page in enumerate(pdf.pages):
        if len(page.chars) > 0 or page.width <= page.height:
            continue
        best = None
        for dpi, med in HELVETIA_OCR_VARIANTS:
            im = page.to_image(resolution=dpi).original.convert("L")
            if med:
                im = im.filter(ImageFilter.MedianFilter(med))
            text = pytesseract.image_to_string(im, lang="deu", config="--psm 6")
            rows, control = _helvetia_parse_text(text)
            if not rows:
                continue
            ok = control is not None and abs(round(sum(r[1] for r in rows), 2) - control) <= 0.02
            score = (1 if ok else 0, len(rows))
            if best is None or score > best[0]:
                best = (score, rows, control, ok)
            if ok:
                break
        if best is None:
            continue
        _, rows, control, ok = best
        all_ok = all_ok and ok
        control_total += control or 0.0
        for name, value, line in rows:
            all_rows.append((pidx, name, value, line, "ocr"))
    return all_rows, round(control_total, 2), all_ok


def find_total_in_text(full_text):
    """Sucht best-effort einen Gesamt-/Endbetrag im Freitext (fuer
    Sammelbelege ohne Kundendetail, zur Kontrolle)."""
    candidates = []
    for pat in TOTAL_LINE_PATTERNS:
        for m in re.finditer(pat, full_text, re.IGNORECASE | re.MULTILINE):
            val = parse_amount(m.group(1))
            if val is not None:
                candidates.append(val)
    if candidates:
        return candidates[-1]
    return None


def is_mirrored_text(full_text):
    return any(marker in full_text for marker in REVERSED_MARKERS)


# ---------------------------------------------------------------------------
# Kontoauszug-Abgleich (VR Bank Rhein-Neckar): erkennt Zahlungseingaenge ohne
# zugehoerige Abrechnungs-PDF fuer den Monat.
# ---------------------------------------------------------------------------

BANK_TX_START_RE = re.compile(
    r"^(\d{2}\.\d{2})\.\s+(\d{2}\.\d{2})\.\s+(.+?)\s+PN:\d+\s+(-?[\d.,]+)\s+([HS])\s*$"
)
# Wiederkehrende Brief-/Fusszeilen (Bankname, Adresse, Kontostand-Zusammen-
# fassung, Seitenumbruch-Hinweise) - koennen an beliebiger Stelle in der
# Zeile stehen (z.B. "K00062862 Bitte beachten Sie ..."), daher ohne
# Zeilenanfang-Anker.
BANK_BOILERPLATE_RE = re.compile(
    r"(VR Bank Rhein-Neckar|www\.vrbank\.de|Kontokorrent|EUR-Konto Kontonummer|"
    r"Postfach \d|neuer Kontostand|Gesamtumsatz:|Ihr Kreditrahmen|\*{5,}|"
    r"Augustaanlage|Ihr Berater:|Telefon: 0|E-Mail:|Bu-Tag Wert Vorgang|"
    r"alter Kontostand|Übertrag (auf|von) Blatt|Bitte beachten Sie|"
    r"Kontoauszug Nr\.|erstellt am)"
)
# Zeilen, die NUR aus einem dieser Fragmente bestehen (Absender-/Empfaenger-
# Adressblock von SSH selbst, wiederholt sich auf jeder Seite) - werden nur
# bei exakter Volltreffer-Gleichheit gefiltert, damit ein echter Absendername
# mit z.B. "GmbH & Co.KG" als Namensbestandteil nicht faelschlich verworfen
# wird.
BANK_BOILERPLATE_WHOLE_LINE_RE = re.compile(
    r"^(K\d{8}|\d{4}|000|0522|SSH Versicherungsmakler|GmbH & Co\.KG|"
    r"Anne-Frank-Str\. 8|\d{5} Viernheim)$"
)

# Bekannte Absender-/Verwendungszweck-Stichworte je Versicherer-"Familie"
# (normalisiert: klein geschrieben, Umlaute uebersetzt, alles ausser a-z0-9
# entfernt - siehe normalize_for_match()). Ein Versicherer-Dateiname wird
# einer Familie zugeordnet, wenn sein normalisierter Name mit dem
# normalisierten Familienschluessel beginnt (siehe insurer_bank_family()) -
# so fallen z.B. "R+V-VM-076965" und "R+V-VM-214574" beide unter "r+v".
# Diese Liste ist zwangslaeufig unvollstaendig (neue Versicherer, geaenderte
# Bank-Bezeichnungen) - unbekannte Zahlungseingaenge werden deshalb nicht
# stillschweigend ignoriert, sondern explizit als "nicht zugeordnet"
# ausgewiesen (siehe reconcile_bank_credits()).
INSURER_BANK_ALIASES = {
    "allianz": ["allianz"],
    "axa": ["axa"],
    "gothaer-allgemeine": ["gothaerallgemeine"],
    "gothaer-leben": ["gothaerleben", "gothaerlebensversicherung"],
    "fondsfinanz": ["fondsfinanz"],
    "r+v": ["rvallgemeine", "zentralesdirektinkasso"],
    "mannheimer": ["mannheimer"],
    "alte-leipziger": ["alteleipziger"],
    "vema-pool": ["vemaversicherungsmaklergenossenschaft"],
    "wuerttembergische": ["wuerttversicherg", "wuerttembergische"],
    "hiscox": ["hiscox"],
    "deurag": ["deurag"],
    "signal-iduna": ["signaliduna"],
    "hdi-leben": ["hdileben"],
    "hdi": ["hdiversicherung"],
    "swiss-life": ["swisslife"],
    "vhv": ["vhv"],
    "itzehoer": ["itzehoerversvereinag", "itzehoer"],
    "markel": ["markel"],
    "haftpflichtkasse": ["haftpflichtkasse"],
    "herzenssache": ["herzenssache"],
    "sv-sparkassenversicherung": ["svholding", "svsparkassen"],
    "aig": ["aigeurope"],
    "amex-pool": ["amex", "qualitypool"],
    "interrisk": ["interrisk"],
    "dialog": ["dialogversicherung"],
    "arag": ["aragse"],
    "auxilia": ["kraftfahrerschutz"],
    "concordia": ["concordia"],
    "bsg": ["barmenia"],
    # Dateiname "Konzept-u-Marketing-RH/-TS", Bank-Absender "Konzept und
    # Marketing GmbH" - ohne Alias fiel die Zahlung als "fehlende
    # Abrechnung" auf, obwohl beide Dateien vorlagen.
    "konzept-u-marketing": ["konzeptundmarketing"],
    "bayerische": ["bayerische"],
}


def normalize_for_match(s):
    """Wie normalize(), zusaetzlich werden alle Nicht-alphanumerischen
    Zeichen entfernt (Leerzeichen, Punkte, Bindestriche, '+', ...) - macht
    Versicherer-Dateinamen und Bank-Freitext direkt vergleichbar."""
    return re.sub(r"[^a-z0-9]", "", normalize(s))


def insurer_bank_family(insurer_key):
    """Ordnet einen Versicherer-Dateinamen (z.B. 'R+V-VM-076965') seiner
    Bank-Abgleichs-'Familie' zu (z.B. 'r+v'). Unbekannte Versicherer liefern
    ihren eigenen normalisierten Namen zurueck (kein Fehler - wird beim
    Abgleich dann einfach nicht in INSURER_BANK_ALIASES gefunden)."""
    key_norm = normalize_for_match(insurer_key)
    for family in sorted(INSURER_BANK_ALIASES, key=len, reverse=True):
        if key_norm.startswith(normalize_for_match(family)):
            return family
    return key_norm


# ---------------------------------------------------------------------------
# Kunde -> Betreuer-Zuordnung (CRM-Export "Betreuer.xlsx")
# ---------------------------------------------------------------------------

# Andreas Selle hat selbst keinen CRM-Account - seine Kunden werden dort
# unter "Calvin Keinarth" gefuehrt (Nutzer-Angabe). Umsatz wird trotzdem
# Andreas Selle zugerechnet, nicht Calvin.
BETREUER_ALIASES = {"calvin keinarth": "Andreas Selle"}
PARTNER_NAMES = ["Robin Heckmann", "Tim Selle", "Andreas Selle"]
# Farbcodes je Partner fuer die farbige Markierung in Excel/PDF (Nutzer-
# Vorgabe: Rot=RH, Gelb=TS, Blau=AS).
PARTNER_COLORS_XLSX = {
    "Robin Heckmann": "FFFFC7CE",
    "Tim Selle": "FFFFEB9C",
    "Andreas Selle": "FFBDD7EE",
}
PARTNER_COLORS_PDF = {
    "Robin Heckmann": (200, 30, 30),
    "Tim Selle": (170, 140, 0),
    "Andreas Selle": (30, 80, 170),
}
# Kraeftigere Variante nur fuer die kleinen Legenden-Farbkaestchen (dort
# spielt Lesbarkeit von Text keine Rolle, ein sattes Gelb ist als "Gelb"
# aber deutlich eindeutiger erkennbar als das gedeckte PARTNER_COLORS_PDF-Gelb,
# das auf weissem Hintergrund lesbar bleiben muss).
PARTNER_SWATCH_COLORS_PDF = {
    "Robin Heckmann": (230, 50, 50),
    "Tim Selle": (255, 200, 0),
    "Andreas Selle": (50, 110, 220),
}

# Manuelle Korrekturen fuer Kunden, die match_betreuer() ueber die
# Betreuer.xlsx nicht (mehr) zuverlaessig finden kann - z.B. weil der Name
# im Quell-PDF stark abgekuerzt/abweichend ist, weil es sich um einen
# Handelsnamen statt des CRM-Firmennamens handelt, oder weil sich der
# Nachname zwischenzeitlich geaendert hat (Heirat) und daher keine noch so
# lockere Schreibvarianz je zum CRM-Namen passen wuerde. Vom Nutzer per
# Hand bestaetigt (Stand Juli 2026) - Key ist der ORIGINALE, im PDF
# auftauchende Kundenname (wird ueber normalize_for_match() abgeglichen,
# siehe apply_betreuer()). Bei neuen Faellen einfach ergaenzen; langfristig
# sauberer ist es, die Ursache direkt in der Betreuer.xlsx zu pflegen
# (z.B. Namensaenderung/Status "inaktiv" bei Rosen Adrian).
MANUAL_BETREUER_OVERRIDES_RAW = {
    "HVP": "Robin Heckmann",  # High Voltage Projects GmbH
    "Brötzmann Tanja": "Robin Heckmann",
    "Heckmann, Oliver": "Robin Heckmann",
    "Aghaie Pour": "Robin Heckmann",  # Hamid Aghaiepour
    "Shalamanova Snezhanka Dan": "Tim Selle",  # Snezhana Danailova Shalamanova
    "Calhanoglu Muhammed": "Robin Heckmann",
    "Rieker Anagbogu": "Robin Heckmann",  # Elvira Rieker Anagbogu
    "Rosen Adrian": "Robin Heckmann",  # in CRM als inaktiv markiert, aber aktiv
    "Rosen": "Robin Heckmann",  # bloss "Rosen" (ohne Vorname) auf AXA-Umbuchungszeilen - derselbe Kunde
    "Dos Santos Card": "Robin Heckmann",  # Paulo Dos Santos Cardoso
    "Die Peperonis Heddesh": "Robin Heckmann",  # Die Peperonis Heddesheim Inh. Tuerkmen & Dornsbach
    "Pension Wohlgelegen": "Tim Selle",  # Pension Wohlgelegen Natascha San Miguel Teran
    "Turbo team Hatice Sengönül": "Robin Heckmann",  # Turbo Team Umzuege & Transporte
    "Hoffmann Holding Gmb": "Tim Selle",  # Hoffmann Holding GmbH
    "Gutperle & Czech": "Andreas Selle",  # Gutperle & Czech Projektentwicklungs GmbH
    "Katzir-Shimo": "Robin Heckmann",  # Alon Katzir-Shimoni
    "Saßmann-Herz": "Robin Heckmann",  # Gunther Saßmann
    "Van der": "Tim Selle",  # Karin Van der Raaij
    "Stasiak, Alexander": "Tim Selle",  # jetzt Alexander Weickel (Heirat)
    "Erbrecht, Felix": "Robin Heckmann",  # jetzt Felix Aleman (Heirat)

    # August-2026: von Robin manuell hergeleitete Zuordnung fuer Kunden, die
    # match_betreuer() nicht (mehr) auflösen konnte - meist durch Extraktions-
    # Artefakte (abgeschnittene Namen wg. Spaltenbreite, oder bei Concordia
    # ein an den abgeschnittenen Namen gehaengtes Datum, siehe
    # extract_generic_table()/VHV_NAME_COLUMN_X0_MIN). Die Concordia-Zeilen
    # mit angehaengtem "01.08.26" werden sich in kuenftigen Monaten mit
    # anderem Datum wiederholen und brauchen dann jeweils einen neuen
    # Override-Eintrag, bis die Extraktion selbst repariert ist.
    "Golfplatz Heddesheim, Gut Neuz": "Andreas Selle",  # Mannheimer
    "Golfplatz Kurpfalz, GmbH & Co.": "Andreas Selle",  # Mannheimer
    "Rainer Marzenell & Botho Finck": "Robin Heckmann",  # Mannheimer
    "Petri + Söhne GmbH": "Robin Heckmann",  # Mannheimer
    "M.S. Car Logistic UG": "Robin Heckmann",  # Mannheimer + VHV, dasselbe Unternehmen
    "NETLINE Gesellschaft für": "Tim Selle",  # Hiscox, abgeschnitten
    "Schäfer, Brigi 01.08.26": "Robin Heckmann",  # Concordia
    "Gumbel, Patric 01.08.26": "Tim Selle",  # Concordia
    "Katzir- Shimon 01.08.26": "Robin Heckmann",  # Concordia, derselbe Kunde wie "Katzir-Shimo"
    "7 Aygan T&O Gm 01.08.26": "Robin Heckmann",  # Concordia
    "Take Time Bist 01.08.26": "Robin Heckmann",  # Concordia
    "Mercante, Ales 01.08.26": "Tim Selle",  # Concordia
    "Mohedien, Jiea 01.08.26": "Robin Heckmann",  # Concordia
    "Oppermann, Thi 01.08.26": "Robin Heckmann",  # Concordia
    "Kircher, Gisel 01.08.26": "Robin Heckmann",  # Concordia
    "Becker, Marina 01.08.26": "Robin Heckmann",  # Concordia
    "Bonvissuto, Do 01.08.26": "Robin Heckmann",  # Concordia
    "Garcia, Oscar 01.08.26": "Robin Heckmann",  # Concordia
    "Heckmann, Devo 01.08.26": "Robin Heckmann",  # Concordia
    "Pavelic, Phili 01.08.26": "Robin Heckmann",  # Concordia
    "Domingos, Yona 01.08.26": "Robin Heckmann",  # Concordia
    "Krasniqi, Just 01.08.26": "Robin Heckmann",  # Concordia
    "Motejat, Vanes 01.08.26": "Tim Selle",  # Concordia
    "Hustiuc, Mihai 01.08.26": "Tim Selle",  # Concordia
    "Badescu, Vili- 01.08.26": "Tim Selle",  # Concordia
    "Marin, Constan 01.08.26": "Tim Selle",  # Concordia
    "Da Costa, Dani 01.08.26": "Robin Heckmann",  # Concordia
    "Martirosjan, D 01.08.26": "Tim Selle",  # Concordia
    "Burke, Katarin 01.08.26": "Robin Heckmann",  # Concordia
    "Beeken, Martin 01.08.26": "Robin Heckmann",  # Concordia
    "Holzapfel": "Robin Heckmann",  # Alte-Leipziger, nur Nachname
    "Dipl.-Ing. Jürgen Daske Daske Architekten+Ingenieure": "Robin Heckmann",  # VHV
    "TAIN-KIM-HENG IMPORT-EXPORT": "Robin Heckmann",  # VHV
    "GmbH RET Automotive und Logistik": "Tim Selle",  # VHV
    "Benjamin Sebastian Brendle": "Tim Selle",  # VHV
    "Motor Hub Global GmbH": "Robin Heckmann",  # VHV
    "Sabrina Dana Hofstätter": "Tim Selle",  # VHV
    "Enrico Marco Pennino": "Tim Selle",  # VHV
    "Kiet Nho Tieu": "Robin Heckmann",  # VHV
    "Angela Ignazia Costa": "Tim Selle",  # VHV
    "GmbH INTE Informationstechnologie": "Robin Heckmann",  # VHV, "RH " abgeschnitten
    "Leticia Natascha Viviani": "Robin Heckmann",  # VHV
    "Aaliyah Angelina Viviani": "Robin Heckmann",  # VHV
    "Heizung Sanitär Klimatechnik Nico Geschwill Meisterbetrieb": "Robin Heckmann",  # VHV
    "Slike Brecht": "Robin Heckmann",  # VHV
    "Said Omar Saidi": "Tim Selle",  # VHV
    "Patrizia Göhlich": "Robin Heckmann",  # VHV
    "Vanessa Patrizia Di Stefano": "Tim Selle",  # VHV
    "PFANDLEIHHAUS G": "Tim Selle",  # Württembergische, abgeschnitten
    "VENDING GBR CMB": "Tim Selle",  # Württembergische, abgeschnitten
    "Heckmann": "Robin Heckmann",  # Auxilia, nur Nachname - Vorsicht bei Kollision mit anderen "Heckmann"-Kunden
    "Marcello": "Robin Heckmann",  # Auxilia, nur Nachname
    "Naser": "Robin Heckmann",  # Auxilia, nur Nachname
    "Baziari": "Robin Heckmann",  # Auxilia, nur Nachname
    "Seyhan*Bekyigat": "Tim Selle",  # Deurag
    "Mohammad*Mohammad": "Tim Selle",  # Deurag
    "Hertel*Daniela": "Robin Heckmann",  # Deurag
    "AHNOGLZEALPOFEL, PATRIK": "Robin Heckmann",  # Haftpflichtkasse, OCR-verzerrt
    "BONVISSUTO": "Robin Heckmann",  # Haftpflichtkasse
    "WEG Behrensstr. 20 v.d.": "Robin Heckmann",  # Haftpflichtkasse
    "WEG Industriestr. 24, v.D.": "Robin Heckmann",  # Haftpflichtkasse
    "WEG Zimmermann/ Sattler": "Robin Heckmann",  # Domcura (WEG Zimmermann/Sattler)
    "Duong": "Robin Heckmann",  # Domcura - passt zu Nutzer-Zielwert (Domcura komplett RH)
    # Dialog, August-2026: extract_dialog() liest die 4 Kundennamen wegen
    # des Wasserzeichen-Punktrasters nur verlesen (siehe extract_dialog()) -
    # per Betrag eindeutig zugeordnet: 148,81+73,09+23,28=245,18 EUR passt
    # exakt zum Nutzer-Zielwert TS, 82,89 EUR exakt zum Nutzer-Zielwert RH.
    "JAngolo Gatto Ro": "Tim Selle",  # eigentlich Angelo Gatto Rotondo
    "Popodopoulom F": "Tim Selle",  # eigentlich Papadopoulou, Maria
    "S Innereausb": "Tim Selle",  # eigentlich G & S Innenausbau GmbH
    "MNEG Prinzregent": "Robin Heckmann",  # eigentlich WEG Prinzregentenstr. 18
    "Duc": "Robin Heckmann",  # Domcura, nur Vor-/Nachname - Kollisionsrisiko
    "MERLING ANNEMAR": "Robin Heckmann",  # Württembergische, abgeschnitten
    "040.054.0460573.0 | Simon,Maxime": "Robin Heckmann",  # Helvetia
    "Krawczyk, Philipp": "Tim Selle",  # Fondsfinanz
    "Sauer-Segmann, Allianz": "Robin Heckmann",  # Fondsfinanz
    "Born": "Tim Selle",  # Fondsfinanz, nur Nachname - Kollisionsrisiko
    "Oscar Andres Garcia 01.08.2026": "Robin Heckmann",  # Württembergische-Adam-Riese
    "Heckmann*Willi": "Robin Heckmann",  # Signal-Iduna
    "Heinze*Simon": "Tim Selle",  # Signal-Iduna
    "Vohrmann Claus Benedikt": "Robin Heckmann",  # Allianz
    "Leitner,Dr. Hans": "Tim Selle",  # BSG - Hans Leitner
    "Wilhelm": "Tim Selle",  # Auxilia - Vorname-Fragment von Gerhard Wilhelm Martyniak
    "Martyniak": "Tim Selle",  # Auxilia - Nachname-Fragment von Gerhard Wilhelm Martyniak
    # Betreuer.xlsx ist fuer Mandy Bechtel veraltet (Privatkunden-Blatt sagt
    # "Robin Heckmann", Firmenkunden-GbR-Zeile "Tim Selle") - laut Nutzer
    # (2026-09-14) ist sie tatsaechlich Andreas Selles Kundin, ebenso die GbR
    # "Rene Gutperle/Thomas Czech/Angela Czech/Mandy Bechtel ... GbR" (passt
    # zum bereits bestehenden "Gutperle & Czech"-Override oben). Betreuer.xlsx
    # sollte an der Quelle korrigiert werden, dieser Override greift bis
    # dahin.
    "Bechtel, Mandy": "Andreas Selle",
    "Bechtel": "Andreas Selle",
    # Mannheimer: laesst sich exakt (154,39 EUR) auf die RH->AS-Differenz in
    # dieser Abrechnung zurueckfuehren, passt zum Gutperle/Czech-Cluster
    # (siehe "Gutperle & Czech" oben) - bestaetigt durch exakten AS-Treffer
    # (1.060,57 EUR) in der August-2026-Abstimmung.
    "Werner Gutperle GmbH & Co.KG": "Andreas Selle",
    # Wuerttembergische (VM 0821-2534-5-01), August-2026: vom Nutzer
    # bestaetigt (2026-09-14). Kovacic Edita bleibt laut Nutzer TS (NICHT
    # RH, wie zunaechst - falsch - aus der 4,31-EUR-Restdifferenz vermutet) -
    # die Ursache dieser Restdifferenz ist noch ungeklaert.
    "Czech Thomas": "Tim Selle",

    # September-2026: von Robin manuell hergeleitete Zuordnung. Concordia-
    # Zeilen mit angehaengtem Buchungsdatum (z.B. "Lamitschka, Ph") werden ab
    # jetzt ueber strip_trailing_date() in apply_betreuer() datumsunabhaengig
    # abgeglichen - der Key hier steht daher OHNE Datum und gilt fuer alle
    # Monate, nicht nur September.
    "Marcel Martin Wolfgang Bajer": "Tim Selle",  # VHV
    "Bau & Elektrotechnik GmbH Czech": "Tim Selle",  # VHV
    "und Netzbau GmbH autmatec Leitungs-": "Robin Heckmann",  # VHV, Wortreihenfolge vertauscht - eigentlich "autmatec Leitungs- und Netzbau GmbH"
    "Raymond Oliver Ruppert Mügge": "Tim Selle",  # VHV
    "MSC GmbH in Gründung": "Tim Selle",  # VHV
    "Jennifer Nkiruka Anagbogu": "Robin Heckmann",  # VHV
    "schaft mbH FGH Zertifizierungsgesell-": "Tim Selle",  # VHV, umgebrochen/abgeschnitten - eigentlich "FGH Zertifizierungsgesellschaft mbH"
    "Immobilien GmbH, Empire Home": "Robin Heckmann",  # HDI, Wortreihenfolge vertauscht - eigentlich "Empire Home Immobilien GmbH"
    "D. Bellgardt & M. Sößer GbR": "Robin Heckmann",  # HDI (abgekuerzte Form; volle Form "Daniel Bellgardt und Michael Sößer GbR" matcht bereits ueber Betreuer.xlsx)
    "7 Aygan T&Ö GmbH": "Robin Heckmann",  # HDI
    "Restaurant Artischocke": "Robin Heckmann",  # HDI
    "Aladinco Baudienstleistungen": "Tim Selle",  # HDI
    "Back, Heinrich": "Robin Heckmann",
    "VIP Baudienstleistungen & Handels GmbH": "Tim Selle",
    "VIP Baudienstleistungen & Hand": "Tim Selle",  # dieselbe Firma, abgeschnitten (VHV)
    "Rudolf Wiedemann Bärbel Henika": "Robin Heckmann",
    "Jukadio Sports": "Robin Heckmann",
    "Schadenszentrum Rhein-Neckar": "Robin Heckmann",
    "Michelhans Isa": "Robin Heckmann",
    "CARVISION": "Tim Selle",  # Württembergische, abgeschnitten

    # Concordia, abgeschnittener Name + Buchungsdatum (siehe strip_trailing_date()
    # oben - Key hier bewusst OHNE Datum, damit er monatsunabhaengig greift):
    "Lamitschka, Ph": "Tim Selle",
    "Dos Santos Car": "Robin Heckmann",  # andere Abschneidestelle als bestehendes "Dos Santos Card"
    "Katzir-Shimoni": "Robin Heckmann",  # andere Abschneidestelle als bestehendes "Katzir-Shimo"
    "7 Aygan T&O Gm": "Robin Heckmann",  # Concordia-Truncation derselben Firma wie "7 Aygan T&Ö GmbH" oben
    "Take Time Bist": "Robin Heckmann",  # dieselbe Firma wie "Take Time Bistro" unten, abgeschnitten
    "Take Time Bistro": "Robin Heckmann",
    "Mohedien, Jiea": "Robin Heckmann",
    "Oppermann, Thi": "Robin Heckmann",
    "Mercante, Ales": "Tim Selle",
    "Kircher, Gisel": "Robin Heckmann",
    "Becker, Marina": "Robin Heckmann",
    "Bonvissuto, Do": "Robin Heckmann",  # anderer Kunde als bestehendes "BONVISSUTO" (Haftpflichtkasse)
    "Garcia, Oscar": "Robin Heckmann",
    "Heckmann, Devo": "Robin Heckmann",
    "Pavelic, Phili": "Robin Heckmann",
    "Domingos, Yona": "Robin Heckmann",
    "Krasniqi, Just": "Robin Heckmann",  # Fondsfinanz-Truncation von Krasniqi, Justin
    "Badescu, Vili-": "Tim Selle",
    "Marin, Constan": "Robin Heckmann",
    "Motejat, Vanes": "Tim Selle",
    "Hustiuc, Mihai": "Tim Selle",
    "Da Costa, Dani": "Robin Heckmann",
    "Burke, Katarin": "Robin Heckmann",
    "Martirosjan, D": "Tim Selle",
    "Beeken, Martin": "Robin Heckmann",

    "Roman": "Robin Heckmann",  # nur Vorname/Nachname - Kollisionsrisiko
    "Lewandowski, HDI": "Robin Heckmann",  # Quelle unklar (kein Klartext-Treffer in den September-Dateien) - Key wie vom Nutzer gemeldet
    "ACAR FAHRZEUGTE": "Robin Heckmann",  # abgeschnitten
    "HASHEMI, SAID KABIR": "Tim Selle",  # Haftpflichtkasse
    "IMMHOLD TREUHAN": "Tim Selle",  # abgeschnitten
    "SIEGL VERWALTUN": "Tim Selle",  # abgeschnitten
    "Gro~*Thomas": "Tim Selle",  # OCR-verzerrt, vermutlich "Groß*Thomas" (Deurag-Namensmuster)
    "ABNOGNEVLISOSUTO": "Robin Heckmann",  # Haftpflichtkasse, Vertrag 54486580 - vom Nutzer identifiziert als Domenico Savio Angelo Bonvissuto; Name im PDF-Textlayer selbst verstuemmelt (kein OCR-Problem)
    "Pension Wohligelegen": "Tim Selle",  # Alte-Leipziger, andere OCR-Schreibweise als bestehendes "Pension Wohlgelegen" (fehlendes "i")
    "HeckmannxWilli": "Robin Heckmann",  # Signal-Iduna, andere OCR-Schreibweise als bestehendes "Heckmann*Willi" ("*" als "x" gelesen)
    "Michaela u. Christine Seipp": "Robin Heckmann",  # Alte-Leipziger
    "Seipp Michaela": "Robin Heckmann",
    "Seipp Christine": "Robin Heckmann",
    # Alte-Leipziger September-2026: laut Excel des Nutzers in der RH-Spalte
    # (62,87 Turbo team + 50,21 Oppermann + 10,98 Walla = 124,06 RH).
    "Alexandra, Walla": "Robin Heckmann",
    "Walla, Alexandra": "Robin Heckmann",
    # Dialog September-2026: OCR-verstuemmelte Namen, per Betrag aus den
    # Nutzer-Zielwerten TS 20,51 (23,28 - 102,72 + 99,95) und RH 263,31
    # (53,79 + 209,52) hergeleitet; "WESG Industriest" = WEG Industriestr.,
    # passend zum bestehenden RH-Override fuer "WEG Industriestr. 24, v.D.".
    "S InnerausblF": "Tim Selle",  # G & S Innenausbau GmbH
    "Z aCbR A": "Tim Selle",  # GZ GbR
    "K und S Iemobili": "Tim Selle",  # K und S Immobilien und ...
    "WESG Industriest": "Robin Heckmann",  # WEG Industriestr.
    # Baloise (ehem. Basler): laut Excel des Nutzers in September (34,50) und
    # August (29,19) komplett RH; OCR-Schreibvarianten desselben Kunden.
    "ANNATRANSPORT": "Robin Heckmann",
    "ANMATRANSPORT": "Robin Heckmann",
    "AMNATRANSPORT": "Robin Heckmann",
    "SAS GABRI": "Robin Heckmann",
    "LUTEC GMBH": "Robin Heckmann",
    # Helvetia September-2026, Vermittlerliste 00026/2026: Nutzer bestaetigt
    # Yasar/Aysar, Mehmet als TS-Kunden; die Liste (261,16 - 189,33 + 509,74 -
    # 331,33 = 250,24) ist in der Excel komplett TS gebucht, daher auch
    # Tavan, Ismael.
    "040.090.0504595.6 | Yasar,Mehmet": "Tim Selle",
    "040.090.0504618.0 | Tavan,Ismael": "Tim Selle",
    # Continentale Protokoll EUROPA-Reisekrankenversicherung: Buchungstext
    # "RV-200926540 SASSMANN-H" (2,40 EUR) - laut Nutzer RH-Kunde.
    "SASSMANN-H": "Robin Heckmann",
}

# Kennzeichen fuer nicht-kundenspezifische Sammelabzuege innerhalb einer
# Abrechnung (z.B. Fondsfinanz-Stornoreserve, Swiss-Life-VSV-Beitrag, siehe
# extract_swiss_life_vsv_deduction()) - werden in apply_betreuer() separat
# behandelt (Zuordnung zur groessten Kundenposition derselben Datei), nicht
# ueber die normale Namenszuordnung.
COLLECTIVE_DEDUCTION_MARKER = "nicht kundenspezifisch"

# Kennzeichen fuer Betraege, die AUCH NICHT ueber COLLECTIVE_DEDUCTION_MARKER
# (Zuordnung zur groessten Kundenposition derselben Datei) behandelt werden
# sollen, weil sie sich ueberhaupt keinem einzelnen Kunden/Partner zuordnen
# lassen - z.B. die Auszahlung einer freien Stornoreserve bei Amex-Pool
# (siehe extract_amex_stornoreserve_auszahlung()): Lebensversicherungs-
# Provisionen werden bei SSH vollstaendig verbucht, obwohl ein Teil davon
# zunaechst als Stornoreserve einbehalten wird - die Betreuer-Zuordnung
# erfolgt also schon beim urspruenglichen Buchungsdatum. Wird die Reserve
# spaeter (manuell angestossen) ausgezahlt, betrifft das rueckwirkend viele
# verschiedene, nicht mehr einzeln nachvollziehbare Kunden/Partner (Nutzer-
# Bestaetigung) - die Zeile bleibt daher bewusst ohne Betreuer UND wird aus
# dem "Kunde_ohne_Betreuer"-Blatt ausgeblendet (siehe write_excel()/app.py),
# damit sie dort nicht faelschlich als offener Pruef-Fall auftaucht.
NICHT_ZUORDENBAR_MARKER = "nicht zuordenbar"

GEB_RE = re.compile(r"\(geb\.?[^)]*\)|\bgeb\.?:?\s.*$", re.IGNORECASE)
LEGAL_FORM_SUFFIXES = ["gmbhcokg", "gmbh", "cokg", "ug", "ev", "ag", "se", "kg", "ohg", "gbr"]


def _strip_geb(s):
    """Entfernt Geburtsname-Zusaetze wie '(geb. Bock)' oder 'geb.: Pleli'
    aus einem CRM-Namen, damit z.B. 'Dorn (geb. Bock)' auf 'Dorn' passt."""
    if not s:
        return s
    return GEB_RE.sub("", str(s)).strip()


def _strip_legal_form(s):
    for suf in LEGAL_FORM_SUFFIXES:
        if s.endswith(suf) and len(s) > len(suf):
            return s[: -len(suf)]
    return s


def _levenshtein_le1(a, b):
    """True, wenn a und b sich in hoechstens einem Zeichen unterscheiden
    (Einfuegen/Loeschen/Ersetzen) - fuer sehr enge Tippfehler-Varianten wie
    'Philip'/'Phillip'. Bewusst kein allgemeiner Fuzzy-Match (Risiko, zwei
    verschiedene Personen zu verwechseln), siehe _names_match()."""
    if abs(len(a) - len(b)) > 1:
        return False
    if a == b:
        return True
    # gleiche Laenge: genau ein ersetztes Zeichen erlaubt
    if len(a) == len(b):
        diffs = sum(1 for x, y in zip(a, b) if x != y)
        return diffs <= 1
    # Laengendifferenz 1: ein eingefuegtes/geloeschtes Zeichen erlaubt
    shorter, longer = (a, b) if len(a) < len(b) else (b, a)
    for i in range(len(longer)):
        if longer[:i] + longer[i + 1:] == shorter:
            return True
    return False


def _names_match(a, b):
    """Vergleicht zwei bereits normalisierte Vornamen. Exakt gleich,
    ODER einer ist Praefix des anderen (typisch fuer im Quell-PDF
    abgeschnittene Namen, z.B. 'wolfgan' vs 'wolfgang', 'seb' vs
    'sebastian' - siehe Nutzer-Bestaetigung), ODER Levenshtein-Distanz <= 1
    bei Namen ab 4 Zeichen (enge Schreibvarianten wie 'philip'/'phillip').
    Der Nachname muss IMMER exakt uebereinstimmen (siehe match_betreuer) -
    diese Unschaerfe gilt nur fuer den Vornamen, um keine unterschiedlichen
    Personen zu verwechseln."""
    if not a or not b:
        return False
    if a == b:
        return True
    if len(a) >= 3 and len(b) >= 3 and (a.startswith(b) or b.startswith(a)):
        return True
    if len(a) >= 4 and len(b) >= 4 and _levenshtein_le1(a, b):
        return True
    return False


def load_betreuer_lookup(path):
    """Laedt den CRM-Export 'Betreuer.xlsx' (Blaetter 'Privatkunden' mit
    Anrede/Vorname/Nachname/Kundenverantwortliche und 'Firmenkunden' mit
    u.a. Firmenname/Kundenverantwortliche) und baut daraus Nachschlage-
    Strukturen fuer match_betreuer(). 'Calvin Keinarth' wird auf 'Andreas
    Selle' umgemappt (siehe BETREUER_ALIASES)."""
    import openpyxl

    wb = openpyxl.load_workbook(path, data_only=True)
    by_surname = {}  # normalisierter Nachname -> [(normalisierter Vorname, Betreuer)]
    people = []  # flache Liste (normalisierter Nachname, Vorname, Betreuer) - fuer Fuzzy-Nachname-Scan
    companies = {}  # normalisierter (Rechtsform-bereinigter) Firmenname -> Betreuer

    if "Privatkunden" in wb.sheetnames:
        ws = wb["Privatkunden"]
        header = [c.value for c in ws[1]]
        if "Vorname" in header and "Nachname" in header and "Kundenverantwortliche" in header:
            i_vor, i_nach, i_bet = (header.index(k) for k in
                                     ("Vorname", "Nachname", "Kundenverantwortliche"))
            for row in ws.iter_rows(min_row=2, values_only=True):
                vor, nach, bet = row[i_vor], row[i_nach], row[i_bet]
                if not bet or not nach:
                    continue
                bet = BETREUER_ALIASES.get(normalize(str(bet)), bet)
                nnach = normalize_for_match(_strip_geb(nach))
                nvor = normalize_for_match(_strip_geb(vor))
                if nnach:
                    by_surname.setdefault(nnach, []).append((nvor, bet))
                    people.append((nnach, nvor, bet))

    if "Firmenkunden" in wb.sheetnames:
        ws = wb["Firmenkunden"]
        header = [c.value for c in ws[1]]
        if "Firmenname" in header and "Kundenverantwortliche" in header:
            i_firma, i_bet = header.index("Firmenname"), header.index("Kundenverantwortliche")
            for row in ws.iter_rows(min_row=2, values_only=True):
                firma, bet = row[i_firma], row[i_bet]
                if not bet or not firma:
                    continue
                bet = BETREUER_ALIASES.get(normalize(str(bet)), bet)
                key = _strip_legal_form(normalize_for_match(firma))
                if key:
                    companies[key] = bet

    return {"by_surname": by_surname, "people": people, "companies": companies}


TRAILING_DATE_RE = re.compile(r"\s+\d{2}\.\d{2}\.(?:\d{2}|\d{4})$")


def strip_trailing_date(kunde):
    """Entfernt ein an einen (meist wegen Spaltenbreite abgeschnittenen)
    Kundennamen angehaengtes Buchungsdatum, z.B. 'Lamitschka, Ph 01.09.26' ->
    'Lamitschka, Ph' (siehe extract_generic_table() - betrifft v.a.
    Concordia). Ohne das wuerde fuer denselben wiederkehrenden Kunden jeden
    Monat ein neuer MANUAL_BETREUER_OVERRIDES_RAW-Eintrag mit dem jeweils
    aktuellen Datum noetig sein, obwohl der (abgeschnittene) Name selbst
    monatsunabhaengig stabil bleibt. Wird sowohl auf den extrahierten Namen
    als auch beim Aufbau der Override-Lookup-Keys angewandt (apply_betreuer()),
    damit alte UND neue Override-Eintraege unabhaengig davon greifen, ob ihr
    Key das Datum noch enthaelt."""
    return TRAILING_DATE_RE.sub("", str(kunde))


def match_betreuer(kunde, lookup):
    """Ordnet einen extrahierten Kundennamen (z.B. 'Nachname, Vorname',
    'Vorname Nachname' oder ein Firmenname) einem Betreuer aus dem CRM-
    Export zu (siehe load_betreuer_lookup()).

    Geprueft wird stufenweise, von streng zu locker (exakt vor unscharf,
    Nachname vor Vorname), und bei jeder Stufe gilt: sobald genau EIN
    Betreuer zu den Kandidaten passt, ist das Ergebnis "eindeutig genug"
    und wird zurueckgegeben (Nutzer-Vorgabe: auch bei nicht exakter
    Schreibweise zuordnen, wenn sonst klar ist, wer gemeint ist). Passen
    mehrere unterschiedliche Betreuer auf derselben Stufe, gilt der Fall als
    mehrdeutig - dann wird NICHT auf eine lockerere Stufe weiter versucht
    (sonst waechst das Verwechslungsrisiko), sondern None geliefert (siehe
    Kunde in Manuelle_Pruefung/unmatched-Uebersicht).

    Stufen: 1) Nachname exakt + Vorname exakt, 2) Nachname exakt + Vorname
    unscharf, 3) Nachname unscharf + Vorname exakt, 4) Nachname unscharf +
    Vorname unscharf."""
    if not kunde:
        return None
    kunde = str(kunde).strip()

    candidate_pairs = []
    if "," in kunde:
        a, b = (p.strip() for p in kunde.split(",", 1))
        candidate_pairs.append((normalize_for_match(a), normalize_for_match(b)))  # Nachname, Vorname
    tokens = kunde.replace(",", " ").split()
    if len(tokens) >= 2:
        first = normalize_for_match(tokens[0])
        rest = normalize_for_match(" ".join(tokens[1:]))
        candidate_pairs.append((rest, first))  # Nachname=rest, Vorname=erstes Token
        candidate_pairs.append((first, rest))  # falls Vorname zuerst steht

    for nachname, vorname in candidate_pairs:
        if not nachname or not vorname:
            continue

        exact = [bet for vn, bet in lookup["by_surname"].get(nachname, []) if vn == vorname]
        distinct = set(exact)
        if len(distinct) == 1:
            return distinct.pop()
        if len(distinct) > 1:
            continue  # mehrdeutig auf dieser Stufe -> nicht lockerer weiterversuchen

        fuzzy_vor = [bet for vn, bet in lookup["by_surname"].get(nachname, []) if _names_match(vn, vorname)]
        distinct = set(fuzzy_vor)
        if len(distinct) == 1:
            return distinct.pop()
        if len(distinct) > 1:
            continue

        exact_vor_fuzzy_nach = [bet for nn, vn, bet in lookup["people"]
                                 if vn == vorname and nn != nachname and _names_match(nn, nachname)]
        distinct = set(exact_vor_fuzzy_nach)
        if len(distinct) == 1:
            return distinct.pop()
        if len(distinct) > 1:
            continue

        fuzzy_both = [bet for nn, vn, bet in lookup["people"]
                      if nn != nachname and _names_match(nn, nachname) and _names_match(vn, vorname)]
        distinct = set(fuzzy_both)
        if len(distinct) == 1:
            return distinct.pop()

    company_key = _strip_legal_form(normalize_for_match(kunde))
    if company_key in lookup["companies"]:
        return lookup["companies"][company_key]
    return None


def extract_bank_credits(path, password=""):
    """Liest einen VR-Bank-Rhein-Neckar-Kontoauszug (PDF mit normalem
    Textlayer) und liefert alle Buchungen (nicht nur Gutschriften) als Liste
    von Dicts: datum, vorgang, betrag (vorzeichenbehaftet), sender,
    verwendungszweck.

    Die Seiten wiederholen Brief-Kopf/Fusszeilen und einen 'Uebertrag auf/
    von Blatt N'-Saldo-Hinweis auf jeder Seite - werden vorab herausgefiltert
    und alle Seiten zu einem durchgehenden Zeilenstrom zusammengefuegt, damit
    eine Buchung, deren Absenderzeile durch einen Seitenumbruch von ihrer
    Buchungszeile getrennt ist, nicht faelschlich der naechsten Buchung oder
    einer Kopfzeile zugeordnet wird."""
    all_lines = []
    with pdfplumber.open(path, password=password) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            for line in text.split("\n"):
                line = line.strip()
                if not line:
                    continue
                if BANK_BOILERPLATE_RE.search(line) or BANK_BOILERPLATE_WHOLE_LINE_RE.match(line):
                    continue
                all_lines.append(line)

    txs = []
    i, n = 0, len(all_lines)
    while i < n:
        m = BANK_TX_START_RE.match(all_lines[i])
        if not m:
            i += 1
            continue
        bu_tag, _wert, vorgang, amt_s, flag = m.groups()
        amt = parse_amount(amt_s)
        amt = -amt if flag == "S" else amt
        sender = all_lines[i + 1] if i + 1 < n else ""
        j = i + 2
        detail_lines = []
        while j < n and not BANK_TX_START_RE.match(all_lines[j]):
            detail_lines.append(all_lines[j])
            j += 1
        txs.append({
            "datum": bu_tag, "vorgang": vorgang.strip(), "betrag": round(amt, 2),
            "sender": sender, "verwendungszweck": " ".join(detail_lines),
        })
        i = j
    return txs


def previous_month_folder(month_folder):
    """'September-2026' -> 'August-2026', 'Januar-2026' -> 'Dezember-2025'.
    Liefert None, wenn der Ordnername nicht dem Muster 'Monat-Jahr' folgt."""
    try:
        name, year = month_folder.rsplit("-", 1)
        idx = MONTHS_DE.index(name)
        year = int(year)
    except (ValueError, AttributeError):
        return None
    if idx == 0:
        return f"{MONTHS_DE[11]}-{year - 1}"
    return f"{MONTHS_DE[idx - 1]}-{year}"


def load_prior_month_totals(source):
    """Liest das Ergebnis-Excel eines Vormonats (von diesem Tool erzeugt,
    Datei oder Datei-Objekt) und liefert je Abrechnungsdatei die Betraege, mit
    denen ein Zahlungseingang des laufenden Monats erklaert werden kann:
    [{'insurer': ..., 'amounts': [Summe_extrahiert, Summe_lt_PDF, ...]}].
    Fehlende Blaetter/Spalten werden toleriert (leere Liste)."""
    entries = {}
    try:
        sheets = pd.read_excel(
            source, sheet_name=["Kontrolle", "Sammelbelege_ohne_Details", "Manuelle_Pruefung"]
        )
    except Exception:
        try:
            if hasattr(source, "seek"):
                source.seek(0)
            sheets = {"Kontrolle": pd.read_excel(source, sheet_name="Kontrolle")}
        except Exception:
            return []
    for sheet, cols in (
        ("Kontrolle", ("Summe_extrahiert", "Summe_lt_PDF")),
        ("Sammelbelege_ohne_Details", ("Betrag_lt_PDF",)),
        ("Manuelle_Pruefung", ("Betrag_lt_PDF",)),
    ):
        df = sheets.get(sheet)
        if df is None or "Versicherer" not in df.columns:
            continue
        for _, row in df.iterrows():
            amounts = entries.setdefault(str(row["Versicherer"]), [])
            for col in cols:
                if col in df.columns and pd.notna(row[col]):
                    try:
                        amounts.append(round(float(row[col]), 2))
                    except (TypeError, ValueError):
                        pass
    return [{"insurer": k, "amounts": v} for k, v in entries.items()]


def reconcile_bank_credits(bank_credits, insurer_keys, prior_totals=None, explained_out=None):
    """Vergleicht Gutschriften aus dem Kontoauszug mit den in diesem Monat
    tatsaechlich verarbeiteten Versicherer-Dateien. Liefert eine Liste nicht
    zuordenbarer Gutschriften (Dicts wie extract_bank_credits(), zusaetzlich
    ohne 'vorgang') - das sind Zahlungseingaenge, zu denen keine passende
    Abrechnungs-PDF gefunden wurde, also ein Hinweis auf eine fehlende
    Abrechnung. Es wird zuerst nur der Absendername geprueft, erst wenn das
    nichts findet, zusaetzlich der Verwendungszweck (manche Versicherer
    zahlen ueber einen Pool-Absender wie 'Qualitypool' aus, der eigentliche
    Versicherer - z.B. 'AMEX' - steht dann nur im Verwendungszweck). Reine
    Absender-Suche zuerst zu versuchen vermeidet Fehltreffer durch
    Referenznummern im Verwendungszweck, die zufaellig wie ein anderer
    Versicherer-Name aussehen (z.B. eine VEMA-Referenznummer in einer
    Baloise-Zahlung).

    prior_totals (optional, siehe load_prior_month_totals()): Abrechnungen
    des Vormonats. Eine Abrechnung wird oft erst im Folgemonat bezahlt
    (Statement im Monatsordner M, Zahlungseingang in M+1) - ein sonst
    unzuordenbarer Zahlungseingang gilt daher als erklaert, wenn der
    Versicherer im Vormonat eine Abrechnung hatte UND der Betrag genau einer
    ihrer Summen (extrahiert oder aufgedruckt) entspricht. Nur Betrag UND
    Versicherer zusammen zaehlen - Versicherer allein wuerde ein fehlendes
    Statement des laufenden Monats bei regelmaessig abrechnenden
    Versicherern immer verdecken.

    explained_out (optional, Liste): erhaelt die so erklaerten Zahlungen
    (Dicts + 'erklaert_durch') zur Anzeige."""
    families_present = {insurer_bank_family(k) for k in insurer_keys}

    def try_match(text_norm, families):
        for family in families:
            aliases = INSURER_BANK_ALIASES.get(family, [family])
            if any(alias in text_norm for alias in aliases):
                return family
        return None

    prior_by_family = {}
    for entry in prior_totals or []:
        fam = insurer_bank_family(entry["insurer"])
        prior_by_family.setdefault(fam, []).extend(entry["amounts"])
    prior_families = set(prior_by_family)

    unmatched = []
    for credit in bank_credits:
        if credit["vorgang"] != "GUTSCHRIFT":
            continue
        sender_norm = normalize_for_match(credit["sender"])
        full_norm = normalize_for_match(credit["sender"] + " " + credit["verwendungszweck"])
        if try_match(sender_norm, families_present):
            continue
        if try_match(full_norm, families_present):
            continue
        if prior_families:
            fam = try_match(sender_norm, prior_families) or try_match(full_norm, prior_families)
            if fam and any(abs(a - credit["betrag"]) <= 0.01 for a in prior_by_family[fam]):
                if explained_out is not None:
                    explained_out.append(dict(credit, erklaert_durch=f"Abrechnung {fam} aus dem Vormonat"))
                continue
        unmatched.append(credit)
    return unmatched


DIALOG_AMOUNT_RE = re.compile(r"-?\d{1,3}(?:\.\d{3})*,\d{2}")


def find_dialog_total_hint(pdf):
    """Dialog: die eigentliche Kundenpositionen-Tabelle wird ueber
    extract_dialog() gelesen (zeilenweiser Zuschnitt + Median-Filter gegen
    das Wasserzeichen-Punktraster). Diese Funktion liest stattdessen die
    hochkant-Seiten (Deckblatt, Abrechnungsuebersicht, 'PG-Uebersicht') -
    der Gesamtbetrag ("Summe gesamt") steht dort mehrfach (Uebersichts-
    UND Deckblatt-Seite), oft zusammen mit viel Bildrauschen durch dasselbe
    Wasserzeichen-Punktraster wie auf der Detailseite. Verlaesslicher
    Text-Anker wie "Summe" ueberlebt die OCR dabei nicht immer (selbst mit
    Schwellwert-Vorverarbeitung + Sparse-Text-Modus --psm 11) - stattdessen
    werden alle erkannten Dezimalbetraege ueber saemtliche hochkant-Seiten
    gesammelt und der insgesamt haeufigste zurueckgegeben (der echte
    Gesamtbetrag taucht erfahrungsgemaess auf mehreren Seiten auf, zufaellige
    OCR-Fehlwerte dagegen nicht). Dient als Pflicht-Gegenprobe fuer
    extract_dialog()'s Ergebnis (siehe process_file())."""
    if not OCR_AVAILABLE:
        return None

    # Bevorzugter Weg: Deckblatt-Zeilen "Provision DIV  283,82" und
    # "Auszahlung dd.mm.yyyy : 283,82-". Mit Median-Filter vor der OCR werden
    # sie meist lesbar, einzelne Aufloesungs-/Filter-Kombinationen liefern
    # aber verwechselte Ziffern (z.B. 981,77 statt 441,77) - daher
    # Mehrheitsentscheid ueber mehrere Varianten. Pro Zeile zaehlt nur der
    # ERSTE Betrag (die zweite Spalte ist der kumulierte Jahreswert).
    votes = []
    for page in pdf.pages:
        if len(page.chars) > 0 or page.width > page.height:
            continue
        for med, dpi in ((3, 200), (5, 300), (3, 300), (5, 200)):
            im = page.to_image(resolution=dpi).original.convert("L").filter(ImageFilter.MedianFilter(med))
            text = pytesseract.image_to_string(im, lang="deu", config="--psm 6")
            for line in text.splitlines():
                low = line.lower()
                if "provision" in low or "auszahlung" in low:
                    nums = DIALOG_AMOUNT_RE.findall(line)
                    if nums:
                        v = parse_amount(nums[0])
                        if v is not None and abs(v) > 0.004:
                            votes.append(round(abs(v), 2))
    if votes:
        best, count = Counter(votes).most_common(1)[0]
        if count >= 2:
            return best

    all_amounts = []
    for page in pdf.pages:
        if len(page.chars) > 0 or page.width > page.height:
            continue  # keine gescannte Seite bzw. die quer liegende Detailseite
        im = page.to_image(resolution=400).original.convert("L")
        im = ImageOps.autocontrast(im, cutoff=1)
        im = im.point(lambda p: 255 if p > 150 else 0)
        ocr_text = pytesseract.image_to_string(im, lang="deu", config="--psm 11")
        all_amounts.extend(a for a in (parse_amount(m) for m in DIALOG_AMOUNT_RE.findall(ocr_text)) if a)
    if not all_amounts:
        return None
    return Counter(all_amounts).most_common(1)[0][0]


# Versicherer, die ihre Betragsspalte aus ihrer eigenen Soll/Haben-Sicht
# drucken statt aus Maklersicht - ein bei ihnen negativer Betrag ist fuer
# SSH eine Gutschrift und umgekehrt (Nutzer-Bestaetigung fuer Hiscox/
# Mannheimer, urspruenglich bei AIG anhand der aufgedruckten Summenzeile
# "...zu Ihren Gunsten..." entdeckt). Siehe SIGN_INVERTED_INSURERS-Nutzung
# in process_file().
SIGN_INVERTED_INSURERS = {"aig", "hiscox", "mannheimer"}


def insurer_name_from_filename(filename, month_folder):
    # NFC-Normalisierung: macOS zerlegt Umlaute in Dateinamen standardmaessig
    # in NFD (z.B. "u" + combining diaeresis statt des einzelnen Zeichens
    # "ü") - visuell identisch, aber ein anderer Byte-/Codepoint-Verlauf.
    # fpdf's Kernschriftart (Latin-1-basiert) kennt das freistehende
    # Kombinationszeichen nicht und wirft FPDFUnicodeEncodingException,
    # sobald der so abgeleitete Versicherer-Name in der PDF-Uebersicht
    # ausgegeben wird (beobachtet: Kollege laedt Dateien von einem Mac
    # hoch). NFC vereinheitlicht auf die zusammengesetzte Form, die sowohl
    # fpdf als auch Excel/Betreuer.xlsx-Abgleich zuverlaessig verarbeiten.
    base = unicodedata.normalize("NFC", os.path.splitext(os.path.basename(filename))[0])
    base = re.sub(r"^Abrechnung-\d+-", "", base)
    for m in MONTHS_DE:
        base = re.sub(rf"-{m}-\d{{4}}$", "", base)
    base = base.rstrip("-")
    return base


# ---------------------------------------------------------------------------
# Hauptverarbeitung je Datei
# ---------------------------------------------------------------------------

# Bekannte Passwoerter fuer verschluesselte Versicherer-PDFs (vom Nutzer
# mitgeteilt). Schluessel = klein geschriebener insurer-Name (Substring-
# Match), Wert = Passwort.
PDF_PASSWORDS = {
    "itzehoer": "685195211",
}


def process_file(filepath, month_folder):
    insurer = insurer_name_from_filename(filepath, month_folder)
    filename = os.path.basename(filepath)

    password = next((p for k, p in PDF_PASSWORDS.items() if k in insurer.lower()), "")
    try:
        pdf = pdfplumber.open(filepath, password=password)
    except Exception as e:
        return {
            "insurer": insurer, "file": filename, "status": "fehler",
            "reason": f"PDF konnte nicht geoeffnet werden ({type(e).__name__}) - "
                      f"evtl. passwortgeschuetzt.",
            "rows": [], "total_hint": None,
        }

    try:
        full_text_parts = []
        for page in pdf.pages:
            t = page.extract_text() or ""
            full_text_parts.append(t)
        full_text = "\n".join(full_text_parts)
        insurer_lower = insurer.lower()

        if "allianz" in insurer_lower or "axa" in insurer_lower or "gothaer" in insurer_lower:
            # Gespiegelter Text wie bei is_mirrored_text() weiter unten,
            # aber mit eigenem, koordinatenbasiertem Parser statt manueller
            # Pruefung - siehe extract_allianz()/extract_axa()/extract_gothaer().
            if "allianz" in insurer_lower:
                rows = extract_allianz(pdf)
            elif "axa" in insurer_lower:
                rows = extract_axa(pdf)
            else:
                rows = extract_gothaer(pdf)
            total_hint = find_total_in_text(full_text[::-1])
            if not rows:
                return {
                    "insurer": insurer, "file": filename, "status": "sammelbeleg",
                    "reason": "Kein Kunden-Positionsdetail im PDF gefunden - "
                              "vermutlich reiner Sammelbeleg/Kontoauszug ohne "
                              "Einzelaufstellung.",
                    "rows": [], "total_hint": total_hint,
                }
            return {
                "insurer": insurer, "file": filename, "status": "ok",
                "reason": "", "rows": rows, "total_hint": total_hint,
            }

        if is_mirrored_text(full_text):
            reversed_text = full_text[::-1]
            total_hint = find_total_in_text(reversed_text)
            return {
                "insurer": insurer, "file": filename, "status": "sonderformat",
                "reason": "PDF-Text ist pro Seite gespiegelt/verdreht (bekannter "
                          "Sonderfall bei diesem Versicherer-Layout). "
                          "Automatische Kundenzuordnung in dieser Version nicht "
                          "moeglich - bitte manuell pruefen.",
                "rows": [], "total_hint": total_hint,
            }

        if insurer_lower == "dialog" and OCR_AVAILABLE:
            # Dialog 'Provisionseinzelnachweis': dichte Tabelle mit
            # gepunktetem Wasserzeichen-Hintergrund, der normale OCR fast
            # vollstaendig verwirrt - eigener zeilenweiser Zuschnitt-Pfad
            # mit Median-Filter, siehe extract_dialog(). Ergebnis wird
            # zwingend gegen die "Summe gesamt" von der PG-Uebersichtsseite
            # geprueft (find_dialog_total_hint) - bei Abweichung faellt die
            # Datei automatisch auf manuelle Pruefung zurueck statt falsche
            # Zahlen zu liefern (gleiches Muster wie extract_continentale()).
            rows = extract_dialog(pdf.pages)
            target_total = find_dialog_total_hint(pdf)
            rows = repair_signs_to_control(rows, target_total)
            extracted_sum = round(sum(r[2] for r in rows), 2)
            # Toleranz 0,10 EUR: einzelne verlesene Ziffern in der (durch das
            # Wasserzeichen teils verdeckten) Summenzeile eines Blocks sind
            # kein Zuordnungsrisiko (beobachtet Juli-2026: 38,94 als 38,99 ->
            # 0,05 EUR Abweichung), eine groessere Abweichung dagegen schon.
            if not rows or target_total is None or abs(extracted_sum - target_total) > 0.10:
                return {
                    "insurer": insurer, "file": filename, "status": "sonderformat",
                    "reason": "OCR-Ergebnis der Kundenpositionen stimmt nicht "
                              "(oder konnte nicht geprueft werden) mit der im "
                              "PDF aufgedruckten 'Summe gesamt' ueberein "
                              f"(extrahiert: {extracted_sum}, Soll: {target_total}) "
                              "- bitte manuell pruefen, um keine falschen "
                              "Kundenzuordnungen zu riskieren.",
                    "rows": [], "total_hint": target_total,
                }
            return {
                "insurer": insurer, "file": filename, "status": "ok",
                "reason": "OCR verwendet, gegen 'Summe gesamt' verifiziert"
                          + ("" if abs(extracted_sum - target_total) <= 0.01 else
                             f" (Abweichung {round(extracted_sum - target_total, 2)} EUR)"),
                "rows": rows, "total_hint": target_total,
            }

        if insurer_lower == "continentale" and OCR_AVAILABLE:
            # Continentale 'Provisionsnote Einzelergebnisse': sehr dichte,
            # OCR-mehrdeutige Tabelle - siehe extract_continentale(). Das
            # Ergebnis wird zwingend gegen den aufgedruckten "Neuer Saldo"
            # geprueft; bei Abweichung faellt die Datei automatisch auf
            # manuelle Pruefung zurueck statt falsche Zahlen zu liefern.
            rows, target_total, all_majority = extract_continentale(pdf)
            extracted_sum = round(sum(r[2] for r in rows), 2)
            # "Neuer Saldo" enthaelt zusaetzlich Alt-Saldo/Stornoreserve-
            # Bewegungen und ist deshalb nicht in jedem Monat identisch mit
            # der Summe der Einzelpositionen (September-2026: Positionen
            # 30,23, Neuer Saldo 32,63; August: 30,23 vs. 60,46). Zusaetzlich
            # zur exakten Uebereinstimmung wird daher akzeptiert, wenn JEDE
            # Zeile von mindestens zwei der drei OCR-Aufloesungen identisch
            # gelesen wurde und die Positionssumme den Saldo nicht
            # uebersteigt - dann ist der Saldo-Unterschied ein Vortrag, kein
            # Lesefehler (Hinweis steht im Kontrolle-Blatt).
            carry_over_ok = (
                bool(rows) and target_total is not None and all_majority
                and 0 < extracted_sum <= target_total + 0.01
            )
            if not rows or target_total is None or (
                abs(extracted_sum - target_total) > 0.01 and not carry_over_ok
            ):
                return {
                    "insurer": insurer, "file": filename, "status": "sonderformat",
                    "reason": "OCR-Ergebnis der Kundenpositionen stimmt nicht "
                              "(oder konnte nicht geprueft werden) mit dem im "
                              "PDF aufgedruckten 'Neuer Saldo' ueberein "
                              f"(extrahiert: {extracted_sum}, Soll: {target_total}) "
                              "- bitte manuell pruefen, um keine falschen "
                              "Kundenzuordnungen zu riskieren.",
                    "rows": [], "total_hint": target_total,
                }
            return {
                "insurer": insurer, "file": filename, "status": "ok",
                "reason": ("OCR verwendet, gegen 'Neuer Saldo' verifiziert"
                           if abs(extracted_sum - target_total) <= 0.01 else
                           f"OCR verwendet (Mehrheitsentscheid ueber 3 Aufloesungen); "
                           f"'Neuer Saldo' {target_total} enthaelt Vortrag/Stornoreserve "
                           f"({round(target_total - extracted_sum, 2)} EUR mehr als die Positionen)"),
                "rows": rows, "total_hint": target_total,
            }

        if insurer_lower == "barmenia" and OCR_AVAILABLE:
            # Barmenia/Gothaer 'Vergluetungsnachweis': gescanntes PDF, dessen
            # Detailtabelle mit dem Standard-OCR-Pfad (300dpi, wortpositions-
            # basiert) die Betragsspalten verliert - eigener Pfad mit
            # hoeherer Aufloesung, siehe extract_barmenia(). Die "Laufende &
            # Folgeprovisionen"-Seiten haben eine ANDERE, noch dichtere
            # Tabellenvorlage und brauchen eine eigene zeilenweise Zuschnitt-
            # Extraktion, siehe extract_barmenia_laufende().
            rows = (extract_barmenia(pdf) + extract_barmenia_abschluss_gl(pdf)
                    + extract_barmenia_laufende(pdf))
            # Kontrollsumme: Gesamtergebnis Abschluss- + Laufende Verguetungen
            # der Gesamtuebersicht (Kontrolle-Blatt zeigt die Differenz).
            b_abschluss, b_laufende = find_barmenia_overview_totals(pdf)
            if b_abschluss is not None and b_laufende is not None:
                total_hint = round(b_abschluss + b_laufende, 2)
            else:
                total_hint = find_total_in_text(full_text)
            if not rows:
                return {
                    "insurer": insurer, "file": filename, "status": "sammelbeleg",
                    "reason": "Kein Kunden-Positionsdetail im PDF gefunden - "
                              "vermutlich reiner Sammelbeleg/Kontoauszug ohne "
                              "Einzelaufstellung.",
                    "rows": [], "total_hint": total_hint,
                }
            return {
                "insurer": insurer, "file": filename, "status": "ok",
                "reason": "OCR verwendet", "rows": rows, "total_hint": total_hint,
            }

        if ("huasbesitzer" in insurer_lower or "hausbesitzer" in insurer_lower) and OCR_AVAILABLE:
            # Bayerische Hausbesitzer-Versicherung: gescanntes Einzelblatt pro
            # Vertrag, siehe extract_bayerische_hausbesitzer(). Ergebnis wird
            # zwingend gegen den aufgedruckten "Provisionsanspruch von
            # insgesamt" geprueft (gleiches Muster wie extract_dialog()).
            rows, target_total = extract_bayerische_hausbesitzer(pdf)
            extracted_sum = round(sum(r[2] for r in rows), 2)
            if not rows or target_total is None or abs(extracted_sum - target_total) > 0.01:
                return {
                    "insurer": insurer, "file": filename, "status": "sonderformat",
                    "reason": "OCR-Ergebnis der Kundenposition(en) stimmt nicht "
                              "(oder konnte nicht geprueft werden) mit dem im "
                              "PDF aufgedruckten 'Provisionsanspruch von "
                              f"insgesamt' ueberein (extrahiert: {extracted_sum}, "
                              f"Soll: {target_total}) - bitte manuell pruefen, "
                              "um keine falschen Kundenzuordnungen zu riskieren.",
                    "rows": [], "total_hint": target_total,
                }
            return {
                "insurer": insurer, "file": filename, "status": "ok",
                "reason": "OCR verwendet, gegen 'Provisionsanspruch von "
                          "insgesamt' verifiziert",
                "rows": rows, "total_hint": target_total,
            }

        if insurer_lower == "helvetia" and OCR_AVAILABLE:
            # Helvetia 'Vermittlerliste Komposit': eigener Pfad, siehe
            # extract_helvetia(); nur uebernommen, wenn die Summe jeder Seite
            # zur gedruckten Kontrollsumme passt - sonst unveraendert der
            # generische Pfad weiter unten.
            h_rows, h_control, h_ok = extract_helvetia(pdf)
            if h_rows and h_ok:
                return {
                    "insurer": insurer, "file": filename, "status": "ok",
                    "reason": "OCR verwendet, je Liste gegen 'Summe' verifiziert",
                    "rows": h_rows, "total_hint": h_control,
                }

        if ("baloise" in insurer_lower or "basler" in insurer_lower) and OCR_AVAILABLE:
            # Baloise/Basler 'Kontoauszug': eigener Pfad mit Median-Filter,
            # siehe extract_baloise(). Kontrolle gegen die ausgezahlte Summe
            # (Buchungsart 420, 'Prov-Zahlung') unter Beruecksichtigung von
            # Saldo-Vortrag/End-Saldo.
            rows, payout, vortrag, endsaldo = extract_baloise(pdf)
            if rows:
                extracted_sum = round(sum(r[2] for r in rows), 2)
                expected = [payout] if payout is not None else []
                if payout is not None:
                    expected.append(round(payout + endsaldo - vortrag, 2))
                ok = any(abs(extracted_sum - e) <= 0.05 for e in expected)
                if ok:
                    return {
                        "insurer": insurer, "file": filename, "status": "ok",
                        "reason": "OCR verwendet, gegen ausgezahlte Summe verifiziert",
                        "rows": rows, "total_hint": payout,
                    }
                return {
                    "insurer": insurer, "file": filename, "status": "sonderformat",
                    "reason": "OCR-Ergebnis der Kundenpositionen stimmt nicht "
                              "(oder konnte nicht geprueft werden) mit der "
                              "ausgezahlten Summe ueberein "
                              f"(extrahiert: {extracted_sum}, Soll: {payout}) "
                              "- bitte manuell pruefen.",
                    "rows": [], "total_hint": payout,
                }

        if insurer_lower == "arag" and OCR_AVAILABLE:
            # ARAG-Einzelaufstellung: gescanntes PDF, dessen Tabelle mit dem
            # Standard-OCR-Pfad (300dpi, wortpositionsbasiert) zu schlecht
            # erkannt wird - eigener Pfad mit hoeherer Aufloesung, siehe
            # extract_arag().
            rows = extract_arag(pdf)
            total_hint = find_total_in_text(full_text)
            if not rows:
                return {
                    "insurer": insurer, "file": filename, "status": "sammelbeleg",
                    "reason": "Kein Kunden-Positionsdetail im PDF gefunden - "
                              "vermutlich reiner Sammelbeleg/Kontoauszug ohne "
                              "Einzelaufstellung.",
                    "rows": [], "total_hint": total_hint,
                }
            return {
                "insurer": insurer, "file": filename, "status": "ok",
                "reason": "OCR verwendet", "rows": rows, "total_hint": total_hint,
            }

        target_month_num = None
        for m_idx, m_name in enumerate(MONTHS_DE, start=1):
            if month_folder.lower().startswith(m_name.lower()):
                target_month_num = m_idx
                break

        pages_words = []
        used_ocr = False
        for pidx, page in enumerate(pdf.pages):
            page_full_text = page.extract_text() or ""
            page_text_start = page_full_text[:60].lower()
            if "enstandsliste" in page_text_start or "vm-konto" in page_text_start:
                # "Außenstandsliste" (z.B. Mannheimer): Liste noch nicht
                # bezahlter/offener Posten, noch keine realisierte Courtage.
                # "VM-Konto" (z.B. Itzehoer): Kontobewegungs-/Saldo-Uebersicht
                # in einem anderen, oft zeichenweise gespreizten Layout ohne
                # eigene Kopfzeilen-Erkennung - wuerde sonst mit der zuletzt
                # erkannten Spaltenaufteilung Muell als Buchungszeilen liefern.
                pages_words.append((pidx, [], "text"))
                continue

            words, source = get_words_for_page(page)
            if source == "ocr":
                used_ocr = True
            elif source == "text" and "itzehoer" in insurer_lower:
                words = _merge_touching_words(words)

            if target_month_num is not None:
                # Manche PDFs, die im Monatsordner X liegen, enthalten eine
                # eingebettete Teilabrechnung fuer einen ANDEREN Monat (z.B.
                # SV Sparkassenversicherung: der Juni-PDF-Ordner enthaelt auch
                # einen "vorlaeufige Buchungen"-Abschnitt fuer 01.07.-05.07.).
                # Anhand von "Vom: DD.MM.YYYY" erkennen und ausschliessen,
                # damit diese Betraege nicht dem falschen Monat zugerechnet
                # werden. Bei gescannten Seiten (kein Text-Layer) aus den
                # ersten OCR-Woertern rekonstruieren.
                #
                # WICHTIG: "Erstellungsdatum:" eignet sich NICHT fuer diesen
                # Zweck (Versuch verworfen) - bei Alte Leipziger steht auf
                # ALLEN Seiten eines mehrseitigen Sammel-Scans dasselbe
                # Erstellungsdatum (Druckdatum des gesamten Stapels, oft
                # erst Anfang des Folgemonats), unabhaengig vom tatsaechlichen
                # Buchungsdatum der einzelnen Positionen - haette sonst ganze
                # Dateien faelschlich auf 0 gefiltert.
                period_text = page_full_text[:300]
                if not period_text.strip():
                    period_text = " ".join(w["text"] for w in words[:100])
                m = re.search(r"Vom:?\s*(\d{2})\.(\d{2})\.(\d{4})", period_text)
                if m and int(m.group(2)) != target_month_num:
                    pages_words.append((pidx, [], "text"))
                    continue

            pages_words.append((pidx, words, source))

        if insurer_lower in SIGN_INVERTED_INSURERS:
            # AIG, Hiscox und Mannheimer drucken ihre Betragsspalte aus Sicht
            # des Versicherers (dessen eigenes Soll/Haben), nicht aus Sicht
            # von SSH als Makler - ein beim Versicherer negativer Saldo ist
            # fuer uns eine Gutschrift und umgekehrt (Nutzer-Bestaetigung).
            # Bei AIG zusaetzlich explizit belegt durch die Summenzeile
            # "Summe Vermittler (Provision/Courtage zu Ihren Gunsten): ...
            # 112,39-". Vorzeichen daher fuer alle drei explizit umdrehen.
            rows = [(p, k, -amt, l, s) for (p, k, amt, l, s) in extract_generic_table(pages_words)]
        elif "alte-leipziger" in insurer_lower:
            rows = extract_alte_leipziger(pages_words)
        elif "fondsfinanz" in insurer_lower:
            rows = extract_fondsfinanz(pages_words)
        elif "vhv" in insurer_lower:
            rows = extract_vhv(pages_words)
        elif "swiss" in insurer_lower:
            # Swiss Life liefert zwei Formate im selben PDF: eine normale
            # Tabelle (Bestandspflege) und ein Block-Format
            # (Einzelaufstellung, "VN/VP ... Summe Vertrag").
            rows = extract_generic_table(pages_words) + extract_vn_summevertrag_block(pages_words)
            vsv = extract_swiss_life_vsv_deduction(list(enumerate(full_text_parts)))
            if vsv:
                vsv_page, vsv_amount, vsv_line = vsv
                rows.append((
                    vsv_page,
                    "Vertrauensschadenversicherung-Beitrag (nicht kundenspezifisch, siehe Abrechnungsuebersicht)",
                    round(vsv_amount, 2), vsv_line, "text",
                ))
        elif "amex" in insurer_lower:
            rows = extract_generic_table(pages_words)
            auszahlung = extract_amex_stornoreserve_auszahlung(list(enumerate(full_text_parts)))
            if auszahlung:
                a_page, a_amount, a_line = auszahlung
                rows.append((
                    a_page,
                    f"Auszahlung freie Stornoreserve ({NICHT_ZUORDENBAR_MARKER}, "
                    "siehe Sonstige Positionen)",
                    round(a_amount, 2), a_line, "text",
                ))
        elif "sparkassenversicherung" in insurer_lower:
            rows = extract_sv_sparkasse(pages_words)
        elif "domcura" in insurer_lower:
            rows = extract_domcura(pages_words)
        elif "ergo" in insurer_lower:
            rows = extract_ergo(pages_words)
        else:
            rows = extract_generic_table(pages_words)

        total_hint = find_total_in_text(full_text)
        if total_hint is not None and insurer_lower in SIGN_INVERTED_INSURERS:
            total_hint = -total_hint
        if not full_text.strip() and used_ocr:
            ocr_text_parts = []
            for pidx, words, source in pages_words:
                ocr_text_parts.append(" ".join(w["text"] for w in words))
            total_hint = find_total_in_text("\n".join(ocr_text_parts)) or total_hint

        if not rows:
            return {
                "insurer": insurer, "file": filename, "status": "sammelbeleg",
                "reason": "Kein Kunden-Positionsdetail im PDF gefunden - "
                          "vermutlich reiner Sammelbeleg/Kontoauszug ohne "
                          "Einzelaufstellung.",
                "rows": [], "total_hint": total_hint,
            }

        return {
            "insurer": insurer, "file": filename, "status": "ok",
            "reason": "OCR verwendet" if used_ocr else "",
            "rows": rows, "total_hint": total_hint,
        }
    finally:
        pdf.close()


def process_csv_file(filepath, month_folder):
    """Verarbeitet eine CSV-Datei. Bisher zwei bekannte Formate, beide per
    Spaltenpruefung (nicht Dateiname) unterschieden, siehe extract_vema_csv()
    und extract_covomo_csv(). Liefert dasselbe Ergebnis-Dict-Format wie
    process_file(), damit process_files() PDFs und CSVs einheitlich
    behandeln kann."""
    filename = os.path.basename(filepath)
    rows = extract_vema_csv(filepath)
    if rows:
        return {
            "insurer": "VEMA-Pool", "file": filename, "status": "ok",
            "reason": "", "rows": rows, "total_hint": None,
        }

    rows = extract_covomo_csv(filepath)
    if rows:
        insurer = insurer_name_from_filename(filepath, month_folder)
        return {
            "insurer": insurer, "file": filename, "status": "ok",
            "reason": "", "rows": rows, "total_hint": None,
        }

    return {
        "insurer": "VEMA-Pool", "file": filename, "status": "sammelbeleg",
        "reason": "Keine erkennbaren Buchungszeilen in der CSV-Datei "
                  "gefunden (falsches Spaltenformat?).",
        "rows": [], "total_hint": None,
    }


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def process_files(files, month_label, progress_callback=None):
    """Verarbeitet eine Liste von PDF- und/oder CSV-Pfaden und liefert die
    vier Ergebnis-Tabellen (df_rows, df_control, df_agg, df_problem) als
    DataFrames. CSV-Dateien (z.B. VEMA-Poolabrechnung-N.csv) werden ueber
    process_csv_file() statt process_file() verarbeitet, siehe
    extract_vema_csv().

    progress_callback (optional): wird nach jeder Datei mit
    (index, gesamt, dateiname) aufgerufen - fuer Fortschrittsanzeigen in
    einer Oberflaeche (CLI-print oder z.B. Streamlit-Progressbar)."""
    all_rows = []
    aggregate_rows = []
    problem_rows = []
    control_rows = []

    for i, f in enumerate(files):
        if progress_callback:
            progress_callback(i, len(files), os.path.basename(f))
        if f.lower().endswith(".csv"):
            result = process_csv_file(f, month_label)
        else:
            result = process_file(f, month_label)
        insurer = result["insurer"]
        filename = result["file"]
        status = result["status"]

        if status == "ok":
            for page_idx, kunde, provision, raw_line, source in result["rows"]:
                all_rows.append({
                    "Versicherer": insurer,
                    "Kunde": kunde,
                    "Provision": provision,
                    "Datei": filename,
                    "Seite": page_idx + 1,
                    "Quelle": {"ocr": "OCR", "csv": "CSV"}.get(source, "Text"),
                    "Rohzeile": raw_line,
                })
            extracted_sum = sum(r[2] for r in result["rows"])
            control_rows.append({
                "Versicherer": insurer, "Datei": filename,
                "Anzahl_Buchungen": len(result["rows"]),
                "Summe_extrahiert": round(extracted_sum, 2),
                "Summe_lt_PDF": result["total_hint"],
                "Differenz": (round(extracted_sum - result["total_hint"], 2)
                              if result["total_hint"] is not None else None),
                "Hinweis": result["reason"],
            })
        elif status == "sammelbeleg":
            aggregate_rows.append({
                "Versicherer": insurer, "Datei": filename,
                "Betrag_lt_PDF": result["total_hint"],
                "Hinweis": result["reason"],
            })
        elif status == "sonderformat":
            problem_rows.append({
                "Versicherer": insurer, "Datei": filename,
                "Betrag_lt_PDF": result["total_hint"],
                "Grund": result["reason"],
            })
        else:
            problem_rows.append({
                "Versicherer": insurer, "Datei": filename,
                "Betrag_lt_PDF": None,
                "Grund": result["reason"],
            })

    df_rows = pd.DataFrame(all_rows)
    df_agg = pd.DataFrame(aggregate_rows)
    df_problem = pd.DataFrame(problem_rows)
    df_control = pd.DataFrame(control_rows)
    return df_rows, df_control, df_agg, df_problem


def apply_betreuer(df_rows, lookup):
    """Ergaenzt df_rows um eine Spalte 'Betreuer' (siehe match_betreuer(),
    MANUAL_BETREUER_OVERRIDES_RAW) sowie je eine Spalte pro Partner
    (PARTNER_NAMES) mit dem jeweiligen Betrag der Zeile (sonst leer) - fuer
    die farbige Kunde_Provision-Ansicht bzw. die PDF-Uebersicht. Optionaler
    Schritt: wird nur aufgerufen, wenn eine Betreuer.xlsx verfuegbar ist
    (siehe main()/app.py) - ohne diesen Aufruf bleibt df_rows unveraendert."""
    if df_rows is None or df_rows.empty:
        return df_rows
    df_rows = df_rows.copy()

    overrides = {
        normalize_for_match(strip_trailing_date(k)): v
        for k, v in MANUAL_BETREUER_OVERRIDES_RAW.items()
    }

    def resolve(kunde):
        kunde_stripped = strip_trailing_date(kunde)
        override = overrides.get(normalize_for_match(kunde_stripped))
        return override if override else match_betreuer(kunde_stripped, lookup)

    df_rows["Betreuer"] = df_rows["Kunde"].apply(resolve)

    # Nicht-kundenspezifische Sammelabzuege (Stornoreserve, VSV-Beitrag,
    # siehe COLLECTIVE_DEDUCTION_MARKER) koennen nicht ueber einen Namen
    # zugeordnet werden - sie werden stattdessen automatisch dem Betreuer
    # der betragsmaessig groessten (bereits zugeordneten) Kundenposition
    # derselben Abrechnungsdatei zugerechnet (Nutzer-Beispiel: die
    # Fondsfinanz-Stornoreserve gehoert zu Robin Heckmann, weil dessen
    # Kunde Heckmann, Oliver die mit Abstand groesste Position in
    # derselben Abrechnung ist).
    collective_mask = df_rows["Kunde"].str.contains(COLLECTIVE_DEDUCTION_MARKER, na=False)
    for idx in df_rows[collective_mask].index:
        datei = df_rows.at[idx, "Datei"]
        same_file = df_rows[(df_rows["Datei"] == datei) & ~collective_mask & df_rows["Betreuer"].notna()]
        if not same_file.empty:
            top_row = same_file.loc[same_file["Provision"].abs().idxmax()]
            df_rows.at[idx, "Betreuer"] = top_row["Betreuer"]

    for partner in PARTNER_NAMES:
        df_rows[partner] = df_rows.apply(
            lambda r, p=partner: r["Provision"] if r["Betreuer"] == p else None, axis=1
        )
    return df_rows


def write_excel(df_rows, df_control, df_agg, df_problem, out_target, df_bank_unmatched=None,
                df_bank_explained=None):
    """Schreibt die Ergebnis-Tabellen als (optisch aufbereitete) Excel-Datei
    mit den vier Standard-Blaettern (siehe README: Aufbau der Ausgabe-Excel).
    out_target ist entweder ein Dateipfad oder ein Datei-artiges Objekt
    (z.B. io.BytesIO fuer die Web-Oberflaeche). df_bank_unmatched (optional):
    Ergebnis von reconcile_bank_credits() als DataFrame - wird als fuenftes
    Blatt "Fehlende_Abrechnungen" ergaenzt, falls angegeben."""
    if isinstance(out_target, (str, os.PathLike)):
        out_dir = os.path.dirname(out_target)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)

    has_betreuer = df_rows is not None and "Betreuer" in df_rows.columns
    kunde_cols = ["Versicherer", "Kunde", "Provision", "Datei", "Seite", "Quelle", "Rohzeile"]
    kunde_currency = {"Provision"}
    if has_betreuer:
        kunde_cols = kunde_cols + ["Betreuer"] + PARTNER_NAMES
        kunde_currency = kunde_currency | set(PARTNER_NAMES)

    sheets = [
        ("Kunde_Provision", df_rows, kunde_cols, kunde_currency),
        ("Kontrolle", df_control,
         ["Versicherer", "Datei", "Anzahl_Buchungen", "Summe_extrahiert",
          "Summe_lt_PDF", "Differenz", "Hinweis"],
         {"Summe_extrahiert", "Summe_lt_PDF", "Differenz"}),
        ("Sammelbelege_ohne_Details", df_agg,
         ["Versicherer", "Datei", "Betrag_lt_PDF", "Hinweis"],
         {"Betrag_lt_PDF"}),
        ("Manuelle_Pruefung", df_problem,
         ["Versicherer", "Datei", "Betrag_lt_PDF", "Grund"],
         {"Betrag_lt_PDF"}),
    ]
    if df_bank_unmatched is not None:
        sheets.append((
            "Fehlende_Abrechnungen", df_bank_unmatched,
            ["datum", "betrag", "sender", "verwendungszweck"],
            {"betrag"},
        ))
    if df_bank_explained is not None and not df_bank_explained.empty:
        sheets.append((
            "Zahlung_Vormonat", df_bank_explained,
            ["datum", "betrag", "sender", "erklaert_durch"],
            {"betrag"},
        ))
    if has_betreuer:
        unmatched_mask = df_rows["Betreuer"].isna() & ~df_rows["Kunde"].str.contains(
            NICHT_ZUORDENBAR_MARKER, na=False
        )
        df_unmatched = df_rows[unmatched_mask][
            ["Versicherer", "Kunde", "Provision", "Datei"]
        ].drop_duplicates()
        sheets.append((
            "Kunde_ohne_Betreuer", df_unmatched,
            ["Versicherer", "Kunde", "Provision", "Datei"],
            {"Provision"},
        ))

    df_insurer_sum, df_partner_sum = _build_summary_frames(df_rows)

    with pd.ExcelWriter(out_target, engine="openpyxl") as writer:
        df_insurer_sum.to_excel(writer, sheet_name="Zusammenfassung", index=False, startrow=0)
        _style_worksheet(writer.sheets["Zusammenfassung"], df_insurer_sum,
                          ["Versicherer", "Summe"], {"Summe"})
        if df_partner_sum is not None:
            _write_partner_summary(writer.sheets["Zusammenfassung"], df_insurer_sum, df_partner_sum)

        for sheet_name, df, default_cols, currency_cols in sheets:
            (df if not df.empty else pd.DataFrame(columns=default_cols)).to_excel(
                writer, sheet_name=sheet_name, index=False, columns=default_cols
            )
            _style_worksheet(writer.sheets[sheet_name], df, default_cols, currency_cols)
            if sheet_name == "Kunde_Provision" and has_betreuer:
                _color_betreuer_rows(writer.sheets[sheet_name], df, default_cols)


def _build_summary_frames(df_rows):
    """Baut die beiden Tabellen fuer das Blatt 'Zusammenfassung': Summe je
    Versicherer (immer) sowie Summe je Partner inkl. 'Nicht zugeordnet'
    (nur wenn eine Betreuer-Zuordnung vorliegt, sonst None).

    NICHT_ZUORDENBAR_MARKER-Zeilen (z.B. Amex-Pool-Stornoreserve-Auszahlung)
    fliessen NICHT in diese Summen ein - sie wurden beim urspruenglichen
    Buchungsdatum bereits voll verbucht/einem Partner zugerechnet (siehe
    NICHT_ZUORDENBAR_MARKER-Definition), die spaetere Auszahlung darf den
    Monatsumsatz daher nicht ein zweites Mal erhoehen (Nutzer-Bestaetigung:
    die eigene Monatsuebersicht fuehrt diese Auszahlung gar nicht erst als
    Position)."""
    if df_rows is None or df_rows.empty:
        return pd.DataFrame(columns=["Versicherer", "Summe"]), None

    df_summary_rows = df_rows[~df_rows["Kunde"].str.contains(NICHT_ZUORDENBAR_MARKER, na=False)]

    df_insurer_sum = (
        df_summary_rows.groupby("Versicherer", as_index=False)["Provision"].sum()
        .rename(columns={"Provision": "Summe"})
        .sort_values("Summe", ascending=False)
    )

    df_partner_sum = None
    if "Betreuer" in df_rows.columns:
        rows = [{"Partner": p, "Summe": df_summary_rows.loc[df_summary_rows["Betreuer"] == p, "Provision"].sum()}
                for p in PARTNER_NAMES]
        rows.append({"Partner": "Nicht zugeordnet",
                      "Summe": df_summary_rows.loc[df_summary_rows["Betreuer"].isna(), "Provision"].sum()})
        df_partner_sum = pd.DataFrame(rows)

    return df_insurer_sum, df_partner_sum


def _write_partner_summary(ws, df_insurer_sum, df_partner_sum):
    """Schreibt die 'Summe je Partner'-Tabelle unterhalb der Versicherer-
    Summen in dasselbe Blatt (eigene Kopfzeile/Formatierung, da
    _style_worksheet() von genau einer Tabelle ab Zeile 1 ausgeht)."""
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    start_row = len(df_insurer_sum) + 3
    ws.cell(row=start_row, column=1, value="Partner")
    ws.cell(row=start_row, column=2, value="Summe")

    header_fill = PatternFill(start_color="FF1F4E78", end_color="FF1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    for col_idx in (1, 2):
        cell = ws.cell(row=start_row, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(vertical="center")

    fills = {p: PatternFill(start_color=c, end_color=c, fill_type="solid")
             for p, c in PARTNER_COLORS_XLSX.items()}
    for i, row in enumerate(df_partner_sum.itertuples(), start=1):
        row_idx = start_row + i
        ws.cell(row=row_idx, column=1, value=row.Partner)
        cell = ws.cell(row=row_idx, column=2, value=row.Summe)
        cell.number_format = '#,##0.00 "€"'
        fill = fills.get(row.Partner)
        if fill:
            ws.cell(row=row_idx, column=1).fill = fill
            ws.cell(row=row_idx, column=2).fill = fill

    ws.column_dimensions[get_column_letter(1)].width = max(ws.column_dimensions[get_column_letter(1)].width or 10, 20)
    ws.column_dimensions[get_column_letter(2)].width = max(ws.column_dimensions[get_column_letter(2)].width or 10, 16)


def _color_betreuer_rows(ws, df, columns):
    """Faerbt jede Zeile der Kunde_Provision-Tabelle je nach zugeordnetem
    Betreuer ein (siehe PARTNER_COLORS_XLSX, Nutzer-Vorgabe Rot/Gelb/Blau) -
    so ist beim Ueberfliegen der Tabelle auf einen Blick erkennbar, wer
    fuer welchen Umsatz verantwortlich ist. Zeilen ohne Zuordnung
    (Betreuer leer) bleiben ungefaerbt (siehe Kunde_ohne_Betreuer-Blatt)."""
    from openpyxl.styles import PatternFill

    if df is None or df.empty or "Betreuer" not in df.columns:
        return
    fills = {p: PatternFill(start_color=c, end_color=c, fill_type="solid")
             for p, c in PARTNER_COLORS_XLSX.items()}
    n_cols = len(columns)
    for row_idx, betreuer in enumerate(df["Betreuer"], start=2):
        fill = fills.get(betreuer)
        if not fill:
            continue
        for col_idx in range(1, n_cols + 1):
            ws.cell(row=row_idx, column=col_idx).fill = fill


def _style_worksheet(ws, df, columns, currency_cols):
    """Optische Aufbereitung eines Ergebnis-Blatts: fette, farbige
    Kopfzeile, fixierte erste Zeile, Auto-Filter, sinnvolle Spaltenbreiten,
    Waehrungsformat fuer Betragsspalten sowie eine rote Hervorhebung
    abweichender "Differenz"-Werte im Kontrolle-Blatt."""
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.formatting.rule import CellIsRule
    from openpyxl.utils import get_column_letter

    n_rows = len(df) if df is not None and not df.empty else 0
    n_cols = len(columns)
    if n_cols == 0:
        return

    header_fill = PatternFill(start_color="FF1F4E78", end_color="FF1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    for col_idx in range(1, n_cols + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(vertical="center")

    ws.freeze_panes = "A2"
    if n_rows > 0:
        ws.auto_filter.ref = f"A1:{get_column_letter(n_cols)}{n_rows + 1}"

    for col_idx, col_name in enumerate(columns, start=1):
        letter = get_column_letter(col_idx)
        if n_rows > 0 and col_name in df.columns:
            longest_value = df[col_name].map(lambda v: len(str(v))).max()
        else:
            longest_value = 0
        width = max(len(str(col_name)), longest_value) + 2
        ws.column_dimensions[letter].width = min(max(width, 10), 60)

        if col_name in currency_cols:
            for row_idx in range(2, n_rows + 2):
                ws.cell(row=row_idx, column=col_idx).number_format = '#,##0.00 "€"'

    if n_rows > 0 and "Differenz" in columns:
        diff_letter = get_column_letter(columns.index("Differenz") + 1)
        rng = f"{diff_letter}2:{diff_letter}{n_rows + 1}"
        red_fill = PatternFill(start_color="FFFFC7CE", end_color="FFFFC7CE", fill_type="solid")
        ws.conditional_formatting.add(
            rng,
            CellIsRule(operator="notEqual", formula=["0"], fill=red_fill),
        )


def _fmt_eur(x):
    """Deutsches Zahlenformat (Komma als Dezimaltrennzeichen, Punkt als
    Tausendertrennzeichen), z.B. 1234.5 -> '1.234,50 EUR'. 'EUR' statt '€'
    ausgeschrieben, da die Standard-PDF-Schrift (helvetica) das
    Euro-Zeichen nicht unterstuetzt."""
    return f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") + " EUR"


def _pdf_safe_text(s):
    """Macht einen String sicher fuer die fpdf-Kernschriftart 'helvetica'
    (nur Latin-1/WinAnsi). Zwei Faelle:
    1. NFC-Normalisierung: macOS liefert Umlaute in Datei-/manchmal auch
       Formularnamen standardmaessig in NFD (z.B. 'u' + freistehendes
       Kombinations-Trema statt des einzelnen Zeichens 'ü') - visuell
       identisch, aber Latin-1 kennt das freistehende Kombinationszeichen
       nicht. NFC setzt das wieder zum einzelnen Zeichen zusammen, das
       Latin-1 abdeckt (beobachtet: FPDFUnicodeEncodingException beim
       Hochladen von einem Mac).
    2. Fallback fuer alles, was auch nach NFC nicht in Latin-1 passt (z.B.
       Emoji oder sonstiger OCR-Ausschuss in Kundennamen) - wird durch '?'
       ersetzt, statt die komplette (rein informative) PDF-Uebersicht
       abstuerzen zu lassen; das Excel-Ergebnis ist davon nie betroffen."""
    s = unicodedata.normalize("NFC", str(s))
    try:
        s.encode("latin-1")
        return s
    except UnicodeEncodeError:
        return s.encode("latin-1", errors="replace").decode("latin-1")


def build_summary_pdf(df_rows, out_target, month_label):
    """Erstellt eine einseitige (bzw. mehrseitige) PDF-Uebersicht: je
    Versicherer absteigend nach Gesamtumsatz sortiert, mit der jeweiligen
    Kunden-/Vertragsaufstellung (ebenfalls absteigend nach Betrag) direkt
    daneben. Rein informativ/optisch - keine neue Extraktionslogik, nutzt
    nur die bereits extrahierten Zeilen aus process_files().

    out_target: Dateipfad (str) oder datei-artiges Objekt (z.B. io.BytesIO
    fuer die Web-Oberflaeche)."""
    from fpdf import FPDF

    pdf = FPDF(format="A4", unit="mm")
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("helvetica", "B", 16)
    pdf.cell(0, 10, f"Courtage-Uebersicht {_pdf_safe_text(month_label)}", new_x="LMARGIN", new_y="NEXT")

    if df_rows is None or df_rows.empty:
        pdf.set_font("helvetica", "", 11)
        pdf.cell(0, 8, "Keine Buchungszeilen vorhanden.")
        pdf.output(out_target)
        return

    has_betreuer = "Betreuer" in df_rows.columns
    if has_betreuer:
        per_customer = df_rows.groupby(["Versicherer", "Kunde", "Betreuer"], as_index=False,
                                        dropna=False)["Provision"].sum()
    else:
        per_customer = df_rows.groupby(["Versicherer", "Kunde"], as_index=False)["Provision"].sum()
        per_customer["Betreuer"] = None
    insurer_totals = per_customer.groupby("Versicherer")["Provision"].sum().sort_values(ascending=False)

    pdf.set_font("helvetica", "", 11)
    pdf.cell(0, 7, f"Gesamtsumme: {_fmt_eur(insurer_totals.sum())}", new_x="LMARGIN", new_y="NEXT")

    if has_betreuer:
        pdf.set_font("helvetica", "", 9)
        y_legend = pdf.get_y() + 1
        x_legend = pdf.l_margin
        for partner in PARTNER_NAMES:
            r, g, b = PARTNER_SWATCH_COLORS_PDF[partner]
            pdf.set_fill_color(r, g, b)
            pdf.rect(x_legend, y_legend + 0.5, 4, 4, "F")
            pdf.set_xy(x_legend + 5, y_legend)
            pdf.set_text_color(0, 0, 0)
            pdf.cell(55, 5, partner)
            x_legend += 60
        pdf.set_y(y_legend + 6)

    pdf.ln(3)

    pdf.set_font("helvetica", "B", 11)
    pdf.cell(0, 7, "Summe je Versicherer", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("helvetica", "", 9)
    for versicherer, total in insurer_totals.items():
        pdf.cell(100, 5, _pdf_safe_text(versicherer))
        pdf.cell(0, 5, _fmt_eur(total), new_x="LMARGIN", new_y="NEXT")
    pdf.ln(3)

    if has_betreuer:
        pdf.set_font("helvetica", "B", 11)
        pdf.cell(0, 7, "Monatsuebersicht", new_x="LMARGIN", new_y="NEXT")
        pdf.set_font("helvetica", "", 10)
        for partner in PARTNER_NAMES:
            total = df_rows.loc[df_rows["Betreuer"] == partner, "Provision"].sum()
            r, g, b = PARTNER_COLORS_PDF[partner]
            pdf.set_text_color(r, g, b)
            pdf.cell(60, 6, partner)
            pdf.cell(0, 6, _fmt_eur(total), new_x="LMARGIN", new_y="NEXT")
        unmatched_total = df_rows.loc[df_rows["Betreuer"].isna(), "Provision"].sum()
        pdf.set_text_color(120, 120, 120)
        pdf.cell(60, 6, "Nicht zugeordnet")
        pdf.cell(0, 6, _fmt_eur(unmatched_total), new_x="LMARGIN", new_y="NEXT")
        pdf.set_text_color(0, 0, 0)
        pdf.ln(3)

    left_w, gap, right_w = 55, 5, 125
    left_margin = pdf.l_margin

    for versicherer, total in insurer_totals.items():
        customers = per_customer[per_customer["Versicherer"] == versicherer].sort_values(
            "Provision", ascending=False
        )

        # Vor jedem Block pruefen, ob der Versicherer-Name + mindestens die
        # erste Kundenzeile noch auf die aktuelle Seite passen - sonst
        # Seitenumbruch VOR dem Block (statt mittendrin) fuer ein saubereres
        # Layout.
        if pdf.get_y() + 16 > pdf.page_break_trigger:
            pdf.add_page()

        y_start = pdf.get_y()
        page_start = pdf.page_no()
        x_left = left_margin
        x_right = left_margin + left_w + gap
        right_line_h = 5

        pdf.set_xy(x_left, y_start)
        pdf.set_font("helvetica", "B", 11)
        pdf.multi_cell(left_w, 6, _pdf_safe_text(versicherer))
        pdf.set_x(x_left)
        pdf.set_font("helvetica", "", 9)
        pdf.multi_cell(left_w, 5, f"Gesamt:\n{_fmt_eur(total)}")
        y_left_end, page_left_end = pdf.get_y(), pdf.page_no()

        # Rechte Spalte zeilenweise mit manueller Seitenumbruch-Kontrolle
        # ausgeben (statt multi_cell()): so bleiben linke/rechte Spalte auch
        # bei einem Umbruch mitten in einer langen Kundenliste konsistent
        # positioniert, statt dass darunterliegender Inhalt auf der falschen
        # Seite landet und grosse Luecken entstehen.
        pdf.set_xy(x_right, y_start)
        pdf.set_font("helvetica", "", 9)
        y = y_start
        for row in customers.itertuples():
            if y + right_line_h > pdf.page_break_trigger:
                pdf.add_page()
                y = pdf.t_margin
            pdf.set_xy(x_right, y)
            pdf.set_text_color(*PARTNER_COLORS_PDF.get(row.Betreuer, (0, 0, 0)))
            pdf.multi_cell(right_w, right_line_h, f"{_pdf_safe_text(row.Kunde)} - {_fmt_eur(row.Provision)}")
            pdf.set_text_color(0, 0, 0)
            y = pdf.get_y()
        y_right_end, page_right_end = y, pdf.page_no()

        # Falls die rechte Spalte einen Seitenumbruch ausgeloest hat, ist
        # y_left_end (von der vorherigen Seite) fuer die Positionierung von
        # Trennlinie/naechstem Block irrelevant - sonst wuerde eine grosse,
        # falsch platzierte Luecke auf der neuen Seite entstehen. Es zaehlt
        # immer die Spalte, die zuletzt (auf der spaeteren Seite) endet.
        if page_right_end != page_left_end:
            y_end = y_right_end if page_right_end > page_left_end else y_left_end
        else:
            y_end = max(y_left_end, y_right_end)

        pdf.set_xy(x_left, y_end + 2)
        pdf.set_draw_color(200, 200, 200)
        pdf.line(x_left, pdf.get_y(), x_right + right_w, pdf.get_y())
        pdf.ln(3)

    pdf.output(out_target)


def find_betreuer_path(base_dir):
    """Liefert den Pfad zur aktuellen Betreuer.xlsx in base_dir. Bevorzugt
    die exakte Datei 'Betreuer.xlsx', faellt aber auf die zuletzt
    geaenderte Datei zurueck, die zu 'Betreuer*.xlsx' passt (z.B.
    'Betreuer-12-10-2026.xlsx') - der CRM-Export wird manchmal mit
    Datumsstempel im Dateinamen abgelegt statt die bestehende Datei zu
    ueberschreiben (beobachtet 2026-10-12, main() fand dadurch tagelang
    keine Betreuer.xlsx mehr und lief stillschweigend ohne Zuordnung
    weiter). Gibt bei keinem Treffer trotzdem den (nicht existierenden)
    Standardpfad zurueck - os.path.isfile() beim Aufrufer bleibt die
    eigentliche Pruefung."""
    exact = os.path.join(base_dir, "Betreuer.xlsx")
    if os.path.isfile(exact):
        return exact
    candidates = glob.glob(os.path.join(base_dir, "Betreuer*.xlsx"))
    if not candidates:
        return exact
    return max(candidates, key=os.path.getmtime)


def main():
    if len(sys.argv) < 2:
        print("Aufruf: python courtage_extraktor.py <Monatsordner z.B. Juni-2026> [--pdf-uebersicht]")
        sys.exit(1)

    month_folder = sys.argv[1]
    make_pdf_summary = "--pdf-uebersicht" in sys.argv[2:]
    input_dir = os.path.join(BASE_DIR, month_folder)
    if not os.path.isdir(input_dir):
        print(f"Ordner nicht gefunden: {input_dir}")
        sys.exit(1)

    pdf_files = sorted(glob.glob(os.path.join(input_dir, "Abrechnung-*.pdf")))
    # Nicht nur "VEMA-Poolabrechnung-N.csv" (urspruenglicher Dateiname beim
    # Download) - der Nutzer benennt die Datei manchmal wie die PDFs um
    # (z.B. "Abrechnung-1-VEMA-Pool-August-2026.csv", siehe August-2026).
    # extract_vema_csv() prueft ohnehin selbst anhand der Spalten (Betrag-
    # Spalte vorhanden?), ob es sich um eine echte VEMA-Pool-CSV handelt -
    # jede .csv im Ordner ist daher ein sicherer Kandidat.
    csv_files = sorted(glob.glob(os.path.join(input_dir, "*.csv")))
    files = pdf_files + csv_files
    print(f"Gefunden: {len(pdf_files)} Abrechnungs-PDFs und {len(csv_files)} "
          f"VEMA-Pool-CSV(s) in {month_folder}")

    def report_progress(i, total, filename):
        print(f"  verarbeite {filename} ...")

    df_rows, df_control, df_agg, df_problem = process_files(files, month_folder, report_progress)

    betreuer_path = find_betreuer_path(BASE_DIR)
    if os.path.isfile(betreuer_path):
        lookup = load_betreuer_lookup(betreuer_path)
        df_rows = apply_betreuer(df_rows, lookup)
        n_unmatched = int(df_rows["Betreuer"].isna().sum()) if not df_rows.empty else 0
        print(f"Betreuer.xlsx gefunden - Kunden zugeordnet ({n_unmatched} ohne Zuordnung, "
              "siehe Blatt 'Kunde_ohne_Betreuer').")
    else:
        print("Hinweis: Betreuer.xlsx nicht gefunden - keine Betreuer-Zuordnung erstellt.")

    df_bank_unmatched = None
    df_bank_explained = None
    bank_files = sorted(
        f for f in glob.glob(os.path.join(input_dir, "*.pdf"))
        if "kontoauszug" in os.path.basename(f).lower()
        and "verwaltung" not in os.path.basename(f).lower()
    )
    if bank_files:
        all_credits = []
        for bf in bank_files:
            try:
                all_credits.extend(extract_bank_credits(bf))
            except Exception as exc:
                print(f"Hinweis: Kontoauszug {os.path.basename(bf)} konnte nicht gelesen "
                      f"werden ({type(exc).__name__}) - uebersprungen.")
        if all_credits:
            insurer_keys = sorted(
                set(df_control["Versicherer"].tolist() if "Versicherer" in df_control.columns else [])
                | set(df_agg["Versicherer"].tolist() if "Versicherer" in df_agg.columns else [])
                | set(df_problem["Versicherer"].tolist() if "Versicherer" in df_problem.columns else [])
            )
            prior = []
            prev = previous_month_folder(month_folder)
            if prev:
                prev_xlsx = os.path.join(TOOL_DIR, "output", prev, f"Kunde_Provision_{prev}.xlsx")
                if os.path.isfile(prev_xlsx):
                    prior = load_prior_month_totals(prev_xlsx)
                    print(f"Bankabgleich: Vormonat {prev} beruecksichtigt ({len(prior)} Versicherer).")
                else:
                    print(f"Bankabgleich: kein Ergebnis fuer Vormonat {prev} gefunden "
                          f"({prev_xlsx}) - nur der laufende Monat wird abgeglichen.")
            explained = []
            unmatched = reconcile_bank_credits(all_credits, insurer_keys, prior, explained)
            df_bank_unmatched = pd.DataFrame(
                unmatched, columns=["datum", "betrag", "sender", "verwendungszweck"]
            )
            df_bank_explained = pd.DataFrame(
                explained, columns=["datum", "betrag", "sender", "verwendungszweck", "vorgang", "erklaert_durch"]
            )
            print(f"Bankabgleich: {len(unmatched)} Zahlungseingang/-eingaenge ohne Abrechnung, "
                  f"{len(explained)} durch Vormonat erklaert.")

    out_dir = os.path.join(TOOL_DIR, "output", month_folder)
    out_path = os.path.join(out_dir, f"Kunde_Provision_{month_folder}.xlsx")
    write_excel(df_rows, df_control, df_agg, df_problem, out_path, df_bank_unmatched, df_bank_explained)

    print()
    print(f"Fertig. {len(df_rows)} Buchungszeilen aus {len(files)} Dateien extrahiert.")
    print(f"  - {len(df_agg)} Datei(en) als Sammelbeleg ohne Kundendetail")
    print(f"  - {len(df_problem)} Datei(en) zur manuellen Pruefung")
    print(f"Ausgabe: {out_path}")

    if make_pdf_summary:
        pdf_path = os.path.join(out_dir, f"Courtage-Uebersicht_{month_folder}.pdf")
        build_summary_pdf(df_rows, pdf_path, month_folder)
        print(f"PDF-Uebersicht: {pdf_path}")


if __name__ == "__main__":
    main()
