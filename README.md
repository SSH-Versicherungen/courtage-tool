# Courtage-Extraktor

Liest die Courtageabrechnungs-PDFs eines Monats ein und extrahiert je
Buchungszeile: **Versicherer, Kunde, Provision**. Optional (wenn eine
`Betreuer.xlsx` bereitgestellt wird) ordnet es jede Buchung zusätzlich einem
der drei Partner zu und markiert sie farbig - siehe "Betreuer-Zuordnung"
unten.

Es gibt zwei Wege, das Tool monatlich zu nutzen: eine **Web-Oberfläche im
Browser** (empfohlen, kein Terminal nötig) oder die **Kommandozeile**
(praktisch, wenn die PDFs schon in `Umsatz/<Monat>/` liegen).

## Monatliche Nutzung: Web-Oberfläche (empfohlen)

Einmal starten:
```bash
cd Courtage-Tool
streamlit run app.py
```
Es öffnet sich automatisch ein Browserfenster (falls nicht:
`http://localhost:8501` manuell aufrufen). Dort:

1. Monat und Jahr auswählen.
2. Alle Courtageabrechnungs-PDFs des Monats per Drag & Drop hochladen oder
   auswählen (Mehrfachauswahl möglich).
3. Optional: den/die Kontoauszug/-züge (VR Bank Rhein-Neckar) desselben
   Monats zusätzlich hochladen - siehe "Kontoauszug-Abgleich" unten.
4. Optional: `Betreuer.xlsx` hochladen - siehe "Betreuer-Zuordnung" unten.
5. Optional: Häkchen bei "PDF-Übersicht zusätzlich erstellen" setzen - siehe
   "PDF-Übersicht" unten.
6. Auf **Verarbeiten** klicken.
7. Ergebnis direkt im Browser ansehen (Kontrolltabelle, Sammelbelege,
   manuelle Prüfung, alle Buchungszeilen, ggf. fehlende Abrechnungen, ggf.
   Kunden ohne Betreuer-Zuordnung) und über **Excel-Ergebnis herunterladen**
   als `.xlsx`-Datei speichern (bzw. zusätzlich **PDF-Übersicht
   herunterladen**, falls angehakt).

Um die Web-Oberfläche zu beenden, im Terminal-Fenster `Strg+C` drücken (oder
das Fenster schließen). Beim nächsten Monat einfach erneut `streamlit run
app.py` ausführen - die hochgeladenen PDFs werden nur temporär für die
Verarbeitung gespeichert und danach wieder gelöscht, nichts wird dauerhaft
auf der Festplatte abgelegt außer der heruntergeladenen Excel-Datei.

## Monatliche Nutzung: Kommandozeile

```bash
cd Courtage-Tool
python courtage_extraktor.py Juli-2026
```

Der Parameter ist der Name des Monatsordners in `Umsatz/` (z.B. `Juli-2026`,
`August-2026`, ...). Das Skript sucht darin automatisch alle Dateien nach dem
Muster `Abrechnung-*.pdf` (Kontoauszüge, Rechnungen-Unterordner etc. werden
ignoriert).

Ergebnis liegt danach hier:
```
Courtage-Tool/output/<Monat>/Kunde_Provision_<Monat>.xlsx
```

Mit dem optionalen Flag `--pdf-uebersicht` wird zusätzlich die PDF-Übersicht
erstellt (siehe unten):
```bash
python courtage_extraktor.py Juli-2026 --pdf-uebersicht
```

Liegt eine Datei `Betreuer.xlsx` direkt im `Umsatz`-Ordner (eine Ebene über
`Courtage-Tool/`), wird sie automatisch erkannt und für die Betreuer-
Zuordnung verwendet (siehe "Betreuer-Zuordnung" unten) - kein zusätzliches
Flag nötig.

Für einen neuen Monat ist **keine Anpassung am Skript nötig** - nur der
Ordnername als Argument (CLI) bzw. die Monat/Jahr-Auswahl (Web-Oberfläche)
ändert sich. Die Erkennungslogik (Spaltenüberschriften, Versicherer-
Sonderfälle) ist dauerhaft im Skript hinterlegt und wird von beiden Wegen
gemeinsam genutzt (`app.py` ruft nur `process_files()`/`courtage_extraktor.py`
auf, enthält selbst keine Extraktionslogik).

## Aufbau der Ausgabe-Excel

- **Kunde_Provision** - die eigentliche Rohdaten-Tabelle: Versicherer, Kunde,
  Provision, Datei, Seite, Quelle (Text/OCR), Rohzeile (Originaltext zur
  Kontrolle/Nachvollziehbarkeit).
- **Kontrolle** - je Datei: Anzahl Buchungen, extrahierte Summe, eine
  (best-effort) aus dem PDF-Text erkannte Vergleichssumme, Differenz. Dient
  als Plausibilitätscheck - die Vergleichssumme ist ein heuristischer
  Texttreffer und nicht immer die "richtige" Kennzahl (siehe Hinweis unten),
  bei größeren Differenzen lohnt ein Blick in die Rohzeile/das Original-PDF.
- **Sammelbelege_ohne_Details** - Versicherer, deren PDF wirklich nur einen
  Sammelbetrag ohne jede Kundenaufschlüsselung enthält: aktuell nur ARAG
  (reiner Kontoauszug/Avise, keine Einzelposten in diesem PDF) und VEMA-Pool
  (verweist im PDF-Text explizit auf eine separat ausgelieferte CSV-Datei mit
  den Einzelverträgen - siehe "VEMA-Pool-CSV" unten, diese CSV wird
  inzwischen automatisch mitverarbeitet, wenn sie im Monatsordner liegt).
  Wichtig: mehrseitige
  Sammelabrechnungen wie Mannheimer, VHV, Swiss Life, HDI-Leben oder SV
  Sparkassenversicherung haben ihre Kundenpositionen oft erst auf hinteren
  Seiten ("Einzelposten", "Anhang", "Provisionsberechnungsnachweis") - diese
  werden mittlerweile korrekt ausgewertet, nicht als Sammelbeleg eingestuft.
- **Manuelle_Pruefung** - Dateien, die aktuell nicht automatisch verarbeitet
  werden können, OBWOHL sie echte Kundenpositionen enthalten (siehe Grenzen
  unten). Nicht mit "Sammelbelege_ohne_Details" verwechseln: dort fehlen die
  Positionsdaten im PDF komplett, hier sind sie vorhanden, aber (noch) nicht
  zuverlässig automatisch auslesbar.
- **Fehlende_Abrechnungen** - nur vorhanden, wenn beim Verarbeiten mindestens
  ein Kontoauszug mit hochgeladen wurde (siehe "Kontoauszug-Abgleich" unten).
- **Kunde_ohne_Betreuer** - nur vorhanden, wenn eine `Betreuer.xlsx`
  bereitgestellt wurde (siehe "Betreuer-Zuordnung" unten): Kunden, die
  keinem der drei Partner eindeutig zugeordnet werden konnten.
- **Zusammenfassung** - immer das erste Blatt: Summe je Versicherer
  (absteigend), darunter (falls eine `Betreuer.xlsx` bereitgestellt wurde)
  Summe je Partner inkl. "Nicht zugeordnet".

Alle Blätter sind formatiert (fette blaue Kopfzeile, Auto-Filter, passende
Spaltenbreiten, Währungsformat für Beträge); im Blatt "Kontrolle" wird eine
von 0 abweichende "Differenz" rot hervorgehoben. Ist eine `Betreuer.xlsx`
bereitgestellt, sind im Blatt "Kunde_Provision" zusätzlich alle Zeilen eines
zugeordneten Kunden farbig markiert (siehe unten).

## PDF-Übersicht (optional)

Erstellt eine kompakte PDF-Zusammenfassung (`Courtage-Uebersicht_<Monat>.pdf`,
`build_summary_pdf()`): zu Beginn eine Summe je Versicherer sowie (bei
vorhandener `Betreuer.xlsx`) eine Monatsübersicht je Partner, danach je
Versicherer **absteigend nach Gesamtumsatz** sortiert, mit der zugehörigen
Kunden-/Vertragsaufstellung (ebenfalls absteigend nach Betrag) direkt
daneben - bei vorhandener Betreuer-Zuordnung farbig je Partner. Rein
informativ/optisch - nutzt nur die bereits extrahierten Zeilen, keine neue
Extraktionslogik. Kundenbeträge werden je Versicherer aufsummiert (mehrere
Buchungen desselben Kunden innerhalb eines Versicherers erscheinen als eine
Zeile).

Web-Oberfläche: Häkchen "PDF-Übersicht zusätzlich erstellen" vor dem Klick
auf "Verarbeiten". Kommandozeile: Flag `--pdf-uebersicht` (siehe oben).

## Betreuer-Zuordnung (optional)

Ordnet jede Buchungszeile einem der drei Partner (Robin Heckmann, Tim Selle,
Andreas Selle) zu, basierend auf dem "Kundenverantwortliche"-Feld im
internen CRM (`https://ssh-versicherungen.insurgo.cloud/dashboard`). Da
Andreas Selle dort keinen eigenen Account hat, werden seine Kunden im CRM
unter "Calvin Keinarth" geführt - das Tool rechnet diese automatisch
Andreas Selle zu (`BETREUER_ALIASES`).

Voraussetzung: eine `Betreuer.xlsx` (CRM-Export mit den Blättern
"Privatkunden" und "Firmenkunden", Spalten u.a. Vorname/Nachname bzw.
Firmenname und "Kundenverantwortliche"). Diese Datei enthält echte
Kundendaten und liegt bewusst **außerhalb** des Git-Repos, direkt im
`Umsatz`-Ordner - CLI: wird dort automatisch gefunden (siehe oben);
Web-Oberfläche: muss pro Sitzung hochgeladen werden (Streamlit Cloud hat
keinen Zugriff auf die lokale Festplatte).

Ist eine `Betreuer.xlsx` vorhanden, ergänzt das Tool im Blatt
"Kunde_Provision" die Spalte "Betreuer" sowie je eine Spalte pro Partner mit
dessen Anteil der jeweiligen Zeile, und markiert die ganze Zeile farbig:
**Rot = Robin Heckmann, Gelb = Tim Selle, Blau = Andreas Selle** (in der
PDF-Übersicht entsprechend der Kundentext eingefärbt, mit Legende oben auf
der Seite) - so sind Zuordnung bzw. Abweichungen beim Überfliegen sofort
sichtbar.

Namensabgleich (`match_betreuer()`): zuerst exakt, sonst mit kontrollierter
Unschärfe (abgeschnittene oder leicht abweichend geschriebene Vor-/Nachnamen,
z.B. weil das Quell-PDF eines Versicherers Namen abschneidet) - aber **nur**,
wenn dabei eindeutig genau ein Partner infrage kommt. Ist das Ergebnis
mehrdeutig (mehrere Partner mit ähnlich passendem Namen) oder findet sich
gar kein Kandidat, bleibt der Kunde absichtlich unzugeordnet (kein Fill,
taucht im Blatt "Kunde_ohne_Betreuer" auf) statt geraten zu werden.
Geburtsname-Zusätze wie "(geb. Bock)" oder "geb.: Pleli" werden vor dem
Abgleich entfernt. Firmenkunden werden zusätzlich über den (um Rechtsform-
Endungen wie GmbH/UG/KG bereinigten) Firmennamen abgeglichen.

Hinweis: Der CRM-Export bildet den Stand zum Zeitpunkt des Exports ab - bei
sehr aktuellen Änderungen im CRM (neue Kunden, Betreuer-Wechsel) empfiehlt
sich vor der monatlichen Abrechnung ein frischer Export.

Für Fälle, die auch mit Unschärfe nie zuverlässig auffindbar wären (z.B.
Heiratsnamen wie "Stasiak" -> "Weickel", oder Handelsnamen, die vom
CRM-Firmennamen stark abweichen) gibt es zusätzlich eine von Hand gepflegte
Korrekturliste `MANUAL_BETREUER_OVERRIDES_RAW` in `courtage_extraktor.py`
(Kundenname aus dem PDF -> Partner). Neue Fälle einfach melden, dann wird
die Liste ergänzt - langfristig sauberer ist es, die Ursache direkt in der
Betreuer.xlsx zu pflegen. Nicht-kundenspezifische Sammelabzüge innerhalb
einer Abrechnung (z.B. Fondsfinanz-Stornoreserve) werden automatisch der
größten Kundenposition derselben Datei zugerechnet, da sie selbst keinen
Namen tragen.

## Kontoauszug-Abgleich (optional)

Wird beim Verarbeiten zusätzlich ein oder mehrere Kontoauszüge der VR Bank
Rhein-Neckar (PDF, wie sie die Bank direkt bereitstellt) hochgeladen, prüft
das Tool zusätzlich: **Gibt es einen Zahlungseingang (Gutschrift) auf dem
Konto, zu dem keine passende Abrechnungs-PDF verarbeitet wurde?** Das ist ein
Hinweis darauf, dass für diesen Versicherer entweder die PDF-Abrechnung noch
fehlt (z.B. noch nicht per Post/E-Mail angekommen oder vergessen
hochzuladen) oder dass ein neuer, dem Tool noch unbekannter Versicherer
gezahlt hat.

Funktionsweise (`extract_bank_credits()`/`reconcile_bank_credits()` in
`courtage_extraktor.py`):
1. Der Kontoauszug (normaler PDF-Textlayer, keine OCR nötig) wird in
   einzelne Buchungen zerlegt (Datum, Vorgangsart, Betrag, Absender/
   Verwendungszweck) - Brief-Kopf-/Fusszeilen und der pro Seite wiederholte
   "Übertrag auf/von Blatt N"-Saldo-Hinweis werden herausgefiltert, alle
   Seiten zu einem durchgehenden Buchungsstrom zusammengefügt (wichtig, da
   eine Buchung sonst an einem Seitenumbruch auseinandergerissen würde).
2. Nur Buchungen vom Typ "GUTSCHRIFT" werden betrachtet (keine Lastschriften,
   Daueraufträge oder Eigenüberweisungen).
3. Jede Gutschrift wird - zuerst über den Absendernamen, sonst zusätzlich
   über den Verwendungszweck - versucht, einem der in diesem Monat
   tatsächlich verarbeiteten Versicherer zuzuordnen (`INSURER_BANK_ALIASES`:
   eine kleine, von Hand gepflegte Stichwortliste, z.B. "WUERTT.VERSICHERG."
   → Württembergische, "Qualitypool" im Verwendungszweck "AMEX Abr." →
   Amex-Pool). Diese Liste ist zwangsläufig unvollständig; taucht ein neuer
   Versicherer im Kontoauszug auf, dessen Bank-Bezeichnung nicht erkannt
   wird, erscheint er - statt stillschweigend ignoriert zu werden -
   ebenfalls im Ergebnis, und die Alias-Liste kann ergänzt werden.
4. Nicht zuordenbare Gutschriften landen im Excel-Blatt
   "Fehlende_Abrechnungen" (Datum, Betrag, Absender, Verwendungszweck).

Wichtig: das ist ein **heuristischer Namensabgleich, kein Betragsabgleich**
(prüft nicht, ob die Höhe der Gutschrift zur extrahierten Provisionssumme
passt) und unterstützt bisher nur das Kontoauszug-Format der VR Bank
Rhein-Neckar. Vor einer Rückfrage beim Versicherer den Treffer in der
Rohspalte "Absender"/"Verwendungszweck" gegenprüfen.

## VEMA-Pool-CSV

VEMA-Pool liefert die Kundenpositionen nicht im PDF (das bleibt ein reiner
Sammelbeleg), sondern als separate CSV-Datei(en)
(`VEMA-Poolabrechnung-<N>.csv`, Semikolon-getrennt). Liegt eine solche CSV
im selben Monatsordner wie die PDFs, wird sie automatisch mitverarbeitet
(CLI: `python courtage_extraktor.py <Monat>` findet sie über das Muster
`VEMA-Poolabrechnung-*.csv`; Web-Oberfläche: einfach mit hochladen, die
Endung `.csv` wird erkannt). Verwendet wird die Spalte "Betrag" (Courtage
NACH Abzug des VEMA-Pool-Verwaltungsbeitrags/"Einbehalt", z.B. 10% - dieser
Einbehalt ist anders als z.B. Fondsfinanz' Stornoreserve ein dauerhafter
Abzug, der SSH nie zufließt, daher zählt hier der Netto- statt der
Brutto-Betrag "Courtage"). Siehe `extract_vema_csv()`.

## Grenzen der aktuellen Version

- **AIG**: gelöst - AIG druckt die Provisions-/Courtage-Spalte immer mit
  nachgestelltem Minus, auch wenn der Betrag laut eigener Summenzeile
  ("Summe Vermittler (Provision/Courtage zu Ihren Gunsten): ... 112,39-")
  tatsächlich eine Gutschrift für den Makler ist - umgekehrte Vorzeichenlogik
  zu allen anderen Versicherern. Wird im Code explizit umgedreht (siehe
  `insurer_lower == "aig"`-Zweig in `process_file()`).
- **Allianz, AXA, Gothaer (Allgemeine + Leben)**: gelöst - siehe
  `extract_allianz()`/`extract_axa()`/`extract_gothaer()` im technischen
  Aufbau unten. Alle drei speichern den Text pro Seite gespiegelt/verdreht,
  mit echten Kundenpositionen in einem kartenartigen Layout. Wichtig fürs
  Verständnis der Rohzeilen: bei Allianz und AXA können für denselben
  Vertrag/Kunden mehrere, auch vom Betrag her identisch aussehende Zeilen
  im selben Dokument auftauchen (z.B. eine Storno- und eine Neu-Buchung,
  oder bei AXA sogar zwei exakt identische Buchungen) - das ist **kein
  Extraktionsfehler**, sondern entspricht exakt dem im PDF selbst
  aufgedruckten Kontrollbetrag (Allianz: "Sach-Provisionen"-Summe im
  Anschreiben; AXA: "Betreu.-prov."-Summe auf der letzten Seite; Gothaer:
  "Endsumme" auf dem Deckblatt) - siehe Docstrings der drei Funktionen für
  die Herleitung. Gothaer ist bisher nur an kleinen Ein-Kunden-Dateien
  verifiziert (Juni 2026 enthielt keine größeren Testfälle) - bei
  auffälligen Differenzen in der Kontrolle-Tabelle die Rohzeile prüfen.
- **Itzehoer**: PDF ist passwortgeschützt, kann so nicht geöffnet werden.
- **Alte Leipziger**: gelöst - siehe `extract_alte_leipziger()`. Der Scan ist
  sauber, die einzige Eigenheit ist, dass der Betrag auf der ersten Zeile
  eines Kundenblocks steht, der lesbare (oft mehrzeilige) Name aber erst auf
  der/den Folgezeile(n) - wird pro Block gesammelt und bei "Zwischensumme"
  zugeordnet. Summe je Teil-Abrechnung stimmt exakt mit der "Zahlungsausgang"-
  Zeile im PDF überein.
- **Barmenia (bzw. Barmenia/Gothaer)**: gelöst - siehe `extract_barmenia()`.
  Gescanntes "Vergütungsnachweis Abschlussvergütungen"-PDF, bei dem der
  Standard-OCR-Pfad (300dpi) die Betragsspalten auf der Detailseite verliert;
  bei 400dpi + tabellenorientiertem OCR-Modus (`--psm 6`) sind sie lesbar.
  Der Vergütungsbetrag steht pro Kundenzeile zweimal identisch hintereinander
  (Vergütungsbetrag = Abgerechneter Betrag) - dieses Zahlenpaar wird als
  Betrag genommen.
- **Continentale**: gelöst - siehe `extract_continentale()`. Sehr dichtes,
  OCR-mehrdeutiges Tabellenformat: derselbe Betrag wird je nach
  OCR-Auflösung mal sauber ("+19,45"), mal mit verlorenem Komma ("+502"
  statt "+5,02") oder mit Buchstaben-Ersatzzeichen ("H021" statt "+0,21")
  gelesen. Der Parser probiert daher mehrere Auflösungen (400/500/600dpi)
  und nimmt je Position den saubersten Treffer, mit einer Ziffern-Rückfall-
  Heuristik für hartnäckige Fälle. Weil diese Heuristik unvermeidbar
  fehleranfällig ist, wird das Ergebnis **immer** gegen den auf der
  Kontoauszug-Seite aufgedruckten "Neuer Saldo" geprüft (siehe
  `process_file()`): stimmt die Summe nicht exakt überein, fällt die Datei
  automatisch auf manuelle Prüfung zurück, statt falsche Kundenzuordnungen
  zu riskieren - passend zum Grundsatz "Summe der Einzelpositionen muss
  immer dem Gesamtsaldo entsprechen, im Zweifel Rückmeldung statt Raten".
- **Swiss Life - Vertrauensschadenversicherung**: Swiss Life behält pro
  Abrechnung einen pauschalen (nicht kundenbezogenen) Beitrag zur
  Vertrauensschadenversicherung ein, der direkt von der Gesamt-Courtage
  abzuziehen ist (Nutzer-Rückmeldung - frühere Annahme, dieser Beitrag werde
  separat privat verrechnet, war falsch). Er steht explizit auf der
  "Abrechnungsübersicht zum ..."-Seite (nicht auf der "Kumulierte
  Jahreswerte"-Seite, die denselben Zeilennamen fürs gesamte Jahr zeigt) als
  eigene Zeile "Beiträge zur Vertrauensschadenversicherung" - wird als
  eigene, ausdrücklich nicht-kundenspezifische Abzugszeile ergänzt, siehe
  `extract_swiss_life_vsv_deduction()`.
- **Dialog**: bleibt manuelle Prüfung. Die Kundenpositions-Tabelle
  (L0100 "Provisionseinzelnachweis") ist im Scan zu klein/dicht gedruckt
  (teils zusätzlich durch Textmarker-Anmerkungen überdeckt) - mehrere OCR-
  Auflösungen/Modi wurden ausprobiert, aber Kommastellen kippen dabei
  unzuverlässig (z.B. "12,01" vs. "1201"), was einen Betrag um Faktor 100
  verfälschen könnte. Um nicht versehentlich falsche Zahlen zu liefern,
  bleibt das Kundendetail Handarbeit. Als Kontrollhinweis wird aber der
  Gesamtbetrag von der (großformatigeren, zuverlässiger lesbaren)
  "PG-Übersicht"-Seite per OCR ermittelt und in der Spalte "Betrag_lt_PDF"
  im Blatt "Manuelle_Pruefung" angezeigt (siehe `find_dialog_total_hint()`)
  - das ist ein heuristischer Best-effort-Wert, keine belastbare Zahl.
- Andere gescannte PDFs (Signal Iduna, SV Sparkassenversicherung) werden
  per OCR gelesen und funktionieren inzwischen gut (SV Sparkassenversicherung
  hat einen eigenen Parser, da das PDF keine klassische Tabellen-Kopfzeile
  hat); bei SV weicht die extrahierte Summe wegen einzelner OCR-Lesefehler
  (z.B. verlorenes Komma in einem Betrag) minimal (üblicherweise <1%) vom
  echten Gesamtbetrag ab - bei Bedarf gegen die Rohzeile/das Original prüfen.
- Bei manchen Versicherern gibt es zwei plausible Beträge pro Position (z.B.
  brutto vor Stornoreserve vs. netto danach) - das Skript nimmt im Zweifel
  den Betrag unter der am weitesten rechts stehenden "Courtage"/"Provision"-
  Spalte (i.d.R. der Auszahlungsbetrag).
- Die "Summe_lt_PDF"-Vergleichsspalte im Blatt "Kontrolle" ist ein
  heuristischer Texttreffer und bei manchen Versicherern (Concordia, Amex-
  Pool, Württembergische) ungenau/falsch, während die eigentliche Extraktion
  daneben nachweislich korrekt ist (manuell gegen die PDFs geprüft) - eine
  Differenz dort bedeutet nicht automatisch einen Fehler in der Extraktion.

Diese Fälle sind in der Excel-Datei klar gekennzeichnet (Blatt
"Manuelle_Pruefung" bzw. "Sammelbelege_ohne_Details") statt stillschweigend
falsche oder erfundene Zahlen zu liefern.

## Technischer Aufbau (für Weiterentwicklung)

- `courtage_extraktor.py` enthält:
  - eine generische, koordinatenbasierte Tabellen-Engine (erkennt
    Spaltenüberschriften wie "Kunde"/"Versicherungsnehmer" und
    "Courtage"/"Provision"/"Vergütung" automatisch anhand ihrer Position im
    PDF, auch mehrfach pro Seite bei mehreren Teiltabellen) - funktioniert
    für die meisten Versicherer ohne Anpassung, inkl. Mannheimer (mit einer
    Blacklist für Buchungsart-Label wie "Bestandspflege", die sonst als
    Kundenname missverstanden würden),
  - einen eigenen Parser für Block-Formate ohne Tabellenkopf: Fondsfinanz und
    Swiss Life ("VN/VP Name ... Summe Vertrag X"),
  - einen eigenen Parser für VHV (dort steht die Name+Adresse-Zeile nach den
    zugehörigen Buchungszeilen, nicht davor),
  - einen eigenen Parser für SV Sparkassenversicherung (OCR, keine
    Tabellen-Kopfzeile im PDF),
  - einen eigenen Parser für Alte Leipziger (`extract_alte_leipziger`, OCR):
    Betrag steht vor dem lesbaren (oft mehrzeiligen) Namen, umgekehrte
    Reihenfolge zur generischen Engine - Betraege werden pro Kundenblock
    gesammelt und erst bei "Zwischensumme" zugeordnet,
  - einen Best-effort-Kontrollbetrag fuer Dialog (`find_dialog_total_hint`,
    OCR mit Schwellwert-Vorverarbeitung auf der PG-Uebersicht-Seite) -
    liefert keine Kundenpositionen, nur einen Hinweiswert fuer die manuelle
    Pruefung,
  - einen eigenen Parser für Allianz (`extract_allianz`), AXA
    (`extract_axa`) und Gothaer (`extract_gothaer`): alle drei PDFs
    speichern den Text pro Seite gespiegelt/verdreht, wodurch pdfplumber
    vertauschte Koordinatenachsen liefert ("top" wirkt wie die horizontale,
    "x0" wie die vertikale Achse) - die Parser gruppieren Wörter über die
    x0-Achse zu Zeilen und werten z.B. die Spaltenzugehörigkeit eines
    Betrags (Allianz) über seine "top"-Position relativ zu den
    Kopfzeilen-Wörtern aus. Alle drei sind gegen die im jeweiligen PDF
    selbst aufgedruckte Kontrollsumme exakt verifiziert (nicht nur gegen
    die externe Vergleichs-Exceltabelle).
  - einen OCR-Fallback (Tesseract, deutsches Sprachpaket unter
    `Courtage-Tool/tessdata/`) für PDFs ohne Text-Ebene,
  - Erkennung von Sammelbelegen und der bekannten Sonderformate,
  - `write_excel()`/`_style_worksheet()`: gemeinsame Excel-Ausgabe (Formate,
    Kopfzeile, Auto-Filter, Spaltenbreiten, Differenz-Hervorhebung) für CLI
    und Web-Oberfläche,
  - `extract_bank_credits()`/`reconcile_bank_credits()`: der optionale
    Kontoauszug-Abgleich (siehe eigener Abschnitt oben).
  - `load_betreuer_lookup()`/`match_betreuer()`/`apply_betreuer()`: die
    optionale Betreuer-Zuordnung (siehe eigener Abschnitt oben).
- Neue Versicherer mit "normaler" Tabellenstruktur sollten ohne Codeänderung
  funktionieren. Taucht ein neuer Versicherer in "Sammelbelege_ohne_Details"
  auf, obwohl er eigentlich Positionsdaten haben müsste, oder ein neues
  Sonderformat auftritt: kurz Bescheid geben, dann kann das gezielt
  nachgebessert werden.

## Einrichtung (einmalig erledigt)

Für dieses Tool wurden auf diesem Rechner installiert: Python 3.12 (winget),
die Pakete `pdfplumber`, `pandas`, `openpyxl`, `pytesseract`, `pypdfium2`,
`streamlit` (für die Web-Oberfläche, `app.py`), sowie Tesseract-OCR (winget,
UB-Mannheim-Build) mit deutschem Sprachpaket (liegt in
`Courtage-Tool/tessdata/`, da kein Schreibzugriff auf
`C:\Program Files\Tesseract-OCR` bestand). Für einen neuen Rechner müsste das
einmalig wiederholt werden.
