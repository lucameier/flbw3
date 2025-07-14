"""
FLBW Daten Transformation für die neue SAP-Struktur
=================================================

Diese Anwendung transformiert FLBW-Daten aus dem neuen SAP-Exportformat in ein standardisiertes Analyseformat.
Die Transformation umfasst folgende Hauptfunktionen:
- ICT-Klassifizierung basierend auf Kontierungsbeschreibung und spezifischen Auftragsnummern
- Unterkategorie-Ableitung mit find_keyword (prüft nur den Textanfang)
- Pivotierung nach definierten Feldern
- Automatische Status-Erkennung (Arbeit, Arbeit Unproduktiv, Abwesend)

@author: Luca Meier
@date: 09.05.2025
"""

import pandas as pd
import streamlit as st
import io
import re

# Streamlit-Konfiguration
# ----------------------
# Setzt die grundlegenden Einstellungen für die Streamlit-Webanwendung
st.set_page_config(
    page_title="FLBW Daten Transformation (Neue SAP-Struktur)",
    page_icon=":chart_with_upwards_trend:",
    menu_items={
        'Report a Bug': 'mailto:luca.meier@sbb.ch',
        'About': "Made with :heart: by [Luca Meier](mailto:luca.meier@sbb.ch)"
    }
)

def transform_data(file_buffer):
    """
    Transformiert die FLBW-Daten aus dem SAP-Export in ein standardisiertes Analyseformat.
    
    Args:
        file_buffer: Der Inhalt der hochgeladenen Excel-Datei
        
    Returns:
        DataFrame: Die transformierten und pivotierten Daten
    """
    # Excel-Datei einlesen
    df = pd.read_excel(file_buffer, sheet_name="Sheet1", header=0)
    
    # Spaltenumbenennung für bessere Lesbarkeit und Standardisierung
    df.rename(columns={
        "OE": "Organisationseinheit",
        "Personalnummer": "U-Nummer",
        "Name des Mitarbeiters bzw. Bewerbers": "Name",
        "Datum": "Datum",
        "Kontierungstext": "Kontierungsbeschreibung",
        "Kontierung (Empf.)": "Kontierungstyp",
        "Allgemeiner Empfänger": "Kontierungsnummer",
        "Kurztext": "Leistung Kurztext",
        "Leistungsart": "Leistungsart",
        "EmpfKostenstelle": "EmpfKostenstelle",
        "Empfänger-PSP-Element": "Projektdefinition",
        "Lohnart-Langtext": "Lohnart-Langtext",
        "Anzahl (Maßeinheit)": "Betrag",
        "Text AnAbArt": "Text AnAbArt"
    }, inplace=True)
    
    # Datumsverarbeitung
    df["Datum"] = pd.to_datetime(df["Datum"], format="%d.%m.%Y", errors="coerce")
    df["Monat"] = df["Datum"].dt.month

    # Initialisierung der Kategorisierungsfelder
    df["Kategorie"] = ""
    df["Unterkategorie"] = ""
    df["Unterkategorie Name"] = ""
    
    # Konvertierung des Betrags in numerische Werte
    df["Betrag"] = pd.to_numeric(df["Betrag"], errors="coerce").fillna(0)
    
    # Definition der ICT-Auftragsnummern für die Kategorisierung
    ict_order_numbers = {
        "170232862", "170232863", "170232864", "170232865", "170232866",
        "170232867", "170232869", "170233584", "170423823", "170423824",
        "170423825", "170423826", "170423827", "170423828", "170423829",
        "170424380", "170424465", "170424663", "170322127"
    }
    
    # Definition der FLBW-Auftragsnummern für die Kategorisierung
    flbw_order_numbers = {
        "170320380", "170320381", "170320382", "170320373", "170319824",
        "152319457", "152318870", "170223872", "170319986", "170424330", "170320783"
    }
    
    # Definition der Abwesenheitsarten für die Kategorisierung
    # Format: {code: (beschreibung, anwesend_flag)}
    abwesenheitsarten = {
        "0100": ("Ferien", 1), "0110": ("Treueprämie (Zeit)", 1), "0120": ("Flexa-Urlaub", 1), "0130": ("Vorruhestandurlaub", 1),
        "0200": ("Krankheit", 1), "0201": ("Krank teilarbeitsfähig", 1), "0210": ("Krankheit (Militär)", 1), "0211": ("K (Mil) teilarbeitsfähig", 1),
        "0220": ("Quarantäne", 1), "0222": ("COVID Kinderbetreuung", 1), "0225": ("COVID Risikogruppe", 1), "0300": ("Militär/Zivilschutz", 1),
        "0400": ("Unfall (NBU)", 1), "0401": ("NBU teilarbeitsfähig", 1), "0405": ("Unfall (BU)", 1), "0406": ("BU teilarbeitsfähig", 1),
        "0410": ("Unfall (Militär)", 1), "0411": ("Unfall (Mil) teilarbeit", 1), "0800": ("Kompensation Pikett", 1), "0900": ("Kompensation Überzeit", 1),
        "0910": ("Komp. Überzeit 2", 1), "0915": ("Komp. Überstunden (ÜS)", 1), "0917": ("Komp. Ausdeh. Höchst-AZ", 1), "0920": ("Komp. Nachtdienst 3", 1),
        "0925": ("Komp. Gleitzeit", 1), "0940": ("Überzeit", 1), "0950": ("Seminar/Kurs", 1), "0960": ("Reisezeitgutschrift", 1), "0970": ("Freistellung PeKo", 1),
        "1000": ("ab.freier Tag (VJ)", 1), "1010": ("Ausgleichstag", 1), "1020": ("Ruhetag", 1), "1030": ("Komp. CTS", 1), "1040": ("Teilzeittag \"80/20\"", 1),
        "1050": ("Teilzeittag", 1), "1060": ("Teilzeittag", 1), "1100": ("Auszeit 25", 1), "1110": ("Auszeit 40", 1), "1120": ("Auszeit in Tagen", 1),
        "1125": ("Auszeit pro rata", 1), "1140": ("Ind. bez. Urlaub IBU", 1), "1150": ("Selbstlernzeit", 1), "8710": ("Weiterbildungsurlaub", 1),
        "8712": ("ausserschul.Jugendarbeit", 1), "8713": ("Arbeitsenthebung", 1), "8714": ("Freisetzung", 1), "8715": ("Mutterschutz", 1),
        "8716": ("Adoptionsurlaub", 1), "8717": ("Urlaub Berufsbildung", 1), "8718": ("Frei nach Personenunfall", 1), "8719": ("Untersuch SUVA", 1),
        "8720": ("Unbezahlter Urlaub", 1), "8721": ("Freistellung", 1), "875A": ("Hochzeit", 1), "875B": ("Vaterschaftsurlaub", 1),
        "875C": ("Tod Ehegatte,Eltern,Kind", 1), "875D": ("Tod SchEltern,Geschwister", 1), "875E": ("Tod GrEltern/UrGrEltern", 1),
        "875F": ("Familiäre Gründe", 1), "875G": ("Pflege der Kinder", 1), "875H": ("Stellenbewerbungen", 1), "875I": ("Wohnungswechsel", 1),
        "875J": ("Vorsprache b. Behörden", 1), "875K": ("Ausübung öffentl. Ämter", 1), "875L": ("Arbeitsjubiläum", 1), "875M": ("Entlassung Wehrpflicht", 1),
        "875N": ("Vaterschaftsurlaub", 1), "875O": ("Orientierungstag Militär", 1), "876A": ("Feuerwehr/Einsatz b.Alarm", 1), "876B": ("Feuerwehr/Kurs", 1),
        "876C": ("Aktiver Spitzensport", 1), "876D": ("Ltg. Behindertensport", 1), "876E": ("Wohnungssuche", 1), "876F": ("1.Mai Veranstaltungen", 0),
        "876G": ("Bild.Veranstalt.Gewerksch", 0), "876H": ("Freiwilliger Zivilschutz", 0), "876I": ("Jugend u.Sport", 0), "877A": ("Urlaubsscheck 1/1", 0),
        "877B": ("Urlaubsscheck 1/2", 0), "878A": ("Mutterschaftsurlaub", 0), "879A": ("Erziehungsurlaub", 0), "879B": ("Gesetzl. Betr.Urlaub Kind", 0),
        "9001": ("Leistung (SPX)", 0), "9070": ("Piketteinsatz (SPX)", 0), "9100": ("Pause bez. (30%) (SPX)", 0), "9150": ("Pause unbezahlt (SPX)", 0),
        "9270": ("Arbeitsunterbrech. (SPX)", 0), "9U00": ("Unterbruch ARF (Ums)", 0), "9U01": ("Unterbruch FLU (Ums)", 0), "I-12": ("Führung/Coaching", 0),
        "I-14": ("f.-int. Team-/Fachs./KT", 0), "I-15": ("f.-üb. Team-/Fachs./KT", 0), "I-16": ("Offerten für Dritte", 0), "I-17": ("IT-Störungen/Unterbrüche", 0),
        "I-18": ("PEKO", 0), "I-19": ("Uebriges", 0)
    }
    
    # Arbeitszeit-Codes (alle Codes mit anwesend_flag = 1)
    arbeitszeit_codes = {code: name for code, (name, flag) in abwesenheitsarten.items() if flag == 1}
    
    # Abwesenheits-Codes (alle Codes mit anwesend_flag = 0)
    abwesenheits_codes = {code: name for code, (name, flag) in abwesenheitsarten.items() if flag == 0}
    
    # Definition der Leistungsarten für Arbeitszeit/Abwesenheit-Kategorisierung
    arbeitszeit_leistungsarten = {
        "LA1620", "LA1611", "LA1402", "LA1619", "LA1617", "LA1002", "LA1615", "LA1824", "LA1825"
    }
    
    # Kategorisierung der Einträge
    df["Kategorie"] = df.apply(
        lambda row: "ICT" if str(row["Kontierungsbeschreibung"]).startswith("PP-UHR ICT") or str(row["Kontierungsnummer"]) in ict_order_numbers else (
            "FLBW" if "FLBW" in str(row["Kontierungsbeschreibung"]) or str(row["Kontierungsnummer"]) in flbw_order_numbers else (
                "PSP" if "PSP" in str(row["Kontierungstyp"]) else (
                    "Anwesenheit" if pd.notna(row["Lohnart-Langtext"]) and str(row["Lohnart-Langtext"]).strip() != "" else (
                        "Arbeitsleistung" if str(row["Leistungsart"]) in arbeitszeit_leistungsarten and str(row["Text AnAbArt"]) in arbeitszeit_codes else (
                            "Abwesenheit" if str(row["Leistungsart"]) in arbeitszeit_leistungsarten and str(row["Text AnAbArt"]) in abwesenheits_codes else "Anderes"
                        )
                    )
                )
            )
        ),
        axis=1
    )
    
    # Hilfsfunktionen für die Unterkategorie-Ableitung
    def extract_number(text, num_digits=7):
        """
        Extrahiert eine Zahl mit der angegebenen Stellenzahl aus einem Text.
        
        Args:
            text: Der zu analysierende Text
            num_digits: Die gewünschte Stellenzahl (Standard: 7)
            
        Returns:
            str: Die extrahierte Zahl oder "Unbekannte Kontierungsnummer"
        """
        matches = re.findall(r"\d+", str(text))
        for m in matches:
            if len(m) == num_digits:
                return m
            elif len(m) > num_digits:
                return m[-num_digits:]
        return "Unbekannte Kontierungsnummer"
    
    def find_keyword(text):
        """
        Sucht nach definierten Schlüsselwörtern am Anfang des Textes.
        
        Args:
            text: Der zu analysierende Text
            
        Returns:
            str: Das gefundene Schlüsselwort oder "XXX"
        """
        possible_keywords = [
            "ABW", "ÄAUF", "EINK", "INNO", "IHE", "MFK", "MDBI", "MON", "NORM", 
            "OBS", "RCM", "REST", "SICH", "STADA", "STUD", "SUE", "SYM", "PSUP",
            "PROD", "IND", "MDG", "ANA", "INST", "ADM", "KURE", "CLR", "CAD", "IHS"
        ]
        text_upper = str(text).upper().strip()
        for keyword in possible_keywords:
            if text_upper.startswith(keyword):
                return keyword
        return "XXX"
    
    # Unterkategorie-Ableitung basierend auf der Kategorie
    df["Unterkategorie"] = df.apply(
        lambda row: (
            extract_number(row["Kontierungsnummer"], num_digits=8) if row["Kategorie"] == "ICT" else (
                find_keyword(row["Leistung Kurztext"]) if row["Kategorie"] == "FLBW" else (
                    extract_number(row["Kontierungsnummer"]) if row["Kategorie"] == "PSP" else (
                        str(row["Kontierungsbeschreibung"]) if row["Kategorie"] in ["Arbeitsleistung", "Abwesenheit", "Anwesenheit"] else ""
                    )
                )
            )
        ),
        axis=1
    )
    
    # Erstellung des Unterkategorie-Namens
    df["Unterkategorie Name"] = df.apply(
        lambda row: row["Unterkategorie"] + " " + row["Projektdefinition"] if row["Kategorie"] == "PSP" else row["Unterkategorie"],
        axis=1
    )
    
    # Status-Erkennung basierend auf Lohnart und AnAbArt
    def status_logik(row):
        """
        Bestimmt den Status eines Eintrags basierend auf Lohnart und AnAbArt.
        
        Args:
            row: Die zu analysierende Zeile
            
        Returns:
            str: Der ermittelte Status (Arbeit Unproduktiv, Arbeit oder Abwesend)
        """
        if pd.notna(row["Lohnart-Langtext"]) and str(row["Lohnart-Langtext"]).strip() != "":
            return "Arbeit Unproduktiv"
        if re.fullmatch(r"2\d{3}", str(row["Text AnAbArt"]).strip()):
            return "Arbeit"
        return "Abwesend"
    # Status direkt in die Spalte 'Text AnAbArt' schreiben
    df["Text AnAbArt"] = df.apply(status_logik, axis=1)
    
    # Definition der statischen Spalten für die Pivotierung
    static_cols = [
        "Organisationseinheit", "U-Nummer", "Name", "Kontierungsbeschreibung",
        "Kontierungstyp", "Kontierungsnummer", "Leistung Kurztext", "Leistungsart",
        "EmpfKostenstelle", "Projektdefinition", "Lohnart-Langtext",
        "Text AnAbArt", "Kategorie", "Unterkategorie", "Unterkategorie Name"
    ] 
    
    # Ersetzung fehlender Werte in den Gruppierungsspalten
    for col in static_cols:
        df[col] = df[col].fillna("Unbekannt")
    
    # Konvertierung der EmpfKostenstelle zu String
    df["EmpfKostenstelle"] = df["EmpfKostenstelle"].fillna("").astype(str)
    
    # Pivotierung der Daten
    pivot_df = df.pivot_table(
        index=static_cols,
        columns="Monat",
        values="Betrag",
        aggfunc="sum",
        fill_value=0
    ).reset_index()

    # Monatsnamen-Mapping
    month_names = {
        1: "Januar", 2: "Februar", 3: "März", 4: "April",
        5: "Mai", 6: "Juni", 7: "Juli", 8: "August",
        9: "September", 10: "Oktober", 11: "November", 12: "Dezember"
    }
    pivot_df.rename(columns=month_names, inplace=True)
    
    # Ermittlung der vorhandenen Monatsspalten
    existing_months = [month_names[m] for m in month_names if month_names[m] in pivot_df.columns]
    
    # Sortierung der Spalten
    pivot_df = pivot_df[static_cols + sorted(existing_months, key=lambda m: list(month_names.values()).index(m))]
    
    # Berechnung der Year-to-Date Summe
    pivot_df["ytd"] = pivot_df[existing_months].sum(axis=1)

    return pivot_df

# Streamlit-Oberfläche
st.title('📈 FLBW Daten Transformation (Neue SAP-Struktur)')

with st.expander("Erklärung"):
    st.markdown("""
    Diese Web-Anwendung transformiert FLBW-Daten aus dem neuen SAP-Exportformat in ein standardisiertes Analyseformat.
    
    **Detaillierte Transformationsschritte:**
    
    1. **Spaltenumbenennung:**  
       Die Originalspalten werden gemäß folgendem Mapping umbenannt:
       - OE → Organisationseinheit
       - Personalnummer → U-Nummer
       - Name des Mitarbeiters bzw. Bewerbers → Name
       - Kontierungstext → Kontierungsbeschreibung
       - Kontierung (Empf.) → Kontierungstyp
       - Allgemeiner Empfänger → Kontierungsnummer
       - Kurztext → Leistung Kurztext
       - EmpfKostenstelle → EmpfKostenstelle
       - Empfänger-PSP-Element → Projektdefinition
       - Anzahl (Maßeinheit) → Betrag
       - Text AnAbArt → Text AnAbArt (wird später zum Status)
    
    2. **Datumsverarbeitung:**  
       - Konvertierung des Datums in das Format DD.MM.YYYY
       - Extraktion des Monats als numerischer Wert (1-12)
    
    3. **Kategorisierung:**  
       Einträge werden in folgende Kategorien eingeteilt:
       - **ICT:** Wenn die Kontierungsbeschreibung mit "PP-UHR ICT" beginnt ODER die Kontierungsnummer in der Liste der ICT-Auftragsnummern enthalten ist
       - **FLBW:** Wenn "FLBW" in der Kontierungsbeschreibung vorkommt
       - **PSP:** Wenn "PSP" im Kontierungstyp enthalten ist
       - **Anderes:** Für alle übrigen Fälle
    
    4. **Unterkategorie-Ableitung:**  
       Je nach Kategorie wird die Unterkategorie wie folgt bestimmt:
       - **ICT:** Extraktion einer 8-stelligen Zahl aus der Kontierungsnummer
       - **FLBW:** Prüfung des Leistung Kurztext auf definierte Schlüsselwörter (z.B. "ABW", "ÄAUF", "EINK", etc.)
       - **PSP:** Extraktion einer 7-stelligen Zahl aus der Kontierungsnummer
       - Für PSP-Einträge wird der Unterkategorie Name als Kombination aus Unterkategorie und Projektdefinition erstellt
    
    5. **Status-Ableitung:**  
       - Die Spalte "Text AnAbArt" wird durch die Status-Logik ersetzt:
         - "Arbeit Unproduktiv" falls Lohnart-Langtext nicht leer ist
         - "Arbeit" falls Text AnAbArt ein vierstelliger Code mit 2 am Anfang ist (z.B. 2001)
         - "Abwesend" in allen anderen Fällen
    
    6. **Datentypen-Bereinigung:**  
       - Die Spalte "EmpfKostenstelle" wird explizit als String behandelt, um Kompatibilitätsprobleme zu vermeiden
    
    7. **Datenaggregation:**  
       - Gruppierung nach allen statischen Feldern (Organisationseinheit, U-Nummer, Name, etc.)
       - Aggregation der Beträge pro Monat
       - Berechnung der Year-to-Date (ytd) Summe über alle Monate
    
    8. **Ausgabeformat:**  
       - Erstellung einer pivotierten Tabelle mit Monatsspalten
       - Umwandlung der Monatsnummern in Monatsnamen (z.B. 1 → Januar)
       - Sortierung der Spalten: zuerst statische Felder, dann Monate chronologisch
    
    **Hinweis:** Die Transformation berücksichtigt fehlende Werte und ersetzt diese durch "Unbekannt" in den Gruppierungsspalten.
    """)

uploaded_file = st.file_uploader("Bitte wählen Sie die Excel-Datei aus", type=["xlsx", "xls"])

if uploaded_file:
    with st.spinner('Daten werden transformiert. Bitte warten...'):
        transformed_df = transform_data(uploaded_file)
        # Excel-Datei im Speicher vorbereiten
        buffer = io.BytesIO()
        transformed_df.to_excel(buffer, index=False)
        excel_data = buffer.getvalue()
    
    st.success('Die Daten wurden erfolgreich transformiert.', icon="✅")
    st.balloons()

    # Excel-Download bereitstellen
    st.download_button(
        label="Transformierte Daten herunterladen",
        data=excel_data,
        file_name="transformed_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.header("Transformierte Daten")
    st.dataframe(transformed_df)