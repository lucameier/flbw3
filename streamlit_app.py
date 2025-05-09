"""
Created on Fri Aug 18 13:52:13 2023
Angepasster Transformer für die neue SAP-Struktur inkl.:
 - ICT-Klassifizierung (über Kontierungsbeschreibung und spezifische Auftragsnummern)
 - Unterkategorie-Ableitung (mit find_keyword, das nur den Textanfang prüft)
 - Pivotierung nach den korrekten Feldern
@author: Luca Meier
"""

import pandas as pd
import streamlit as st
import io
import re

# Streamlit-Seiteneinstellungen
st.set_page_config(
    page_title="FLBW Daten Transformation (Neue SAP-Struktur)",
    page_icon=":chart_with_upwards_trend:",
    menu_items={
        'Report a Bug': 'mailto:luca.meier@sbb.ch',
        'About': "Made with :heart: by [Luca Meier](mailto:luca.meier@sbb.ch)"
    }
)

def transform_data(file_buffer):
    # Excel-Datei aus dem Blatt "Sheet1" einlesen
    df = pd.read_excel(file_buffer, sheet_name="Sheet1", header=0)
    
    # Spalten umbenennen (Mapping an die neuen Bezeichnungen)
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
        "Text AnAbArt": "Text AnAbArt",
        "Ab-/Anwesenheitsart": "Abwesenheitsart"
    }, inplace=True)
    
    # Datum in datetime konvertieren und den Monat extrahieren
    df["Datum"] = pd.to_datetime(df["Datum"], format="%d.%m.%Y", errors="coerce")
    df["Monat"] = df["Datum"].dt.month

    # Zusätzliche Felder initialisieren
    df["Kategorie"] = ""
    df["Unterkategorie"] = ""
    df["Unterkategorie Name"] = ""
    
    # Sicherstellen, dass "Betrag" numerisch ist
    df["Betrag"] = pd.to_numeric(df["Betrag"], errors="coerce").fillna(0)
    
    # Mapping für Abwesenheitsart: Codes werden in ausgeschriebene Form umgewandelt
    abwesenheitsart_mapping = {
        "876F": "1. Mai Veranstaltungen",
        "": "ADM, Führung und Administration",
        "8716": "Adoptionsurlaub",
        "876C": "Aktiver Spitzensport",
        "8713": "Arbeitsenthebung",
        "875L": "Arbeitsjubiläum",
        "2270": "Arbeitsunterbrechung",
        "2000": "Arbeitszeit",
        "8712": "ausserschul. Jugendarbeit",
        "875K": "Ausübung öffentl. Ämter",
        "1120": "Auszeit in Tagen",
        "2002": "BBD",
        "2004": "BBD Nachtdienst 3",
        "2003": "BBD Überzeit 1",
        "876G": "Bild. Veranstaltung Gewerksch.",
        "406": "BU teilarbeitsfähig",
        "875M": "Entlassung Wehrpflicht",
        "879A": "Erziehungsurlaub",
        "875F": "Familiäre Gründe",
        "100": "Ferien",
        "876A": "Feuerwehr / Einsatz b. Alarm",
        "876B": "Feuerwehr / Kurs",
        "970": "Freistellung PeKo",
        "8714": "Freistetzung",
        "876H": "Freiwilliger Zivilschutz",
        "879B": "Gestzl. Betr. Urlaub Kind",
        "875A": "Hochzeit",
        "1140": "Ind. Bez. Urlaub IBU",
        "876I": "Jugend u. Sport",
        "211": "K (Mil) teilarbeitsfähig",
        "9": "Kasernierung",
        "2090": "Kasernierung Überzeit",
        "925": "Komp. Gleitzeit",
        "920": "Komp. Nachtdienst 3",
        "800": "Komp. Pikett",
        "900": "Komp. Überzeit 1",
        "201": "Krank teilarbeitsfähig",
        "200": "Krankheit",
        "210": "Krankheit (Militär)",
        "": "KVP, Kaizen / KVP",
        "": "LA1624, P-UHR RIE/FSY Prod-STD",
        "876D": "Ltg. Behindertensport",
        "300": "Militär / Zivilschutz",
        "878A": "Mutterschaftsurlaub",
        "8715": "Mutterschutz",
        "401": "NBU teilarbeitsfähig",
        "2110": "Pause / Nachtzu / ArbOrt",
        "2100": "Pause AZG auswärts",
        "875G": "Pflege der Kinder",
        "5": "Pikett Mittel",
        "1": "Pikettdienst 1 streng",
        "2": "Pikettdienst 2 normal",
        "3": "Pikettdienst 3 leicht",
        "2070": "Piketteinsatz",
        "960": "Reisezeitgutschrift",
        "2201": "Schichtlage ArG",
        "2202": "Schichtlage AZG",
        "": "SCHULU, Schulungen",
        "950": "Seminar / Kurs",
        "": "SITZ, Sitzungen",
        "875H": "Stellenbewerbungen",
        "875C": "Tod Ehegatte, Eltern, Kind",
        "875E": "Tod GrEltern, UrGrEltern",
        "875D": "Tod SchEltern, Geschwister",
        "110": "Treueprämie (Zeit)",
        "940": "Überzeit",
        "405": "Unfall (BU)",
        "411": "Unfall (Mil) teilarbeit",
        "410": "Unfall (Militär)",
        "400": "Unfall (NBU)",
        "8719": "Untersuch SUVA",
        "8717": "Urlaub Berufsbildung",
        "877A": "Urlaubsscheck 1/1",
        "877B": "Urlaubsscheck 1/2",
        "875N": "Vaterschaftsurlaub",
        "875J": "Vorsprache b. Behörden",
        "8710": "Weiterbildungsurlaub",
        "876E": "Wohnungssuche",
        "875I": "Wohnungswechsel"
    }
    df["Abwesenheitsart"] = df["Abwesenheitsart"].fillna("").str.strip()
    df["Abwesenheitsart"] = df["Abwesenheitsart"].map(lambda x: abwesenheitsart_mapping.get(x, x))
    
    # Kategorisierung:
    # - ICT: wenn die Kontierungsbeschreibung mit "PP-UHR ICT" beginnt
    #        oder wenn die Kontierungsnummer (als Auftragsnummer) in der folgenden Liste enthalten ist.
    ict_order_numbers = {
        "170232862", "170232863", "170232864", "170232865", "170232866",
        "170232867", "170232869", "170233584", "170423823", "170423824",
        "170423825", "170423826", "170423827", "170423828", "170423829",
        "170424380", "170424465", "170424663"
    }
    df["Kategorie"] = df.apply(
        lambda row: "ICT" if str(row["Kontierungsbeschreibung"]).startswith("PP-UHR ICT") or str(row["Kontierungsnummer"]) in ict_order_numbers else (
            "FLBW" if "FLBW" in str(row["Kontierungsbeschreibung"]) else (
                "PSP" if "PSP" in str(row["Kontierungstyp"]) else "Anderes"
            )
        ),
        axis=1
    )
    
    # Ableitung der Unterkategorie:
    # - ICT: 8-stellige Zahl aus der Kontierungsnummer
    # - FLBW: Schlüsselwort aus dem "Leistung Kurztext" (nur am Anfang geprüft)
    # - PSP: 7-stellige Zahl aus der Kontierungsnummer
    def extract_number(text, num_digits=7):
        matches = re.findall(r"\d+", str(text))
        for m in matches:
            if len(m) == num_digits:
                return m
            elif len(m) > num_digits:
                return m[-num_digits:]
        return "Unbekannte Kontierungsnummer"
    
    def find_keyword(text):
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
    
    df["Unterkategorie"] = df.apply(
        lambda row: (
            extract_number(row["Kontierungsnummer"], num_digits=8) if row["Kategorie"] == "ICT" else (
                find_keyword(row["Leistung Kurztext"]) if row["Kategorie"] == "FLBW" else (
                    extract_number(row["Kontierungsnummer"]) if row["Kategorie"] == "PSP" else ""
                )
            )
        ),
        axis=1
    )
    
    df["Unterkategorie Name"] = df.apply(
        lambda row: row["Unterkategorie"] + " " + row["Projektdefinition"] if row["Kategorie"] == "PSP" else row["Unterkategorie"],
        axis=1
    )
    df
    # Definiere die statischen Spalten, die als Index in der Pivotierung genutzt werden sollen
    static_cols = [
        "Organisationseinheit", "U-Nummer", "Name", "Kontierungsbeschreibung",
        "Kontierungstyp", "Kontierungsnummer", "Leistung Kurztext", "Leistungsart",
        "EmpfKostenstelle", "Projektdefinition", "Lohnart-Langtext",
        "Text AnAbArt", "Abwesenheitsart", "Kategorie", "Unterkategorie", "Unterkategorie Name"
    ] 
    
    static_cols2 = [
        "Organisationseinheit", "U-Nummer", "Name", "Kontierungsbeschreibung",
        "Kontierungstyp", "Kontierungsnummer", "Leistung Kurztext", "Leistungsart",
        "EmpfKostenstelle", "Projektdefinition", "Lohnart-Langtext",
        "Text AnAbArt", "Abwesenheitsart", "Kategorie", "Unterkategorie", "Unterkategorie Name"
    ]
    
    # Fehlende Werte in Gruppierungsspalten ersetzen
    for col in static_cols2:
        df[col] = df[col].fillna("Unbekannt")
    
    # Pivotierung: Aggregiere den Betrag pro Gruppe (definiert durch die statischen Felder) und Monat
    pivot_df = df.pivot_table(
        index=static_cols2,
        columns="Monat",
        values="Betrag",
        aggfunc="sum",
        fill_value=0
    ).reset_index()

    # Optional: Mapping von Monatsnummern zu Monatsnamen
    month_names = {
        1: "Januar", 2: "Februar", 3: "März", 4: "April",
        5: "Mai", 6: "Juni", 7: "Juli", 8: "August",
        9: "September", 10: "Oktober", 11: "November", 12: "Dezember"
    }
    pivot_df.rename(columns=month_names, inplace=True)
    
    # Ermittele, welche Monats-Spalten tatsächlich vorhanden sind
    existing_months = [month_names[m] for m in month_names if month_names[m] in pivot_df.columns]
    
    # Sortiere die Spalten: zuerst die statischen Spalten, dann die Monats-Spalten in chronologischer Reihenfolge
    pivot_df = pivot_df[static_cols + sorted(existing_months, key=lambda m: list(month_names.values()).index(m))]
    
    # Berechne die "ytd" (Year-to-Date)-Summe über alle vorhandenen Monats-Spalten
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
    
    2. **Datumsverarbeitung:**  
       - Konvertierung des Datums in das Format DD.MM.YYYY
       - Extraktion des Monats als numerischer Wert (1-12)
    
    3. **Abwesenheitsart-Mapping:**  
       Standardisierung der Abwesenheitsarten auf einheitliche Codes (z.B. "Ferien" → "100", "Krankheit" → "200")
    
    4. **Kategorisierung:**  
       Einträge werden in folgende Kategorien eingeteilt:
       - **ICT:** Wenn die Kontierungsbeschreibung mit "PP-UHR ICT" beginnt ODER die Kontierungsnummer in der Liste der ICT-Auftragsnummern enthalten ist
       - **FLBW:** Wenn "FLBW" in der Kontierungsbeschreibung vorkommt
       - **PSP:** Wenn "PSP" im Kontierungstyp enthalten ist
       - **Anderes:** Für alle übrigen Fälle
    
    5. **Unterkategorie-Ableitung:**  
       Je nach Kategorie wird die Unterkategorie wie folgt bestimmt:
       - **ICT:** Extraktion einer 8-stelligen Zahl aus der Kontierungsnummer
       - **FLBW:** Prüfung des Leistung Kurztext auf definierte Schlüsselwörter (z.B. "ABW", "ÄAUF", "EINK", etc.)
       - **PSP:** Extraktion einer 7-stelligen Zahl aus der Kontierungsnummer
       - Für PSP-Einträge wird der Unterkategorie Name als Kombination aus Unterkategorie und Projektdefinition erstellt
    
    6. **Datenaggregation:**  
       - Gruppierung nach allen statischen Feldern (Organisationseinheit, U-Nummer, Name, etc.)
       - Aggregation der Beträge pro Monat
       - Berechnung der Year-to-Date (ytd) Summe über alle Monate
    
    7. **Ausgabeformat:**  
       - Erstellung einer pivotierten Tabelle mit Monatsspalten
       - Umwandlung der Monatsnummern in Monatsnamen (z.B. 1 → Januar)
       - Sortierung der Spalten: zuerst statische Felder, dann Monate chronologisch
    
    **Hinweis:** Die Transformation berücksichtigt fehlende Werte und ersetzt diese durch "Unbekannt" in den Gruppierungsspalten.
    """)

uploaded_file = st.file_uploader("Bitte wählen Sie die Excel-Datei aus", type=["xlsx", "xls"])

if uploaded_file:
    with st.spinner('Daten werden transformiert. Bitte warten...'):
        transformed_df = transform_data(uploaded_file)
    st.success('Die Daten wurden erfolgreich transformiert.', icon="✅")
    st.balloons()

    # Excel-Download bereitstellen
    buffer = io.BytesIO()
    transformed_df.to_excel(buffer, index=False)
    st.download_button(
        label="Transformierte Daten herunterladen",
        data=buffer,
        file_name="transformed_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.header("Transformierte Daten")
    st.dataframe(transformed_df)