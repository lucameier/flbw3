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
        "170424380", "170424465", "170424663"
    }
    
    # Kategorisierung der Einträge
    df["Kategorie"] = df.apply(
        lambda row: "ICT" if str(row["Kontierungsbeschreibung"]).startswith("PP-UHR ICT") or str(row["Kontierungsnummer"]) in ict_order_numbers else (
            "FLBW" if "FLBW" in str(row["Kontierungsbeschreibung"]) else (
                "PSP" if "PSP" in str(row["Kontierungstyp"]) else "Anderes"
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
                    extract_number(row["Kontierungsnummer"]) if row["Kategorie"] == "PSP" else ""
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