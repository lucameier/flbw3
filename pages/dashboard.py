import pandas as pd
import streamlit as st
import altair as alt
from datetime import datetime

# Streamlit-Seiteneinstellungen
st.set_page_config(
    page_title="FLBW Arbeitszeiten Dashboard",
    page_icon=":chart_with_upwards_trend:",
    layout="wide",
    menu_items={
        'Report a Bug': 'mailto:luca.meier@sbb.ch',
        'About': "Made with :heart: by [Luca Meier](mailto:luca.meier@sbb.ch)"
    }
)

# Sidebar: Datenquelle und Filter
st.sidebar.header("Datenquelle & Filter")
uploaded_file = st.sidebar.file_uploader("Excel-Datei mit Arbeitszeitdaten hochladen", type=["xlsx"])
if not uploaded_file:
    st.sidebar.warning("Bitte eine Excel-Datei hochladen, um das Dashboard zu nutzen.")
    st.stop()

df = pd.read_excel(uploaded_file)

# Filter: Organisationseinheit, Mitarbeiter, Kategorie
org_units = sorted(df['Organisationseinheit'].unique())
selected_org = st.sidebar.multiselect("Organisationseinheit", options=org_units, default=org_units)

names = sorted(df['Name'].unique())
selected_names = st.sidebar.multiselect("Mitarbeiter", options=names, default=names)

kategorien = sorted(df['Kategorie'].unique())
selected_kat = st.sidebar.multiselect("Kategorie", options=kategorien, default=kategorien)

# Filter anwenden
df_filtered = df[
    df['Organisationseinheit'].isin(selected_org) &
    df['Name'].isin(selected_names) &
    df['Kategorie'].isin(selected_kat)
]

# Monatsdaten lange Form
monthly_cols = ['Januar','Februar','März','April','Mai','Juni','Juli','August','September','Oktober','November','Dezember']
if set(monthly_cols).issubset(df_filtered.columns):
    df_melt = df_filtered.melt(
        id_vars=['Name','Organisationseinheit','Kategorie','ytd'],
        value_vars=monthly_cols,
        var_name='Monat',
        value_name='Stunden'
    )
    # Monat in Zahl und Datum umwandeln für Charts
    monat_mapping = {m:i+1 for i,m in enumerate(monthly_cols)}
    df_melt['MonatNum'] = df_melt['Monat'].map(monat_mapping)
    df_melt['Datum'] = df_melt['MonatNum'].apply(lambda m: datetime(datetime.now().year, m, 1))
else:
    st.error("Die erwarteten Monats-Spalten fehlen.")
    st.stop()

# Tabs für Layout
tab_overview, tab_trends, tab_employees, tab_data = st.tabs([
    "📊 Übersicht", "📈 Trends", "👥 Mitarbeiter", "🗂️ Rohdaten"
])

# Übersicht-Tab
with tab_overview:
    st.header("Kennzahlen")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        total_ytd = df_filtered['ytd'].sum()
        st.metric("Gesamtstunden (YTD)", f"{total_ytd:,.0f}")
    with c2:
        avg_record = df_melt.groupby('Name')['Stunden'].sum().mean()
        st.metric("Ø Stunden pro Mitarbeiter", f"{avg_record:,.1f}")
    with c3:
        num_employees = df_filtered['U-Nummer'].nunique()
        st.metric("Mitarbeiter gesamt", num_employees)
    with c4:
        num_projects = df_filtered['Projektdefinition'].nunique()
        st.metric("Projekte gesamt", num_projects)
    
    st.markdown("---")
    st.subheader("Stundensumme nach Organisationseinheit")
    bar_org = alt.Chart(
        df_filtered.groupby('Organisationseinheit')['ytd'].sum().reset_index()
    ).mark_bar().encode(
        x=alt.X('ytd:Q', title='Stunden (YTD)'),
        y=alt.Y('Organisationseinheit:N', sort='-x'),
        tooltip=['Organisationseinheit','ytd']
    ).interactive()
    st.altair_chart(bar_org, use_container_width=True)

# Trends-Tab
with tab_trends:
    st.header("Monatliche Entwicklung")
    line = alt.Chart(
        df_melt.groupby('Datum')['Stunden'].sum().reset_index()
    ).mark_line(point=True).encode(
        x=alt.X('Datum:T', title='Monat'),
        y=alt.Y('Stunden:Q', title='Summe Stunden'),
        tooltip=[alt.Tooltip('Datum:T', title='Monat'), alt.Tooltip('Stunden:Q', title='Stunden')]
    ).interactive()
    st.altair_chart(line, use_container_width=True)

    st.subheader("Heatmap: Stunden nach Kategorie und Monat")
    heatmap = alt.Chart(
        df_melt.groupby(['Monat','Kategorie'])['Stunden'].sum().reset_index()
    ).mark_rect().encode(
        x=alt.X('Monat:N', sort=monthly_cols),
        y=alt.Y('Kategorie:N'),
        color=alt.Color('Stunden:Q', title='Stunden'),
        tooltip=['Monat','Kategorie','Stunden']
    ).interactive()
    st.altair_chart(heatmap, use_container_width=True)

# Mitarbeiter-Tab
with tab_employees:
    st.header("Top-Mitarbeiter nach YTD")
    top_n = st.slider("Anzahl der Top-Mitarbeiter", min_value=5, max_value=20, value=10)
    top_emps = df_filtered.groupby('Name')['ytd'].sum().reset_index().nlargest(top_n, 'ytd')
    bar_emp = alt.Chart(top_emps).mark_bar().encode(
        x=alt.X('ytd:Q', title='Stunden (YTD)'),
        y=alt.Y('Name:N', sort='-x'),
        tooltip=['Name','ytd']
    )
    st.altair_chart(bar_emp, use_container_width=True)

    with st.expander("Details zu Top-Mitarbeitern anzeigen"):
        st.table(top_emps)

# Rohdaten-Tab
with tab_data:
    st.header("Gefilterte Rohdaten")
    with st.expander("Daten anzeigen"):
        st.dataframe(df_filtered.reset_index(drop=True), height=400)

    # Download der CSV
    csv = df_filtered.to_csv(index=False, sep=';').encode('utf-8-sig')
    st.download_button(
        label="📥 Als CSV herunterladen",
        data=csv,
        file_name='FLBW_Arbeitszeiten_Filtered.csv',
        mime='text/csv'
    )
