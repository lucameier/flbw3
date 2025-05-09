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
uploaded_file = st.sidebar.file_uploader(
    "Excel-Datei mit Arbeitszeitdaten hochladen", type=["xlsx"]
)
if not uploaded_file:
    st.sidebar.warning("Bitte eine Excel-Datei hochladen, um das Dashboard zu nutzen.")
    st.stop()

df = pd.read_excel(uploaded_file)

# Filteroptionen
org_units = sorted(df['Organisationseinheit'].unique())
selected_org = st.sidebar.multiselect("Organisationseinheit", options=org_units, default=org_units)

names = sorted(df['Name'].unique())
selected_names = st.sidebar.multiselect("Mitarbeiter", options=names, default=names)

kategorien = sorted(df['Kategorie'].unique())
selected_kat = st.sidebar.multiselect("Kategorie", options=kategorien, default=kategorien)

# Gefilterte Daten
df_filtered = df[
    df['Organisationseinheit'].isin(selected_org) &
    df['Name'].isin(selected_names) &
    df['Kategorie'].isin(selected_kat)
]

# Monatsdaten reformattieren
monthly_cols = ['Januar','Februar','März','April','Mai','Juni','Juli','August','September','Oktober','November','Dezember']
if set(monthly_cols).issubset(df_filtered.columns):
    df_melt = df_filtered.melt(
        id_vars=[col for col in df_filtered.columns if col not in monthly_cols],
        value_vars=monthly_cols,
        var_name='Monat',
        value_name='Stunden'
    )
    mapping = {m: i+1 for i, m in enumerate(monthly_cols)}
    df_melt['MonatNum'] = df_melt['Monat'].map(mapping)
    df_melt['Datum'] = df_melt['MonatNum'].apply(lambda m: datetime(datetime.now().year, m, 1))
else:
    st.error("Monats-Spalten fehlen im Datensatz. Überprüfe die Datei.")
    st.stop()

# Tabs für Navigation
tab_overview, tab_trends, tab_employees, tab_analysis, tab_data = st.tabs([
    "📊 Übersicht", "📈 Trends", "👥 Mitarbeiter", "🔧 Analysen", "🗂️ Rohdaten"
])

# Übersicht-Tab
with tab_overview:
    st.header("Kennzahlen")
    cols = st.columns(6)
    total_ytd = df_filtered['ytd'].sum()
    avg_hours = df_filtered.groupby('Name')['ytd'].sum().mean()
    num_emp = df_filtered['U-Nummer'].nunique()
    num_proj = df_filtered['Projektdefinition'].nunique()
    num_leistung = df_filtered['Leistungsart'].nunique()
    num_kostenstelle = df_filtered['EmpfKostenstelle'].nunique()
    metrics = [
        ("Gesamtstunden (YTD)", f"{total_ytd:,.0f}"),
        ("Ø Stunden/Mitarbeiter", f"{avg_hours:,.1f}"),
        ("Mitarbeiter gesamt", num_emp),
        ("Projekte gesamt", num_proj),
        ("Leistungsarten gesamt", num_leistung),
        ("Kostenstellen gesamt", num_kostenstelle)
    ]
    for col, (label, value) in zip(cols, metrics):
        col.metric(label, value)

    st.markdown("---")
    st.subheader("Stundensumme nach Organisationseinheit")
    org_data = df_filtered.groupby('Organisationseinheit')['ytd'].sum().reset_index()
    chart_org = alt.Chart(org_data).mark_bar().encode(
        x=alt.X('ytd:Q', title='Stunden'),
        y=alt.Y('Organisationseinheit:N', sort='-x'),
        tooltip=['Organisationseinheit','ytd']
    ).interactive()
    st.altair_chart(chart_org, use_container_width=True)

# Trends-Tab
with tab_trends:
    st.header("Monatliche Entwicklung (YTD)")
    trend_data = df_melt.groupby('Datum')['Stunden'].sum().reset_index()
    chart_line = alt.Chart(trend_data).mark_line(point=True).encode(
        x=alt.X('Datum:T', title='Monat'),
        y=alt.Y('Stunden:Q', title='Summe Stunden'),
        tooltip=[alt.Tooltip('Datum:T', title='Monat'), alt.Tooltip('Stunden:Q', title='Stunden')]
    ).interactive()
    st.altair_chart(chart_line, use_container_width=True)

    st.subheader("Heatmap: Stunden nach Kategorie & Monat")
    heat_data = df_melt.groupby(['Monat','Kategorie'])['Stunden'].sum().reset_index()
    chart_heat = alt.Chart(heat_data).mark_rect().encode(
        x=alt.X('Monat:N', sort=monthly_cols),
        y=alt.Y('Kategorie:N'),
        color=alt.Color('Stunden:Q', title='Stunden'),
        tooltip=['Monat','Kategorie','Stunden']
    ).interactive()
    st.altair_chart(chart_heat, use_container_width=True)

# Mitarbeiter-Tab
with tab_employees:
    st.header("Top-Mitarbeiter nach YTD-Stunden")
    top_n = st.slider("Anzahl Top-Mitarbeiter", 5, 20, 10)
    emp_data = df_filtered.groupby('Name')['ytd'].sum().reset_index().nlargest(top_n, 'ytd')
    chart_emp = alt.Chart(emp_data).mark_bar().encode(
        x=alt.X('ytd:Q', title='Stunden'),
        y=alt.Y('Name:N', sort='-x'),
        tooltip=['Name','ytd']
    )
    st.altair_chart(chart_emp, use_container_width=True)

    with st.expander("Details zu Top-Mitarbeitern"):
        st.table(emp_data)

# Analysen-Tab
with tab_analysis:
    st.header("Vertiefte Analysen")
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Stunden nach Leistungsart")
        art_data = df_filtered.groupby('Leistungsart')['ytd'].sum().reset_index().sort_values('ytd', ascending=False)
        chart_art = alt.Chart(art_data).mark_bar().encode(
            x=alt.X('ytd:Q', title='Stunden'),
            y=alt.Y('Leistungsart:N', sort='-x'),
            tooltip=['Leistungsart','ytd']
        ).interactive()
        st.altair_chart(chart_art, use_container_width=True)

    with col2:
        st.subheader("Stunden nach EmpfKostenstelle")
        cost_data = df_filtered.groupby('EmpfKostenstelle')['ytd'].sum().reset_index().sort_values('ytd', ascending=False)
        chart_cost = alt.Chart(cost_data).mark_bar().encode(
            x=alt.X('ytd:Q', title='Stunden'),
            y=alt.Y('EmpfKostenstelle:N', sort='-x'),
            tooltip=['EmpfKostenstelle','ytd']
        ).interactive()
        st.altair_chart(chart_cost, use_container_width=True)

    st.subheader("Stunden nach Kontierungstyp")
    type_data = df_filtered.groupby('Kontierungstyp')['ytd'].sum().reset_index().sort_values('ytd', ascending=False)
    chart_type = alt.Chart(type_data).mark_bar().encode(
        x=alt.X('ytd:Q', title='Stunden'),
        y=alt.Y('Kontierungstyp:N', sort='-x'),
        tooltip=['Kontierungstyp','ytd']
    ).interactive()
    st.altair_chart(chart_type, use_container_width=True)

    st.subheader("Top 10 Projekte nach Stunden")
    proj_data = df_filtered.groupby('Projektdefinition')['ytd'].sum().reset_index().nlargest(10,'ytd')
    chart_proj = alt.Chart(proj_data).mark_bar().encode(
        x=alt.X('ytd:Q', title='Stunden'),
        y=alt.Y('Projektdefinition:N', sort='-x'),
        tooltip=['Projektdefinition','ytd']
    ).interactive()
    st.altair_chart(chart_proj, use_container_width=True)

# Rohdaten-Tab
with tab_data:
    st.header("Gefilterte Rohdaten")
    with st.expander("Tabelle anzeigen"):
        st.dataframe(df_filtered.reset_index(drop=True), height=400)

    csv = df_filtered.to_csv(index=False, sep=';').encode('utf-8-sig')
    st.download_button(
        label="📥 Als CSV herunterladen",
        data=csv,
        file_name='FLBW_Arbeitszeiten_Filtered.csv',
        mime='text/csv'
    )
