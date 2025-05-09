import streamlit as st
import pandas as pd
import plotly.express as px

# Load data
@st.cache_data
def load_data(path: str):
    df = pd.read_excel(path)
    # Melt monthly columns
    months = ['Januar','Februar','März','April','Mai','Juni','Juli','August','September','Oktober','November','Dezember']
    df_melt = df.melt(
        id_vars=[col for col in df.columns if col not in months + ['ytd']],
        value_vars=months,
        var_name='Monat',
        value_name='Wert'
    )
    # Ensure month ordering
    df_melt['Monat'] = pd.Categorical(df_melt['Monat'], categories=months, ordered=True)
    return df, df_melt

# Path to your data file
DATA_PATH = 'transformed_data (9).xlsx'

df, df_melt = load_data(DATA_PATH)

# Sidebar filters
st.sidebar.header("Filter")
unit_filter = st.sidebar.multiselect(
    "Organisationseinheit", sorted(df['Organisationseinheit'].unique()), default=None
)
cat_filter = st.sidebar.multiselect(
    "Kategorie", sorted(df['Kategorie'].unique()), default=None
)
group_by = st.sidebar.selectbox(
    "Gruppieren nach", options=['Kategorie', 'Unterkategorie', 'Organisationseinheit'], index=0
)

# Apply filters
df_filt = df.copy()
df_melt_filt = df_melt.copy()
if unit_filter:
    df_filt = df_filt[df_filt['Organisationseinheit'].isin(unit_filter)]
    df_melt_filt = df_melt_filt[df_melt_filt['Organisationseinheit'].isin(unit_filter)]
if cat_filter:
    df_filt = df_filt[df_filt['Kategorie'].isin(cat_filter)]
    df_melt_filt = df_melt_filt[df_melt_filt['Kategorie'].isin(cat_filter)]

# Title and KPIs
st.title("Kosten- und Leistungsdashboard")
col1, col2 = st.columns(2)
with col1:
    total_ytd = df_filt['ytd'].sum()
    st.metric("Total YTD (Summe)", f"{total_ytd:,.2f}")
with col2:
    # Last non-zero month value\    
    month_sums = df_melt_filt.groupby('Monat')['Wert'].sum()
    nonzero = month_sums[month_sums > 0]
    if not nonzero.empty:
        last_month = nonzero.index[-1]
        last_val = nonzero.iloc[-1]
        st.metric(f"Letzter Monat ({last_month})", f"{last_val:,.2f}")
    else:
        st.metric("Letzter Monat", "Keine Daten")

st.markdown("---")

# Time series
st.header("Monatlicher Verlauf")
ts = df_melt_filt.groupby(['Monat', group_by])['Wert'].sum().reset_index()
fig_ts = px.line(
    ts, x='Monat', y='Wert', color=group_by, markers=True,
    title="Monatliche Entwicklung"
)
st.plotly_chart(fig_ts, use_container_width=True)

# Top Unterkategorien nach YTD
st.header("Top Unterkategorien nach YTD")
top_sub = df_filt.groupby('Unterkategorie')['ytd'].sum().nlargest(10).reset_index()
fig_bar = px.bar(
    top_sub, x='ytd', y='Unterkategorie', orientation='h',
    title="Top 10 Unterkategorien (YTD)",
    labels={'ytd':'YTD Summe', 'Unterkategorie':'Unterkategorie'}
)
st.plotly_chart(fig_bar, use_container_width=True)

# Treemap Kategorie & Unterkategorie
st.header("Strukturübersicht")
tree = df_filt.groupby(['Kategorie','Unterkategorie'])['ytd'].sum().reset_index()
fig_tree = px.treemap(
    tree, path=['Kategorie','Unterkategorie'], values='ytd',
    title="Treemap: Kategorie & Unterkategorie (YTD)"
)
st.plotly_chart(fig_tree, use_container_width=True)

# Heatmap Monate vs. Kategorie
st.header("Heatmap: Monate vs. Kategorie")
heat_data = df_melt_filt.groupby(['Kategorie','Monat'])['Wert'].sum().reset_index()
heat_pivot = heat_data.pivot(index='Kategorie', columns='Monat', values='Wert').fillna(0)
fig_heat = px.imshow(
    heat_pivot,
    labels=dict(x="Monat", y="Kategorie", color="Summe"),
    title="Heatmap der monatlichen Werte nach Kategorie"
)
st.plotly_chart(fig_heat, use_container_width=True)

st.sidebar.markdown("---")
st.sidebar.text("Dashboard erstellt mit Streamlit & Plotly")
