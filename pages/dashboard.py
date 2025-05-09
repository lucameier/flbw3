# pages/2_Arbeitszeiten.py

import pandas as pd
import streamlit as st
import altair as alt
import plotly.express as px

def show_arbeitszeiten():
    st.header("FLBW Arbeitszeiten Dashboard (Erweitert)")

    # --- 1) Datei-Upload
    uploaded = st.file_uploader("Bitte Excel-Datei hochladen", type=["xlsx"])
    if not uploaded:
        st.info("Lade hier deine `transformed_data`-Datei hoch, um das Dashboard zu starten.")
        return
    df = pd.read_excel(uploaded)

    # --- 2) Sidebar-Filter
    st.sidebar.markdown("## Filter")
    def ms_filter(label, col):
        opts = sorted(df[col].dropna().unique())
        return st.sidebar.multiselect(label, opts, default=opts)
    sel = {
        "OE":   ms_filter("Organisationseinheit", "Organisationseinheit"),
        "KT":   ms_filter("Kontierungstyp",        "Kontierungstyp"),
        "LA":   ms_filter("Leistungsart",          "Leistungsart"),
        "Kat":  ms_filter("Kategorie",             "Kategorie"),
        "UKat": ms_filter("Unterkategorie",        "Unterkategorie"),
        "EKSt": ms_filter("EmpfKostenstelle",      "EmpfKostenstelle")
    }
    mask = (
        df["Organisationseinheit"].isin(sel["OE"]) &
        df["Kontierungstyp"].isin(sel["KT"]) &
        df["Leistungsart"].isin(sel["LA"]) &
        df["Kategorie"].isin(sel["Kat"]) &
        df["Unterkategorie"].isin(sel["UKat"]) &
        df["EmpfKostenstelle"].isin(sel["EKSt"])
    )
    dff = df[mask].copy()
    if dff.empty:
        st.warning("Keine Daten für diese Filterkombination.")
        return

    # --- 3) Monatsspalten & Basis-KPIs
    monate = ["Januar","Februar","März","April","Mai","Juni",
              "Juli","August","September","Oktober","November","Dezember"]

    total_ytd       = dff["ytd"].sum()
    emp_ytd         = dff.groupby("U-Nummer")["ytd"].sum().reset_index()
    median_ytd_ma   = emp_ytd["ytd"].median()
    std_ytd_ma      = emp_ytd["ytd"].std()
    cv_ytd_ma       = (std_ytd_ma / median_ytd_ma) if median_ytd_ma else 0
    avg_per_month   = dff[monate].sum(axis=1).mean()
    peak_month      = dff[monate].sum().idxmax()
    low_month       = dff[monate].sum().idxmin()
    emp_count       = emp_ytd["U-Nummer"].nunique()
    proj_count      = dff["Projektdefinition"].nunique()
    # Total Überstunden (>160h/Monat)
    overtime        = dff[monate].sub(160).clip(lower=0).sum().sum()
    # Beispiel: Anteil Leistungsart "Operation"
    if "Operation" in dff["Leistungsart"].unique():
        op_share = dff.query("Leistungsart=='Operation'")["ytd"].sum() / total_ytd
    else:
        op_share = None

    # KPI-Übersicht
    cols1 = st.columns(6)
    cols1[0].metric("Gesamt YTD-Stunden",       f"{total_ytd:,.0f}")
    cols1[1].metric("Median YTD/Stunde pro MA", f"{median_ytd_ma:,.1f}")
    cols1[2].metric("StdDev YTD/Stunde MA",     f"{std_ytd_ma:,.1f}")
    cols1[3].metric("CV YTD/Stunde MA",         f"{cv_ytd_ma:.2%}")
    cols1[4].metric("Ø Monatsstunden",          f"{avg_per_month:,.1f}")
    cols1[5].metric("Mitarbeitende",            f"{emp_count}")

    cols2 = st.columns(5)
    cols2[0].metric("Projekte gesamt",         f"{proj_count}")
    cols2[1].metric("Stärkster Monat",         peak_month)
    cols2[2].metric("Schwächster Monat",       low_month)
    cols2[3].metric("Total Überstunden",       f"{overtime:,.0f}")
    cols2[4].metric("Operation-Anteil",        f"{op_share:.1%}" if op_share is not None else "n/a")

    # --- 4) Tabs für Visualisierungen
    tab_labels = [
        "Advanced Insights",
        "Trend & Wachstum",
        "Verteilung & Boxplot",
        "Kostenstellen & Kategorien",
        "Projekt-Insights"
    ]
    tabs = st.tabs(tab_labels)

    # --- 4.1) Advanced Insights
    with tabs[0]:
        st.subheader("Korrelations-Heatmap der Monatsstunden")
        corr = dff[monate].corr().reset_index().melt(
            id_vars="index", var_name="Monat2", value_name="Corr"
        ).rename(columns={"index":"Monat1"})
        thr = st.slider("Min. Korrelations-Schwelle", 0.0, 1.0, 0.5, 0.05)
        heat = (
            alt.Chart(corr.query("abs(Corr) >= @thr"))
               .mark_rect()
               .encode(
                   x=alt.X("Monat1:N", sort=monate),
                   y=alt.Y("Monat2:N", sort=monate),
                   color=alt.Color("Corr:Q", scale=alt.Scale(scheme="redblue", domain=[-1,1])),
                   tooltip=["Monat1","Monat2","Corr"]
               )
               .properties(height=400, width=400)
        )
        st.altair_chart(heat, use_container_width=False)

        st.markdown("---")
        st.subheader("Sunburst: Kategorie → Unterkategorie → EmpfKostenstelle")
        sun_df = (
            dff.groupby(["Kategorie","Unterkategorie","EmpfKostenstelle"])["ytd"]
               .sum().reset_index()
        )
        fig_sb = px.sunburst(
            sun_df,
            path=["Kategorie","Unterkategorie","EmpfKostenstelle"],
            values="ytd",
            title="Stundenverteilung hierarchisch"
        )
        st.plotly_chart(fig_sb, use_container_width=True)

        st.markdown("---")
        st.subheader("Interaktive Tabellen-Analyse")
        st.info("Sortiere und filtere direkt in der Tabelle.")
        st.dataframe(dff, use_container_width=True)

    # --- 4.2) Trend & Wachstum
    with tabs[1]:
        st.subheader("Monatlicher Verlauf & YoY-Wachstum")
        mon_sum = dff[monate].sum().reset_index()
        mon_sum.columns = ["Monat","Stunden"]
        area = (
            alt.Chart(mon_sum)
               .mark_area(opacity=0.3)
               .encode(
                   x=alt.X("Monat:N", sort=monate),
                   y="Stunden:Q",
                   tooltip=["Monat","Stunden"]
               )
        )
        line_ma = (
            alt.Chart(mon_sum)
               .mark_line(color="steelblue")
               .transform_window(
                   rolling_mean="mean(Stunden)",
                   frame=[-2, 0]
               )
               .encode(
                   x=alt.X("Monat:N", sort=monate),
                   y="rolling_mean:Q",
                   tooltip=alt.Tooltip("rolling_mean:Q", title="3-Monats MA")
               )
        )
        st.altair_chart((area + line_ma).interactive(), use_container_width=True)

        mon_growth = mon_sum.assign(pct=lambda df_: df_["Stunden"].pct_change()*100).dropna()
        bar_growth = (
            alt.Chart(mon_growth)
               .mark_bar(color="orange")
               .encode(
                   x=alt.X("Monat:N", sort=monate),
                   y=alt.Y("pct:Q", title="% Wachstum"),
                   tooltip=["Monat","pct"]
               )
        )
        st.altair_chart(bar_growth.interactive(), use_container_width=True)

    # --- 4.3) Verteilung & Boxplot
    with tabs[2]:
        st.subheader("Verteilung der YTD-Stunden pro Mitarbeitenden")
        hist = (
            alt.Chart(emp_ytd)
               .mark_bar()
               .encode(
                   alt.X("ytd:Q", bin=alt.Bin(maxbins=30), title="YTD-Stunden"),
                   y="count():Q",
                   tooltip=["count()"]
               )
        )
        box = (
            alt.Chart(dff)
               .mark_boxplot(extent="min-max")
               .encode(
                   x=alt.X("Organisationseinheit:N", title="OE"),
                   y=alt.Y("ytd:Q", title="YTD-Stunden"),
                   color="Organisationseinheit:N"
               )
        )
        st.altair_chart(hist & box, use_container_width=True)

    # --- 4.4) Kostenstellen & Kategorien
    with tabs[3]:
        st.subheader("Stunden nach EmpfKostenstelle & Kategorie")
        cat_sum = dff.groupby("Kategorie")["ytd"].sum().reset_index()
        pie = (
            alt.Chart(cat_sum)
               .mark_arc(innerRadius=50)
               .encode(
                   theta=alt.Theta("ytd:Q"),
                   color=alt.Color("Kategorie:N"),
                   tooltip=["Kategorie","ytd"]
               )
        )
        cs_sum = (
            dff.groupby(["EmpfKostenstelle","Leistungsart"])["ytd"]
               .sum().reset_index()
        )
        stack = (
            alt.Chart(cs_sum)
               .mark_bar()
               .encode(
                   x=alt.X("ytd:Q", title="YTD-Stunden"),
                   y=alt.Y("EmpfKostenstelle:N", sort='-x'),
                   color="Leistungsart:N",
                   tooltip=["EmpfKostenstelle","Leistungsart","ytd"]
               )
        )
        st.altair_chart(pie | stack, use_container_width=True)

    # --- 4.5) Projekt-Insights
    with tabs[4]:
        st.subheader("Top-Projekte & Scatter-Analyse")
        topn = st.slider("Anzahl Top Projekte", 5, 30, 10)
        proj_sum = (
            dff.groupby("Projektdefinition")["ytd"]
               .sum().reset_index()
               .nlargest(topn, "ytd")
        )
        bar = (
            alt.Chart(proj_sum)
               .mark_bar(cornerRadiusTopLeft=3, cornerRadiusTopRight=3)
               .encode(
                   x=alt.X("ytd:Q", title="YTD-Stunden"),
                   y=alt.Y("Projektdefinition:N", sort="-x", title="Projekt"),
                   tooltip=["Projektdefinition","ytd"]
               )
        )
        scatter_df = emp_ytd.merge(
            dff[["U-Nummer"]+monate].groupby("U-Nummer").mean().reset_index(),
            on="U-Nummer", suffixes=("_ytd","_avgm")
        )
        scatter = (
            alt.Chart(scatter_df)
               .mark_circle(size=50)
               .encode(
                   x=alt.X("ytd:Q", title="YTD-Stunden"),
                   y=alt.Y("März:Q", title="Ø Monatsstunden (März)"),
                   tooltip=["U-Nummer","ytd","März"]
               )
               .interactive()
        )
        st.altair_chart(bar & scatter, use_container_width=True)

    # --- 5) Rohdaten & Download
    with st.expander("Rohdaten anzeigen"):
        st.dataframe(dff, use_container_width=True)
        csv = dff.to_csv(index=False).encode("utf-8")
        st.download_button("Daten als CSV herunterladen", csv,
                           "arbeitszeiten_export.csv", "text/csv")


# Funktion beim Seitenaufruf ausführen
show_arbeitszeiten()
