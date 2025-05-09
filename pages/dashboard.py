import pandas as pd
import streamlit as st
import altair as alt

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
        "OE":       ms_filter("Organisationseinheit", "Organisationseinheit"),
        "KT":       ms_filter("Kontierungstyp", "Kontierungstyp"),
        "LA":       ms_filter("Leistungsart", "Leistungsart"),
        "Kat":      ms_filter("Kategorie", "Kategorie"),
        "EKSt":     ms_filter("EmpfKostenstelle", "EmpfKostenstelle")
    }
    mask = (
        df["Organisationseinheit"].isin(sel["OE"]) &
        df["Kontierungstyp"].isin(sel["KT"]) &
        df["Leistungsart"].isin(sel["LA"]) &
        df["Kategorie"].isin(sel["Kat"]) &
        df["EmpfKostenstelle"].isin(sel["EKSt"])
    )
    dff = df[mask].copy()
    if dff.empty:
        st.warning("Keine Daten für diese Filterkombination.")
        return

    # --- 3) Monatsspalten & Basis-KPIs
    monate = ["Januar","Februar","März","April","Mai","Juni",
              "Juli","August","September","Oktober","November","Dezember"]
    # Summen und Kennzahlen
    total_ytd       = dff["ytd"].sum()
    median_ytd_ma   = dff.groupby("U-Nummer")["ytd"].sum().median()
    std_ytd_ma      = dff.groupby("U-Nummer")["ytd"].sum().std()
    cv_ytd_ma       = std_ytd_ma / median_ytd_ma if median_ytd_ma else 0
    avg_per_month   = dff[monate].sum(axis=1).mean()
    peak_month      = dff[monate].sum().idxmax()
    low_month       = dff[monate].sum().idxmin()
    emp_count       = dff["U-Nummer"].nunique()
    proj_count      = dff["Projektdefinition"].nunique()

    # KPI-Übersicht
    cols = st.columns(6)
    cols[0].metric("Gesamt YTD-Stunden",       f"{total_ytd:,.0f}")
    cols[1].metric("Median YTD/Stunde pro MA", f"{median_ytd_ma:,.1f}")
    cols[2].metric("StdDev YTD/Stunde MA",     f"{std_ytd_ma:,.1f}")
    cols[3].metric("CV YTD/Stunde MA",         f"{cv_ytd_ma:.2%}")
    cols[4].metric("Ø Monatsstunden",          f"{avg_per_month:,.1f}")
    cols[5].metric("Mitarbeitende",            f"{emp_count}")

    # Zusätzliche KPIs
    cols2 = st.columns(5)
    cols2[0].metric("Projekte gesamt",         proj_count)
    cols2[1].metric("Stärkster Monat",         peak_month)
    cols2[2].metric("Schwächster Monat",       low_month)
    # Berechnung Überstunden (z.B. >160h/Monat)
    overtime = dff[monate].applymap(lambda x: max(x-160,0)).sum().sum()
    cols2[3].metric("Total Überstunden",       f"{overtime:,.0f}")
    # Anteil Leistungsart "Operation" (Beispiel)
    if "Operation" in dff["Leistungsart"].unique():
        op_share = dff.query("Leistungsart=='Operation'")["ytd"].sum() / total_ytd
        cols2[4].metric("Operation-Anteil",      f"{op_share:.1%}")
    else:
        cols2[4].metric("Operation-Anteil",      "n/a")

    # --- 4) Tabs mit Grafiken
    tabs = st.tabs([
        "Trend & Wachstum", 
        "Verteilung & Boxplot", 
        "Kostenstellen & Kategorien", 
        "Projekt-Insights"
    ])

    # 4.1) Trend & Wachstum
    with tabs[0]:
        st.subheader("Monatlicher Verlauf & YoY-Wachstum")
        # Monatsverlauf als Area Chart
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

        # Monat-zu-Monat %-Wachstum
        mon_growth = mon_sum.assign(
            pct=lambda df_: df_["Stunden"].pct_change()*100
        ).dropna()
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

    # 4.2) Verteilung & Boxplot
    with tabs[1]:
        st.subheader("Verteilung der YTD-Stunden pro Mitarbeitenden")
        # Histogramm
        emp_ytd = dff.groupby("U-Nummer")["ytd"].sum().reset_index()
        hist = (
            alt.Chart(emp_ytd)
               .mark_bar()
               .encode(
                   alt.X("ytd:Q", bin=alt.Bin(maxbins=30), title="YTD-Stunden"),
                   y="count():Q",
                   tooltip=["count()"]
               )
        )
        # Boxplot pro OE
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

    # 4.3) Kostenstellen & Kategorien
    with tabs[2]:
        st.subheader("Stunden nach EmpfKostenstelle & Kategorie")
        # Pie Chart Kategorie-Anteil
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
        # Stacked Bar Kostenstellen vs. Leistungsart
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

    # 4.4) Projekt-Insights
    with tabs[3]:
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
        # Scatter: Ø Monatsstunden vs. YTD pro MA
        scatter_df = emp_ytd.merge(
            dff[["U-Nummer"]+monate]
               .groupby("U-Nummer").mean().reset_index(),
            on="U-Nummer", suffixes=("_ytd","_avgm")
        )
        scatter = (
            alt.Chart(scatter_df)
               .mark_circle(size=50)
               .encode(
                   x=alt.X("ytd:Q", title="YTD-Stunden"),
                   y=alt.Y("März:Q", title="Ø Monatsstunden (z.B. März)"),
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
