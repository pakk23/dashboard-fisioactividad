import streamlit as st
import pandas as pd
import plotly.express as px
from procesador_timp import procesar, exportar_excel

st.set_page_config(
    page_title="Dashboard Reservas · Fisioactividad",
    page_icon="🏃",
    layout="wide",
)

st.title("📊 Dashboard de Reservas — Fisioactividad")

# ── Carga de archivo ──────────────────────────────────────────────────────────
uploaded = st.file_uploader(
    "Carga el Excel de reservas exportado desde TIMP (.xls / .xlsx)",
    type=["xls", "xlsx"],
)

if not uploaded:
    st.info("Arrastra aquí el archivo de exportación de TIMP para comenzar.")
    st.stop()

# ── Procesamiento ──────────────────────────────────────────────────────────────
@st.cache_data
def cargar_y_procesar(file_bytes, file_name):
    from io import BytesIO
    return procesar(BytesIO(file_bytes))

with st.spinner("Procesando datos TIMP..."):
    file_bytes = uploaded.read()
    resultado  = cargar_y_procesar(file_bytes, uploaded.name)

df            = resultado["df_main"]
df_eco        = resultado["df_eco"]
df_excl       = resultado["df_excl"]
df_cancel     = resultado["df_cancel"]
df_pendientes = resultado["df_pendientes"]
df_sesiones   = resultado["df_sesiones"]
stats         = resultado["stats"]

# ── Resumen del procesamiento ─────────────────────────────────────────────────
with st.expander("📋 Resumen del procesamiento", expanded=True):
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("Registros originales",    f"{stats['total_original']:,}")
    c2.metric("Fisio Invasiva",          f"{stats['eco_separadas']:,}")
    c3.metric("Excluidas",               f"{stats['excluidas']:,}")
    c4.metric("Canceladas / En cola",    f"{stats['canceladas']:,}")
    c5.metric("Filas editadas",          f"{stats['editadas']:,}")
    c6.metric("Ventas pendientes",       f"{stats['pendientes']:,}")

st.divider()

# ── Sidebar: filtros ──────────────────────────────────────────────────────────
st.sidebar.header("🔍 Filtros")

# Fecha desde columna "Fecha" (texto DD/MM/YYYY) o "Inicio" si existe
if "Fecha" in df.columns:
    fechas_dt = pd.to_datetime(df["Fecha"], format="%d/%m/%Y", errors="coerce")
    fecha_min = fechas_dt.min().date() if pd.notna(fechas_dt.min()) else None
    fecha_max = fechas_dt.max().date() if pd.notna(fechas_dt.max()) else None
else:
    fechas_dt = df["Inicio"] if "Inicio" in df.columns else pd.Series(dtype="datetime64[ns]")
    fecha_min = fechas_dt.min().date() if not fechas_dt.empty else None
    fecha_max = fechas_dt.max().date() if not fechas_dt.empty else None

df["_fecha_dt"] = fechas_dt

if fecha_min and fecha_max:
    rango = st.sidebar.date_input("Rango de fechas", value=(fecha_min, fecha_max),
                                   min_value=fecha_min, max_value=fecha_max)
    f_ini, f_fin = (rango[0], rango[1]) if len(rango) == 2 else (fecha_min, fecha_max)
else:
    f_ini, f_fin = fecha_min, fecha_max

estados      = sorted(df["Estado de reserva"].dropna().unique().tolist())
estados_sel  = st.sidebar.multiselect("Estado de reserva", estados, default=estados)

servicios    = sorted(df["Servicio"].dropna().unique().tolist())
servicios_sel = st.sidebar.multiselect("Servicio", servicios, default=servicios)

profesionales = sorted(df["Profesional"].dropna().unique().tolist()) if "Profesional" in df.columns else []
prof_sel      = st.sidebar.multiselect("Profesional", profesionales, default=profesionales)

origenes     = sorted(df["Origen"].dropna().unique().tolist()) if "Origen" in df.columns else []
orig_sel     = st.sidebar.multiselect("Origen", origenes, default=origenes)

# ── Aplicar filtros ───────────────────────────────────────────────────────────
mask = (
    (df["_fecha_dt"].dt.date >= f_ini) &
    (df["_fecha_dt"].dt.date <= f_fin) &
    (df["Estado de reserva"].isin(estados_sel)) &
    (df["Servicio"].isin(servicios_sel))
)
if prof_sel:
    mask &= df["Profesional"].isin(prof_sel)
if orig_sel:
    mask &= df["Origen"].isin(orig_sel)

dff = df[mask].copy()

# ── KPIs ──────────────────────────────────────────────────────────────────────
total       = len(dff)
aceptadas   = (dff["Estado de reserva"] == "Aceptada").sum()
canceladas  = (dff["Estado de reserva"] == "Cancelada").sum()
tasa_cancel = (canceladas / total * 100) if total > 0 else 0
ingresos    = dff["Precio reserva"].sum() if "Precio reserva" in dff.columns else 0
ingresos_eco = df_eco["Ingreso estimado"].sum() if not df_eco.empty and "Ingreso estimado" in df_eco.columns else 0

k1, k2, k3, k4, k5, k6 = st.columns(6)
k1.metric("Total reservas",      f"{total:,}")
k2.metric("Aceptadas",           f"{aceptadas:,}")
k3.metric("Canceladas",          f"{canceladas:,}")
k4.metric("Tasa cancelación",    f"{tasa_cancel:.1f}%")
k5.metric("Facturación (€)",     f"{ingresos:,.2f}")
k6.metric("Fisio Invasiva (€)",  f"{ingresos_eco:,.2f}")

st.divider()

# ── Pestañas ──────────────────────────────────────────────────────────────────
tab_dash, tab_sesiones, tab_pivot, tab_eco, tab_cancel, tab_excl, tab_pendientes, tab_detalle = st.tabs([
    "📈 Dashboard",
    "📅 Sesiones",
    "📊 Tabla Dinámica",
    "🔬 Fisio Invasiva",
    "❌ Canceladas",
    "🚫 Excluidas",
    "⚠️ Ventas Pendientes",
    "📋 Detalle reservas",
])

# ── TAB: Dashboard ────────────────────────────────────────────────────────────
with tab_dash:
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Reservas por día")
        daily = dff.groupby(dff["_fecha_dt"].dt.date).size().reset_index(name="Reservas")
        daily.columns = ["Fecha", "Reservas"]
        fig = px.bar(daily, x="Fecha", y="Reservas")
        fig.update_layout(margin=dict(t=10, b=0))
        st.plotly_chart(fig, use_container_width=True)

    with col2:
        st.subheader("Estado de reservas")
        estado_cnt = dff["Estado de reserva"].value_counts().reset_index()
        estado_cnt.columns = ["Estado", "Cantidad"]
        fig2 = px.pie(estado_cnt, names="Estado", values="Cantidad", hole=0.4)
        fig2.update_layout(margin=dict(t=10, b=0))
        st.plotly_chart(fig2, use_container_width=True)

    col3, col4 = st.columns(2)
    with col3:
        st.subheader("Top 10 servicios")
        top_srv = dff["Servicio"].value_counts().head(10).reset_index()
        top_srv.columns = ["Servicio", "Reservas"]
        fig3 = px.bar(top_srv, x="Reservas", y="Servicio", orientation="h")
        fig3.update_layout(yaxis={"categoryorder": "total ascending"}, margin=dict(t=10, b=0))
        st.plotly_chart(fig3, use_container_width=True)

    with col4:
        if "Profesional" in dff.columns:
            st.subheader("Facturación por profesional")
            prof_fac = dff.groupby("Profesional")["Precio reserva"].sum().reset_index()
            prof_fac = prof_fac.sort_values("Precio reserva", ascending=False).head(10)
            fig4 = px.bar(prof_fac, x="Precio reserva", y="Profesional", orientation="h",
                          labels={"Precio reserva": "€"})
            fig4.update_layout(yaxis={"categoryorder": "total ascending"}, margin=dict(t=10, b=0))
            st.plotly_chart(fig4, use_container_width=True)

    st.subheader("Evolución mensual")
    dff2 = dff.copy()
    dff2["Mes"] = dff2["_fecha_dt"].dt.to_period("M").astype(str)
    monthly = dff2.groupby(["Mes", "Estado de reserva"]).size().reset_index(name="Reservas")
    fig5 = px.bar(monthly, x="Mes", y="Reservas", color="Estado de reserva", barmode="stack")
    fig5.update_layout(margin=dict(t=10, b=0), xaxis_tickangle=-45)
    st.plotly_chart(fig5, use_container_width=True)

# ── TAB: Sesiones ─────────────────────────────────────────────────────────────
with tab_sesiones:
    st.subheader("📅 Sesiones únicas (Actividad + Fecha + Hora + Profesional)")
    if df_sesiones.empty:
        st.info("No hay datos de sesiones disponibles.")
    else:
        s1, s2, s3, s4 = st.columns(4)
        s1.metric("Total sesiones",     f"{len(df_sesiones):,}")
        s2.metric("Ocupación media",    f"{df_sesiones['Ocupación'].mean()*100:.1f}%")
        s3.metric("€/sesión media",     f"{df_sesiones['€/ses.'].mean():,.2f} €")
        s4.metric("Facturación total",  f"{df_sesiones['€/ses.'].sum():,.2f} €")

        # Filtro por actividad
        acts = ["Todas"] + sorted(df_sesiones["Actividad"].dropna().unique().tolist())
        act_sel = st.selectbox("Filtrar por actividad", acts, key="ses_act")
        df_ses_f = df_sesiones if act_sel == "Todas" else df_sesiones[df_sesiones["Actividad"] == act_sel]

        # Formato visual de ocupación y euros
        df_ses_show = df_ses_f.copy()
        df_ses_show["Ocupación"] = df_ses_show["Ocupación"].apply(lambda x: f"{x*100:.0f}%")
        df_ses_show["€/pers."]   = df_ses_show["€/pers."].apply(lambda x: f"{x:,.2f} €")
        df_ses_show["€/ses."]    = df_ses_show["€/ses."].apply(lambda x: f"{x:,.2f} €")
        st.dataframe(df_ses_show, use_container_width=True, height=450, hide_index=True)

        col_s1, col_s2 = st.columns(2)
        with col_s1:
            ocu_srv = df_sesiones.groupby("Actividad")["Ocupación"].mean().reset_index()
            ocu_srv["Ocupación %"] = (ocu_srv["Ocupación"] * 100).round(1)
            fig_ocu = px.bar(ocu_srv.sort_values("Ocupación %", ascending=False),
                             x="Actividad", y="Ocupación %", title="Ocupación media por actividad (%)")
            st.plotly_chart(fig_ocu, use_container_width=True)
        with col_s2:
            fac_srv = df_sesiones.groupby("Actividad")["€/ses."].sum().reset_index()
            fig_fac = px.bar(fac_srv.sort_values("€/ses.", ascending=False),
                             x="Actividad", y="€/ses.", title="Facturación total por actividad (€)")
            st.plotly_chart(fig_fac, use_container_width=True)

# ── TAB: Tabla Dinámica ───────────────────────────────────────────────────────
with tab_pivot:
    st.subheader("Tabla Dinámica — Profesional × Servicio")
    servs_pivot    = ["Todos"] + sorted(df["Servicio"].dropna().unique().tolist())
    serv_pivot_sel = st.selectbox("Filtrar por servicio", servs_pivot, key="pivot_srv")
    df_pivot = df if serv_pivot_sel == "Todos" else df[df["Servicio"] == serv_pivot_sel]

    pivot = (
        df_pivot.groupby("Profesional")
        .agg(Reservas=("ID", "count"), Facturación=("Precio reserva", "sum"))
        .reset_index()
        .sort_values("Facturación", ascending=False)
    )
    pivot["Facturación"] = pivot["Facturación"].round(2)
    pivot["Facturación (€)"] = pivot["Facturación"].apply(lambda x: f"{x:,.2f} €")
    st.dataframe(pivot[["Profesional", "Reservas", "Facturación (€)"]],
                 use_container_width=True, hide_index=True)

    col_p1, col_p2 = st.columns(2)
    with col_p1:
        fig_p1 = px.bar(pivot, x="Profesional", y="Reservas", title="Reservas por profesional")
        st.plotly_chart(fig_p1, use_container_width=True)
    with col_p2:
        fig_p2 = px.bar(pivot, x="Profesional", y="Facturación", title="Facturación por profesional (€)")
        st.plotly_chart(fig_p2, use_container_width=True)

# ── TAB: Fisio Invasiva ───────────────────────────────────────────────────────
with tab_eco:
    st.subheader("Fisioterapia Manual, Invasiva y Ecográfica 🔬")
    if df_eco.empty:
        st.info("No hay reservas de este servicio en el archivo.")
    else:
        e1, e2 = st.columns(2)
        e1.metric("Total reservas", len(df_eco))
        e2.metric("Ingreso estimado total", f"{df_eco['Ingreso estimado'].sum():,.2f} €")
        eco_cols = [c for c in df_eco.columns if c != "_editada"]
        st.dataframe(df_eco[eco_cols], use_container_width=True, height=400)

# ── TAB: Canceladas ───────────────────────────────────────────────────────────
with tab_cancel:
    st.subheader("❌ Reservas Canceladas y En cola")
    if df_cancel.empty:
        st.info("No hay reservas canceladas en el archivo.")
    else:
        ca1, ca2 = st.columns(2)
        ca1.metric("Total canceladas / en cola", f"{len(df_cancel):,}")
        # Top canceladores
        if "Cliente" in df_cancel.columns:
            top_cancel = df_cancel["Cliente"].value_counts().head(10).reset_index()
            top_cancel.columns = ["Cliente", "Cancelaciones"]
            ca2.metric("Clientes únicos", df_cancel["Cliente"].nunique())
            fig_ca = px.bar(top_cancel, x="Cancelaciones", y="Cliente", orientation="h",
                            title="Top 10 clientes con más cancelaciones")
            fig_ca.update_layout(yaxis={"categoryorder": "total ascending"})
            st.plotly_chart(fig_ca, use_container_width=True)
        cancel_cols = [c for c in df_cancel.columns if c != "_editada"]
        st.dataframe(df_cancel[cancel_cols], use_container_width=True, height=400)

# ── TAB: Excluidas ────────────────────────────────────────────────────────────
with tab_excl:
    st.subheader("🚫 Reservas Excluidas (Cita exterior, Gestión, etc.)")
    if df_excl.empty:
        st.info("No hay reservas excluidas en el archivo.")
    else:
        st.metric("Total excluidas", f"{len(df_excl):,}")
        excl_cols = [c for c in df_excl.columns if c != "_editada"]
        st.dataframe(df_excl[excl_cols], use_container_width=True, height=400)

# ── TAB: Ventas Pendientes ────────────────────────────────────────────────────
with tab_pendientes:
    st.subheader("⚠️ Ventas Pendientes (reservas no canjeadas)")
    if df_pendientes.empty:
        st.success("No hay ventas pendientes.")
    else:
        st.warning(f"{len(df_pendientes)} cliente(s) con reservas sin canjear.")
        st.dataframe(df_pendientes, use_container_width=True, height=400)

# ── TAB: Detalle ──────────────────────────────────────────────────────────────
with tab_detalle:
    st.subheader(f"Detalle de reservas procesadas ({total:,} registros)")
    cols_show = [c for c in [
        "ID", "Servicio", "Recurso", "Fecha", "Hora",
        "Cliente", "Canjeada", "Estado de reserva",
        "Venta", "Precio", "Sesiones del bono", "Precio reserva",
        "Profesional", "Origen",
    ] if c in dff.columns]
    st.dataframe(dff[cols_show].sort_values("Fecha", ascending=False),
                 use_container_width=True, height=500)

# ── Descarga ──────────────────────────────────────────────────────────────────
st.divider()
st.subheader("⬇️ Descargar archivo procesado")
col_dl1, col_dl2 = st.columns(2)

with col_dl1:
    if st.button("Generar Excel completo procesado"):
        with st.spinner("Generando Excel..."):
            excel_bytes = exportar_excel(resultado)
        st.download_button(
            "📥 Descargar Excel (.xlsx)",
            data=excel_bytes,
            file_name="reservas_procesadas_fisioactividad.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

with col_dl2:
    csv = dff[cols_show].to_csv(index=False).encode("utf-8")
    st.download_button(
        "📥 Descargar CSV filtrado",
        data=csv,
        file_name="reservas_filtradas.csv",
        mime="text/csv",
    )
