app.py        color: #0F172A;
        margin-bottom: 0rem;
    }
    .orion-subtitle {
        color: #475569;
        font-size: 1rem;
        margin-bottom: 1rem;
    }
    div[data-testid="stMetric"] {
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        padding: 16px;
        border-radius: 18px;
        box-shadow: 0 6px 20px rgba(15, 23, 42, 0.06);
    }
    .section-card {
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        padding: 18px;
        border-radius: 18px;
        margin-bottom: 14px;
        box-shadow: 0 6px 18px rgba(15, 23, 42, 0.05);
    }
    .risk-high {color:#991B1B; font-weight:800;}
    .risk-mid {color:#92400E; font-weight:800;}
    .risk-low {color:#166534; font-weight:800;}
</style>
""", unsafe_allow_html=True)

# -------------------------
# UTILIDADES
# -------------------------
def clean_col_name(x):
    if pd.isna(x):
        return ""
    return str(x).strip()

def money_to_num(x):
    if pd.isna(x):
        return 0.0
    s = str(x).replace("$", "").replace(",", "").replace(" ", "").strip()
    if s in ["", "-", "nan", "None"]:
        return 0.0
    return pd.to_numeric(s, errors="coerce") if not pd.isna(pd.to_numeric(s, errors="coerce")) else 0.0

def num(x):
    if pd.isna(x):
        return 0.0
    if isinstance(x, str):
        x = x.replace("$", "").replace(",", "").strip()
        if x in ["", "-"]:
            return 0.0
    return float(pd.to_numeric(x, errors="coerce") or 0)

def standardize_text(s):
    if pd.isna(s):
        return "SIN DATO"
    s = str(s).strip()
    return s if s else "SIN DATO"

def safe_metric_delta(value):
    try:
        return f"{value:+.1f}%"
    except Exception:
        return None

def to_excel_bytes(sheets: dict):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, data in sheets.items():
            safe_name = str(name)[:31]
            data.to_excel(writer, sheet_name=safe_name, index=False)
    return output.getvalue()

# -------------------------
# LECTURA DE PRODUCTIVIDAD
# -------------------------
@st.cache_data(show_spinner=False)
def load_productivity(file):
    xls = pd.ExcelFile(file)
    target = None
    for s in xls.sheet_names:
        if "resultados" in s.lower() and "productividad" in s.lower():
            target = s
            break

    if target is None:
        return pd.DataFrame(), None

    df = pd.read_excel(file, sheet_name=target)
    df.columns = [clean_col_name(c) for c in df.columns]

    # Unificar columna Nombre duplicada
    if "Nombre" in df.columns and "nombre" in df.columns:
        df["Nombre"] = df["Nombre"].fillna(df["nombre"])
        df = df.drop(columns=["nombre"])
    elif "nombre" in df.columns and "Nombre" not in df.columns:
        df = df.rename(columns={"nombre": "Nombre"})

    rename_map = {
        "RECORRIDOs": "Recorridos",
        "Fecha s": "Fecha_Corta"
    }
    df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})

    # Crear columnas faltantes si no existen
    expected = ["Fecha", "Ubicación", "Actividad Realizada", "Área", "Número de Piezas",
                "Hora Inicio", "Hora Fin", "Nombre", "Motivo de ingreso", "Recorridos"]
    for c in expected:
        if c not in df.columns:
            df[c] = np.nan

    df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce")
    df["Día"] = df["Fecha"].dt.date
    df["Semana"] = df["Fecha"].dt.isocalendar().week.astype("Int64")
    df["Mes"] = df["Fecha"].dt.strftime("%B")
    df["Ubicación"] = df["Ubicación"].apply(standardize_text)
    df["Actividad Realizada"] = df["Actividad Realizada"].apply(standardize_text)
    df["Área"] = df["Área"].apply(standardize_text)
    df["Nombre"] = df["Nombre"].apply(standardize_text)
    df["Motivo de ingreso"] = df["Motivo de ingreso"].apply(standardize_text)
    df["Número de Piezas"] = pd.to_numeric(df["Número de Piezas"], errors="coerce").fillna(0)
    df["Recorridos"] = pd.to_numeric(df["Recorridos"], errors="coerce").fillna(0)

    # Duración en minutos con manejo flexible
    hi = pd.to_datetime(df["Hora Inicio"].astype(str), errors="coerce")
    hf = pd.to_datetime(df["Hora Fin"].astype(str), errors="coerce")
    df["Minutos"] = (hf - hi).dt.total_seconds() / 60
    df.loc[df["Minutos"] < 0, "Minutos"] = np.nan
    df["Piezas_por_Hora"] = np.where(df["Minutos"] > 0, df["Número de Piezas"] / (df["Minutos"] / 60), np.nan)

    return df, target

# -------------------------
# LECTURA DE HOJAS MENSUALES COMERCIALES
# -------------------------
@st.cache_data(show_spinner=False)
def load_monthly_sales(file):
    xls = pd.ExcelFile(file)
    monthly_sheets = [s for s in xls.sheet_names if re.search(r"\b26\b", s) and "operación semanal" not in s.lower()]
    all_rows = []
    all_long = []

    for sheet in monthly_sheets:
        raw = pd.read_excel(file, sheet_name=sheet, header=None)
        if raw.empty or raw.shape[0] < 2:
            continue

        headers = [clean_col_name(x) for x in raw.iloc[0].tolist()]
        df = raw.iloc[1:].copy()
        df.columns = headers
        df = df.dropna(how="all")

        # Renombrar posiciones base para evitar problemas con encabezados vacíos
        base_cols = [
            "Art Padre", "Id Art", "Marca", "Marca Price", "Modelo", "Modelo Proveedor",
            "Corrida", "Color", "Origen", "Tallas", "Estatus", "Clasificacion", "ObservCI",
            "Linea Price", "Division", "Area Responsable Org", "Familia AP RLN",
            "Gpo Seg AP RLN", "Seccion", "Categoria", "Sub Categoria", "Grupo RLN",
            "Familia RLN", "Precio Mayoreo", "Precio Menudeo", "Tiendas"
        ]
        for i, col in enumerate(base_cols):
            if i < df.shape[1]:
                df = df.rename(columns={df.columns[i]: col})

        # Total columns por posición esperada
        if df.shape[1] > 28:
            total_pzs_col = df.columns[26]
            total_dev_col = df.columns[27]
            total_money_col = df.columns[28]
            df["Total Ventas Netas Pzs"] = df[total_pzs_col].apply(num)
            df["Total Dev Pzs"] = df[total_dev_col].apply(num)
            df["Total Venta Neta $"] = df[total_money_col].apply(num)
        else:
            df["Total Ventas Netas Pzs"] = 0
            df["Total Dev Pzs"] = 0
            df["Total Venta Neta $"] = 0

        for c in ["Precio Mayoreo", "Precio Menudeo"]:
            if c in df.columns:
                df[c] = df[c].apply(money_to_num)
            else:
                df[c] = 0

        for c in ["Id Art", "Modelo", "Color", "Tiendas", "Marca", "Categoria", "Sub Categoria", "Familia RLN", "Estatus"]:
            if c not in df.columns:
                df[c] = "SIN DATO"
            df[c] = df[c].apply(standardize_text)

        df["Mes_Origen"] = sheet
        keep = ["Mes_Origen", "Id Art", "Modelo", "Color", "Marca", "Categoria", "Sub Categoria",
                "Familia RLN", "Estatus", "Precio Mayoreo", "Precio Menudeo", "Tiendas",
                "Total Ventas Netas Pzs", "Total Dev Pzs", "Total Venta Neta $"]
        all_rows.append(df[keep].copy())

        # Transformación larga por fecha: desde col 29 en bloques de 3
        for idx in range(29, df.shape[1], 3):
            if idx + 2 >= df.shape[1]:
                continue
            date_label = raw.iloc[0, idx]
            date_val = pd.to_datetime(date_label, errors="coerce", dayfirst=True)
            if pd.isna(date_val):
                continue

            temp = df[["Mes_Origen", "Id Art", "Modelo", "Color", "Marca", "Categoria", "Sub Categoria",
                       "Familia RLN", "Estatus", "Tiendas"]].copy()
            temp["Fecha_Venta"] = date_val
            temp["Ventas Netas Pzs"] = df.iloc[:, idx].apply(num)
            temp["Dev Pzs"] = df.iloc[:, idx + 1].apply(num)
            temp["Venta Neta $"] = df.iloc[:, idx + 2].apply(num)
            temp = temp[(temp["Ventas Netas Pzs"] != 0) | (temp["Dev Pzs"] != 0) | (temp["Venta Neta $"] != 0)]
            all_long.append(temp)

    sales = pd.concat(all_rows, ignore_index=True) if all_rows else pd.DataFrame()
    sales_long = pd.concat(all_long, ignore_index=True) if all_long else pd.DataFrame()
    return sales, sales_long, monthly_sheets

# -------------------------
# HEADER
# -------------------------
st.markdown('<div class="orion-title">🚀 ORION PRO</div>', unsafe_allow_html=True)
st.markdown('<div class="orion-subtitle">Plataforma Nacional de Recuperación de Mercancía | Operaciones Ropa</div>', unsafe_allow_html=True)

with st.sidebar:
    st.header("📂 Fuente de datos")
    archivo = st.file_uploader("Carga el archivo Excel de ORION", type=["xlsx"])

if archivo is None:
    st.info("Por favor cargue un archivo Excel para iniciar ORION.")
    st.stop()

with st.spinner("Construyendo ORION PRO..."):
    df_prod, prod_sheet = load_productivity(archivo)
    df_sales, df_sales_long, monthly_sheets = load_monthly_sales(archivo)

if df_prod.empty:
    st.error("No encontré la hoja 'Resultados productividad'. Revisa el archivo.")
    st.stop()

# -------------------------
# FILTROS
# -------------------------
with st.sidebar:
    st.header("🎛️ Filtros PRO")

    min_date = df_prod["Fecha"].min()
    max_date = df_prod["Fecha"].max()

    if pd.notna(min_date) and pd.notna(max_date):
        fecha_range = st.date_input(
            "Rango de fechas operación",
            value=(min_date.date(), max_date.date()),
            min_value=min_date.date(),
            max_value=max_date.date()
        )
    else:
        fecha_range = None

    tiendas = sorted(df_prod["Ubicación"].dropna().unique())
    actividades = sorted(df_prod["Actividad Realizada"].dropna().unique())
    areas = sorted(df_prod["Área"].dropna().unique())
    nombres = sorted(df_prod["Nombre"].dropna().unique())

    f_tienda = st.multiselect("Ubicación / tienda", tiendas)
    f_actividad = st.multiselect("Actividad", actividades)
    f_area = st.multiselect("Área", areas)
    f_nombre = st.multiselect("Colaborador", nombres)

    st.divider()
    meta_prod_hora = st.number_input("Meta piezas por hora", min_value=1, value=60, step=5)
    meta_recorridos = st.number_input("Meta recorridos", min_value=1, value=3, step=1)
    umbral_bajo = st.slider("Umbral bajo desempeño vs promedio", 0.1, 1.0, 0.7, 0.05)

# Aplicar filtros operación
dff = df_prod.copy()

if fecha_range and isinstance(fecha_range, tuple) and len(fecha_range) == 2:
    start, end = fecha_range
    dff = dff[(dff["Fecha"].dt.date >= start) & (dff["Fecha"].dt.date <= end)]

if f_tienda:
    dff = dff[dff["Ubicación"].isin(f_tienda)]
if f_actividad:
    dff = dff[dff["Actividad Realizada"].isin(f_actividad)]
if f_area:
    dff = dff[dff["Área"].isin(f_area)]
if f_nombre:
    dff = dff[dff["Nombre"].isin(f_nombre)]

# Filtro comercial por tienda si aplica
sales_f = df_sales.copy()
sales_long_f = df_sales_long.copy()
if f_tienda and not sales_f.empty:
    sales_f = sales_f[sales_f["Tiendas"].isin(f_tienda)]
if f_tienda and not sales_long_f.empty:
    sales_long_f = sales_long_f[sales_long_f["Tiendas"].isin(f_tienda)]

# -------------------------
# KPIs
# -------------------------
piezas = dff["Número de Piezas"].sum()
registros = len(dff)
colabs = dff["Nombre"].nunique()
recorridos = dff["Recorridos"].sum()
minutos = dff["Minutos"].sum(skipna=True)
prod_hora = piezas / (minutos / 60) if minutos and minutos > 0 else 0

venta_total = sales_f["Total Venta Neta $"].sum() if not sales_f.empty else 0
venta_pzs = sales_f["Total Ventas Netas Pzs"].sum() if not sales_f.empty else 0
dev_pzs = sales_f["Total Dev Pzs"].sum() if not sales_f.empty else 0
ticket_promedio = venta_total / venta_pzs if venta_pzs else 0

cumpl_prod = min(prod_hora / meta_prod_hora, 1) if meta_prod_hora else 0
cumpl_rec = min(recorridos / meta_recorridos, 1) if meta_recorridos else 0
score = round(((cumpl_prod * 0.45) + (cumpl_rec * 0.25) + (min(registros / max(colabs,1) / 10, 1) * 0.15) + (min(venta_pzs / max(piezas,1), 1) * 0.15)) * 100, 1)

if score >= 85:
    riesgo = "BAJO"
    riesgo_class = "risk-low"
elif score >= 65:
    riesgo = "MEDIO"
    riesgo_class = "risk-mid"
else:
    riesgo = "ALTO"
    riesgo_class = "risk-high"

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("📦 Piezas operadas", f"{piezas:,.0f}")
k2.metric("⚡ Piezas / hora", f"{prod_hora:,.1f}", safe_metric_delta((prod_hora/meta_prod_hora-1)*100 if meta_prod_hora else 0))
k3.metric("👥 Colaboradores", f"{colabs:,.0f}")
k4.metric("🚶 Recorridos", f"{recorridos:,.0f}", safe_metric_delta((recorridos/meta_recorridos-1)*100 if meta_recorridos else 0))
k5.metric("💰 Venta neta ligada", f"${venta_total:,.0f}")

st.markdown(f"""
<div class="section-card">
    <b>Índice Integral ORION:</b> {score}/100 &nbsp; | &nbsp;
    <b>Riesgo operativo:</b> <span class="{riesgo_class}">{riesgo}</span> &nbsp; | &nbsp;
    <b>Hoja operación:</b> {prod_sheet} &nbsp; | &nbsp;
    <b>Hojas comerciales:</b> {", ".join(monthly_sheets) if monthly_sheets else "No detectadas"}
</div>
""", unsafe_allow_html=True)

# -------------------------
# TABS
# -------------------------
tabs = st.tabs([
    "📊 Panel Ejecutivo",
    "🔁 Recuperación Operativa",
    "💰 Recuperación Comercial",
    "👤 Colaboradores",
    "🏆 Rankings",
    "🚨 Alertas",
    "🧾 Datos / Exportación"
])

# -------------------------
# PANEL EJECUTIVO
# -------------------------
with tabs[0]:
    c1, c2 = st.columns([1.2, 1])

    with c1:
        st.subheader("Tendencia de piezas operadas")
        if not dff.empty:
            trend = dff.groupby("Día", as_index=False)["Número de Piezas"].sum()
            fig = px.line(trend, x="Día", y="Número de Piezas", markers=True, title="Piezas por día")
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.warning("Sin datos para graficar.")

    with c2:
        st.subheader("Distribución por actividad")
        act = dff.groupby("Actividad Realizada", as_index=False)["Número de Piezas"].sum().sort_values("Número de Piezas", ascending=False)
        if not act.empty:
            fig = px.pie(act, names="Actividad Realizada", values="Número de Piezas", hole=0.45)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.warning("Sin datos.")

    st.subheader("Resumen por tienda")
    tienda = dff.groupby("Ubicación", as_index=False).agg(
        Piezas=("Número de Piezas", "sum"),
        Registros=("Número de Piezas", "count"),
        Recorridos=("Recorridos", "sum"),
        Colaboradores=("Nombre", "nunique"),
        Minutos=("Minutos", "sum")
    )
    tienda["Piezas/Hora"] = np.where(tienda["Minutos"] > 0, tienda["Piezas"] / (tienda["Minutos"]/60), 0)
    tienda["Score"] = np.minimum(tienda["Piezas/Hora"] / meta_prod_hora, 1) * 70 + np.minimum(tienda["Recorridos"] / meta_recorridos, 1) * 30
    tienda["Estatus"] = np.where(tienda["Score"] >= 85, "🟢 Óptimo", np.where(tienda["Score"] >= 65, "🟡 Atención", "🔴 Crítico"))
    st.dataframe(tienda.sort_values("Score", ascending=False), use_container_width=True)

# -------------------------
# RECUPERACIÓN OPERATIVA
# -------------------------
with tabs[1]:
    c1, c2 = st.columns(2)

    with c1:
        st.subheader("Piezas por área")
        area = dff.groupby("Área", as_index=False)["Número de Piezas"].sum().sort_values("Número de Piezas")
        fig = px.bar(area, x="Número de Piezas", y="Área", orientation="h")
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        st.subheader("Piezas por motivo de ingreso")
        mot = dff.groupby("Motivo de ingreso", as_index=False)["Número de Piezas"].sum().sort_values("Número de Piezas", ascending=False)
        fig = px.bar(mot, x="Motivo de ingreso", y="Número de Piezas")
        st.plotly_chart(fig, use_container_width=True)

    st.subheader("Mapa operativo tienda / actividad")
    pivot = dff.pivot_table(
        index="Ubicación",
        columns="Actividad Realizada",
        values="Número de Piezas",
        aggfunc="sum",
        fill_value=0
    )
    st.dataframe(pivot, use_container_width=True)

# -------------------------
# RECUPERACIÓN COMERCIAL
# -------------------------
with tabs[2]:
    if df_sales.empty:
        st.warning("No se detectaron hojas comerciales mensuales.")
    else:
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Venta neta $", f"${venta_total:,.0f}")
        c2.metric("Venta neta pzs", f"{venta_pzs:,.0f}")
        c3.metric("Devoluciones pzs", f"{dev_pzs:,.0f}")
        c4.metric("Venta promedio / pieza", f"${ticket_promedio:,.2f}")

        c1, c2 = st.columns(2)
        with c1:
            st.subheader("Venta por tienda")
            vta_tienda = sales_f.groupby("Tiendas", as_index=False)["Total Venta Neta $"].sum().sort_values("Total Venta Neta $", ascending=False).head(20)
            fig = px.bar(vta_tienda, x="Tiendas", y="Total Venta Neta $")
            st.plotly_chart(fig, use_container_width=True)

        with c2:
            st.subheader("Venta por categoría")
            vta_cat = sales_f.groupby("Categoria", as_index=False)["Total Venta Neta $"].sum().sort_values("Total Venta Neta $", ascending=False).head(15)
            fig = px.bar(vta_cat, x="Total Venta Neta $", y="Categoria", orientation="h")
            st.plotly_chart(fig, use_container_width=True)

        if not sales_long_f.empty:
            st.subheader("Tendencia comercial diaria")
            daily = sales_long_f.groupby("Fecha_Venta", as_index=False).agg(
                Venta_Neta=("Venta Neta $", "sum"),
                Piezas=("Ventas Netas Pzs", "sum"),
                Dev=("Dev Pzs", "sum")
            )
            fig = px.line(daily, x="Fecha_Venta", y=["Venta_Neta", "Piezas", "Dev"], markers=True)
            st.plotly_chart(fig, use_container_width=True)

        st.subheader("Top 30 modelos por venta neta")
        top_modelos = sales_f.groupby(["Id Art", "Modelo", "Color", "Tiendas"], as_index=False).agg(
            Venta_Neta=("Total Venta Neta $", "sum"),
            Piezas=("Total Ventas Netas Pzs", "sum"),
            Dev=("Total Dev Pzs", "sum")
        ).sort_values("Venta_Neta", ascending=False).head(30)
        st.dataframe(top_modelos, use_container_width=True)

# -------------------------
# COLABORADORES
# -------------------------
with tabs[3]:
    st.subheader("Productividad por colaborador")
    colab = dff.groupby("Nombre", as_index=False).agg(
        Piezas=("Número de Piezas", "sum"),
        Registros=("Número de Piezas", "count"),
        Recorridos=("Recorridos", "sum"),
        Minutos=("Minutos", "sum"),
        Actividades=("Actividad Realizada", "nunique")
    )
    colab["Piezas/Hora"] = np.where(colab["Minutos"] > 0, colab["Piezas"] / (colab["Minutos"]/60), 0)
    promedio_colab = colab["Piezas"].mean() if not colab.empty else 0
    colab["Estatus"] = np.where(
        colab["Piezas"] >= promedio_colab, "🟢 Sobre promedio",
        np.where(colab["Piezas"] >= promedio_colab * umbral_bajo, "🟡 En observación", "🔴 Bajo")
    )
    colab = colab.sort_values("Piezas", ascending=False)

    fig = px.bar(colab.head(25), x="Nombre", y="Piezas", color="Estatus", title="Top colaboradores por piezas")
    st.plotly_chart(fig, use_container_width=True)
    st.dataframe(colab, use_container_width=True)

# -------------------------
# RANKINGS
# -------------------------
with tabs[4]:
    st.subheader("Ranking nacional ORION")
    ranking = dff.groupby("Ubicación", as_index=False).agg(
        Piezas=("Número de Piezas", "sum"),
        Recorridos=("Recorridos", "sum"),
        Colaboradores=("Nombre", "nunique"),
        Minutos=("Minutos", "sum")
    )
    ranking["Piezas/Hora"] = np.where(ranking["Minutos"] > 0, ranking["Piezas"] / (ranking["Minutos"]/60), 0)

    if not sales_f.empty:
        venta_t = sales_f.groupby("Tiendas", as_index=False).agg(Venta_Neta=("Total Venta Neta $", "sum"), Venta_Pzs=("Total Ventas Netas Pzs", "sum"))
        ranking = ranking.merge(venta_t, left_on="Ubicación", right_on="Tiendas", how="left").drop(columns=["Tiendas"], errors="ignore")
    else:
        ranking["Venta_Neta"] = 0
        ranking["Venta_Pzs"] = 0

    for c in ["Venta_Neta", "Venta_Pzs"]:
        ranking[c] = ranking[c].fillna(0)

    def scale(s):
        mx = s.max()
        return (s / mx * 100) if mx and mx > 0 else 0

    ranking["Score Productividad"] = scale(ranking["Piezas/Hora"])
    ranking["Score Recorridos"] = scale(ranking["Recorridos"])
    ranking["Score Comercial"] = scale(ranking["Venta_Neta"])
    ranking["Índice ORION"] = (ranking["Score Productividad"] * 0.45 + ranking["Score Recorridos"] * 0.25 + ranking["Score Comercial"] * 0.30).round(1)
    ranking["Nivel"] = np.where(ranking["Índice ORION"] >= 85, "🟢 Líder", np.where(ranking["Índice ORION"] >= 65, "🟡 Competitivo", "🔴 Riesgo"))
    ranking = ranking.sort_values("Índice ORION", ascending=False)

    st.dataframe(ranking, use_container_width=True)

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Top 10 tiendas")
        st.dataframe(ranking.head(10), use_container_width=True)
    with c2:
        st.subheader("Bottom 10 tiendas")
        st.dataframe(ranking.tail(10), use_container_width=True)

# -------------------------
# ALERTAS
# -------------------------
with tabs[5]:
    st.subheader("Centro de alertas inteligentes")

    alertas = []

    if prod_hora < meta_prod_hora:
        alertas.append({
            "Tipo": "Productividad",
            "Prioridad": "Alta" if prod_hora < meta_prod_hora * 0.7 else "Media",
            "Alerta": f"La productividad actual es {prod_hora:,.1f} pzs/h vs meta {meta_prod_hora:,.1f}.",
            "Acción sugerida": "Validar tiempos de inicio/fin, reforzar cuadrillas y priorizar actividades de mayor volumen."
        })

    if recorridos < meta_recorridos:
        alertas.append({
            "Tipo": "Recorridos",
            "Prioridad": "Alta",
            "Alerta": f"Los recorridos registrados son {recorridos:,.0f} vs meta {meta_recorridos:,.0f}.",
            "Acción sugerida": "Asegurar captura diaria de recorridos y asignar responsable por franja horaria."
        })

    if colabs == 0:
        alertas.append({
            "Tipo": "Plantilla",
            "Prioridad": "Alta",
            "Alerta": "No hay colaboradores detectados en el periodo filtrado.",
            "Acción sugerida": "Revisar carga de archivo y registros de nombres."
        })

    if not dff.empty:
        by_store = dff.groupby("Ubicación")["Número de Piezas"].sum()
        avg_store = by_store.mean()
        low_stores = by_store[by_store < avg_store * umbral_bajo]
        for tienda_name, val in low_stores.items():
            alertas.append({
                "Tipo": "Tienda",
                "Prioridad": "Media",
                "Alerta": f"{tienda_name} está por debajo del umbral operativo con {val:,.0f} piezas.",
                "Acción sugerida": "Revisar ejecución en piso, probadores, cambios y cajas; confirmar registro completo."
            })

    alert_df = pd.DataFrame(alertas)
    if alert_df.empty:
        st.success("Sin alertas críticas con los filtros actuales.")
    else:
        st.dataframe(alert_df, use_container_width=True)

# -------------------------
# DATOS / EXPORTACIÓN
# -------------------------
with tabs[6]:
    st.subheader("Datos base")
    st.caption("Puedes descargar la información ya transformada para análisis adicional.")

    st.write("Productividad filtrada")
    st.dataframe(dff, use_container_width=True)

    if not sales_f.empty:
        st.write("Comercial consolidado")
        st.dataframe(sales_f, use_container_width=True)

    export_sheets = {"Productividad_filtrada": dff}
    if not sales_f.empty:
        export_sheets["Comercial_consolidado"] = sales_f
    if not sales_long_f.empty:
        export_sheets["Comercial_diario"] = sales_long_f

    st.download_button(
        "⬇️ Descargar consolidado ORION en Excel",
        data=to_excel_bytes(export_sheets),
        file_name="ORION_consolidado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.download_button(
        "⬇️ Descargar productividad CSV",
        data=dff.to_csv(index=False).encode("utf-8-sig"),
        file_name="ORION_productividad.csv",
        mime="text/csv"
    )

st.caption("ORION PRO | Operaciones Ropa | Dashboard generado en Streamlit")
