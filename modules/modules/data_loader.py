import pandas as pd
import unicodedata

MESES = [
    "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"
]

def limpiar_texto(texto):
    texto = str(texto).strip().lower()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return texto

def detectar_hoja_operativa(sheet_names):
    for hoja in sheet_names:
        h = limpiar_texto(hoja)
        if "resultado" in h and "productividad" in h:
            return hoja
    return None

def detectar_hojas_mensuales(sheet_names):
    hojas = []
    for hoja in sheet_names:
        h = limpiar_texto(hoja)
        if "26" in h and any(mes in h for mes in MESES):
            hojas.append(hoja)
    return hojas

def leer_hoja_robusta(excel_file, sheet_name):
    for header in [0, 1, 2, 3]:
        try:
            df = pd.read_excel(excel_file, sheet_name=sheet_name, header=header)
            df.columns = [str(c).strip() for c in df.columns]

            columnas_validas = [
                "Tienda", "Fecha", "Nombre", "Ocurrencia",
                "Dev_pzs", "Vta_Pzs", "Vta_Imp", "Costo_Dev",
                "Habilitado", "Ubicado", "Muertos", "Cajas"
            ]

            if any(col in df.columns for col in columnas_validas):
                return df

        except Exception:
            continue

    return pd.read_excel(excel_file, sheet_name=sheet_name)

def load_excel_file(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    sheet_names = xls.sheet_names

    operativa_name = detectar_hoja_operativa(sheet_names)
    mensuales_names = detectar_hojas_mensuales(sheet_names)

    if operativa_name:
        df_operativa = leer_hoja_robusta(uploaded_file, operativa_name)
    else:
        df_operativa = pd.DataFrame()

    dfs_mensuales = []

    for hoja in mensuales_names:
        df = leer_hoja_robusta(uploaded_file, hoja)
        df["Mes_Origen"] = hoja
        dfs_mensuales.append(df)

    if dfs_mensuales:
        df_recuperacion = pd.concat(dfs_mensuales, ignore_index=True)
    else:
        df_recuperacion = pd.DataFrame()

    return {
        "sheet_names": sheet_names,
        "operativa_name": operativa_name,
        "mensuales_names": mensuales_names,
        "df_operativa": df_operativa,
        "df_recuperacion": df_recuperacion
    }
