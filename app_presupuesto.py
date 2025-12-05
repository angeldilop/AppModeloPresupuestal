import streamlit as st
import pandas as pd
from io import BytesIO

# --------------------- CONFIGURACIÓN BÁSICA --------------------- #

st.set_page_config(
    page_title="Modelo de Asignación Presupuestal",
    layout="wide"
)

MONTH_NAMES = {
    1: "enero",
    2: "febrero",
    3: "marzo",
    4: "abril",
    5: "mayo",
    6: "junio",
    7: "julio",
    8: "agosto",
    9: "septiembre",
    10: "octubre",
    11: "noviembre",
    12: "diciembre",
}

# --------------------- ENCABEZADO --------------------- #

st.markdown(
    """
    <div style='text-align:center; line-height:1.2'>
        <h2 style='margin-bottom:0'>MODELO DE ASIGNACIÓN PRESUPUESTAL</h2>
        <p style='margin:4px 0'><b>Dirección de Distribución</b></p>
        <p style='margin:0'><b>SISTEMA DE INTELIGENCIA COMERCIAL</b></p>
    </div>
    <hr>
    """,
    unsafe_allow_html=True
)

st.sidebar.header("Carga de archivo")

uploaded_file = st.sidebar.file_uploader(
    "Sube el archivo de presupuesto (Excel .xlsx)",
    type=["xlsx"]
)

# --------------------- FUNCIÓN PARA CARGAR DATOS --------------------- #

@st.cache_data
def load_data(file) -> pd.DataFrame:
    df = pd.read_excel(file)
    # Normalizamos algunos nombres de columnas por si cambian en el futuro
    df.columns = [c.strip() for c in df.columns]

    # Nos aseguramos de que Mes sea numérico
    if "Mes" in df.columns:
        df["Mes"] = pd.to_numeric(df["Mes"], errors="coerce").astype("Int64")

    return df


def get_month_name(mes_num):
    try:
        mes_num = int(mes_num)
        return MONTH_NAMES.get(mes_num, str(mes_num))
    except Exception:
        return str(mes_num)


# --------------------- FUNCIÓN PARA PINTAR UN COMPROBANTE --------------------- #

def mostrar_comprobante(df, director, mes):
    df_filtrado = df[(df["Director"] == director) & (df["Mes"] == mes)]

    if df_filtrado.empty:
        st.warning("No hay registros para este director y mes.")
        return

    mes_nombre = get_month_name(mes)

    st.markdown(
        f"""
        <p style='text-align:center; font-size:16px; margin-top:8px'>
        <b>Asignación presupuestal para {director} del mes de {mes_nombre}</b>
        </p>
        """,
        unsafe_allow_html=True
    )

    # Columnas que vamos a mostrar en la tabla principal
    columnas_tabla = [
        "Linea Negocio",
        "Ramo",
        "Zona",
        "Canal",
        "Sub-Canal",
        "Oficina",
        "Líder comercial Esp.",
        "Presupuesto"
    ]

    columnas_presentes = [c for c in columnas_tabla if c in df_filtrado.columns]

    df_tabla = df_filtrado[columnas_presentes].copy()

    # Renombramos para que se vea como lo pediste
    renombres = {
        "Linea Negocio": "Línea de negocio",
        "Líder comercial Esp.": "Líder Equipo",
        "Presupuesto": "Valor"
    }
    df_tabla.rename(columns=renombres, inplace=True)

    # Formato de número
    if "Valor" in df_tabla.columns:
        df_tabla["Valor"] = pd.to_numeric(df_tabla["Valor"], errors="coerce")

    st.markdown("### Detalle de asignación")
    st.dataframe(
        df_tabla.style.format({"Valor": "{:,.0f}"}),
        use_container_width=True
    )

    # Totales por línea de negocio
    if "Línea de negocio" in df_tabla.columns and "Valor" in df_tabla.columns:
        totales_linea = (
            df_tabla.groupby("Línea de negocio", dropna=False)["Valor"]
            .sum()
            .reset_index()
        )
        st.markdown("#### Totales por línea de negocio")
        st.table(totales_linea.style.format({"Valor": "{:,.0f}"}))

        total_general = df_tabla["Valor"].sum()
        st.markdown(f"**Total general del presupuesto: {total_general:,.0f}**")

    # Bloque de firmas
    st.markdown(
        """
        <br><br>
        <table style='width:100%; text-align:center; font-size:12px'>
            <tr>
                <td>__________________________________</td>
                <td>__________________________________</td>
                <td>__________________________________</td>
            </tr>
            <tr>
                <td>
                    <b>Elaboró:</b> SANCHEZ GUERRERO Eduin Danilo<br>
                    Líder SIC VID BTA Torre Colpatria
                </td>
                <td>
                    <b>Revisó:</b> DIAZ LOPEZ Angel Alberto<br>
                    Líder CEAC VID BTA Torre Colpatria
                </td>
                <td>
                    <b>Aprobó:</b> ROMERO FERNANDEZ Guiovanna Andrea<br>
                    Líder Canal Multilinea SEG BTA Torre Colpatria
                </td>
            </tr>
        </table>
        """,
        unsafe_allow_html=True
    )


# --------------------- FUNCIÓN PARA GENERAR EXCEL CON TODAS LAS HOJAS --------------------- #

def generar_excel_todos(df):
    """
    Genera un Excel en memoria donde:
    - Cada hoja = un director + un mes.
    - Cada hoja contiene la tabla de detalle de presupuesto.
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for director in sorted(df["Director"].dropna().unique()):
            df_dir = df[df["Director"] == director]
            for mes in sorted(df_dir["Mes"].dropna().unique()):
                df_mes = df_dir[df_dir["Mes"] == mes]
                if df_mes.empty:
                    continue

                mes_nombre = get_month_name(mes)

                sheet_name = f"{director.split(',')[0][:10]}_{mes_nombre[:3]}"
                sheet_name = sheet_name.replace("/", "-")[:31]  # límite Excel

                columnas_tabla = [
                    "Linea Negocio",
                    "Ramo",
                    "Zona",
                    "Canal",
                    "Sub-Canal",
                    "Oficina",
                    "Líder comercial Esp.",
                    "Presupuesto"
                ]
                columnas_presentes = [c for c in columnas_tabla if c in df_mes.columns]
                df_out = df_mes[columnas_presentes].copy()
                df_out.to_excel(writer, sheet_name=sheet_name, index=False)

    output.seek(0)
    return output


# --------------------- LÓGICA PRINCIPAL DE LA APP --------------------- #

if uploaded_file is None:
    st.info("Sube el archivo **Presupuesto 2023.xlsx** en el panel lateral para comenzar.")
else:
    df = load_data(uploaded_file)

    st.success(
        f"Archivo cargado correctamente. Filas: {len(df):,} | Columnas: {len(df.columns)}"
    )

    # Controles para seleccionar Director y Mes
    directores = sorted(df["Director"].dropna().unique())
    meses_disponibles = sorted(df["Mes"].dropna().unique())

    col1, col2 = st.columns(2)
    with col1:
        director_sel = st.selectbox("Selecciona el Director", directores)
    with col2:
        mes_sel = st.selectbox(
            "Selecciona el mes",
            meses_disponibles,
            format_func=get_month_name
        )

    mostrar_comprobante(df, director_sel, mes_sel)

    st.markdown("---")
    st.subheader("Descargar todos los comprobantes")

    if st.button("Generar libro Excel con todas las hojas (director x mes)"):
        excel_bytes = generar_excel_todos(df)
        st.download_button(
            "Descargar archivo Excel",
            data=excel_bytes,
            file_name="Asignacion_Presupuestal_Directores.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
