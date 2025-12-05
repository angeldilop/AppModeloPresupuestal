import io
import unicodedata
from typing import List, Optional

import pandas as pd
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    Image as RLImage,
    PageBreak,
)
from PIL import Image


def quitar_acentos(texto: str) -> str:
    if not isinstance(texto, str):
        return ""
    nfkd = unicodedata.normalize("NFKD", texto)
    return "".join(c for c in nfkd if not unicodedata.combining(c))


def normalizar_nombre_col(col: str) -> str:
    return quitar_acentos(col).strip().lower()


def buscar_columna(
    df: pd.DataFrame, candidatos: List[str], contiene: Optional[str] = None
) -> Optional[str]:
    norm_map = {normalizar_nombre_col(c): c for c in df.columns}

    for cand in candidatos:
        cand_norm = normalizar_nombre_col(cand)
        if cand_norm in norm_map:
            return norm_map[cand_norm]

    if contiene:
        sub = normalizar_nombre_col(contiene)
        posibles = [
            real
            for norm, real in norm_map.items()
            if sub in norm and not norm.startswith("unnamed")
        ]
        if posibles:
            return sorted(posibles)[0]

    return None


def nombre_mes_es(num_mes: int) -> str:
    meses = [
        "enero",
        "febrero",
        "marzo",
        "abril",
        "mayo",
        "junio",
        "julio",
        "agosto",
        "septiembre",
        "octubre",
        "noviembre",
        "diciembre",
    ]
    if 1 <= num_mes <= 12:
        return meses[num_mes - 1]
    return f"Mes {num_mes}"


def preparar_logo(logo_bytes: Optional[bytes], altura: float = 18 * mm) -> Optional[RLImage]:
    if not logo_bytes:
        return None

    buffer_logo = io.BytesIO(logo_bytes)
    img = RLImage(buffer_logo)

    aspect = img.imageWidth / float(img.imageHeight)
    img.drawHeight = altura
    img.drawWidth = altura * aspect
    return img


def construir_encabezado(story, styles, logo_bytes: Optional[bytes]):
    logo_flowable = preparar_logo(logo_bytes)

    titulo_lines = [
        "PROGRAMA DE ASIGNACIÓN PRESUPUESTAL",
        "AXA COLPATRIA",
        "Dirección de Distribución",
        "SISTEMA DE INTELIGENCIA COMERCIAL",
    ]
    titulo_text = "<br/>".join(titulo_lines)

    estilo_titulo = ParagraphStyle(
        "TituloEncabezado",
        parent=styles["Normal"],
        alignment=1,
        fontSize=12,
        leading=14,
    )
    par_titulo = Paragraph(titulo_text, estilo_titulo)

    width, _ = landscape(letter)
    tabla_data = [[par_titulo, logo_flowable if logo_flowable else ""]]
    tabla = Table(tabla_data, colWidths=[0.75 * width, 0.25 * width])

    tabla.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("ALIGN", (0, 0), (0, 0), "CENTER"),
                ("ALIGN", (1, 0), (1, 0), "RIGHT"),
                ("BOX", (0, 0), (-1, -1), 0, colors.white),
            ]
        )
    )

    story.append(tabla)
    story.append(Spacer(1, 6 * mm))


def construir_tabla_detalle(
    df_subset: pd.DataFrame,
    cols_detalle: List[str],
    story,
):
    header = [
        "Línea de negocio",
        "Ramo",
        "Zona",
        "Canal",
        "Sub-Canal",
        "Oficina",
        "Líder Equipo",
        "Valor",
    ]

    data = [header]
    for _, row in df_subset[cols_detalle].iterrows():
        fila = [
            str(row[cols_detalle[0]]),
            str(row[cols_detalle[1]]),
            str(row[cols_detalle[2]]),
            str(row[cols_detalle[3]]),
            str(row[cols_detalle[4]]),
            str(row[cols_detalle[5]]),
            str(row[cols_detalle[6]]),
            f"{float(row[cols_detalle[7]]):,.0f}",
        ]
        data.append(fila)

    tabla = Table(data, repeatRows=1)
    tabla.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                ("ALIGN", (0, 0), (-1, 0), "CENTER"),
                ("ALIGN", (7, 1), (7, -1), "RIGHT"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
            ]
        )
    )

    story.append(tabla)
    story.append(Spacer(1, 5 * mm))


def construir_resumen(
    df_subset: pd.DataFrame,
    col_linea: str,
    col_valor: str,
    styles,
    story,
):
    resumen = (
        df_subset.groupby(col_linea)[col_valor]
        .sum()
        .reset_index()
        .sort_values(col_linea)
    )

    header = ["Línea de negocio", "Valor total"]
    data = [header]
    for _, row in resumen.iterrows():
        data.append(
            [
                str(row[col_linea]),
                f"{float(row[col_valor]):,.0f}",
            ]
        )

    tabla = Table(data, repeatRows=1, colWidths=[80 * mm, 40 * mm])
    tabla.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                ("ALIGN", (1, 1), (1, -1), "RIGHT"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
            ]
        )
    )

    story.append(tabla)

    total_general = float(df_subset[col_valor].sum())
    estilo_total = ParagraphStyle(
        "TotalPresupuesto",
        parent=styles["Normal"],
        alignment=0,
        fontSize=9,
    )
    texto_total = (
        f"Total general del presupuesto: "
        f"<b>{total_general:,.0f}</b>"
    )
    story.append(Spacer(1, 3 * mm))
    story.append(Paragraph(texto_total, estilo_total))
    story.append(Spacer(1, 8 * mm))


def construir_pie_pagina(styles, story):
    estilo_footer = ParagraphStyle(
        "Footer",
        parent=styles["Normal"],
        alignment=0,
        fontSize=8,
        leading=10,
    )

    textos = [
        (
            "Usuario Elaboró: SANCHEZ GUERRERO Eduin Danilo<br/>"
            "Líder SIC VID BTA Torre Colpatria"
        ),
        (
            "Usuario Revisó: DIAZ LOPEZ Angel Alberto<br/>"
            "Líder CEAC VID BTA Torre Colpatria"
        ),
        (
            "Usuario Aprobó: ROMERO FERNANDEZ Guiovanna Andrea<br/>"
            "Líder Canal Multilinea SEG BTA Torre Colpatria"
        ),
    ]

    pars = [Paragraph(t, estilo_footer) for t in textos]

    tabla = Table([pars], colWidths=[70 * mm, 70 * mm, 70 * mm])
    tabla.setStyle(
        TableStyle(
            [
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ]
        )
    )

    story.append(tabla)


def generar_pdf_presupuesto(
    df: pd.DataFrame,
    logo_bytes: Optional[bytes] = None,
) -> bytes:
    styles = getSampleStyleSheet()
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(letter),
        leftMargin=15 * mm,
        rightMargin=15 * mm,
        topMargin=10 * mm,
        bottomMargin=10 * mm,
    )

    col_director = buscar_columna(
        df,
        ["Director", "Franquicia", "Nombre Director", "Director Comercial"],
        contiene="director",
    )
    col_mes = buscar_columna(df, ["Mes"], contiene="mes")
    col_ano = buscar_columna(df, ["Año", "Ano", "Year"], contiene="an")

    col_linea = buscar_columna(
        df, ["Línea de negocio", "Linea de negocio"], contiene="linea"
    )
    col_ramo = buscar_columna(df, ["Ramo"], contiene="ramo")
    col_zona = buscar_columna(df, ["Zona"], contiene="zona")
    col_canal = buscar_columna(df, ["Canal"], contiene="canal")
    col_subcanal = buscar_columna(
        df, ["Sub-Canal", "Subcanal", "Sub Canal"], contiene="sub"
    )
    col_oficina = buscar_columna(df, ["Oficina"], contiene="oficina")
    col_lider = buscar_columna(
        df, ["Lider Equipo", "Líder Equipo", "Lider"], contiene="lider"
    )
    col_valor = buscar_columna(
        df, ["Valor", "Presupuesto", "Valor Total"], contiene="valor"
    )

    columnas_requeridas = {
        "Director": col_director,
        "Mes": col_mes,
        "Año": col_ano,
        "Línea de negocio": col_linea,
        "Ramo": col_ramo,
        "Zona": col_zona,
        "Canal": col_canal,
        "Sub-Canal": col_subcanal,
        "Oficina": col_oficina,
        "Líder Equipo": col_lider,
        "Valor": col_valor,
    }

    faltantes = [log for log, real in columnas_requeridas.items() if real is None]
    if faltantes:
        raise ValueError(
            "No se pudieron encontrar las columnas requeridas en el archivo Excel. "
            f"Faltan: {', '.join(faltantes)}"
        )

    df[col_valor] = pd.to_numeric(df[col_valor], errors="coerce").fillna(0)

    story = []

    directores = sorted([d for d in df[col_director].dropna().unique()])

    for director in directores:
        df_dir = df[df[col_director] == director].copy()

        meses_unicos = [m for m in df_dir[col_mes].dropna().unique()]
        try:
            meses_ordenados = sorted(meses_unicos, key=lambda x: int(float(x)))
        except Exception:
            meses_ordenados = sorted(meses_unicos, key=lambda x: str(x))

        for mes in meses_ordenados:
            df_dm = df_dir[df_dir[col_mes] == mes].copy()
            if df_dm.empty:
                continue

            construir_encabezado(story, styles, logo_bytes)

            mes_val = df_dm[col_mes].iloc[0]
            try:
                mes_num = int(float(mes_val))
            except Exception:
                mes_num = None

            if mes_num is not None and 1 <= mes_num <= 12:
                nombre_mes = nombre_mes_es(mes_num)
            else:
                nombre_mes = str(mes_val)

            anos = [a for a in df_dm[col_ano].dropna().unique()]
            if anos:
                try:
                    ano_text = str(int(float(anos[0])))
                except Exception:
                    ano_text = str(anos[0])
            else:
                ano_text = ""

            estilo_sub = ParagraphStyle(
                "Subtitulo",
                parent=styles["Normal"],
                alignment=1,
                fontSize=10,
                leading=12,
            )
            texto_sub = (
                f"Asignación presupuestal para {director} "
                f"del mes de {nombre_mes}"
            )
            if ano_text:
                texto_sub += f" del {ano_text}"

            story.append(Paragraph(texto_sub, estilo_sub))
            story.append(Spacer(1, 5 * mm))

            cols_detalle = [
                col_linea,
                col_ramo,
                col_zona,
                col_canal,
                col_subcanal,
                col_oficina,
                col_lider,
                col_valor,
            ]

            construir_tabla_detalle(df_dm, cols_detalle, story)
            construir_resumen(df_dm, col_linea, col_valor, styles, story)
            construir_pie_pagina(styles, story)
            story.append(PageBreak())

    if story and isinstance(story[-1], PageBreak):
        story = story[:-1]

    doc.build(story)
    pdf_bytes = buffer.getvalue()
    buffer.close()
    return pdf_bytes


def main():
    st.set_page_config(
        page_title="Programa de Asignación Presupuestal",
        layout="centered",
    )

    st.title("Programa de Asignación Presupuestal")

    st.write(
        """
        Carga el archivo de presupuesto (2023, 2024 o 2025) y, opcionalmente, 
        el logo de AXA COLPATRIA.  
        La aplicación generará un PDF con una hoja por cada Director y por cada mes.
        """
    )

    archivo = st.file_uploader(
        "Cargar archivo de presupuesto (Excel)",
        type=["xlsx", "xls"],
    )

    logo_file = st.file_uploader(
        "Cargar logo de AXA COLPATRIA (PNG/JPG)",
        type=["png", "jpg", "jpeg"],
    )

    logo_bytes = None
    if logo_file is not None:
        logo_bytes = logo_file.read()
        st.image(logo_bytes, caption="Previsualización del logo", width=160)

    if archivo is not None:
        try:
            df = pd.read_excel(archivo)
        except Exception as e:
            st.error(f"Error al leer el archivo de Excel: {e}")
            return

        st.subheader("Vista previa de los datos")
        st.dataframe(df.head(20))

        st.write("---")
        st.subheader("Generar PDF")

        if st.button("Preparar PDF de asignación presupuestal"):
            try:
                pdf_bytes = generar_pdf_presupuesto(df, logo_bytes)
                st.success("PDF generado correctamente.")
                st.download_button(
                    label="Descargar PDF",
                    data=pdf_bytes,
                    file_name="programa_asignacion_presupuestal.pdf",
                    mime="application/pdf",
                )
            except Exception as e:
                st.error(f"Ocurrió un error al generar el PDF: {e}")


if __name__ == "__main__":
    main()
