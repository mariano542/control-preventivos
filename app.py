# ============================================================
# APP WEB - CONTROL DE MANTENIMIENTO PREVENTIVO
# Desarrollada con Streamlit para acceso desde cualquier PC
# ============================================================

import streamlit as st
import pandas as pd
from datetime import date
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ============================================================
# CONFIGURACION DE LA PAGINA
# ============================================================
st.set_page_config(
    page_title="Control Preventivos",
    page_icon="🔧",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================
# ESTILOS CSS PERSONALIZADOS
# ============================================================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Rajdhani:wght@400;500;600;700&family=Inter:wght@300;400;500;600&display=swap');

    :root {
        --azul-oscuro: #1F3864;
        --azul-medio: #2E75B6;
        --azul-claro: #D6E4F0;
        --naranja: #E25C00;
        --verde: #375623;
        --amarillo: #7D6608;
        --rojo: #C00000;
    }

    .main { background: #0A1628; }
    .stApp { background: linear-gradient(135deg, #0A1628 0%, #1a2744 100%); }

    h1, h2, h3 { font-family: 'Rajdhani', sans-serif !important; }
    p, div, span, label { font-family: 'Inter', sans-serif !important; }

    .titulo-principal {
        font-family: 'Rajdhani', sans-serif;
        font-size: 2.8rem;
        font-weight: 700;
        color: #FFFFFF;
        letter-spacing: 2px;
        text-transform: uppercase;
        margin-bottom: 0;
        line-height: 1;
    }

    .subtitulo {
        font-family: 'Inter', sans-serif;
        font-size: 0.95rem;
        color: #8BA3C7;
        margin-top: 4px;
        letter-spacing: 1px;
    }

    .metric-card {
        background: linear-gradient(135deg, #1F3864 0%, #2E75B6 100%);
        border-radius: 12px;
        padding: 20px;
        text-align: center;
        border: 1px solid rgba(46, 117, 182, 0.3);
        box-shadow: 0 4px 20px rgba(0,0,0,0.3);
    }

    .metric-numero {
        font-family: 'Rajdhani', sans-serif;
        font-size: 2.5rem;
        font-weight: 700;
        color: #FFFFFF;
        line-height: 1;
    }

    .metric-label {
        font-family: 'Inter', sans-serif;
        font-size: 0.75rem;
        color: #8BA3C7;
        text-transform: uppercase;
        letter-spacing: 1px;
        margin-top: 4px;
    }

    .metric-rojo { background: linear-gradient(135deg, #5C1A1A 0%, #C00000 100%); }
    .metric-amarillo { background: linear-gradient(135deg, #5C4A00 0%, #B8860B 100%); }
    .metric-verde { background: linear-gradient(135deg, #1A3A1A 0%, #2E7D32 100%); }
    .metric-azul { background: linear-gradient(135deg, #1F3864 0%, #2E75B6 100%); }

    .badge-pedir {
        background: #FFE699;
        color: #7D6608;
        padding: 2px 8px;
        border-radius: 4px;
        font-weight: 600;
        font-size: 0.75rem;
    }

    .badge-no {
        background: #E2EFDA;
        color: #375623;
        padding: 2px 8px;
        border-radius: 4px;
        font-weight: 600;
        font-size: 0.75rem;
    }

    .badge-revisar {
        background: #FFCCCC;
        color: #C00000;
        padding: 2px 8px;
        border-radius: 4px;
        font-weight: 600;
        font-size: 0.75rem;
    }

    .badge-proximo {
        background: #FFE699;
        color: #7D6608;
        padding: 2px 8px;
        border-radius: 4px;
        font-weight: 600;
        font-size: 0.75rem;
    }

    .badge-ok {
        background: #E2EFDA;
        color: #375623;
        padding: 2px 8px;
        border-radius: 4px;
        font-weight: 600;
        font-size: 0.75rem;
    }

    .stButton > button {
        background: linear-gradient(135deg, #E25C00, #FF7B2C) !important;
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        font-family: 'Rajdhani', sans-serif !important;
        font-size: 1rem !important;
        font-weight: 600 !important;
        letter-spacing: 1px !important;
        padding: 12px 24px !important;
        transition: all 0.2s !important;
        width: 100% !important;
    }

    .stButton > button:hover {
        transform: translateY(-2px) !important;
        box-shadow: 0 6px 20px rgba(226, 92, 0, 0.4) !important;
    }

    .upload-area {
        background: rgba(31, 56, 100, 0.4);
        border: 2px dashed rgba(46, 117, 182, 0.5);
        border-radius: 12px;
        padding: 20px;
        text-align: center;
    }

    .seccion-titulo {
        font-family: 'Rajdhani', sans-serif;
        font-size: 1.4rem;
        font-weight: 600;
        color: #FFFFFF;
        text-transform: uppercase;
        letter-spacing: 1px;
        border-left: 4px solid #E25C00;
        padding-left: 12px;
        margin: 20px 0 10px 0;
    }

    div[data-testid="stDataFrame"] { border-radius: 8px; overflow: hidden; }

    /* Ocultar barra de herramientas de Streamlit */
    #MainMenu { visibility: hidden; }
    header { visibility: hidden; }
    footer { visibility: hidden; }
    .stDeployButton { display: none; }
    [data-testid="stToolbar"] { display: none; }
    [data-testid="stDecoration"] { display: none; }
    [data-testid="stStatusWidget"] { display: none; }

    /* Ajustar el area de carga de archivos */
    [data-testid="stFileUploader"] {
        background: rgba(31, 56, 100, 0.4);
        border: 2px dashed rgba(46, 117, 182, 0.6);
        border-radius: 10px;
        padding: 8px;
    }
    [data-testid="stFileUploader"] label { display: none; }
    [data-testid="stFileUploaderDropzone"] {
        background: transparent !important;
        border: none !important;
        padding: 4px !important;
    }
    [data-testid="stFileUploaderDropzoneInstructions"] {
        color: #8BA3C7 !important;
        font-size: 0.8rem !important;
    }

    .stSelectbox label, .stMultiSelect label {
        color: #8BA3C7 !important;
        font-size: 0.8rem !important;
        text-transform: uppercase !important;
        letter-spacing: 1px !important;
    }

    .info-box {
        background: rgba(31, 56, 100, 0.5);
        border: 1px solid rgba(46, 117, 182, 0.3);
        border-radius: 8px;
        padding: 12px 16px;
        color: #8BA3C7;
        font-size: 0.85rem;
    }
</style>
""", unsafe_allow_html=True)

# ============================================================
# CONFIGURACION FIJA
# ============================================================
FECHA_HOY = date.today()
MARGEN_MATERIALES = 180
DIAS_PROXIMO = 45

# ============================================================
# FUNCIONES AUXILIARES
# ============================================================

def calcular_proximo(dias):
    if pd.isna(dias):
        return ""
    elif dias < 0:
        return "REVISAR"
    elif dias >= DIAS_PROXIMO:
        return "OK"
    else:
        return "PRÓXIMO"

def procesar_csv(archivo):
    """Lee y procesa el CSV de CONSUMAN, calcula todas las columnas."""
    # Intentamos leer con punto y coma primero
    try:
        df = pd.read_csv(archivo, sep=";", encoding="utf-8-sig")
        if len(df.columns) < 5:
            archivo.seek(0)
            df = pd.read_csv(archivo, sep=",", encoding="utf-8-sig")
    except:
        archivo.seek(0)
        df = pd.read_csv(archivo, sep=",", encoding="utf-8-sig")

    # Limpieza
    if "Activo" in df.columns:
        df = df.dropna(subset=["Activo"])
        df = df[df["Activo"].astype(str).str.strip() != ""]
        df = df[df["Activo"].astype(str).str.strip() != "Activo"]

    # Tipos
    df["Frecuencia"] = pd.to_numeric(df.get("Frecuencia", pd.Series()), errors="coerce")
    df["Frecuencia acumulada"] = pd.to_numeric(df.get("Frecuencia acumulada", pd.Series()), errors="coerce")
    df["Fecha planificada"] = pd.to_datetime(df.get("Fecha planificada", pd.Series()), errors="coerce")

    # Columnas calculadas
    df["DIFERENCIA FRECUENCIA"] = df["Frecuencia"] - df["Frecuencia acumulada"]
    df["PEDIR MATERIAL"] = df["DIFERENCIA FRECUENCIA"].apply(
        lambda x: "PEDIR MATERIAL" if pd.notna(x) and x <= MARGEN_MATERIALES else "NO"
    )
    df["FECHA DE PARTIDA"] = FECHA_HOY
    df["fecha_solo"] = df["Fecha planificada"].dt.normalize()
    df["DIAS PARA TAREA"] = (df["fecha_solo"] - pd.Timestamp(FECHA_HOY)).dt.days
    df["PRÓXIMO"] = df["DIAS PARA TAREA"].apply(calcular_proximo)
    df = df.drop(columns=["fecha_solo"], errors="ignore")

    return df.sort_values(["Tipo de Activo", "Activo"]).reset_index(drop=True)

def generar_excel(df):
    """Genera el archivo Excel con formato y colores."""
    columnas = [
        "Tipo de mantenimiento", "Tipo de Activo", "Layout (Activo)", "Activo",
        "Código de Frecuencia", "Parte", "Tarea",
        "Fecha planificada", "FECHA DE PARTIDA", "DIAS PARA TAREA", "PRÓXIMO",
        "Frecuencia", "Frecuencia acumulada", "DIFERENCIA FRECUENCIA", "PEDIR MATERIAL",
        "Estado", "OT"
    ]
    columnas_disponibles = [c for c in columnas if c in df.columns]
    df_export = df[columnas_disponibles]

    wb = Workbook()
    wb.remove(wb.active)

    linea = Side(style="thin", color="B8B8B8")
    borde = Border(left=linea, right=linea, top=linea, bottom=linea)

    anchos = {
        "Tipo de mantenimiento": 18, "Tipo de Activo": 22, "Layout (Activo)": 35,
        "Activo": 40, "Código de Frecuencia": 14, "Parte": 32, "Tarea": 40,
        "Fecha planificada": 18, "FECHA DE PARTIDA": 16, "DIAS PARA TAREA": 13,
        "PRÓXIMO": 12, "Frecuencia": 11, "Frecuencia acumulada": 14,
        "DIFERENCIA FRECUENCIA": 14, "PEDIR MATERIAL": 15, "Estado": 18, "OT": 8
    }
    nuevas = {"DIFERENCIA FRECUENCIA", "PEDIR MATERIAL", "DIAS PARA TAREA", "PRÓXIMO", "FECHA DE PARTIDA"}

    def escribir_hoja(ws, datos):
        ws.freeze_panes = "A2"
        for ci, cn in enumerate(columnas_disponibles, 1):
            c = ws.cell(row=1, column=ci, value=cn)
            c.fill = PatternFill("solid", fgColor="2E75B6" if cn in nuevas else "1F3864")
            c.font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = borde
            ws.column_dimensions[get_column_letter(ci)].width = anchos.get(cn, 14)
        ws.row_dimensions[1].height = 30

        for ri, (_, fila) in enumerate(datos.iterrows(), 2):
            bf = PatternFill("solid", fgColor="DCE6F1" if ri % 2 == 0 else "FFFFFF")
            for ci, cn in enumerate(columnas_disponibles, 1):
                val = fila[cn]
                c = ws.cell(row=ri, column=ci)
                if cn in ("Fecha planificada", "FECHA DE PARTIDA"):
                    if pd.notna(val):
                        c.value = val.date() if hasattr(val, 'date') else val
                        c.number_format = 'DD/MM/YYYY'
                elif cn == "OT":
                    c.value = int(val) if pd.notna(val) else None
                elif cn == "DIAS PARA TAREA":
                    c.value = int(val) if pd.notna(val) else None
                else:
                    c.value = val if pd.notna(val) else None
                c.font = Font(name="Arial", size=9)
                c.border = borde
                if cn == "PEDIR MATERIAL":
                    c.fill = PatternFill("solid", fgColor="FFE699" if val == "PEDIR MATERIAL" else "E2EFDA")
                    c.alignment = Alignment(horizontal="center", vertical="center")
                    c.font = Font(name="Arial", size=9, bold=True)
                elif cn == "PRÓXIMO":
                    col = "FFCCCC" if val == "REVISAR" else ("E2EFDA" if val == "OK" else "FFE699")
                    c.fill = PatternFill("solid", fgColor=col)
                    c.alignment = Alignment(horizontal="center", vertical="center")
                    c.font = Font(name="Arial", size=9, bold=True)
                elif cn == "DIFERENCIA FRECUENCIA":
                    c.fill = PatternFill("solid", fgColor="FFE699" if pd.notna(val) and val <= MARGEN_MATERIALES else "E2EFDA")
                    c.alignment = Alignment(horizontal="center", vertical="center")
                elif cn == "DIAS PARA TAREA":
                    if pd.notna(val):
                        col = "FFCCCC" if val < 0 else ("E2EFDA" if val >= DIAS_PROXIMO else "FFE699")
                        c.fill = PatternFill("solid", fgColor=col)
                    c.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    c.fill = bf
                    c.alignment = Alignment(
                        horizontal="center" if ci in (1,2,5,12,13,16,17) else "left",
                        vertical="center"
                    )
            ws.row_dimensions[ri].height = 16

    ws_all = wb.create_sheet("TODOS")
    escribir_hoja(ws_all, df_export)

    for tipo in sorted(df_export["Tipo de Activo"].dropna().unique()):
        ws = wb.create_sheet(tipo[:31])
        escribir_hoja(ws, df_export[df_export["Tipo de Activo"] == tipo])

    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# ============================================================
# INTERFAZ - SIDEBAR
# ============================================================
with st.sidebar:
    st.markdown("""
    <div style='text-align:center; padding: 20px 0 10px 0;'>
        <div style='font-family: Rajdhani, sans-serif; font-size: 1.5rem; font-weight: 700; color: white; letter-spacing: 2px;'>🔧 PREVENTIVOS</div>
        <div style='font-family: Inter, sans-serif; font-size: 0.7rem; color: #8BA3C7; letter-spacing: 1px;'>SISTEMA DE CONTROL</div>
    </div>
    """, unsafe_allow_html=True)

    st.divider()

    st.markdown("<div style='color: #8BA3C7; font-size: 0.75rem; text-transform: uppercase; letter-spacing: 1px; margin-bottom: 8px;'>📁 Cargar archivo CSV</div>", unsafe_allow_html=True)

    archivo = st.file_uploader(
        "Exportar desde CONSUMAN → Consultar → Planes de Mantenimiento",
        type=["csv"],
        label_visibility="collapsed"
    )

    st.markdown("<div class='info-box'>Exportá desde CONSUMAN:<br><strong>Consultar → Planes de Mantenimiento por Activo → Exportar</strong><br>Guardá como CSV y subí el archivo acá.</div>", unsafe_allow_html=True)

    st.divider()

    st.markdown(f"""
    <div style='color: #8BA3C7; font-size: 0.75rem;'>
        <div>📅 Fecha: <strong style='color: white;'>{FECHA_HOY.strftime('%d/%m/%Y')}</strong></div>
        <div style='margin-top: 4px;'>⚙️ Margen materiales: <strong style='color: white;'>{MARGEN_MATERIALES} hs</strong></div>
        <div style='margin-top: 4px;'>⚙️ Días próximo: <strong style='color: white;'>{DIAS_PROXIMO} días</strong></div>
    </div>
    """, unsafe_allow_html=True)

# ============================================================
# INTERFAZ - CONTENIDO PRINCIPAL
# ============================================================
st.markdown("""
<div style='padding: 10px 0 20px 0;'>
    <div class='titulo-principal'>Control de Mantenimiento</div>
    <div class='subtitulo'>GESTIÓN DE PREVENTIVOS · FLOTA VIAL · SISTEMA CONSUMAN</div>
</div>
""", unsafe_allow_html=True)

if archivo is None:
    # Pantalla de bienvenida cuando no hay archivo cargado
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("""
        <div class='metric-card metric-azul'>
            <div class='metric-numero'>📂</div>
            <div class='metric-label'>Subí el CSV de CONSUMAN</div>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown("""
        <div class='metric-card metric-azul'>
            <div class='metric-numero'>📊</div>
            <div class='metric-label'>Visualizá el tablero</div>
        </div>
        """, unsafe_allow_html=True)
    with col3:
        st.markdown("""
        <div class='metric-card metric-azul'>
            <div class='metric-numero'>⬇️</div>
            <div class='metric-label'>Descargá el Excel</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.info("👈 Subí el archivo CSV desde el panel izquierdo para comenzar")

else:
    # Procesamos el archivo
    with st.spinner("Procesando datos..."):
        df = procesar_csv(archivo)

    # ============================================================
    # METRICAS RESUMEN
    # ============================================================
    total = len(df)
    pedir = len(df[df["PEDIR MATERIAL"] == "PEDIR MATERIAL"])
    revisar = len(df[df["PRÓXIMO"] == "REVISAR"])
    proximo = len(df[df["PRÓXIMO"] == "PRÓXIMO"])
    ok = len(df[df["PRÓXIMO"] == "OK"])

    col1, col2, col3, col4, col5 = st.columns(5)
    with col1:
        st.markdown(f"<div class='metric-card metric-azul'><div class='metric-numero'>{total}</div><div class='metric-label'>Total tareas</div></div>", unsafe_allow_html=True)
    with col2:
        st.markdown(f"<div class='metric-card metric-amarillo'><div class='metric-numero'>{pedir}</div><div class='metric-label'>Pedir material</div></div>", unsafe_allow_html=True)
    with col3:
        st.markdown(f"<div class='metric-card metric-rojo'><div class='metric-numero'>{revisar}</div><div class='metric-label'>Revisar</div></div>", unsafe_allow_html=True)
    with col4:
        st.markdown(f"<div class='metric-card metric-amarillo'><div class='metric-numero'>{proximo}</div><div class='metric-label'>Próximo</div></div>", unsafe_allow_html=True)
    with col5:
        st.markdown(f"<div class='metric-card metric-verde'><div class='metric-numero'>{ok}</div><div class='metric-label'>OK</div></div>", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ============================================================
    # FILTROS
    # ============================================================
    st.markdown("<div class='seccion-titulo'>Filtros</div>", unsafe_allow_html=True)

    col1, col2, col3 = st.columns(3)

    with col1:
        tipos = ["Todos"] + sorted(df["Tipo de Activo"].dropna().unique().tolist())
        tipo_sel = st.selectbox("Tipo de Activo", tipos)

    with col2:
        estado_mat = st.selectbox("Pedir Material", ["Todos", "PEDIR MATERIAL", "NO"])

    with col3:
        estado_prox = st.selectbox("Estado", ["Todos", "REVISAR", "PRÓXIMO", "OK"])

    # Aplicamos filtros
    df_filtrado = df.copy()
    if tipo_sel != "Todos":
        df_filtrado = df_filtrado[df_filtrado["Tipo de Activo"] == tipo_sel]
    if estado_mat != "Todos":
        df_filtrado = df_filtrado[df_filtrado["PEDIR MATERIAL"] == estado_mat]
    if estado_prox != "Todos":
        df_filtrado = df_filtrado[df_filtrado["PRÓXIMO"] == estado_prox]

    # ============================================================
    # TABLA DE DATOS
    # ============================================================
    st.markdown(f"<div class='seccion-titulo'>Datos — {len(df_filtrado)} registros</div>", unsafe_allow_html=True)

    # Columnas para mostrar en el tablero
    cols_mostrar = ["Tipo de Activo", "Activo", "Parte", "Tarea",
                    "Fecha planificada", "DIAS PARA TAREA", "PRÓXIMO",
                    "Frecuencia", "Frecuencia acumulada", "DIFERENCIA FRECUENCIA", "PEDIR MATERIAL", "Estado"]
    cols_disponibles = [c for c in cols_mostrar if c in df_filtrado.columns]

    # Función para colorear las celdas
    def colorear_fila(val, col):
        if col == "PEDIR MATERIAL":
            if val == "PEDIR MATERIAL":
                return "background-color: #FFE699; color: #7D6608; font-weight: bold;"
            return "background-color: #E2EFDA; color: #375623; font-weight: bold;"
        elif col == "PRÓXIMO":
            if val == "REVISAR":
                return "background-color: #FFCCCC; color: #C00000; font-weight: bold;"
            elif val == "PRÓXIMO":
                return "background-color: #FFE699; color: #7D6608; font-weight: bold;"
            elif val == "OK":
                return "background-color: #E2EFDA; color: #375623; font-weight: bold;"
        return ""

    # Mostramos la tabla
    df_display = df_filtrado[cols_disponibles].copy()
    if "Fecha planificada" in df_display.columns:
        df_display["Fecha planificada"] = pd.to_datetime(df_display["Fecha planificada"]).dt.strftime("%d/%m/%Y")

    st.dataframe(
        df_display,
        use_container_width=True,
        height=400,
        hide_index=True
    )

    # ============================================================
    # BOTON DE DESCARGA
    # ============================================================
    st.markdown("<div class='seccion-titulo'>Descargar Excel</div>", unsafe_allow_html=True)

    col1, col2 = st.columns([1, 2])
    with col1:
        with st.spinner("Preparando Excel..."):
            excel_buffer = generar_excel(df)

        nombre_archivo = f"Control_Preventivos_{FECHA_HOY.strftime('%Y%m%d')}.xlsx"

        st.download_button(
            label="⬇️  DESCARGAR EXCEL COMPLETO",
            data=excel_buffer,
            file_name=nombre_archivo,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    with col2:
        tipos_count = df["Tipo de Activo"].nunique()
        maquinas_count = df["Activo"].nunique()
        st.markdown(f"""
        <div class='info-box'>
            El Excel incluye <strong style='color:white;'>{len(df)} registros</strong>,
            <strong style='color:white;'>{maquinas_count} máquinas</strong> y
            <strong style='color:white;'>{tipos_count + 1} hojas</strong>
            (TODOS + una por cada tipo de activo) con todos los colores y formatos.
        </div>
        """, unsafe_allow_html=True)
