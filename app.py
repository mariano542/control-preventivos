import streamlit as st
import pandas as pd
from datetime import date
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Control Preventivos",
    page_icon="🔧",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Rajdhani:wght@400;500;600;700&family=Inter:wght@300;400;500;600&display=swap');

    #MainMenu, header, footer, [data-testid="stToolbar"],
    [data-testid="stDecoration"], [data-testid="stStatusWidget"],
    .stDeployButton, [data-testid="collapsedControl"] { display: none !important; visibility: hidden !important; }

    .stApp { background: linear-gradient(135deg, #0A1628 0%, #1a2744 100%); }
    .main .block-container { padding-top: 2rem; max-width: 1400px; }

    h1, h2, h3 { font-family: 'Rajdhani', sans-serif !important; }
    p, div, span, label { font-family: 'Inter', sans-serif !important; }

    .titulo-principal {
        font-family: 'Rajdhani', sans-serif;
        font-size: 2.8rem;
        font-weight: 700;
        color: #FFFFFF;
        letter-spacing: 2px;
        text-transform: uppercase;
        line-height: 1;
    }
    .subtitulo {
        font-family: 'Inter', sans-serif;
        font-size: 0.9rem;
        color: #8BA3C7;
        margin-top: 4px;
        letter-spacing: 1px;
    }
    .metric-card {
        border-radius: 12px;
        padding: 20px;
        text-align: center;
        border: 1px solid rgba(46,117,182,0.3);
        box-shadow: 0 4px 20px rgba(0,0,0,0.3);
    }
    .metric-numero { font-family: 'Rajdhani', sans-serif; font-size: 2.5rem; font-weight: 700; color: #FFF; line-height: 1; }
    .metric-label { font-family: 'Inter', sans-serif; font-size: 0.72rem; color: #8BA3C7; text-transform: uppercase; letter-spacing: 1px; margin-top: 4px; }
    .metric-rojo { background: linear-gradient(135deg, #5C1A1A 0%, #C00000 100%); }
    .metric-amarillo { background: linear-gradient(135deg, #5C4A00 0%, #B8860B 100%); }
    .metric-verde { background: linear-gradient(135deg, #1A3A1A 0%, #2E7D32 100%); }
    .metric-azul { background: linear-gradient(135deg, #1F3864 0%, #2E75B6 100%); }

    .upload-card {
        background: linear-gradient(135deg, #1F3864 0%, #162a4e 100%);
        border: 2px dashed rgba(46,117,182,0.6);
        border-radius: 16px;
        padding: 32px 24px;
        text-align: center;
        margin-bottom: 24px;
    }
    .upload-titulo {
        font-family: 'Rajdhani', sans-serif;
        font-size: 1.3rem;
        font-weight: 700;
        color: #FFFFFF;
        text-transform: uppercase;
        letter-spacing: 1px;
        margin-bottom: 6px;
    }
    .upload-subtitulo {
        font-family: 'Inter', sans-serif;
        font-size: 0.8rem;
        color: #8BA3C7;
        margin-bottom: 16px;
    }

    .seccion-titulo {
        font-family: 'Rajdhani', sans-serif;
        font-size: 1.3rem;
        font-weight: 600;
        color: #FFF;
        text-transform: uppercase;
        letter-spacing: 1px;
        border-left: 4px solid #E25C00;
        padding-left: 12px;
        margin: 20px 0 10px 0;
    }
    .info-box {
        background: rgba(31,56,100,0.5);
        border: 1px solid rgba(46,117,182,0.3);
        border-radius: 8px;
        padding: 12px 16px;
        color: #8BA3C7;
        font-size: 0.82rem;
    }
    .config-bar {
        background: rgba(31,56,100,0.4);
        border-radius: 10px;
        padding: 10px 20px;
        display: flex;
        gap: 24px;
        align-items: center;
        margin-bottom: 20px;
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
        padding: 10px 24px !important;
        width: 100% !important;
    }
    [data-testid="stFileUploader"] {
        background: rgba(10,22,40,0.4);
        border-radius: 8px;
        padding: 4px;
    }
    [data-testid="stFileUploaderDropzone"] {
        background: transparent !important;
        border: 1px dashed rgba(46,117,182,0.4) !important;
        border-radius: 8px !important;
    }
    .stSelectbox label, .stMultiSelect label {
        color: #8BA3C7 !important;
        font-size: 0.78rem !important;
        text-transform: uppercase !important;
        letter-spacing: 1px !important;
    }
</style>
""", unsafe_allow_html=True)

FECHA_HOY = date.today()
MARGEN_MATERIALES = 180
DIAS_PROXIMO = 45

def calcular_proximo(dias):
    if pd.isna(dias): return ""
    elif dias < 0: return "REVISAR"
    elif dias >= DIAS_PROXIMO: return "OK"
    else: return "PRÓXIMO"

def procesar_csv(archivo):
    try:
        df = pd.read_csv(archivo, sep=";", encoding="utf-8-sig")
        if len(df.columns) < 5:
            archivo.seek(0)
            df = pd.read_csv(archivo, sep=",", encoding="utf-8-sig")
    except:
        archivo.seek(0)
        df = pd.read_csv(archivo, sep=",", encoding="utf-8-sig")

    if "Activo" in df.columns:
        df = df.dropna(subset=["Activo"])
        df = df[df["Activo"].astype(str).str.strip() != ""]
        df = df[df["Activo"].astype(str).str.strip() != "Activo"]

    df["Frecuencia"] = pd.to_numeric(df.get("Frecuencia", pd.Series()), errors="coerce")
    df["Frecuencia acumulada"] = pd.to_numeric(df.get("Frecuencia acumulada", pd.Series()), errors="coerce")
    df["Fecha planificada"] = pd.to_datetime(df.get("Fecha planificada", pd.Series()), dayfirst=True, errors="coerce")

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
    columnas = [
        "Tipo de mantenimiento", "Tipo de Activo", "Layout (Activo)", "Activo",
        "Código de Frecuencia", "Parte", "Tarea",
        "Fecha planificada", "FECHA DE PARTIDA", "DIAS PARA TAREA", "PRÓXIMO",
        "Frecuencia", "Frecuencia acumulada", "DIFERENCIA FRECUENCIA", "PEDIR MATERIAL",
        "Estado", "OT"
    ]
    cols = [c for c in columnas if c in df.columns]
    df_e = df[cols]
    wb = Workbook()
    wb.remove(wb.active)
    ln = Side(style="thin", color="B8B8B8")
    brd = Border(left=ln, right=ln, top=ln, bottom=ln)
    aw = {
        "Tipo de mantenimiento":18,"Tipo de Activo":22,"Layout (Activo)":35,"Activo":40,
        "Código de Frecuencia":14,"Parte":32,"Tarea":40,"Fecha planificada":18,
        "FECHA DE PARTIDA":16,"DIAS PARA TAREA":13,"PRÓXIMO":12,"Frecuencia":11,
        "Frecuencia acumulada":14,"DIFERENCIA FRECUENCIA":14,"PEDIR MATERIAL":15,"Estado":18,"OT":8
    }
    nw = {"DIFERENCIA FRECUENCIA","PEDIR MATERIAL","DIAS PARA TAREA","PRÓXIMO","FECHA DE PARTIDA"}

    def escribir(ws, datos):
        ws.freeze_panes = "A2"
        for ci, cn in enumerate(cols, 1):
            c = ws.cell(row=1, column=ci, value=cn)
            c.fill = PatternFill("solid", fgColor="2E75B6" if cn in nw else "1F3864")
            c.font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = brd
            ws.column_dimensions[get_column_letter(ci)].width = aw.get(cn, 14)
        ws.row_dimensions[1].height = 30
        for ri, (_, fila) in enumerate(datos.iterrows(), 2):
            bf = PatternFill("solid", fgColor="DCE6F1" if ri%2==0 else "FFFFFF")
            for ci, cn in enumerate(cols, 1):
                val = fila[cn]
                c = ws.cell(row=ri, column=ci)
                if cn in ("Fecha planificada","FECHA DE PARTIDA"):
                    if pd.notna(val):
                        c.value = val.date() if hasattr(val,'date') else val
                        c.number_format = 'DD/MM/YYYY'
                elif cn == "OT": c.value = int(val) if pd.notna(val) else None
                elif cn == "DIAS PARA TAREA": c.value = int(val) if pd.notna(val) else None
                else: c.value = val if pd.notna(val) else None
                c.font = Font(name="Arial", size=9)
                c.border = brd
                if cn == "PEDIR MATERIAL":
                    c.fill = PatternFill("solid", fgColor="FFE699" if val=="PEDIR MATERIAL" else "E2EFDA")
                    c.alignment = Alignment(horizontal="center", vertical="center")
                    c.font = Font(name="Arial", size=9, bold=True)
                elif cn == "PRÓXIMO":
                    col = "FFCCCC" if val=="REVISAR" else ("E2EFDA" if val=="OK" else "FFE699")
                    c.fill = PatternFill("solid", fgColor=col)
                    c.alignment = Alignment(horizontal="center", vertical="center")
                    c.font = Font(name="Arial", size=9, bold=True)
                elif cn == "DIFERENCIA FRECUENCIA":
                    c.fill = PatternFill("solid", fgColor="FFE699" if pd.notna(val) and val<=MARGEN_MATERIALES else "E2EFDA")
                    c.alignment = Alignment(horizontal="center", vertical="center")
                elif cn == "DIAS PARA TAREA":
                    if pd.notna(val):
                        col = "FFCCCC" if val<0 else ("E2EFDA" if val>=DIAS_PROXIMO else "FFE699")
                        c.fill = PatternFill("solid", fgColor=col)
                    c.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    c.fill = bf
                    c.alignment = Alignment(horizontal="center" if ci in (1,2,5,12,13,16,17) else "left", vertical="center")
            ws.row_dimensions[ri].height = 16

    ws_all = wb.create_sheet("TODOS")
    escribir(ws_all, df_e)
    for tipo in sorted(df_e["Tipo de Activo"].dropna().unique()):
        ws = wb.create_sheet(tipo[:31])
        escribir(ws, df_e[df_e["Tipo de Activo"]==tipo])
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ============================================================
# INTERFAZ PRINCIPAL
# ============================================================

# Header
st.markdown(f"""
<div style='padding: 0 0 16px 0; border-bottom: 1px solid rgba(46,117,182,0.3); margin-bottom: 20px;'>
    <div class='titulo-principal'>Control de Mantenimiento</div>
    <div class='subtitulo'>GESTIÓN DE PREVENTIVOS · FLOTA VIAL · SISTEMA CONSUMAN</div>
</div>
<div class='config-bar'>
    <span style='color:#8BA3C7; font-size:0.82rem;'>📅 Fecha: <strong style='color:white;'>{FECHA_HOY.strftime('%d/%m/%Y')}</strong></span>
    <span style='color:#8BA3C7; font-size:0.82rem;'>⚙️ Margen materiales: <strong style='color:white;'>{MARGEN_MATERIALES} hs</strong></span>
    <span style='color:#8BA3C7; font-size:0.82rem;'>⚙️ Días próximo: <strong style='color:white;'>{DIAS_PROXIMO} días</strong></span>
</div>
""", unsafe_allow_html=True)

# Upload area en el centro
col_iz, col_centro, col_der = st.columns([1, 2, 1])
with col_centro:
    st.markdown("""
    <div class='upload-card'>
        <div class='upload-titulo'>📂 Subir reporte de CONSUMAN</div>
        <div class='upload-subtitulo'>Consultar → Planes de Mantenimiento por Activo → Exportar → Guardar como CSV</div>
    </div>
    """, unsafe_allow_html=True)
    archivo = st.file_uploader("", type=["csv"], label_visibility="collapsed")

if archivo is None:
    # Pantalla de bienvenida
    st.markdown("<br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("<div class='metric-card metric-azul'><div class='metric-numero'>📂</div><div class='metric-label'>1. Subí el CSV de CONSUMAN</div></div>", unsafe_allow_html=True)
    with col2:
        st.markdown("<div class='metric-card metric-azul'><div class='metric-numero'>📊</div><div class='metric-label'>2. Visualizá el tablero</div></div>", unsafe_allow_html=True)
    with col3:
        st.markdown("<div class='metric-card metric-azul'><div class='metric-numero'>⬇️</div><div class='metric-label'>3. Descargá el Excel</div></div>", unsafe_allow_html=True)
else:
    with st.spinner("Procesando datos..."):
        df = procesar_csv(archivo)

    total = len(df)
    pedir = len(df[df["PEDIR MATERIAL"]=="PEDIR MATERIAL"])
    revisar = len(df[df["PRÓXIMO"]=="REVISAR"])
    proximo = len(df[df["PRÓXIMO"]=="PRÓXIMO"])
    ok = len(df[df["PRÓXIMO"]=="OK"])

    st.markdown("<div class='seccion-titulo'>Resumen</div>", unsafe_allow_html=True)
    col1,col2,col3,col4,col5 = st.columns(5)
    with col1: st.markdown(f"<div class='metric-card metric-azul'><div class='metric-numero'>{total}</div><div class='metric-label'>Total tareas</div></div>", unsafe_allow_html=True)
    with col2: st.markdown(f"<div class='metric-card metric-amarillo'><div class='metric-numero'>{pedir}</div><div class='metric-label'>Pedir material</div></div>", unsafe_allow_html=True)
    with col3: st.markdown(f"<div class='metric-card metric-rojo'><div class='metric-numero'>{revisar}</div><div class='metric-label'>Revisar</div></div>", unsafe_allow_html=True)
    with col4: st.markdown(f"<div class='metric-card metric-amarillo'><div class='metric-numero'>{proximo}</div><div class='metric-label'>Próximo</div></div>", unsafe_allow_html=True)
    with col5: st.markdown(f"<div class='metric-card metric-verde'><div class='metric-numero'>{ok}</div><div class='metric-label'>OK</div></div>", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("<div class='seccion-titulo'>Filtros</div>", unsafe_allow_html=True)

    col1,col2,col3 = st.columns(3)
    with col1:
        tipos = ["Todos"] + sorted(df["Tipo de Activo"].dropna().unique().tolist())
        tipo_sel = st.selectbox("Tipo de Activo", tipos)
    with col2:
        estado_mat = st.selectbox("Pedir Material", ["Todos","PEDIR MATERIAL","NO"])
    with col3:
        estado_prox = st.selectbox("Estado", ["Todos","REVISAR","PRÓXIMO","OK"])

    df_f = df.copy()
    if tipo_sel != "Todos": df_f = df_f[df_f["Tipo de Activo"]==tipo_sel]
    if estado_mat != "Todos": df_f = df_f[df_f["PEDIR MATERIAL"]==estado_mat]
    if estado_prox != "Todos": df_f = df_f[df_f["PRÓXIMO"]==estado_prox]

    st.markdown(f"<div class='seccion-titulo'>Datos — {len(df_f)} registros</div>", unsafe_allow_html=True)

    cols_ver = ["Tipo de Activo","Activo","Parte","Tarea","Fecha planificada",
                "DIAS PARA TAREA","PRÓXIMO","Frecuencia","Frecuencia acumulada",
                "DIFERENCIA FRECUENCIA","PEDIR MATERIAL","Estado"]
    cols_disp = [c for c in cols_ver if c in df_f.columns]
    df_disp = df_f[cols_disp].copy()
    if "Fecha planificada" in df_disp.columns:
        df_disp["Fecha planificada"] = pd.to_datetime(df_disp["Fecha planificada"]).dt.strftime("%d/%m/%Y")

    st.dataframe(df_disp, use_container_width=True, height=420, hide_index=True)

    st.markdown("<div class='seccion-titulo'>Descargar Excel</div>", unsafe_allow_html=True)
    col1, col2 = st.columns([1,2])
    with col1:
        excel_buf = generar_excel(df)
        nombre = f"Control_Preventivos_{FECHA_HOY.strftime('%Y%m%d')}.xlsx"
        st.download_button(
            label="⬇️  DESCARGAR EXCEL COMPLETO",
            data=excel_buf,
            file_name=nombre,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    with col2:
        st.markdown(f"""
        <div class='info-box'>
            El Excel incluye <strong style='color:white;'>{len(df)} registros</strong>,
            <strong style='color:white;'>{df['Activo'].nunique()} máquinas</strong> y
            <strong style='color:white;'>{df['Tipo de Activo'].nunique()+1} hojas</strong>
            con todos los colores y formatos.
        </div>
        """, unsafe_allow_html=True)
