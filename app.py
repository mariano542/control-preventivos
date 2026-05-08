import streamlit as st
import pandas as pd
from datetime import date
from io import BytesIO
import base64
import os
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="APESA · Control Preventivos",
    page_icon="🔧",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@400;600;700;800&family=Inter:wght@300;400;500;600&display=swap');

    #MainMenu, header, footer, [data-testid="stToolbar"],
    [data-testid="stDecoration"], [data-testid="stStatusWidget"],
    .stDeployButton, [data-testid="collapsedControl"] { display: none !important; visibility: hidden !important; }

    /* FONDO BLANCO */
    .stApp { background: #FFFFFF !important; }
    .main .block-container { background: #FFFFFF !important; padding-top: 0; max-width: 1400px; padding-left: 2rem; padding-right: 2rem; }
    section[data-testid="stMain"] { background: #FFFFFF !important; }

    h1,h2,h3 { font-family: 'Barlow Condensed', sans-serif !important; }
    p,div,span,label { font-family: 'Inter', sans-serif !important; }

    /* HEADER */
    .apesa-header {
        background: #1A1A1A;
        border-bottom: 4px solid #C8102E;
        padding: 16px 24px;
        margin: 0 -2rem 24px -2rem;
        display: flex;
        align-items: center;
        gap: 20px;
    }
    .apesa-header-titulo {
        font-family: 'Barlow Condensed', sans-serif;
        font-size: 1.8rem;
        font-weight: 800;
        color: #FFFFFF;
        letter-spacing: 3px;
        text-transform: uppercase;
        line-height: 1;
    }
    .apesa-header-sub {
        font-family: 'Inter', sans-serif;
        font-size: 0.68rem;
        color: #999;
        letter-spacing: 2px;
        text-transform: uppercase;
        margin-top: 3px;
    }
    .apesa-header-divider {
        width: 3px;
        height: 44px;
        background: #C8102E;
        border-radius: 2px;
        flex-shrink: 0;
    }

    /* CONFIG BAR */
    .config-bar {
        background: #F5F5F5;
        border-radius: 8px;
        padding: 10px 20px;
        display: flex;
        gap: 28px;
        align-items: center;
        margin-bottom: 20px;
        border-left: 4px solid #C8102E;
        border: 1px solid #E0E0E0;
        border-left: 4px solid #C8102E;
    }

    /* METRICS */
    .metric-card {
        border-radius: 8px;
        padding: 18px 12px;
        text-align: center;
        border-top: 4px solid transparent;
    }
    .metric-numero { font-family: 'Barlow Condensed', sans-serif; font-size: 2.8rem; font-weight: 800; line-height: 1; }
    .metric-label { font-family: 'Inter', sans-serif; font-size: 0.68rem; text-transform: uppercase; letter-spacing: 1px; margin-top: 4px; }
    .metric-total { background: #1A1A1A; border-color: #333; }
    .metric-total .metric-numero { color: #FFFFFF; }
    .metric-total .metric-label { color: #999; }
    .metric-rojo { background: #FFF0F2; border-color: #C8102E; }
    .metric-rojo .metric-numero { color: #C8102E; }
    .metric-rojo .metric-label { color: #e05060; }
    .metric-naranja { background: #FFF8EE; border-color: #E07000; }
    .metric-naranja .metric-numero { color: #E07000; }
    .metric-naranja .metric-label { color: #c06000; }
    .metric-verde { background: #F0FFF4; border-color: #2E7D32; }
    .metric-verde .metric-numero { color: #2E7D32; }
    .metric-verde .metric-label { color: #388e3c; }
    .metric-amarillo { background: #FFFDE7; border-color: #F9A825; }
    .metric-amarillo .metric-numero { color: #F57F17; }
    .metric-amarillo .metric-label { color: #e65100; }

    /* SECCION TITULO */
    .seccion-titulo {
        font-family: 'Barlow Condensed', sans-serif;
        font-size: 1.2rem;
        font-weight: 700;
        color: #1A1A1A;
        text-transform: uppercase;
        letter-spacing: 2px;
        border-left: 4px solid #C8102E;
        padding-left: 10px;
        margin: 20px 0 12px 0;
    }

    /* UPLOAD */
    .upload-card {
        background: #F8F8F8;
        border: 2px dashed #CCCCCC;
        border-radius: 12px;
        padding: 28px 24px 16px 24px;
        text-align: center;
        margin-bottom: 24px;
        transition: border-color 0.2s;
    }
    .upload-card:hover { border-color: #C8102E; }
    .upload-titulo {
        font-family: 'Barlow Condensed', sans-serif;
        font-size: 1.2rem;
        font-weight: 700;
        color: #1A1A1A;
        text-transform: uppercase;
        letter-spacing: 2px;
        margin-bottom: 4px;
    }
    .upload-sub {
        font-family: 'Inter', sans-serif;
        font-size: 0.75rem;
        color: #888;
        margin-bottom: 12px;
    }

    /* BOTONES */
    .stButton > button, .stDownloadButton > button {
        background: #C8102E !important;
        color: white !important;
        border: none !important;
        border-radius: 6px !important;
        font-family: 'Barlow Condensed', sans-serif !important;
        font-size: 1rem !important;
        font-weight: 700 !important;
        letter-spacing: 2px !important;
        padding: 10px 24px !important;
        width: 100% !important;
        text-transform: uppercase !important;
    }
    .stButton > button:hover, .stDownloadButton > button:hover {
        background: #a50d25 !important;
    }

    /* SELECTBOX */
    .stSelectbox label, .stTextInput label {
        color: #555 !important;
        font-size: 0.72rem !important;
        text-transform: uppercase !important;
        letter-spacing: 1px !important;
    }
    [data-testid="stSelectbox"] > div > div {
        background: #F5F5F5 !important;
        color: #1A1A1A !important;
        border-color: #CCCCCC !important;
    }

    /* FILE UPLOADER */
    [data-testid="stFileUploader"] { background: transparent; }
    [data-testid="stFileUploaderDropzone"] {
        background: #FAFAFA !important;
        border: 1px dashed #CCC !important;
        border-radius: 8px !important;
    }
    [data-testid="stFileUploaderDropzoneInstructions"] { color: #888 !important; }

    /* INFO BOX */
    .info-box {
        background: #F5F5F5;
        border: 1px solid #E0E0E0;
        border-radius: 8px;
        padding: 12px 16px;
        color: #555;
        font-size: 0.82rem;
    }

    /* TABLA */
    [data-testid="stDataFrame"] { border-radius: 8px; overflow: hidden; border: 1px solid #E0E0E0; }

    /* DIVIDER */
    hr { border-color: #E0E0E0 !important; }

    /* SPINNER */
    .stSpinner { color: #C8102E !important; }
</style>
""", unsafe_allow_html=True)

# ============================================================
# CONFIG
# ============================================================
FECHA_HOY = date.today()
MARGEN_MATERIALES = 180
DIAS_PROXIMO = 45

def calcular_proximo(dias):
    if pd.isna(dias): return ""
    elif dias < 0: return "REVISAR"
    elif dias >= DIAS_PROXIMO: return "OK"
    else: return "PRÓXIMO"

def get_logo_base64():
    paths = ["assets/logo.png", "assets/logo.jpg", "assets/logo.PNG"]
    for p in paths:
        if os.path.exists(p):
            ext = p.split(".")[-1].lower()
            mime = "image/png" if ext == "png" else "image/jpeg"
            with open(p, "rb") as f:
                data = base64.b64encode(f.read()).decode()
            return f"data:{mime};base64,{data}"
    return None

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
        "Tipo de mantenimiento","Tipo de Activo","Layout (Activo)","Activo",
        "Código de Frecuencia","Parte","Tarea","Fecha planificada","FECHA DE PARTIDA",
        "DIAS PARA TAREA","PRÓXIMO","Frecuencia","Frecuencia acumulada",
        "DIFERENCIA FRECUENCIA","PEDIR MATERIAL","Estado","OT"
    ]
    cols = [c for c in columnas if c in df.columns]
    df_e = df[cols]
    wb = Workbook()
    wb.remove(wb.active)
    ln = Side(style="thin", color="CCCCCC")
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
            c.fill = PatternFill("solid", fgColor="C8102E" if cn in nw else "1A1A1A")
            c.font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = brd
            ws.column_dimensions[get_column_letter(ci)].width = aw.get(cn, 14)
        ws.row_dimensions[1].height = 30
        for ri, (_, fila) in enumerate(datos.iterrows(), 2):
            bf = PatternFill("solid", fgColor="F5F5F5" if ri%2==0 else "FFFFFF")
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
                c.font = Font(name="Arial", size=9, color="1A1A1A")
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
# HEADER
# ============================================================
logo_src = get_logo_base64()
if logo_src:
    logo_html = f'<img src="{logo_src}" style="height: 44px; object-fit: contain;">'
else:
    logo_html = '<span style="font-family: Barlow Condensed, sans-serif; font-size: 1.8rem; font-weight: 800; color: #C8102E; letter-spacing: 3px;">A.PE.S.A.</span>'

st.markdown(f"""
<div class="apesa-header">
    {logo_html}
    <div class="apesa-header-divider"></div>
    <div>
        <div class="apesa-header-titulo">Control de Mantenimiento</div>
        <div class="apesa-header-sub">Gestión de Preventivos · Flota Vial · Sistema CONSUMAN</div>
    </div>
</div>
""", unsafe_allow_html=True)

# CONFIG BAR
st.markdown(f"""
<div class="config-bar">
    <span style="color:#555; font-size:0.8rem;">📅 Fecha: <strong style="color:#1A1A1A;">{FECHA_HOY.strftime('%d/%m/%Y')}</strong></span>
    <span style="color:#555; font-size:0.8rem;">⚙️ Margen materiales: <strong style="color:#1A1A1A;">{MARGEN_MATERIALES} hs</strong></span>
    <span style="color:#555; font-size:0.8rem;">⚙️ Días próximo: <strong style="color:#1A1A1A;">{DIAS_PROXIMO} días</strong></span>
</div>
""", unsafe_allow_html=True)

# ============================================================
# UPLOAD
# ============================================================
col_iz, col_centro, col_der = st.columns([1, 2, 1])
with col_centro:
    st.markdown("""
    <div class="upload-card">
        <div class="upload-titulo">📂 Subir reporte de CONSUMAN</div>
        <div class="upload-sub">Consultar → Planes de Mantenimiento por Activo → Exportar → Guardar como CSV</div>
    </div>
    """, unsafe_allow_html=True)
    archivo = st.file_uploader("", type=["csv"], label_visibility="collapsed")

# ============================================================
# CONTENIDO PRINCIPAL
# ============================================================
if archivo is None:
    st.markdown("<br>", unsafe_allow_html=True)
    c1,c2,c3 = st.columns(3)
    with c1: st.markdown("<div class='metric-card metric-total'><div class='metric-numero'>📂</div><div class='metric-label'>1. Subí el CSV</div></div>", unsafe_allow_html=True)
    with c2: st.markdown("<div class='metric-card metric-total'><div class='metric-numero'>📊</div><div class='metric-label'>2. Visualizá el tablero</div></div>", unsafe_allow_html=True)
    with c3: st.markdown("<div class='metric-card metric-total'><div class='metric-numero'>⬇️</div><div class='metric-label'>3. Descargá el Excel</div></div>", unsafe_allow_html=True)
else:
    with st.spinner("Procesando datos..."):
        df = procesar_csv(archivo)

    total   = len(df)
    pedir   = len(df[df["PEDIR MATERIAL"]=="PEDIR MATERIAL"])
    revisar = len(df[df["PRÓXIMO"]=="REVISAR"])
    proximo = len(df[df["PRÓXIMO"]=="PRÓXIMO"])
    ok      = len(df[df["PRÓXIMO"]=="OK"])

    st.markdown("<div class='seccion-titulo'>Resumen</div>", unsafe_allow_html=True)
    c1,c2,c3,c4,c5 = st.columns(5)
    with c1: st.markdown(f"<div class='metric-card metric-total'><div class='metric-numero'>{total}</div><div class='metric-label'>Total tareas</div></div>", unsafe_allow_html=True)
    with c2: st.markdown(f"<div class='metric-card metric-amarillo'><div class='metric-numero'>{pedir}</div><div class='metric-label'>Pedir material</div></div>", unsafe_allow_html=True)
    with c3: st.markdown(f"<div class='metric-card metric-rojo'><div class='metric-numero'>{revisar}</div><div class='metric-label'>Revisar</div></div>", unsafe_allow_html=True)
    with c4: st.markdown(f"<div class='metric-card metric-naranja'><div class='metric-numero'>{proximo}</div><div class='metric-label'>Próximo</div></div>", unsafe_allow_html=True)
    with c5: st.markdown(f"<div class='metric-card metric-verde'><div class='metric-numero'>{ok}</div><div class='metric-label'>OK</div></div>", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("<div class='seccion-titulo'>Filtros</div>", unsafe_allow_html=True)

    c1,c2,c3,c4 = st.columns(4)
    with c1:
        tipos = ["Todos"] + sorted(df["Tipo de Activo"].dropna().unique().tolist())
        tipo_sel = st.selectbox("Tipo de Activo", tipos)
    with c2:
        estado_mat = st.selectbox("Pedir Material", ["Todos","PEDIR MATERIAL","NO"])
    with c3:
        estado_prox = st.selectbox("Estado", ["Todos","REVISAR","PRÓXIMO","OK"])
    with c4:
        buscar = st.text_input("Buscar activo", placeholder="Ej: AP17, C42...")

    df_f = df.copy()
    if tipo_sel != "Todos": df_f = df_f[df_f["Tipo de Activo"]==tipo_sel]
    if estado_mat != "Todos": df_f = df_f[df_f["PEDIR MATERIAL"]==estado_mat]
    if estado_prox != "Todos": df_f = df_f[df_f["PRÓXIMO"]==estado_prox]
    if buscar: df_f = df_f[df_f["Activo"].str.contains(buscar, case=False, na=False)]

    st.markdown(f"<div class='seccion-titulo'>Datos — {len(df_f)} registros</div>", unsafe_allow_html=True)

    cols_ver = ["Tipo de Activo","Activo","Parte","Tarea","Fecha planificada",
                "DIAS PARA TAREA","PRÓXIMO","Frecuencia","Frecuencia acumulada",
                "DIFERENCIA FRECUENCIA","PEDIR MATERIAL","Estado"]
    cols_disp = [c for c in cols_ver if c in df_f.columns]
    df_disp = df_f[cols_disp].copy()
    if "Fecha planificada" in df_disp.columns:
        df_disp["Fecha planificada"] = pd.to_datetime(df_disp["Fecha planificada"]).dt.strftime("%d/%m/%Y")

    st.dataframe(df_disp, use_container_width=True, height=420, hide_index=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("<div class='seccion-titulo'>Descargar Excel</div>", unsafe_allow_html=True)

    c1,c2 = st.columns([1,2])
    with c1:
        excel_buf = generar_excel(df)
        nombre = f"APESA_Preventivos_{FECHA_HOY.strftime('%Y%m%d')}.xlsx"
        st.download_button(
            label="⬇️  DESCARGAR EXCEL COMPLETO",
            data=excel_buf,
            file_name=nombre,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    with c2:
        st.markdown(f"""
        <div class='info-box'>
            El Excel incluye <strong style='color:#1A1A1A;'>{len(df)} registros</strong>,
            <strong style='color:#1A1A1A;'>{df['Activo'].nunique()} máquinas</strong> y
            <strong style='color:#1A1A1A;'>{df['Tipo de Activo'].nunique()+1} hojas</strong>
            con encabezados APESA y semáforos de color.
        </div>
        """, unsafe_allow_html=True)
