import streamlit as st
import pandas as pd
from datetime import date
from io import BytesIO
import base64
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="APESA · Control Preventivos", page_icon="🔧", layout="wide", initial_sidebar_state="collapsed")

LOGO_B64 = "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"

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
    df["PEDIR MATERIAL"] = df["DIFERENCIA FRECUENCIA"].apply(lambda x: "PEDIR MATERIAL" if pd.notna(x) and x <= MARGEN_MATERIALES else "NO")
    df["FECHA DE PARTIDA"] = FECHA_HOY
    df["fecha_solo"] = df["Fecha planificada"].dt.normalize()
    df["DIAS PARA TAREA"] = (df["fecha_solo"] - pd.Timestamp(FECHA_HOY)).dt.days
    df["PRÓXIMO"] = df["DIAS PARA TAREA"].apply(calcular_proximo)
    df = df.drop(columns=["fecha_solo"], errors="ignore")
    return df.sort_values(["Tipo de Activo", "Activo"]).reset_index(drop=True)

def generar_excel(df):
    columnas = ["Tipo de mantenimiento","Tipo de Activo","Layout (Activo)","Activo","Código de Frecuencia","Parte","Tarea","Fecha planificada","FECHA DE PARTIDA","DIAS PARA TAREA","PRÓXIMO","Frecuencia","Frecuencia acumulada","DIFERENCIA FRECUENCIA","PEDIR MATERIAL","Estado","OT"]
    cols = [c for c in columnas if c in df.columns]
    df_e = df[cols]
    wb = Workbook()
    wb.remove(wb.active)
    ln = Side(style="thin", color="CCCCCC")
    brd = Border(left=ln, right=ln, top=ln, bottom=ln)
    aw = {"Tipo de mantenimiento":18,"Tipo de Activo":22,"Layout (Activo)":35,"Activo":40,"Código de Frecuencia":14,"Parte":32,"Tarea":40,"Fecha planificada":18,"FECHA DE PARTIDA":16,"DIAS PARA TAREA":13,"PRÓXIMO":12,"Frecuencia":11,"Frecuencia acumulada":14,"DIFERENCIA FRECUENCIA":14,"PEDIR MATERIAL":15,"Estado":18,"OT":8}
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
                        c.value = val.date() if hasattr(val,"date") else val
                        c.number_format = "DD/MM/YYYY"
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

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@400;600;700;800&family=Inter:wght@300;400;500;600&display=swap');
    #MainMenu,header,footer,[data-testid="stToolbar"],[data-testid="stDecoration"],[data-testid="stStatusWidget"],.stDeployButton,[data-testid="collapsedControl"]{display:none!important;visibility:hidden!important;}
    .stApp{background:#111111!important;}
    .main .block-container{background:#111111!important;padding-top:0!important;max-width:100%!important;padding-left:0!important;padding-right:0!important;}
    section[data-testid="stMain"]{background:#111111!important;}
    h1,h2,h3{font-family:'Barlow Condensed',sans-serif!important;}
    p,div,span,label{font-family:'Inter',sans-serif!important;}
    .apesa-header{background:#FFFFFF;border-bottom:5px solid #C8102E;padding:20px 40px;display:flex;align-items:center;gap:24px;}
    .apesa-divider{width:3px;height:56px;background:#C8102E;border-radius:2px;flex-shrink:0;}
    .apesa-titulo{font-family:'Barlow Condensed',sans-serif;font-size:2.2rem;font-weight:800;color:#C8102E;letter-spacing:3px;text-transform:uppercase;line-height:1;}
    .apesa-sub{font-family:'Inter',sans-serif;font-size:0.68rem;color:#888;letter-spacing:2px;text-transform:uppercase;margin-top:4px;}
    .config-bar{background:#1A1A1A;padding:10px 40px;display:flex;gap:40px;align-items:center;border-bottom:1px solid #2A2A2A;}
    .content-pad{padding:24px 32px;}
    .metric-card{border-radius:10px;padding:20px 12px;text-align:center;border-top:4px solid transparent;box-shadow:0 2px 12px rgba(0,0,0,0.4);}
    .metric-numero{font-family:'Barlow Condensed',sans-serif;font-size:3rem;font-weight:800;line-height:1;}
    .metric-label{font-family:'Inter',sans-serif;font-size:0.68rem;text-transform:uppercase;letter-spacing:1px;margin-top:5px;}
    .metric-total{background:#1E1E1E;border-color:#444;}.metric-total .metric-numero{color:#FFF;}.metric-total .metric-label{color:#777;}
    .metric-rojo{background:#2A0008;border-color:#C8102E;}.metric-rojo .metric-numero{color:#C8102E;}.metric-rojo .metric-label{color:#C8102E;opacity:0.7;}
    .metric-naranja{background:#2A1500;border-color:#E07000;}.metric-naranja .metric-numero{color:#E07000;}.metric-naranja .metric-label{color:#E07000;opacity:0.7;}
    .metric-verde{background:#0A2A10;border-color:#2E7D32;}.metric-verde .metric-numero{color:#4CAF50;}.metric-verde .metric-label{color:#4CAF50;opacity:0.7;}
    .metric-amarillo{background:#2A2000;border-color:#F9A825;}.metric-amarillo .metric-numero{color:#F9A825;}.metric-amarillo .metric-label{color:#F9A825;opacity:0.7;}
    .seccion-titulo{font-family:'Barlow Condensed',sans-serif;font-size:1.15rem;font-weight:700;color:#FFFFFF;text-transform:uppercase;letter-spacing:2px;border-left:4px solid #C8102E;padding-left:10px;margin:24px 0 12px 0;}
    .upload-card{background:#1E1E1E;border:2px dashed #444;border-radius:14px;padding:36px 28px 20px 28px;text-align:center;margin-bottom:8px;transition:border-color 0.2s;}
    .upload-icono{font-size:3rem;margin-bottom:8px;}
    .upload-titulo{font-family:'Barlow Condensed',sans-serif;font-size:1.3rem;font-weight:700;color:#FFFFFF;text-transform:uppercase;letter-spacing:2px;margin-bottom:6px;}
    .upload-sub{font-family:'Inter',sans-serif;font-size:0.77rem;color:#666;margin-bottom:4px;}
    .upload-hint{font-family:'Inter',sans-serif;font-size:0.68rem;color:#444;margin-bottom:14px;}
    .stButton > button,.stDownloadButton > button{background:#C8102E!important;color:white!important;border:none!important;border-radius:7px!important;font-family:'Barlow Condensed',sans-serif!important;font-size:1rem!important;font-weight:700!important;letter-spacing:2px!important;padding:12px 28px!important;width:100%!important;text-transform:uppercase!important;box-shadow:0 3px 10px rgba(200,16,46,0.3)!important;}
    .stButton > button:hover,.stDownloadButton > button:hover{background:#a50d25!important;}
    .stSelectbox label,.stTextInput label{color:#777!important;font-size:0.72rem!important;text-transform:uppercase!important;letter-spacing:1px!important;}
    [data-testid="stSelectbox"] > div > div{background:#1E1E1E!important;color:#FFFFFF!important;border-color:#444!important;border-radius:7px!important;}
    [data-testid="stTextInput"] input{background:#1E1E1E!important;color:#FFFFFF!important;border-color:#444!important;border-radius:7px!important;}
    [data-testid="stFileUploader"]{background:transparent;}
    [data-testid="stFileUploaderDropzone"]{background:#1A1A1A!important;border:1px dashed #C8102E!important;border-radius:10px!important;padding:16px!important;}
    [data-testid="stFileUploaderDropzoneInstructions"]{display:none!important;}
    [data-testid="stFileUploader"] > label{display:none!important;}
    [data-testid="stFileUploaderDropzone"] button{
        font-size: 0 !important;
        width: 48px !important;
        height: 48px !important;
        border-radius: 50% !important;
        background: #C8102E !important;
        border: none !important;
        position: relative !important;
        padding: 0 !important;
        min-width: unset !important;
    }
    [data-testid="stFileUploaderDropzone"] button::after{
        content: "⬆️" !important;
        font-size: 1.4rem !important;
        position: absolute !important;
        top: 50% !important;
        left: 50% !important;
        transform: translate(-50%, -50%) !important;
    }
    [data-testid="stDataFrame"]{border-radius:10px;overflow:hidden;border:1px solid #333;box-shadow:0 2px 12px rgba(0,0,0,0.4);}
    .info-box{background:#1E1E1E;border:1px solid #333;border-radius:10px;padding:14px 18px;color:#888;font-size:0.82rem;}
    hr{border-color:#2A2A2A!important;}
</style>
""", unsafe_allow_html=True)

st.markdown(f"""
<div class="apesa-header">
    <img src="data:image/png;base64,{LOGO_B64}" style="height:56px;object-fit:contain;">
    <div class="apesa-divider"></div>
    <div>
        <div class="apesa-titulo">Control de Mantenimiento</div>
        <div class="apesa-sub">Gestión de Preventivos · Flota Vial · Sistema CONSUMAN</div>
    </div>
</div>
""", unsafe_allow_html=True)

fecha_str = FECHA_HOY.strftime('%d/%m/%Y')
st.markdown(f"""
<div class="config-bar">
    <span style="color:#777;font-size:0.82rem;">📅 Fecha: <strong style="color:#FFF;">{fecha_str}</strong></span>
    <span style="color:#777;font-size:0.82rem;">⚙️ Margen materiales: <strong style="color:#FFF;">{MARGEN_MATERIALES} hs</strong></span>
    <span style="color:#777;font-size:0.82rem;">⚙️ Días próximo: <strong style="color:#FFF;">{DIAS_PROXIMO} días</strong></span>
</div>
""", unsafe_allow_html=True)

st.markdown("<div class='content-pad'>", unsafe_allow_html=True)

# Hide native uploader and show custom button
st.markdown("""
<style>
    /* Ocultar TODO el componente nativo de uploader */
    [data-testid="stFileUploader"] {
        position: absolute !important;
        opacity: 0 !important;
        width: 1px !important;
        height: 1px !important;
        overflow: hidden !important;
        pointer-events: none !important;
    }
    /* Boton custom de upload */
    .upload-btn-custom {
        display: block;
        width: 100%;
        padding: 40px 20px;
        background: #1E1E1E;
        border: 2px dashed #C8102E;
        border-radius: 14px;
        text-align: center;
        cursor: pointer;
        transition: all 0.2s;
        margin-bottom: 8px;
    }
    .upload-btn-custom:hover {
        background: #2A2A2A;
        border-color: #FF3050;
    }
    .upload-btn-custom .upload-arrow {
        font-size: 3rem;
        display: block;
        margin-bottom: 10px;
    }
    .upload-btn-custom .upload-main {
        font-family: 'Barlow Condensed', sans-serif;
        font-size: 1.3rem;
        font-weight: 700;
        color: #FFFFFF;
        text-transform: uppercase;
        letter-spacing: 2px;
        display: block;
        margin-bottom: 6px;
    }
    .upload-btn-custom .upload-sec {
        font-family: 'Inter', sans-serif;
        font-size: 0.82rem;
        color: #C8102E;
        font-weight: 600;
        display: block;
        margin-bottom: 4px;
    }
    .upload-btn-custom .upload-hint-text {
        font-family: 'Inter', sans-serif;
        font-size: 0.68rem;
        color: #444;
        display: block;
    }
</style>
""", unsafe_allow_html=True)

col_iz, col_centro, col_der = st.columns([1, 2, 1])
with col_centro:
    # Native uploader (hidden but functional)
    archivo = st.file_uploader("Subir CSV", type=["csv"], label_visibility="hidden")
    
    # Custom visual button (shown on top via JS click)
    if archivo is None:
        st.markdown("""
        <label for="fileInput" class="upload-btn-custom" onclick="document.querySelector('[data-testid=stFileUploaderDropzone]').click()">
            <span class="upload-arrow">⬆️</span>
            <span class="upload-main">Subir reporte de CONSUMAN</span>
            <span class="upload-sec">Adjuntá el archivo CSV</span>
            <span class="upload-hint-text">Consultar → Planes de Mantenimiento por Activo → Exportar → Guardar como CSV</span>
        </label>
        """, unsafe_allow_html=True)

if archivo is None:
    st.markdown("<br>", unsafe_allow_html=True)
    c1,c2,c3 = st.columns(3)
    with c1: st.markdown("<div class='metric-card metric-total'><div class='metric-numero'>📂</div><div class='metric-label'>1. Subí el CSV</div></div>", unsafe_allow_html=True)
    with c2: st.markdown("<div class='metric-card metric-total'><div class='metric-numero'>📊</div><div class='metric-label'>2. Visualizá el tablero</div></div>", unsafe_allow_html=True)
    with c3: st.markdown("<div class='metric-card metric-total'><div class='metric-numero'>⬇️</div><div class='metric-label'>3. Descargá el Excel</div></div>", unsafe_allow_html=True)
else:
    with st.spinner("⚙️ Procesando datos..."):
        df = procesar_csv(archivo)
    total=len(df)
    pedir=len(df[df["PEDIR MATERIAL"]=="PEDIR MATERIAL"])
    revisar=len(df[df["PRÓXIMO"]=="REVISAR"])
    proximo=len(df[df["PRÓXIMO"]=="PRÓXIMO"])
    ok=len(df[df["PRÓXIMO"]=="OK"])
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
        tipos=["Todos"]+sorted(df["Tipo de Activo"].dropna().unique().tolist())
        tipo_sel=st.selectbox("Tipo de Activo",tipos)
    with c2: estado_mat=st.selectbox("Pedir Material",["Todos","PEDIR MATERIAL","NO"])
    with c3: estado_prox=st.selectbox("Estado",["Todos","REVISAR","PRÓXIMO","OK"])
    with c4: buscar=st.text_input("Buscar activo",placeholder="Ej: AP17, C42...")
    df_f=df.copy()
    if tipo_sel!="Todos": df_f=df_f[df_f["Tipo de Activo"]==tipo_sel]
    if estado_mat!="Todos": df_f=df_f[df_f["PEDIR MATERIAL"]==estado_mat]
    if estado_prox!="Todos": df_f=df_f[df_f["PRÓXIMO"]==estado_prox]
    if buscar: df_f=df_f[df_f["Activo"].str.contains(buscar,case=False,na=False)]
    st.markdown(f"<div class='seccion-titulo'>Datos — {len(df_f)} registros</div>", unsafe_allow_html=True)
    cols_ver=["Tipo de Activo","Activo","Parte","Tarea","Fecha planificada","DIAS PARA TAREA","PRÓXIMO","Frecuencia","Frecuencia acumulada","DIFERENCIA FRECUENCIA","PEDIR MATERIAL","Estado"]
    cols_disp=[c for c in cols_ver if c in df_f.columns]
    df_disp=df_f[cols_disp].copy()
    if "Fecha planificada" in df_disp.columns:
        df_disp["Fecha planificada"]=pd.to_datetime(df_disp["Fecha planificada"]).dt.strftime("%d/%m/%Y")
    st.dataframe(df_disp,use_container_width=True,height=420,hide_index=True)
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("<div class='seccion-titulo'>Descargar Excel</div>", unsafe_allow_html=True)
    c1,c2=st.columns([1,2])
    with c1:
        excel_buf=generar_excel(df)
        nombre=f"APESA_Preventivos_{FECHA_HOY.strftime('%Y%m%d')}.xlsx"
        st.download_button(label="⬇️  DESCARGAR EXCEL COMPLETO",data=excel_buf,file_name=nombre,mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with c2:
        maquinas=df['Activo'].nunique()
        tipos_count=df['Tipo de Activo'].nunique()
        st.markdown(f"<div class='info-box'>El Excel incluye <strong style='color:white;'>{total} registros</strong>, <strong style='color:white;'>{maquinas} máquinas</strong> y <strong style='color:white;'>{tipos_count+1} hojas</strong> con encabezados APESA y semáforos de colores.</div>", unsafe_allow_html=True)

st.markdown("</div>", unsafe_allow_html=True)
