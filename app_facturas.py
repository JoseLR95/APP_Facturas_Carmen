import os
import re
import shutil
import tempfile
import zipfile
import streamlit as st
import pdfplumber
import openpyxl
from collections import Counter

# ===========================================================================
# ESTILOS
# ===========================================================================

st.set_page_config(page_title="Renombrador de Facturas", page_icon="📄", layout="centered")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=DM+Sans:wght@300;400;600&display=swap');

html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif;
}

.stApp {
    background: #0f0f0f;
    color: #e8e8e8;
}

h1 {
    font-family: 'DM Mono', monospace !important;
    font-size: 1.6rem !important;
    color: #f0e040 !important;
    letter-spacing: -0.02em;
    margin-bottom: 0 !important;
}

.subtitle {
    font-family: 'DM Mono', monospace;
    font-size: 0.8rem;
    color: #555;
    margin-top: 0.2rem;
    margin-bottom: 2rem;
}

.block-label {
    font-family: 'DM Mono', monospace;
    font-size: 0.72rem;
    color: #f0e040;
    text-transform: uppercase;
    letter-spacing: 0.1em;
    margin-bottom: 0.3rem;
}

.stat-box {
    background: #1a1a1a;
    border: 1px solid #2a2a2a;
    border-radius: 8px;
    padding: 1rem 1.2rem;
    text-align: center;
}

.stat-num {
    font-family: 'DM Mono', monospace;
    font-size: 2rem;
    font-weight: 500;
    color: #f0e040;
    line-height: 1;
}

.stat-label {
    font-size: 0.72rem;
    color: #666;
    margin-top: 0.3rem;
    text-transform: uppercase;
    letter-spacing: 0.05em;
}

.log-box {
    background: #111;
    border: 1px solid #222;
    border-radius: 8px;
    padding: 1rem 1.2rem;
    font-family: 'DM Mono', monospace;
    font-size: 0.72rem;
    color: #888;
    max-height: 320px;
    overflow-y: auto;
    line-height: 1.7;
}

.log-box .ok   { color: #6fcf97; }
.log-box .warn { color: #f2994a; }
.log-box .err  { color: #eb5757; }
.log-box .info { color: #56ccf2; }

section[data-testid="stFileUploadDropzone"] {
    background: #1a1a1a !important;
    border: 1px dashed #333 !important;
    border-radius: 8px !important;
}

div[data-testid="stDownloadButton"] button {
    background: #f0e040 !important;
    color: #0f0f0f !important;
    font-family: 'DM Mono', monospace !important;
    font-weight: 500 !important;
    border: none !important;
    border-radius: 6px !important;
    padding: 0.6rem 1.5rem !important;
    font-size: 0.85rem !important;
    width: 100%;
    margin-top: 1rem;
}

div[data-testid="stDownloadButton"] button:hover {
    background: #fff176 !important;
}

.stProgress > div > div {
    background: #f0e040 !important;
}

hr {
    border-color: #1e1e1e !important;
    margin: 1.5rem 0 !important;
}
</style>
""", unsafe_allow_html=True)

# ===========================================================================
# AUTENTICACIÓN
# ===========================================================================

def check_password():
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False

    if st.session_state.authenticated:
        return True

    st.markdown("<h1>📄 Renombrador de Facturas</h1>", unsafe_allow_html=True)
    st.markdown("<div class='subtitle'>// Introduce la contraseña para acceder</div>", unsafe_allow_html=True)
    st.markdown("<div class='block-label'>Contraseña de acceso</div>", unsafe_allow_html=True)

    password = st.text_input("", type="password", placeholder="Introduce la contraseña", label_visibility="collapsed")

    if st.button("Acceder", use_container_width=True):
        if password == st.secrets["PASSWORD"]:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("Contraseña incorrecta.")

    return False

if not check_password():
    st.stop()

# ===========================================================================
# FUNCIONES DE PROCESAMIENTO
# ===========================================================================

def cargar_excel(excel_bytes, hoja, fila_inicio=2):
    wb = openpyxl.load_workbook(excel_bytes, data_only=True)
    ws = wb[hoja]

    filas = []
    for fila in ws.iter_rows(min_row=fila_inicio, values_only=True):
        codigo = fila[6]  # Columna G
        if codigo is None:
            continue
        codigo = str(codigo).strip()
        col_c  = str(fila[2]).strip() if fila[2] is not None else ""
        col_d  = str(fila[3]).strip() if fila[3] is not None else ""
        col_e  = str(fila[4]).strip() if fila[4] is not None else ""
        col_h  = str(fila[7]).strip() if fila[7] is not None else ""
        filas.append((codigo, col_c, col_d, col_e, col_h))

    conteo_excel     = Counter(f[0] for f in filas)
    duplicados_excel = {cod for cod, cnt in conteo_excel.items() if cnt > 1}

    mapeo = {}
    for codigo, col_c, col_d, col_e, col_h in filas:
        if codigo not in mapeo:
            mapeo[codigo] = (col_c, col_d, col_e, col_h)

    mapeo_duplicados = {}
    for codigo, col_c, col_d, col_e, col_h in filas:
        if codigo in duplicados_excel:
            if codigo not in mapeo_duplicados:
                mapeo_duplicados[codigo] = []
            mapeo_duplicados[codigo].append((col_c, col_d, col_e, col_h))

    return mapeo, duplicados_excel, mapeo_duplicados


def extraer_dnis_celda(col_h):
    return [d.strip().upper() for d in re.split(r'[\s/\-]+', col_h) if d.strip()]

def normalizar_dni(dni):
    if re.match(r'^0\d{8}[A-Z]$', dni):
        return dni[1:]
    return dni

def extraer_texto_pdf(ruta_pdf):
    texto = ""
    try:
        with pdfplumber.open(ruta_pdf) as pdf:
            for pagina in pdf.pages:
                t = pagina.extract_text()
                if t:
                    texto += t + "\n"
    except:
        pass
    return texto

def es_anexo(texto):
    return "Anexo Factura Tarjeta Cheque Gourmet" in texto

def extraer_codigo_cliente(texto):
    match = re.search(r'(\d{5})\s*N[oº°]\s*PROVEEDOR', texto)
    return match.group(1).strip() if match else None

def extraer_numero_factura_factura(texto):
    match = re.search(r'N[oº°]\s*PEDIDO[:\s]+(\w+)', texto)
    return match.group(1).strip() if match else None

def extraer_numero_factura_anexo(texto):
    lineas = texto.splitlines()
    for i, linea in enumerate(lineas):
        if "Nº FACTURA:" in linea and i > 0:
            return lineas[i - 1].strip()
    return None

def extraer_dnis_anexo(texto):
    return [m.upper() for m in re.findall(r'\b(?:[0-9]{9}[A-Z]|[XYZ][0-9]{7}[A-Z])\b', texto)]

def sanitizar_nombre(nombre):
    return re.sub(r'[\\/*?:"<>|]', "", nombre).strip()


def procesar_facturas(pdf_files, excel_bytes, hoja="RECARGAS", fila_inicio=2):
    logs  = []
    stats = {"unicos": 0, "repetidos_pdf": 0, "duplicados_resueltos": 0,
             "duplicados_no_resueltos": 0, "no_encontrados": 0, "anexos": 0}

    def log(msg, tipo=""):
        logs.append((msg, tipo))

    # Cargar Excel
    try:
        mapeo, duplicados_excel, mapeo_duplicados = cargar_excel(excel_bytes, hoja, fila_inicio)
        log(f"Excel cargado — {len(mapeo)} códigos, {len(duplicados_excel)} duplicados", "info")
    except Exception as e:
        log(f"Error cargando Excel: {e}", "err")
        return None, logs, stats

    # Guardar PDFs en carpeta temporal
    tmpdir = tempfile.mkdtemp()
    for f in pdf_files:
        ruta = os.path.join(tmpdir, f.name)
        with open(ruta, "wb") as out:
            out.write(f.read())

    pdfs = sorted(os.listdir(tmpdir))
    log(f"{len(pdfs)} PDFs recibidos", "info")

    # PASO 1: Clasificar en facturas y anexos
    facturas = {}
    anexos   = {}

    for nombre_pdf in pdfs:
        ruta  = os.path.join(tmpdir, nombre_pdf)
        texto = extraer_texto_pdf(ruta)

        if es_anexo(texto):
            num_factura = extraer_numero_factura_anexo(texto)
            dnis        = [normalizar_dni(d) for d in extraer_dnis_anexo(texto)]
            if num_factura:
                anexos[num_factura] = dnis
            log(f"[ANEXO] {nombre_pdf} → factura: {num_factura}, {len(dnis)} DNIs", "info")
            stats["anexos"] += 1
        else:
            codigo      = extraer_codigo_cliente(texto)
            num_factura = extraer_numero_factura_factura(texto)
            facturas[nombre_pdf] = {"codigo": codigo, "num_factura": num_factura}
            log(f"[FACTURA] {nombre_pdf} → código: {codigo}", "")

    # PASO 2: Contar repeticiones por código
    conteo_pdfs = Counter(v["codigo"] for v in facturas.values() if v["codigo"])

    # PASO 3: Primera pasada — no duplicados en Excel
    outdir          = tempfile.mkdtemp()
    contador_vistos = {}
    pendientes      = {}

    for nombre_pdf, datos in facturas.items():
        codigo      = datos["codigo"]
        num_factura = datos["num_factura"]

        if not codigo:
            log(f"⚠ {nombre_pdf} → sin código, ignorado", "warn")
            continue

        contador_vistos[codigo] = contador_vistos.get(codigo, 0) + 1
        n     = contador_vistos[codigo]
        total = conteo_pdfs[codigo]

        if codigo in duplicados_excel:
            pendientes[nombre_pdf] = {"codigo": codigo, "num_factura": num_factura, "n": n}
            continue

        if codigo not in mapeo:
            nuevo_nombre = sanitizar_nombre(f"{codigo}-No encontrado") + ".pdf"
            stats["no_encontrados"] += 1
        elif total == 1:
            col_c, col_d, col_e, _ = mapeo[codigo]
            partes = [p for p in [col_c, col_d, col_e] if p]
            nuevo_nombre = sanitizar_nombre("-".join(partes)) + ".pdf"
            stats["unicos"] += 1
        else:
            col_c, col_d, col_e, _ = mapeo[codigo]
            partes = [p for p in [col_c, col_d, col_e] if p]
            nuevo_nombre = sanitizar_nombre("-".join(partes)) + f"-{n}.pdf"
            stats["repetidos_pdf"] += 1

        shutil.copy2(os.path.join(tmpdir, nombre_pdf), os.path.join(outdir, nuevo_nombre))
        log(f"✓ {nombre_pdf} → {nuevo_nombre}", "ok")

    # PASO 4: Segunda pasada — duplicados Excel
    for nombre_pdf, datos in pendientes.items():
        codigo      = datos["codigo"]
        num_factura = datos["num_factura"]
        n           = datos["n"]
        nuevo_nombre = None

        log(f"[DUP] {nombre_pdf} → código: {codigo}, factura: {num_factura}", "warn")

        if num_factura and num_factura in anexos:
            dnis_anexo   = [normalizar_dni(d) for d in anexos[num_factura]]
            filas_codigo = mapeo_duplicados.get(codigo, [])

            for col_c, col_d, col_e, col_h in filas_codigo:
                dnis_excel = [normalizar_dni(d) for d in extraer_dnis_celda(col_h)]
                if any(dni in dnis_excel for dni in dnis_anexo):
                    partes = [p for p in [col_c, col_d, col_e] if p]
                    nuevo_nombre = sanitizar_nombre("-".join(partes)) + ".pdf"
                    stats["duplicados_resueltos"] += 1
                    log(f"  ✓ DNI coincide → {nuevo_nombre}", "ok")
                    break

        if not nuevo_nombre:
            nuevo_nombre = sanitizar_nombre(f"{codigo}-Duplicado-{n}") + ".pdf"
            stats["duplicados_no_resueltos"] += 1
            log(f"  → {nuevo_nombre}", "warn")

        shutil.copy2(os.path.join(tmpdir, nombre_pdf), os.path.join(outdir, nuevo_nombre))

    # Crear ZIP
    zip_path = os.path.join(tempfile.mkdtemp(), "facturas_renombradas.zip")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in os.listdir(outdir):
            zf.write(os.path.join(outdir, f), f)

    return zip_path, logs, stats


# ===========================================================================
# INTERFAZ PRINCIPAL
# ===========================================================================

st.markdown("<h1>📄 Renombrador de Facturas</h1>", unsafe_allow_html=True)
st.markdown("<div class='subtitle'>// Carga los PDFs y el Excel para renombrar automáticamente</div>", unsafe_allow_html=True)

st.markdown("<div class='block-label'>1 · Fichero Excel</div>", unsafe_allow_html=True)
excel_file = st.file_uploader("", type=["xlsx"], key="excel", label_visibility="collapsed")

st.markdown("<div class='block-label'>2 · PDFs de facturas y anexos</div>", unsafe_allow_html=True)
pdf_files = st.file_uploader("", type=["pdf"], accept_multiple_files=True, key="pdfs", label_visibility="collapsed")

col1, col2 = st.columns(2)
with col1:
    hoja = st.text_input("Hoja del Excel", value="RECARGAS")
with col2:
    fila_inicio = st.number_input("Fila de inicio (datos)", min_value=1, value=2)

st.markdown("<hr>", unsafe_allow_html=True)

if st.button("▶ Procesar facturas", use_container_width=True):
    if not excel_file:
        st.error("Sube el fichero Excel.")
    elif not pdf_files:
        st.error("Sube al menos un PDF.")
    else:
        with st.spinner("Procesando..."):
            zip_path, logs, stats = procesar_facturas(
                pdf_files, excel_file, hoja=hoja, fila_inicio=int(fila_inicio)
            )

        if zip_path:
            st.markdown("<hr>", unsafe_allow_html=True)

            # Stats
            total = sum(v for k, v in stats.items() if k != "anexos")
            c1, c2, c3, c4, c5 = st.columns(5)
            for col, label, key in [
                (c1, "Únicos",       "unicos"),
                (c2, "Repetidos",    "repetidos_pdf"),
                (c3, "Dup. resuel.", "duplicados_resueltos"),
                (c4, "Dup. pend.",   "duplicados_no_resueltos"),
                (c5, "No encontr.",  "no_encontrados"),
            ]:
                with col:
                    st.markdown(f"""
                    <div class='stat-box'>
                        <div class='stat-num'>{stats[key]}</div>
                        <div class='stat-label'>{label}</div>
                    </div>""", unsafe_allow_html=True)

            st.markdown(
                f"<br><div style='text-align:center;font-family:DM Mono,monospace;font-size:0.8rem;color:#555'>"
                f"TOTAL GESTIONADOS: <span style='color:#f0e040'>{total}</span> · "
                f"ANEXOS IGNORADOS: <span style='color:#555'>{stats['anexos']}</span></div>",
                unsafe_allow_html=True
            )

            # Log
            st.markdown("<br><div class='block-label'>Log de procesamiento</div>", unsafe_allow_html=True)
            log_html = ""
            for msg, tipo in logs:
                css = {"ok": "ok", "warn": "warn", "err": "err", "info": "info"}.get(tipo, "")
                log_html += f"<div class='{css}'>{msg}</div>"
            st.markdown(f"<div class='log-box'>{log_html}</div>", unsafe_allow_html=True)

            # Descarga
            with open(zip_path, "rb") as f:
                st.download_button(
                    label="⬇ Descargar ZIP con facturas renombradas",
                    data=f.read(),
                    file_name="facturas_renombradas.zip",
                    mime="application/zip"
                )
