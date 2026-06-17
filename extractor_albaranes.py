# -*- coding: utf-8 -*-
"""
Extractor de albaranes (Sección Salsas) por OCR.
- Lee todos los PDF de una carpeta.
- En cada PDF localiza SOLO las paginas de ALBARAN (ignora COA, lavado EFTCO y bascula).
- Extrae: Nº albaran, fecha, cliente, descripcion, lote, cantidad y Nº de pedido (S/Pedido).
- Cruza la descripcion / codigo de proveedor con un MAESTRO para asignar VUESTRO codigo interno.
- Vuelca una linea por articulo a un Excel, con enlace directo a la pagina del PDF.
"""
import os, re, io, csv, glob, sys, shutil
import fitz                      # PyMuPDF: rasteriza el PDF
import pytesseract               # motor OCR (Tesseract)
from PIL import Image
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ============================================================
# CONFIGURACION  (lo unico que normalmente tocaras)
# ============================================================
CARPETA_PDF   = "."                      # carpeta donde caen los PDF (en produccion: ruta de OneDrive)
CARPETA_PROCESADOS = "ALBARANES PROCESADOS"  # subcarpeta a la que se mueve el PDF ya leido
MAESTRO_CSV   = "maestro_codigos.csv"     # tabla clave -> codigo interno (la mantienes tu)
SALIDA_XLSX   = "Albaranes_Salsas.xlsx"   # Excel resultado
IDIOMA_OCR    = "spa"
DPI           = 300
# Si Tesseract no esta en el PATH (Windows), descomenta y ajusta:
# pytesseract.pytesseract.tesseract_cmd = r"C:\Tesseract\tesseract.exe"

def _resolver_tesseract():
    """Localiza Tesseract tanto si va incrustado en el .exe como si esta instalado en el PC."""
    base = getattr(sys, "_MEIPASS", "")
    candidatos = [
        os.path.join(base, "Tesseract-OCR", "tesseract.exe") if base else "",
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    ]
    for c in candidatos:
        if c and os.path.exists(c):
            pytesseract.pytesseract.tesseract_cmd = c
            td = os.path.join(os.path.dirname(c), "tessdata")
            if os.path.isdir(td):
                os.environ["TESSDATA_PREFIX"] = td
            return
    # si no se encuentra, se confia en que 'tesseract' este en el PATH


# ============================================================
# CLASIFICACION DE PAGINAS POR PALABRAS CLAVE
# ============================================================
def tipo_de_pagina(texto):
    t = texto.upper()
    if "ALBARAN" in t or "ALBARÁN" in t:
        return "ALBARAN"
    if "INFORME ANAL" in t:           # COA / Informe Analitico
        return "COA"
    if "EFTCO" in t or "DOCUMENTO DE LAVADO" in t:
        return "LAVADO"
    if "BASCULA" in t or "BÁSCULA" in t or "Nº TICKET" in t or "N? TICKET" in t:
        return "BASCULA"
    return "OTRO"

# ============================================================
# OCR
# ============================================================
def ocr_pagina(page, dpi=DPI, psm=None):
    pix = page.get_pixmap(dpi=dpi)
    img = Image.open(io.BytesIO(pix.tobytes("png")))
    cfg = f"--psm {psm}" if psm else ""
    return pytesseract.image_to_string(img, lang=IDIOMA_OCR, config=cfg), img

def ocr_banda(img, top, bot):
    """OCR de una banda horizontal (fraccion de alto) — util para tablas con bordes."""
    w, h = img.size
    crop = img.crop((0, int(h*top), w, int(h*bot)))
    return pytesseract.image_to_string(crop, lang=IDIOMA_OCR, config="--psm 6")

# ============================================================
# MAESTRO  clave -> (codigo_interno, descripcion_interna)
# ============================================================
def cargar_maestro(ruta):
    m = {}
    if not os.path.exists(ruta):
        return m
    with open(ruta, encoding="utf-8-sig") as f:
        for fila in csv.DictReader(f, delimiter=";"):
            clave = fila["clave"].strip().upper()
            m[clave] = (fila["codigo_interno"].strip(), fila.get("descripcion_interna", "").strip())
    return m

def asignar_codigo(maestro, articulo_prov, descripcion):
    """1º intenta el codigo de proveedor exacto; 2º busca cualquier clave dentro de la descripcion."""
    if articulo_prov and articulo_prov.upper() in maestro:
        return maestro[articulo_prov.upper()]
    desc = (descripcion or "").upper()
    for clave, val in maestro.items():
        if clave in desc:
            return val
    return ("", "")   # sin coincidencia -> se marcara en rojo

# ============================================================
# EXTRACCION DE CAMPOS DE UN ALBARAN
# ============================================================
def buscar(patron, texto, grupo=1, flags=0):
    m = re.search(patron, texto, flags)
    return m.group(grupo).strip() if m else ""

def extraer_albaran(texto_full, img):
    cab = ocr_banda(img, 0.20, 0.28)   # banda FECHA/NUMERO/CLIENTE/NIF

    fecha   = buscar(r"\b(\d{2}/\d{2}/\d{2,4})\b", cab)
    if len(fecha) == 8:                # 08/06/26 -> 08/06/2026
        fecha = fecha[:6] + "20" + fecha[6:]
    num_alb = buscar(r"\b(SF\s*\d{2,4})\b", cab).replace("  ", " ")
    cliente = buscar(r"\b(\d{8})\b", cab)
    nif     = buscar(r"\b([A-Z]\d{8})\b", cab)

    # S/Pedido (numero de pedido, suele empezar por 45)
    s_pedido = buscar(r"S\s*/?\s*PEDIDO\s+(\d{6,})", texto_full, flags=re.I) \
               or buscar(r"\b(45\d{6,})\b", texto_full)

    # Linea de articulo: CODIGO_PROVEEDOR  DESCRIPCION ... LOTE  CANTIDAD ...
    articulo = buscar(r"\b([A-Z]{3,}[A-Z0-9]*\d{2,})\b\s+VINAGRE", texto_full)
    descripcion = buscar(r"\b[A-Z]{3,}[A-Z0-9]*\d{2,}\s+(VINAGRE[^\n]+)", texto_full)
    descripcion = re.sub(r"\s+", " ", descripcion).strip()

    lote     = buscar(r"\b(SF-\d{4,})\b", texto_full)          # formato de lote de este proveedor
    cantidad = buscar(r"(\d{1,3}(?:\.\d{3})*,\d{2})(?!\d)", texto_full)  # 26.140,00
    impreso  = buscar(r"(?<!\d)(4\d{5})(?!\d)", texto_full)    # codigo 4xxxxx si viniera impreso

    return [{
        "num_albaran": num_alb, "fecha": fecha, "cliente": cliente, "nif": nif,
        "articulo_prov": articulo, "descripcion": descripcion,
        "lote": lote, "cantidad": cantidad, "s_pedido": s_pedido, "cod_impreso": impreso,
    }]

# ============================================================
# PROCESO PRINCIPAL
# ============================================================
def destino_unico(carpeta, nombre):
    """Devuelve una ruta libre en 'carpeta' (evita machacar si ya existe ese nombre)."""
    base, ext = os.path.splitext(nombre)
    dest = os.path.join(carpeta, nombre); n = 2
    while os.path.exists(dest):
        dest = os.path.join(carpeta, f"{base} ({n}){ext}"); n += 1
    return dest

def procesar():
    maestro = cargar_maestro(MAESTRO_CSV)
    carpeta_proc = os.path.join(CARPETA_PDF, CARPETA_PROCESADOS)
    os.makedirs(carpeta_proc, exist_ok=True)
    filas = []
    # glob no es recursivo: los PDF ya movidos a la subcarpeta NO se vuelven a leer
    for ruta_pdf in sorted(glob.glob(os.path.join(CARPETA_PDF, "*.pdf"))):
        doc = fitz.open(ruta_pdf)
        filas_pdf = []
        for i, page in enumerate(doc):
            texto, img = ocr_pagina(page)
            if tipo_de_pagina(texto) != "ALBARAN":
                continue
            for art in extraer_albaran(texto, img):
                cod, desc_int = asignar_codigo(maestro, art["articulo_prov"], art["descripcion"])
                filas_pdf.append({"pagina": i + 1, "codigo_interno": cod,
                                  "desc_interna": desc_int, **art})
        doc.close()

        if not filas_pdf:
            print(f"[AVISO] Sin albaranes en {os.path.basename(ruta_pdf)} -> se deja sin mover para revisar")
            continue

        # primero calculo el destino, luego muevo, y la ruta del Excel ya apunta al destino
        destino = destino_unico(carpeta_proc, os.path.basename(ruta_pdf))
        for f in filas_pdf:
            f.update({"archivo": os.path.basename(destino), "ruta": os.path.abspath(destino)})
            filas.append(f)
        shutil.move(ruta_pdf, destino)

    escribir_excel(filas)
    return filas

# ============================================================
# EXCEL
# ============================================================
def escribir_excel(filas):
    wb = Workbook(); ws = wb.active; ws.title = "Albaranes"
    cabeceras = ["Código MMAA/MP", "Descripción (maestro)", "Lote", "Cantidad",
                 "Nº Pedido", "Nº Albarán", "Fecha", "Cliente", "Descripción albarán",
                 "Cód. proveedor", "Archivo", "Pág.", "Enlace"]
    AZUL = PatternFill("solid", start_color="1F4E78")
    ROJO = PatternFill("solid", start_color="FFC7CE")
    borde = Border(*[Side(style="thin", color="BFBFBF")]*4)
    ws.append(cabeceras)
    for c in range(1, len(cabeceras)+1):
        cel = ws.cell(row=1, column=c)
        cel.font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        cel.fill = AZUL; cel.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cel.border = borde
    ws.row_dimensions[1].height = 30

    for f in filas:
        ws.append([f["codigo_interno"], f["desc_interna"], f["lote"], f["cantidad"],
                   f["s_pedido"], f["num_albaran"], f["fecha"], f["cliente"], f["descripcion"],
                   f["articulo_prov"], f["archivo"], f["pagina"], "Ver albarán"])
        r = ws.max_row
        for c in range(1, len(cabeceras)+1):
            ws.cell(row=r, column=c).font = Font(name="Arial", size=10)
            ws.cell(row=r, column=c).border = borde
        if not f["codigo_interno"]:                      # sin codigo -> resaltar para revisar
            ws.cell(row=r, column=1).fill = ROJO
        enlace = ws.cell(row=r, column=len(cabeceras))   # enlace al PDF, saltando a la pagina
        enlace.hyperlink = f'{f["ruta"]}#page={f["pagina"]}'
        enlace.font = Font(name="Arial", size=10, color="0563C1", underline="single")

    anchos = [16, 26, 14, 12, 14, 12, 12, 12, 36, 14, 26, 6, 14]
    for i, w in enumerate(anchos, 1):
        ws.column_dimensions[chr(64+i) if i <= 26 else "A"].width = w
    ws.freeze_panes = "A2"
    wb.save(SALIDA_XLSX)

if __name__ == "__main__":
    try:
        _resolver_tesseract()
        res = procesar()
        print(f"\nProceso terminado. Albaranes/articulos detectados: {len(res)}")
        print(f"Excel generado: {SALIDA_XLSX}")
        print(f"Los PDF leidos se han movido a: {CARPETA_PROCESADOS}")
    except Exception:
        import traceback
        print("\n*** Ha ocurrido un error ***")
        traceback.print_exc()
    input("\nPulsa Enter para cerrar...")
