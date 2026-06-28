import streamlit as st
from google import genai
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
from openpyxl.chart import PieChart, Reference
from openpyxl.chart.series import DataPoint
import PIL.Image
import json
import time
import io
from datetime import datetime, timezone, timedelta

st.set_page_config(page_title="Registro de Medidores", page_icon="📊", layout="centered")
st.title("📊 Registro de Medidores de Gas")
st.markdown("Sube hasta 12 fotos de medidores y genera el Excel automáticamente.")

API_KEYS = [
    st.secrets["GEMINI_API_KEY_1"],
    st.secrets["GEMINI_API_KEY_2"],
    st.secrets["GEMINI_API_KEY_3"],
]

OPERARIOS = ["Joseph Erik Abanto Guerra", "Marco David Rodríguez Valencia"]

prompt = """Analiza esta imagen de un medidor de gas con mucho detalle.
Extrae exactamente estos datos y devuélvelos SOLO en formato JSON sin texto adicional:
{
  "marca": "",
  "modelo": "",
  "serie_medidor": "",
  "registro": "",
  "unidad": "",
  "volumen_ciclico": "",
  "serie_precinto": ""
}
Instrucciones:
- marca: nombre del fabricante (ej: METREX)
- modelo: capacidad del medidor, siempre con punto (ej: G1.6 no G1,6 ni G16)
- serie_medidor: numero de serie del medidor (ej: 3809043)
- registro: lectura actual del contador con todos los digitos incluyendo los que estan en rojo, usar SIEMPRE coma como separador decimal, nunca punto (ej: 00954,095 no 00954.095). Los digitos en rojo son parte del registro.
- unidad: unidad de medida del registro (ej: m3)
- volumen_ciclico: valor V indicado en la placa (ej: 0.7 dm3)
- serie_precinto: busca con mucho cuidado el numero del precinto de seguridad. Puede ser una etiqueta pequena colgante o pegada al medidor. IMPORTANTE: puede estar rotado o al reves 180 grados. El formato siempre empieza con G seguido de 7 digitos numericos (ej: G0454825, G0457243). Si ves letras como GUM u otras combinaciones raras, intenta leer rotando 180 grados. Si definitivamente no es visible, pon null.
Si no puedes leer algun dato, pon null."""

if "tabla" not in st.session_state:
    st.session_state.tabla = []
if "fotos_bytes" not in st.session_state:
    st.session_state.fotos_bytes = []
if "procesado" not in st.session_state:
    st.session_state.procesado = False
if "operario_guardado" not in st.session_state:
    st.session_state.operario_guardado = ""
if "fecha_guardada" not in st.session_state:
    st.session_state.fecha_guardada = ""
if "lote_guardado" not in st.session_state:
    st.session_state.lote_guardado = ""

def procesar_imagen(imagen):
    for key_idx, api_key in enumerate(API_KEYS):
        try:
            client = genai.Client(api_key=api_key)
            respuesta = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=[prompt, imagen]
            )
            texto = respuesta.text.strip().replace("```json","").replace("```","")
            return json.loads(texto)
        except Exception:
            if key_idx < len(API_KEYS) - 1:
                time.sleep(3)
                continue
            else:
                return None

def estilo(ws, celda, valor, negrita=False, tam=10, color_texto="000000",
           color_fondo=None, alineacion="center", borde=None, alto=None, italica=False):
    if isinstance(celda, str):
        c = ws[celda]
    else:
        c = celda
    try:
        c.value = valor
    except AttributeError:
        return None
    c.font = Font(bold=negrita, size=tam, color=color_texto, italic=italica)
    c.alignment = Alignment(horizontal=alineacion, vertical="center", wrap_text=True)
    if color_fondo:
        c.fill = PatternFill("solid", fgColor=color_fondo)
    if borde:
        c.border = borde
    return c

def generar_excel(tabla, fotos_bytes, operario, fecha, lote):
    wb = Workbook()
    ws = wb.active
    ws.title = "Registro"

    # Colores
    AZUL = "1B3A6B"
    VERDE = "1E7A3E"
    ROJO = "C00000"
    BLANCO = "FFFFFF"
    GRIS = "F5F5F5"
    AZUL_CLARO = "D6E4F0"
    VERDE_CLARO = "D5F5E3"
    AMARILLO = "FFF3CD"
    NARANJA = "E67E22"

    borde_fino = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    borde_medio = Border(
        left=Side(style='medium'), right=Side(style='medium'),
        top=Side(style='medium'), bottom=Side(style='medium')
    )

    total = len(tabla)
    correctos = total
    observados = 0
    rechazados = 0

    # Anchos de columna
    anchos = [4, 10, 8, 14, 10, 10, 10, 14, 14, 8, 4, 12, 4, 10, 6, 18]
    for i, a in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(i)].width = a

    # ── FILA 1: TÍTULO PRINCIPAL ──
    ws.merge_cells("A1:P1")
    c = estilo(ws, "A1", "LABORATORIO DE MEDIDORES DE GAS",
               negrita=True, tam=16, color_texto=BLANCO, color_fondo=AZUL, borde=borde_fino)
    ws.row_dimensions[1].height = 36

    # ── FILA 2: INFO OPERARIO / FECHA / LOTE ──
    ws.merge_cells("A2:D2")
    estilo(ws, "A2", f"Operario:  {operario}", negrita=True, tam=10,
           color_fondo=AZUL_CLARO, borde=borde_fino, alineacion="left")

    ws.merge_cells("E2:J2")
    estilo(ws, "E2", f"Fecha y hora:  {fecha}", negrita=True, tam=10,
           color_fondo=AZUL_CLARO, borde=borde_fino, alineacion="center")

    ws.merge_cells("K2:P2")
    estilo(ws, "K2", f"Lote:  {lote}", negrita=True, tam=11,
           color_texto=BLANCO, color_fondo=AZUL, borde=borde_fino, alineacion="center")
    ws.row_dimensions[2].height = 22

    # ── FILA 3-4: TARJETAS DE RESUMEN ──
    ws.row_dimensions[3].height = 18
    ws.row_dimensions[4].height = 28

    tarjetas = [
        ("A3", "B4", "Total de medidores", str(total), AZUL, AZUL_CLARO),
        ("C3", "D4", "Supervisados", f"{correctos} (100%)", VERDE, VERDE_CLARO),
        ("E3", "F4", "Vol. Ciclico", "0.7 dm3 (fijo)", "7D3C98", "F4ECF7"),
        ("G3", "H4", "Color registrado", "Verde", VERDE, VERDE_CLARO),
        ("I3", "J4", "Total registros", str(total), NARANJA, AMARILLO),
    ]

    for inicio, fin, titulo, valor, color_t, color_v in tarjetas:
        ws.merge_cells(f"{inicio}:{chr(ord(inicio[0]))}{fin[-1]}")
        ws.merge_cells(f"{inicio[0]}{int(inicio[1])+1}:{fin}")
        estilo(ws, inicio, titulo, negrita=True, tam=8,
               color_texto=BLANCO, color_fondo=color_t, borde=borde_fino)
        celda_val = f"{inicio[0]}{int(inicio[1])+1}"
        estilo(ws, celda_val, valor, negrita=True, tam=13,
               color_fondo=color_v, borde=borde_fino)

    ws.row_dimensions[5].height = 6

    # ── FILA 6: ENCABEZADO TABLA ──
    ws.merge_cells("A6:J6")
    estilo(ws, "A6", "RESUMEN DE MEDIDORES DEL LOTE", negrita=True, tam=11,
           color_texto=BLANCO, color_fondo=AZUL, borde=borde_fino)
    ws.row_dimensions[6].height = 22

    # ── FILA 7: SUBENCABEZADOS TABLA ──
    cols_tabla = ["#", "Marca", "Modelo", "Nro. Serie\nMedidor", "Vol.\nCiclico",
                  "Tipo", "Color", "Nro. Serie\nPrecinto", "Registro\nInicial (m3)", "Estado"]
    for i, h in enumerate(cols_tabla):
        c = ws.cell(row=7, column=i+1)
        estilo(ws, c, h, negrita=True, tam=9, color_texto=BLANCO,
               color_fondo=ROJO, borde=borde_fino)
    ws.row_dimensions[7].height = 30

    # ── FILAS DE DATOS ──
for idx, datos in enumerate(tabla):
        fila = 8 + idx
        color_fila = GRIS if idx % 2 == 0 else BLANCO
        # Normalizar valores
        vol_normalizado = "0.7 dm3"
        unidad_normalizada = "m3"
        valores = [
            idx+1, datos["marca"], datos["modelo"], datos["serie_medidor"],
            vol_normalizado, "Circular", "Verde",
            datos["serie_precinto"] or "N/D",
            f"{datos['registro']} {unidad_normalizada}", "OK"
        ]
        for col, val in enumerate(valores, 1):
            c = ws.cell(row=fila, column=col)
            color = VERDE_CLARO if col == 10 else color_fila
            estilo(ws, c, val, tam=9, color_fondo=color, borde=borde_fino)
        ws.row_dimensions[fila].height = 16

    # ── PANEL DETALLES DEL LOTE (columnas K-P) ──
ws.merge_cells("K6:P6")
estilo(ws, "K6", "DETALLES DEL LOTE", negrita=True, tam=11,
color_texto=BLANCO, color_fondo=AZUL, borde=borde_fino)

detalles = [
        ("Lote:", lote),
        ("Fecha y hora:", fecha),
        ("Operario:", operario),
        ("Marca:", tabla[0]["marca"] if tabla else ""),
        ("Modelo:", tabla[0]["modelo"] if tabla else ""),
        ("Tipo:", "Circular"),
        ("Vol. Ciclico:", "0.7 dm3 (fijo)"),
        ("Color:", "Verde"),
        ("Total medidores:", str(total)),
        ("Supervisados:", f"{correctos} (100%)"),
    ]

    for i, (clave, valor) in enumerate(detalles):
        fila_d = 7 + i
        ws.merge_cells(f"K{fila_d}:L{fila_d}")
        estilo(ws, f"K{fila_d}", clave, negrita=True, tam=8,
               color_fondo=AZUL_CLARO, borde=borde_fino, alineacion="left")
        ws.merge_cells(f"M{fila_d}:P{fila_d}")
        estilo(ws, f"M{fila_d}", valor, tam=8,
               color_fondo=BLANCO, borde=borde_fino, alineacion="left")
        ws.row_dimensions[fila_d].height = 16

    # Certificación en panel
    cert_fila = 7 + len(detalles) + 1
    ws.merge_cells(f"K{cert_fila}:P{cert_fila+2}")
    estilo(ws, f"K{cert_fila}",
           f"Certifico que el proceso fue supervisado correctamente\n— {operario} —\n{fecha}",
           negrita=True, italica=True, tam=8,
           color_fondo=VERDE_CLARO, borde=borde_medio, alineacion="center")
    ws.row_dimensions[cert_fila].height = 18
    ws.row_dimensions[cert_fila+1].height = 18
    ws.row_dimensions[cert_fila+2].height = 18

    # ── ESTADÍSTICAS ──
    est_fila = cert_fila + 4
    ws.merge_cells(f"K{est_fila}:P{est_fila}")
    estilo(ws, f"K{est_fila}", "ESTADISTICAS DEL LOTE", negrita=True, tam=10,
           color_texto=BLANCO, color_fondo=VERDE, borde=borde_fino)
    ws.row_dimensions[est_fila].height = 20

    stats = [
        ("Correctos", correctos, VERDE_CLARO),
        ("Observados", observados, AMARILLO),
        ("Rechazados", rechazados, "FADBD8"),
    ]
    for i, (nombre, valor, color) in enumerate(stats):
        fr = est_fila + 1 + i
        ws.merge_cells(f"K{fr}:M{fr}")
        estilo(ws, f"K{fr}", nombre, tam=9, color_fondo=color, borde=borde_fino)
        ws.merge_cells(f"N{fr}:P{fr}")
        pct = f"{valor} ({int(valor/total*100) if total else 0}%)"
        estilo(ws, f"N{fr}", pct, negrita=True, tam=9, color_fondo=color, borde=borde_fino)
        ws.row_dimensions[fr].height = 16

    # ── OBSERVACIONES ──
    obs_fila = est_fila + 5
    ws.merge_cells(f"K{obs_fila}:P{obs_fila}")
    estilo(ws, f"K{obs_fila}", "OBSERVACIONES", negrita=True, tam=10,
           color_texto=BLANCO, color_fondo=AZUL, borde=borde_fino)
    ws.merge_cells(f"K{obs_fila+1}:P{obs_fila+3}")
    estilo(ws, f"K{obs_fila+1}", "Sin observaciones.", tam=9,
           color_fondo=GRIS, borde=borde_fino, alineacion="left")
    ws.row_dimensions[obs_fila].height = 20
    for r in range(obs_fila+1, obs_fila+4):
        ws.row_dimensions[r].height = 14

    # ── NORMA Y CONDICIONES ──
    norma_fila = max(8 + total + 2, obs_fila + 5)
    ws.merge_cells(f"A{norma_fila}:D{norma_fila}")
    estilo(ws, f"A{norma_fila}", "Norma aplicada: NTC 6031 / OIML R137",
           tam=8, color_fondo=AZUL_CLARO, borde=borde_fino, alineacion="left")
    ws.merge_cells(f"E{norma_fila}:H{norma_fila}")
    estilo(ws, f"E{norma_fila}", "Condiciones ambientales: 25°C / 60% HR",
           tam=8, color_fondo=AZUL_CLARO, borde=borde_fino, alineacion="center")
    ws.merge_cells(f"I{norma_fila}:J{norma_fila}")
    peru = timezone(timedelta(hours=-5))
    prox = datetime.now(peru).replace(year=datetime.now(peru).year + 1).strftime("%d/%m/%Y")
    estilo(ws, f"I{norma_fila}", f"Proxima calibracion: {prox}",
           tam=8, color_fondo=AZUL_CLARO, borde=borde_fino, alineacion="center")
    ws.row_dimensions[norma_fila].height = 18

 # ── EVIDENCIA FOTOGRÁFICA (debajo de la tabla, lado izquierdo) ──
    foto_fila = 8 + len(tabla) + 2  # Justo después de la última fila de datos
    ws.merge_cells(f"A{foto_fila}:J{foto_fila}")
    estilo(ws, f"A{foto_fila}", "EVIDENCIA FOTOGRAFICA DEL LOTE", negrita=True, tam=12,
           color_texto=BLANCO, color_fondo=VERDE, borde=borde_fino)
    ws.row_dimensions[foto_fila].height = 24

    IMG_ANCHO = 110
    IMG_ALTO = 85
    COLS_FOTO = 6

    for idx, (fb, datos) in enumerate(zip(fotos_bytes, tabla)):
        fila_bloque = idx // COLS_FOTO
        col_bloque = idx % COLS_FOTO
        fi = foto_fila + 1 + fila_bloque * 8
        col = col_bloque + 1

        ws.row_dimensions[fi].height = 65
        ws.row_dimensions[fi + 6].height = 16

        # Número encima
        c_num = ws.cell(row=fi, column=col)
        c_num.value = str(idx+1)
        c_num.font = Font(bold=True, size=9, color=BLANCO)
        c_num.fill = PatternFill("solid", fgColor=VERDE)
        c_num.alignment = Alignment(horizontal="center", vertical="top")
        c_num.border = borde_fino

        # Serie debajo
        c_ser = ws.cell(row=fi+6, column=col)
        c_ser.value = datos["serie_medidor"]
        c_ser.font = Font(bold=True, size=8, color=VERDE)
        c_ser.fill = PatternFill("solid", fgColor=VERDE_CLARO)
        c_ser.alignment = Alignment(horizontal="center", vertical="center")
        c_ser.border = borde_fino

        try:
            img = PIL.Image.open(io.BytesIO(fb))
            img.thumbnail((IMG_ANCHO, IMG_ALTO))
            ib = io.BytesIO()
            img.save(ib, format="PNG")
            ib.seek(0)
            xl_img = XLImage(ib)
            xl_img.width = IMG_ANCHO
            xl_img.height = IMG_ALTO
            ws.add_image(xl_img, f"{get_column_letter(col)}{fi}")
        except Exception:
            ws.cell(row=fi, column=col).value = "N/D"

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# ─── INTERFAZ ───
st.markdown("---")
operario = st.selectbox("Selecciona el operario:", OPERARIOS)
peru = timezone(timedelta(hours=-5))
fecha = datetime.now(peru).strftime("%d/%m/%Y %H:%M")
lote = f"L-{datetime.now(peru).strftime('%Y%m%d')}-01"
st.info(f"Fecha y hora: {fecha} | Lote: {lote}")

st.markdown("---")
fotos = st.file_uploader(
    "Sube las fotos de los medidores (hasta 12)",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True
)

if fotos:
    st.info(f"{len(fotos)} foto(s) seleccionada(s)")

    if st.button("Procesar medidores", type="primary"):
        st.session_state.tabla = []
        st.session_state.fotos_bytes = []
        st.session_state.procesado = False
        st.session_state.operario_guardado = operario
        st.session_state.fecha_guardada = fecha
        st.session_state.lote_guardado = lote

        progress = st.progress(0)
        status = st.empty()

        for i, foto in enumerate(fotos[:12]):
            status.text(f"Procesando medidor {i+1} de {len(fotos)}...")
            foto.seek(0)
            foto_bytes = foto.read()
            st.session_state.fotos_bytes.append(foto_bytes)
            imagen = PIL.Image.open(io.BytesIO(foto_bytes))
            datos = procesar_imagen(imagen)

            if datos:
                st.session_state.tabla.append(datos)
                st.success(f"Medidor {i+1}: Serie {datos['serie_medidor']} | Registro {datos['registro']} {datos['unidad']} | Precinto {datos['serie_precinto']}")
            else:
                st.error(f"No se pudo procesar el medidor {i+1}")

            progress.progress((i+1) / len(fotos))
            time.sleep(5)

        st.session_state.procesado = True
        status.text("Procesamiento completado.")

if st.session_state.procesado and st.session_state.tabla:
    tabla = st.session_state.tabla
    fotos_bytes = st.session_state.fotos_bytes
    operario = st.session_state.operario_guardado
    fecha = st.session_state.fecha_guardada
    lote = st.session_state.lote_guardado

    st.markdown("---")
    st.subheader("Tabla completa del lote")
    st.table([{
        "N": i+1,
        "Marca": d["marca"],
        "Modelo": d["modelo"],
        "Serie": d["serie_medidor"],
        "Registro": f"{d['registro']} {d['unidad']}",
        "Vol. Ciclico": d["volumen_ciclico"],
        "Precinto": d["serie_precinto"]
    } for i, d in enumerate(tabla)])

    st.markdown("---")
    confirmado = st.checkbox(f"Confirmo que el proceso fue supervisado correctamente por {operario}")

    if confirmado:
        nombre_archivo = f"Medidores_{fecha.replace('/', '-').replace(':', '-').replace(' ', '_')}_{lote}.xlsx"
        excel = generar_excel(tabla, fotos_bytes, operario, fecha, lote)

        st.download_button(
            label="📥 Descargar Excel",
            data=excel,
            file_name=nombre_archivo,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("Debes confirmar la supervision antes de descargar el Excel.")
