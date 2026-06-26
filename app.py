import streamlit as st
from google import genai
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
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

def generar_excel(tabla, fotos_bytes, operario, fecha):
    wb = Workbook()
    ws = wb.active
    ws.title = "Registro de Medidores"

    azul_oscuro = "1F3864"
    rojo = "C00000"
    blanco = "FFFFFF"
    gris_claro = "F2F2F2"
    azul_claro = "D9E1F2"

    borde = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    borde_medio = Border(
        left=Side(style='medium'), right=Side(style='medium'),
        top=Side(style='medium'), bottom=Side(style='medium')
    )

    # Título principal
    ws.merge_cells("A1:H1")
    c = ws["A1"]
    c.value = "LABORATORIO DE MEDIDORES DE GAS"
    c.font = Font(bold=True, color=blanco, size=14)
    c.fill = PatternFill("solid", fgColor=azul_oscuro)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = borde
    ws.row_dimensions[1].height = 28

    # Info operario y fecha
    ws.merge_cells("A2:D2")
    c = ws["A2"]
    c.value = f"Operario: {operario}"
    c.font = Font(bold=True, size=10)
    c.fill = PatternFill("solid", fgColor=azul_claro)
    c.alignment = Alignment(horizontal="left", vertical="center")
    c.border = borde

    ws.merge_cells("E2:H2")
    c = ws["E2"]
    c.value = f"Fecha: {fecha}"
    c.font = Font(bold=True, size=10)
    c.fill = PatternFill("solid", fgColor=azul_claro)
    c.alignment = Alignment(horizontal="right", vertical="center")
    c.border = borde
    ws.row_dimensions[2].height = 20

    # Encabezados de grupo
    ws.merge_cells("A3:D3")
    ws.merge_cells("E3:H3")
    for celda, texto_enc in [("A3", "Datos del medidor"), ("E3", "Precintos")]:
        c = ws[celda]
        c.value = texto_enc
        c.font = Font(bold=True, color=blanco, size=11)
        c.fill = PatternFill("solid", fgColor=azul_oscuro)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = borde
    ws.row_dimensions[3].height = 20

    # Subencabezados
    subencabezados = ["Marca", "Modelo", "Nro. de serie", "Volumen Ciclico",
                      "Tipo", "Color", "Nro. de serie", "Registro Inicial"]
    for i, titulo in enumerate(subencabezados, 1):
        c = ws.cell(row=4, column=i)
        c.value = titulo
        c.font = Font(bold=True, color=blanco, size=10)
        c.fill = PatternFill("solid", fgColor=rojo)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = borde
    ws.row_dimensions[4].height = 18

    # Datos
    for fila_idx, datos in enumerate(tabla, 5):
        fila = [
            datos["marca"], datos["modelo"], datos["serie_medidor"],
            datos["volumen_ciclico"], "Circular", "Verde",
            datos["serie_precinto"], f"{datos['registro']} {datos['unidad']}"
        ]
        fill = PatternFill("solid", fgColor=gris_claro) if fila_idx % 2 == 0 else None
        for col_idx, valor in enumerate(fila, 1):
            c = ws.cell(row=fila_idx, column=col_idx)
            c.value = valor
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = borde
            if fill:
                c.fill = fill

    # Firma de certificación
    firma_fila = len(tabla) + 6
    ws.merge_cells(f"A{firma_fila}:H{firma_fila}")
    c = ws[f"A{firma_fila}"]
    c.value = f"Certifico que el proceso fue supervisado correctamente — {operario} — {fecha}"
    c.font = Font(bold=True, italic=True, size=10, color=azul_oscuro)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.fill = PatternFill("solid", fgColor=azul_claro)
    c.border = borde
    ws.row_dimensions[firma_fila].height = 20

    # Anchos de columna
    anchos = [12, 10, 15, 16, 12, 10, 15, 16]
    for i, ancho in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(i)].width = ancho

    # ─── SECCIÓN DE IMÁGENES (misma hoja, debajo) ───
    img_inicio_fila = firma_fila + 3

    # Título sección imágenes
    ws.merge_cells(f"A{img_inicio_fila}:H{img_inicio_fila}")
    c = ws[f"A{img_inicio_fila}"]
    c.value = "EVIDENCIA FOTOGRÁFICA DE MEDIDORES"
    c.font = Font(bold=True, color=blanco, size=12)
    c.fill = PatternFill("solid", fgColor=azul_oscuro)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = borde
    ws.row_dimensions[img_inicio_fila].height = 24

    # Imágenes en cuadrícula 3 por fila
    IMG_ANCHO = 160
    IMG_ALTO = 120
    COLS_POR_FILA = 3
    columnas_img = ["A", "C", "E"]
    alto_fila_img = 95
    alto_fila_label = 18

    for idx, (foto_bytes, datos) in enumerate(zip(fotos_bytes, tabla), 0):
        fila_bloque = idx // COLS_POR_FILA
        col_bloque = idx % COLS_POR_FILA

        fila_img = img_inicio_fila + 2 + fila_bloque * 8
        fila_label = fila_img + 6

        ws.row_dimensions[fila_img].height = alto_fila_img
        ws.row_dimensions[fila_label].height = alto_fila_label

        col_letra = columnas_img[col_bloque]

        # Número y serie debajo de imagen
        ws.merge_cells(f"{col_letra}{fila_label}:{chr(ord(col_letra)+1)}{fila_label}")
        c = ws[f"{col_letra}{fila_label}"]
        c.value = f"#{idx+1} — Serie: {datos['serie_medidor']}"
        c.font = Font(bold=True, size=9, color=azul_oscuro)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.fill = PatternFill("solid", fgColor=azul_claro)
        c.border = borde

        try:
            img = PIL.Image.open(io.BytesIO(foto_bytes))
            img.thumbnail((IMG_ANCHO, IMG_ALTO))
            img_buffer = io.BytesIO()
            img.save(img_buffer, format="PNG")
            img_buffer.seek(0)
            xl_img = XLImage(img_buffer)
            xl_img.width = IMG_ANCHO
            xl_img.height = IMG_ALTO
            ws.add_image(xl_img, f"{col_letra}{fila_img}")
        except Exception:
            ws[f"{col_letra}{fila_img}"] = "imagen no disponible"

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# ─── INTERFAZ ───
st.markdown("---")
operario = st.selectbox("Selecciona el operario:", OPERARIOS)
peru = timezone(timedelta(hours=-5))
fecha = datetime.now(peru).strftime("%d/%m/%Y %H:%M")
st.info(f"Fecha y hora: {fecha}")

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
        nombre_archivo = f"Medidores_{fecha.replace('/', '-').replace(':', '-').replace(' ', '_')}.xlsx"
        excel = generar_excel(tabla, fotos_bytes, operario, fecha)

        st.download_button(
            label="📥 Descargar Excel",
            data=excel,
            file_name=nombre_archivo,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("Debes confirmar la supervision antes de descargar el Excel.")
