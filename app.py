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
from datetime import datetime
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from google.oauth2 import service_account

st.set_page_config(page_title="Registro de Medidores", page_icon="📊", layout="centered")
st.title("📊 Registro de Medidores de Gas")
st.markdown("Sube hasta 12 fotos de medidores y genera el Excel automáticamente.")

API_KEYS = [
    st.secrets["GEMINI_API_KEY_1"],
    st.secrets["GEMINI_API_KEY_2"],
    st.secrets["GEMINI_API_KEY_3"],
]

OPERARIOS = ["Joseph Erik Abanto Guerra", "Marco David Rodríguez Valencia"]
FOLDER_ID = "1hPilAiAhOVBF2GIh6Y4WldLPks_wcqJI"

prompt = """Analiza esta imagen de un medidor de gas.
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
- serie_medidor: número de serie del medidor (ej: 3809043)
- registro: lectura actual del contador con todos los dígitos, usar SIEMPRE coma como separador decimal, nunca punto (ej: 00954,095 no 00954.095)
- unidad: unidad de medida del registro (ej: m³)
- volumen_ciclico: valor V indicado en la placa (ej: 0.7 dm³)
- serie_precinto: número del precinto de seguridad si es visible (ej: G0359688)
Si no puedes leer algún dato, pon null."""

# Inicializar sesión
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

def get_drive_service():
    creds_dict = json.loads(st.secrets["GOOGLE_DRIVE_CREDENTIALS"])
    creds = service_account.Credentials.from_service_account_info(
        creds_dict,
        scopes=["https://www.googleapis.com/auth/drive"]
    )
    return build("drive", "v3", credentials=creds)

def subir_a_drive(buffer, nombre_archivo):
    try:
        service = get_drive_service()
        media = MediaIoBaseUpload(buffer, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        file_metadata = {"name": nombre_archivo, "parents": [FOLDER_ID]}
        service.files().create(body=file_metadata, media_body=media).execute()
        return True
    except Exception as e:
        st.warning(f"No se pudo guardar en Drive: {e}")
        return False

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
    ws.title = "Datos"

    azul_oscuro = "1F3864"
    rojo = "C00000"
    blanco = "FFFFFF"
    gris_claro = "F2F2F2"

    borde = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    ws.merge_cells("A1:H1")
    c = ws["A1"]
    c.value = f"Laboratorio de Medidores — Operario: {operario} — Fecha: {fecha}"
    c.font = Font(bold=True, color=blanco, size=11)
    c.fill = PatternFill("solid", fgColor=azul_oscuro)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = borde
    ws.row_dimensions[1].height = 20

    ws.merge_cells("A2:D2")
    ws.merge_cells("E2:H2")
    for celda, texto_enc in [("A2", "Datos del medidor"), ("E2", "Precintos")]:
        c = ws[celda]
        c.value = texto_enc
        c.font = Font(bold=True, color=blanco, size=11)
        c.fill = PatternFill("solid", fgColor=azul_oscuro)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = borde
    ws.row_dimensions[2].height = 18

    subencabezados = ["Marca", "Modelo", "Nro. de serie", "Volumen Cíclico",
                      "Tipo", "Color", "Nro. de serie", "Registro Inicial"]
    for i, titulo in enumerate(subencabezados, 1):
        c = ws.cell(row=3, column=i)
        c.value = titulo
        c.font = Font(bold=True, color=blanco, size=10)
        c.fill = PatternFill("solid", fgColor=rojo)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = borde
    ws.row_dimensions[3].height = 18

    for fila_idx, datos in enumerate(tabla, 4):
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

    ultima_fila = len(tabla) + 5
    ws.merge_cells(f"A{ultima_fila}:H{ultima_fila}")
    c = ws[f"A{ultima_fila}"]
    c.value = f"Certifico que el proceso fue supervisado correctamente — {operario} — {fecha}"
    c.font = Font(bold=True, italic=True, size=10)
    c.alignment = Alignment(horizontal="center")

    anchos = [12, 10, 15, 16, 12, 10, 15, 16]
    for i, ancho in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(i)].width = ancho

    # Hoja de imágenes
    ws2 = wb.create_sheet(title="Imágenes")
    ws2.column_dimensions["A"].width = 5
    ws2.column_dimensions["B"].width = 40
    ws2.column_dimensions["C"].width = 20

    for col, titulo in [("A1", "N°"), ("B1", "Foto del medidor"), ("C1", "Serie")]:
        ws2[col] = titulo
        ws2[col].font = Font(bold=True, color=blanco)
        ws2[col].fill = PatternFill("solid", fgColor=azul_oscuro)
        ws2[col].alignment = Alignment(horizontal="center")

    for idx, (foto_bytes, datos) in enumerate(zip(fotos_bytes, tabla), 1):
        fila = 2 + (idx - 1) * 22
        ws2.row_dimensions[fila].height = 150
        ws2.cell(row=fila, column=1).value = idx
        ws2.cell(row=fila, column=3).value = datos["serie_medidor"]
        try:
            img = PIL.Image.open(io.BytesIO(foto_bytes))
            img.thumbnail((280, 200))
            img_buffer = io.BytesIO()
            img.save(img_buffer, format="PNG")
            img_buffer.seek(0)
            xl_img = XLImage(img_buffer)
            xl_img.width = 280
            xl_img.height = 200
            ws2.add_image(xl_img, f"B{fila}")
        except Exception:
            ws2.cell(row=fila, column=2).value = "imagen no disponible"

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# ─── INTERFAZ ───
st.markdown("---")
operario = st.selectbox("👤 Selecciona el operario:", OPERARIOS)
fecha = datetime.now().strftime("%d/%m/%Y %H:%M")
st.info(f"📅 Fecha y hora: {fecha}")

st.markdown("---")
fotos = st.file_uploader(
    "📷 Sube las fotos de los medidores (hasta 12)",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True
)

if fotos:
    st.info(f"📂 {len(fotos)} foto(s) seleccionada(s)")

    if st.button("🚀 Procesar medidores", type="primary"):
        st.session_state.tabla = []
        st.session_state.fotos_bytes = []
        st.session_state.procesado = False
        st.session_state.operario_guardado = operario
        st.session_state.fecha_guardada = fecha

        progress = st.progress(0)
        status = st.empty()

        for i, foto in enumerate(fotos[:12]):
            status.text(f"⏳ Procesando medidor {i+1} de {len(fotos)}...")
            foto.seek(0)
            foto_bytes = foto.read()
            st.session_state.fotos_bytes.append(foto_bytes)
            imagen = PIL.Image.open(io.BytesIO(foto_bytes))
            datos = procesar_imagen(imagen)

            if datos:
                st.session_state.tabla.append(datos)
                st.success(f"✅ Medidor {i+1}: Serie {datos['serie_medidor']} | Registro {datos['registro']} {datos['unidad']} | Precinto {datos['serie_precinto']}")
            else:
                st.error(f"❌ No se pudo procesar el medidor {i+1}")

            progress.progress((i+1) / len(fotos))
            time.sleep(5)

        st.session_state.procesado = True
        status.text("✅ Procesamiento completado.")

# Mostrar resultados si ya se procesó
if st.session_state.procesado and st.session_state.tabla:
    tabla = st.session_state.tabla
    fotos_bytes = st.session_state.fotos_bytes
    operario = st.session_state.operario_guardado
    fecha = st.session_state.fecha_guardada

    st.markdown("---")
    st.subheader("📋 Tabla completa del lote")
    st.table([{
        "N°": i+1,
        "Marca": d["marca"],
        "Modelo": d["modelo"],
        "Serie": d["serie_medidor"],
        "Registro": f"{d['registro']} {d['unidad']}",
        "Vol. Cíclico": d["volumen_ciclico"],
        "Precinto": d["serie_precinto"]
    } for i, d in enumerate(tabla)])

    st.markdown("---")
    confirmado = st.checkbox(f"✅ Confirmo que el proceso fue supervisado correctamente por {operario}")

    if confirmado:
        nombre_archivo = f"Medidores_{fecha.replace('/', '-').replace(':', '-').replace(' ', '_')}.xlsx"
        excel = generar_excel(tabla, fotos_bytes, operario, fecha)

        col1, col2 = st.columns(2)
        with col1:
            st.download_button(
                label="📥 Descargar Excel",
                data=excel,
                file_name=nombre_archivo,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with col2:
            if st.button("☁️ Guardar en Drive"):
                excel.seek(0)
                with st.spinner("Guardando en Google Drive..."):
                    if subir_a_drive(excel, nombre_archivo):
                        st.success("✅ Guardado en Google Drive correctamente.")
    else:
        st.warning("⚠️ Debes confirmar la supervisión antes de descargar el Excel.")
