import streamlit as st
from google import genai
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import PIL.Image
import json
import time
import io

st.set_page_config(page_title="Registro de Medidores", page_icon="📊", layout="centered")
st.title("📊 Registro de Medidores de Gas")
st.markdown("Sube hasta 12 fotos de medidores y genera el Excel automáticamente.")

API_KEY = st.secrets["GEMINI_API_KEY"]
client = genai.Client(api_key=API_KEY)

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

def procesar_imagen(imagen):
    for intento in range(3):
        try:
            respuesta = client.models.generate_content(
                model="gemini-2.0-flash-lite",
                contents=[prompt, imagen]
            )
            texto = respuesta.text.strip().replace("```json","").replace("```","")
            return json.loads(texto)
        except Exception as e:
            if intento < 2:
                time.sleep(10)
            else:
                return None

def generar_excel(tabla):
    wb = Workbook()
    ws = wb.active
    ws.title = "Medidores"

    azul_oscuro = "1F3864"
    rojo = "C00000"
    blanco = "FFFFFF"
    gris_claro = "F2F2F2"

    borde = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )

    ws.merge_cells("A1:D1")
    ws.merge_cells("E1:H1")

    for celda, texto_enc in [("A1", "Datos del medidor"), ("E1", "Precintos")]:
        c = ws[celda]
        c.value = texto_enc
        c.font = Font(bold=True, color=blanco, size=11)
        c.fill = PatternFill("solid", fgColor=azul_oscuro)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = borde

    subencabezados = ["Marca", "Modelo", "Nro. de serie", "Volumen Cíclico",
                      "Tipo", "Color", "Nro. de serie", "Registro Inicial"]

    for i, titulo in enumerate(subencabezados, 1):
        c = ws.cell(row=2, column=i)
        c.value = titulo
        c.font = Font(bold=True, color=blanco, size=10)
        c.fill = PatternFill("solid", fgColor=rojo)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = borde

    for fila_idx, datos in enumerate(tabla, 3):
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

    anchos = [12, 10, 15, 16, 12, 10, 15, 16]
    for i, ancho in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(i)].width = ancho

    ws.row_dimensions[1].height = 20
    ws.row_dimensions[2].height = 18

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# ─── INTERFAZ ───
fotos = st.file_uploader(
    "📷 Sube las fotos de los medidores (hasta 12)",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True
)

if fotos:
    st.info(f"📂 {len(fotos)} foto(s) seleccionada(s)")

    if st.button("🚀 Procesar medidores", type="primary"):
        tabla = []
        progress = st.progress(0)
        status = st.empty()

        for i, foto in enumerate(fotos[:12]):
            status.text(f"⏳ Procesando medidor {i+1} de {len(fotos)}...")
            imagen = PIL.Image.open(foto)
            datos = procesar_imagen(imagen)

            if datos:
                tabla.append(datos)
                st.success(f"✅ Medidor {i+1}: Serie {datos['serie_medidor']} | Registro {datos['registro']} {datos['unidad']} | Precinto {datos['serie_precinto']}")
            else:
                st.error(f"❌ No se pudo procesar el medidor {i+1}")

            progress.progress((i+1) / len(fotos))
            time.sleep(6)

        status.text("✅ Procesamiento completado.")

        if tabla:
            st.markdown("---")
            st.subheader("📋 Tabla completa del lote")
            st.table([{
                "Marca": d["marca"],
                "Modelo": d["modelo"],
                "Serie": d["serie_medidor"],
                "Registro": f"{d['registro']} {d['unidad']}",
                "Vol. Cíclico": d["volumen_ciclico"],
                "Precinto": d["serie_precinto"]
            } for d in tabla])

            excel = generar_excel(tabla)
            st.download_button(
                label="📥 Descargar Excel",
                data=excel,
                file_name="medidores_lote.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
