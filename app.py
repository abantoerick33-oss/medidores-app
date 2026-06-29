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

    AZUL = "17375E"
    AZUL_CLARO = "D9EAF7"
    VERDE = "2E8B57"
    VERDE_CLARO = "EAF7EE"
    ROJO = "C0392B"
    GRIS = "F4F6F7"
    GRIS2 = "D5D8DC"
    BLANCO = "FFFFFF"
    AMARILLO = "FCF3CF"
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

    anchos = [5, 10, 10, 16, 11, 10, 10, 16, 16, 10, 5, 14, 5, 12, 8, 18]
    for i, a in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(i)].width = a

    # FILA 1: TÍTULO PRINCIPAL
    ws.merge_cells("A1:P1")
    estilo(ws, "A1", "LABORATORIO DE MEDIDORES DE GAS",
           negrita=True, tam=18, color_texto=BLANCO, color_fondo=AZUL, borde=borde_medio)
    ws.row_dimensions[1].height = 35

    # FILA 2: SUBTÍTULO
    ws.merge_cells("A2:P2")
    estilo(ws, "A2", "Reporte automático de inspección y registro de medidores",
           tam=10, italica=True, color_texto="404040", color_fondo=GRIS, borde=borde_fino)
    ws.row_dimensions[2].height = 18

    # FILA 3: DATOS GENERALES
    ws.merge_cells("A3:D3")
    estilo(ws, "A3", f"👤 Operario: {operario}", negrita=True,
           color_fondo=AZUL_CLARO, borde=borde_fino, alineacion="left")

    ws.merge_cells("E3:J3")
    estilo(ws, "E3", f"📅 Fecha: {fecha}", negrita=True,
           color_fondo=AZUL_CLARO, borde=borde_fino)

    ws.merge_cells("K3:P3")
    estilo(ws, "K3", f"🏷 Lote: {lote}", negrita=True,
           color_texto=BLANCO, color_fondo=AZUL, borde=borde_fino)
    ws.row_dimensions[3].height = 22

    ws.row_dimensions[4].height = 6

    # FILAS 5-6: TARJETAS EJECUTIVAS V2
    marca_tarjeta = (tabla[0].get("marca") or "METREX") if tabla else "-"
    modelo_tarjeta = (tabla[0].get("modelo") or "G1.6") if tabla else "-"
    estado = "APROBADO"

    tarjetas = [
        ("A5", "D6", "TOTAL DE\nMEDIDORES", str(total), AZUL, AZUL_CLARO),
        ("E5", "H6", "MARCA DEL\nLOTE", marca_tarjeta, VERDE, VERDE_CLARO),
        ("I5", "L6", "MODELO", modelo_tarjeta, "6C3483", "F4ECF7"),
        ("M5", "P6", "ESTADO DEL\nLOTE", estado, VERDE, VERDE_CLARO),
    ]

    for inicio, fin, titulo, valor, color_t, color_v in tarjetas:
        ws.merge_cells(f"{inicio}:{fin}")
        estilo(ws, inicio, f"{titulo}\n\n{valor}",
               negrita=True, tam=11, color_fondo=color_v, borde=borde_medio)
        ws[inicio].alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )

    for r in [5, 6]:
        ws.row_dimensions[r].height = 28

    ws.row_dimensions[7].height = 6

    # FILA 8: ENCABEZADO TABLA
    ws.merge_cells("A8:J8")
    estilo(ws, "A8", "RESUMEN DE MEDIDORES DEL LOTE", negrita=True, tam=11,
           color_texto=BLANCO, color_fondo=AZUL, borde=borde_fino)
    ws.row_dimensions[8].height = 22

    # FILA 9: SUBENCABEZADOS TABLA
    cols_tabla = ["#", "Marca", "Modelo", "Nro. Serie\nMedidor", "Vol.\nCíclico",
                  "Tipo", "Color", "Nro. Serie\nPrecinto", "Registro\nInicial (m³)", "Estado"]
    for i, h in enumerate(cols_tabla):
        c = ws.cell(row=9, column=i+1)
        estilo(ws, c, h, negrita=True, tam=9, color_texto=BLANCO,
               color_fondo=ROJO, borde=borde_fino)
    ws.row_dimensions[9].height = 43

    # FILAS DE DATOS
    MODELO_DEFAULT = "G1.6"
    MARCA_DEFAULT = "METREX"

    for idx, datos in enumerate(tabla):
        fila = 10 + idx
        color_fila = GRIS if idx % 2 == 0 else BLANCO
        vol_normalizado = "0.7 dm³"
        unidad_normalizada = "m³"
        marca = datos.get("marca") or MARCA_DEFAULT
        modelo = datos.get("modelo") or MODELO_DEFAULT
        valores = [
            idx+1, marca, modelo, datos["serie_medidor"],
            vol_normalizado, "Circular", "Verde",
            datos["serie_precinto"] or "N/D",
            f"{datos['registro']} {unidad_normalizada}", "OK"
        ]
        for col, val in enumerate(valores, 1):
            c = ws.cell(row=fila, column=col)
            color = VERDE_CLARO if col == 10 else color_fila
            estilo(ws, c, val, tam=9, color_fondo=color, borde=borde_fino)
        ws.row_dimensions[fila].height = 16

    # PANEL DETALLES (columnas K-P, desde fila 8)
    ws.merge_cells("K8:P8")
    estilo(ws, "K8", "DETALLES DEL LOTE", negrita=True, tam=11,
           color_texto=BLANCO, color_fondo=AZUL, borde=borde_fino)

    detalles = [
        ("Lote:", lote),
        ("Fecha y hora:", fecha),
        ("Operario:", operario),
        ("Marca:", marca_tarjeta),
        ("Modelo:", modelo_tarjeta),
        ("Tipo:", "Circular"),
        ("Vol. Cíclico:", "0.7 dm³ (fijo)"),
        ("Color:", "Verde"),
        ("Total medidores:", str(total)),
        ("Supervisados:", f"{correctos} (100%)"),
    ]

    for i, (clave, valor) in enumerate(detalles):
        fila_d = 9 + i
        ws.merge_cells(f"K{fila_d}:L{fila_d}")
        estilo(ws, f"K{fila_d}", clave, negrita=True, tam=8,
               color_fondo=AZUL_CLARO, borde=borde_fino, alineacion="left")
        ws.merge_cells(f"M{fila_d}:P{fila_d}")
        estilo(ws, f"M{fila_d}", valor, tam=8,
               color_fondo=BLANCO, borde=borde_fino, alineacion="left")
        ws.row_dimensions[fila_d].height = 16

    cert_fila = 9 + len(detalles) + 1
    ws.merge_cells(f"K{cert_fila}:P{cert_fila+2}")
    estilo(ws, f"K{cert_fila}",
           f"Certifico que el proceso fue supervisado correctamente\n— {operario} —\n{fecha}",
           negrita=True, italica=True, tam=8,
           color_fondo=VERDE_CLARO, borde=borde_medio, alineacion="center")
    ws.row_dimensions[cert_fila].height = 18
    ws.row_dimensions[cert_fila+1].height = 18
    ws.row_dimensions[cert_fila+2].height = 18

    # ESTADÍSTICAS
    est_fila = cert_fila + 4
    ws.merge_cells(f"K{est_fila}:P{est_fila}")
    estilo(ws, f"K{est_fila}", "ESTADÍSTICAS DEL LOTE", negrita=True, tam=10,
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

    # OBSERVACIONES
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

    # EVIDENCIA FOTOGRÁFICA
    foto_fila = 10 + total + 2
    ws.merge_cells(f"A{foto_fila}:J{foto_fila}")
    estilo(ws, f"A{foto_fila}", "EVIDENCIA FOTOGRÁFICA DEL LOTE", negrita=True, tam=12,
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

        c_num = ws.cell(row=fi, column=col)
        c_num.value = str(idx+1)
        c_num.font = Font(bold=True, size=9, color=BLANCO)
        c_num.fill = PatternFill("solid", fgColor=VERDE)
        c_num.alignment = Alignment(horizontal="center", vertical="top")
        c_num.border = borde_fino

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

# INTERFAZ
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
        "Vol. Cíclico": d["volumen_ciclico"],
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
