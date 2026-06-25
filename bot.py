"""
RecorridosIA — Bot completo en un solo archivo
@RecorridosIA_bot

Variables de entorno (configurar en Render):
    BOT_TOKEN       = token del bot de Telegram
    GEMINI_API_KEY  = clave de API de Google Gemini AI
    MAPILLARY_TOKEN = token de acceso a Mapillary
    TOTP_SECRET     = clave secreta para acceso al bot (2FA)
"""

import os
import io
import hmac
import struct
import time
import base64
import hashlib
import json
import logging
import httpx
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from telegram import Update, ReplyKeyboardMarkup, ReplyKeyboardRemove
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    ConversationHandler, ContextTypes, filters
)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

BOT_TOKEN       = os.getenv("BOT_TOKEN")
GEMINI_API_KEY  = os.getenv("GEMINI_API_KEY")
MAPILLARY_TOKEN = os.getenv("MAPILLARY_TOKEN")
TOTP_SECRET     = os.getenv("TOTP_SECRET")

GEMINI_URL = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={GEMINI_API_KEY}"


import threading
from http.server import HTTPServer, BaseHTTPRequestHandler
import urllib.request

class PingHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b"RecorridosIA OK")
    def log_message(self, format, *args):
        pass  # silenciar logs del servidor

def keep_alive():
    """Servidor web para que Render no mate el proceso."""
    port = int(os.getenv("PORT", 8080))
    server = HTTPServer(("0.0.0.0", port), PingHandler)
    logger.info(f"🌐 Servidor web en puerto {port}")
    server.serve_forever()

USUARIOS_AUTENTICADOS = set()

# ── Estados ──────────────────────────────────────────────────────────────────
(
    ESPERANDO_TOTP, MENU_PRINCIPAL,
    NOMBRE_RUTA, CODIGO_CUADRILLA, NODO_INICIAL, NODO_FINAL,
    LIDER, AYUDANTE, COORDINADOR, PLACA, DISTANCIA,
    NOVEDADES_AUTO, TAREA_PENDIENTE, FOTO_ANTES, FOTO_DESPUES,
    OBSERVACIONES, PREGUNTA_MANGAS, PREGUNTA_HILOS,
    MANGA_NOMBRE, MANGA_COORDS, MANGA_OBS,
    HILO_ODF, HILO_DATOS,
) = range(23)

# ── Novedades Telconet ────────────────────────────────────────────────────────
REMEDIOS = {
    "VEGETACIÓN SOBRE FIBRA/MANGA.": "REALIZAR LA PODA O RETIRO DE VEGETACIÓN QUE COMPROMETA LA INTEGRIDAD O SEGURIDAD DEL CABLE.",
    "HERRAJES EN MAL ESTADO.": "REALIZAR EL REEMPLAZO INMEDIATO DEL HERRAJE AFECTADO.",
    "POSTES EN MAL ESTADO.": "DOCUMENTAR Y REPORTAR PARA GESTIONAR EL REEMPLAZO DEL POSTE.",
    "POSTES INCLINADOS.": "DOCUMENTAR Y REPORTAR PARA GESTIONAR EL APLOME DEL POSTE.",
    "MANGAS SUELTAS.": "ASEGURAR LA MANGA AL POSTE EN CONFIGURACIÓN TIPO FIGURA 8.",
    "MANGAS ABIERTAS/DAÑADAS.": "REEMPLAZAR TAPAS Y SELLOS GARANTIZANDO EL CIERRE HERMÉTICO.",
    "CABLE LASTIMADO.": "DOCUMENTAR E INFORMAR PARA PROGRAMAR EL CAMBIO DEL TRAMO.",
    "DOCUMENTACIÓN UNIFILAR DE HILOS.": "DOCUMENTAR O SOLICITAR PROGRAMACIÓN DE TRABAJO; UTILIZAR SEGUIDOR DE SEÑAL.",
    "CRUCES DE VÍAS BAJOS.": "AJUSTAR LA ALTURA DEL CABLE A LA DISTANCIA REGLAMENTARIA.",
    "POZO SIN TAPA O EN MAL ESTADO.": "SOLICITAR TRABAJOS DE OBRA CIVIL PARA SU CORRECCIÓN.",
    "LÍNEA ELÉCTRICA EN MAL ESTADO.": "DOCUMENTAR Y SOLICITAR REPORTE AL ÁREA DE REGULATORIO.",
    "AMPLIACIÓN DE VÍA.": "DOCUMENTAR Y COORDINAR MEDIDAS DE MITIGACIÓN CON EL COORDINADOR DE FO.",
    "ELEMENTOS SIN ETIQUETAS ACRÍLICAS.": "VERIFICAR Y COLOCAR ETIQUETA ACRÍLICA CON EL CÓDIGO DE RUTA.",
    "RIESGO DE DERRUMBE O DESLAVE.": "DOCUMENTAR Y SOLICITAR REUBICACIÓN DEL RECORRIDO DEL CABLE.",
    "RIESGO DE INUNDACIONES.": "DOCUMENTAR Y SOLICITAR REUBICACIÓN DEL RECORRIDO DEL CABLE.",
    "RIESGO DE INCENDIO.": "DOCUMENTAR Y SOLICITAR REUBICACIÓN DEL RECORRIDO DEL CABLE.",
    "FALTA DE HERRAJES.": "INSTALAR HERRAJES CONFORME A LA NORMATIVA TÉCNICA.",
    "VANOS POR RETEMPLAR.": "REALIZAR EL RETEMPLADO DEL CABLE PARA RESTABLECER LA TENSIÓN.",
    "RESERVAS SUELTAS.": "REORGANIZAR Y ASEGURAR LA RESERVA EN FIGURA 8.",
}
SIN_NOVEDAD_MOTIVO  = "NO SE REGISTRAN NOVEDADES DURANTE LA INSPECCIÓN."
SIN_NOVEDAD_REMEDIO = "EN ESTE PUNTO LA FIBRA SE ENCUENTRA SIN NOVEDAD."


# ══════════════════════════════════════════════════════════════════════════════
#  TOTP
# ══════════════════════════════════════════════════════════════════════════════

def verificar_codigo_totp(codigo: str) -> bool:
    if not TOTP_SECRET:
        return True
    try:
        secreto = base64.b32decode(TOTP_SECRET.upper().replace(" ", ""), casefold=True)
        ahora   = int(time.time()) // 30
        for delta in [0, -1, 1]:
            contador = struct.pack(">Q", ahora + delta)
            mac      = hmac.new(secreto, contador, hashlib.sha1).digest()
            offset   = mac[-1] & 0x0F
            c        = struct.unpack(">I", mac[offset:offset+4])[0] & 0x7FFFFFFF
            if str(c % 1_000_000).zfill(6) == str(codigo).strip():
                return True
    except Exception:
        pass
    return False


# ══════════════════════════════════════════════════════════════════════════════
#  GEMINI AI
# ══════════════════════════════════════════════════════════════════════════════

async def analizar_imagen_gemini(img_bytes: bytes) -> dict | None:
    img_b64 = base64.b64encode(img_bytes).decode()
    prompt  = """Eres experto en inspección de rutas de fibra óptica Telconet Ecuador.
Analiza esta imagen. Si detectas algún problema responde con JSON:
{"tiene_novedad": true, "motivo": "NOMBRE DEL PROBLEMA EN MAYUSCULAS", "coordenadas": ""}
Si todo está bien: {"tiene_novedad": false, "motivo": "", "coordenadas": ""}
Solo JSON, sin texto extra."""
    payload = {"contents": [{"parts": [
        {"text": prompt},
        {"inline_data": {"mime_type": "image/jpeg", "data": img_b64}}
    ]}]}
    try:
        async with httpx.AsyncClient(timeout=30) as client:
            resp = await client.post(GEMINI_URL, json=payload)
            texto = resp.json()["candidates"][0]["content"]["parts"][0]["text"]
            texto = texto.replace("```json","").replace("```","").strip()
            r = json.loads(texto)
            if not r.get("tiene_novedad"):
                return None
            motivo = r.get("motivo","").upper()
            return {
                "motivo":      motivo,
                "remedio":     REMEDIOS.get(motivo, "DOCUMENTAR Y REPORTAR AL COORDINADOR."),
                "coordenadas": r.get("coordenadas",""),
                "tarea_pendiente": "",
                "foto_antes":  img_bytes,
                "foto_despues": None,
            }
    except Exception as e:
        logger.error(f"Gemini error: {e}")
        return None


# ══════════════════════════════════════════════════════════════════════════════
#  EXCEL
# ══════════════════════════════════════════════════════════════════════════════

def generar_excel(datos: dict) -> bytes:
    from openpyxl import Workbook
    wb  = Workbook()
    r   = datos["recorrido"]

    # ── Hoja 1: REPORTES_DE_RECORRIDOS ───────────────────────────────────────
    ws1 = wb.active
    ws1.title = "REPORTES_DE_RECORRIDOS"
    ws1.column_dimensions["A"].width = 45
    ws1.column_dimensions["B"].width = 20

    header = Font(bold=True, color="FFFFFF")
    fill_h = PatternFill("solid", fgColor="1F4E79")
    fill_n = PatternFill("solid", fgColor="D6E4F0")

    def fila(ws, label, valor, fill=None):
        row = ws.max_row + 1
        ws.cell(row, 1, label).font = Font(bold=True)
        if fill:
            ws.cell(row, 1).fill = fill
        ws.cell(row, 2, valor)

    ws1.append(["REPORTE DE RECORRIDO DE RUTAS INTERURBANAS DE F.O. — FOR FO 02"])
    ws1["A1"].font = Font(bold=True, size=12, color="FFFFFF")
    ws1["A1"].fill = PatternFill("solid", fgColor="1F4E79")

    fila(ws1, "FECHA",            r["fecha"])
    fila(ws1, "HORA INICIO",      r["hora_inicio"])
    fila(ws1, "HORA FIN",         r["hora_fin"])
    fila(ws1, "NOMBRE DE LA RUTA",r["nombre_ruta"])
    fila(ws1, "CÓDIGO CUADRILLA", r["codigo_cuadrilla"])
    fila(ws1, "NODO INICIAL",     r["nodo_inicial"])
    fila(ws1, "NODO FINAL",       r["nodo_final"])
    fila(ws1, "LÍDER",            r["lider"])
    fila(ws1, "AYUDANTE",         r["ayudante"])
    fila(ws1, "COORDINADOR",      r["coordinador"])
    fila(ws1, "DISTANCIA",        datos["ciu"].get("distancia_ruta",""))
    fila(ws1, "PLACA VEHÍCULO",   datos["ciu"].get("vehiculo_placa",""))
    fila(ws1, "OBSERVACIONES",    r["observaciones"])
    ws1.append([])

    for nov in r["novedades"]:
        ws1.append([f"NOVEDAD #{nov.get('numero','')}"])
        ws1.cell(ws1.max_row, 1).fill = fill_n
        ws1.cell(ws1.max_row, 1).font = Font(bold=True)
        fila(ws1, "FECHA/HORA",   f"{nov.get('fecha','')} {nov.get('hora_inicio','')} - {nov.get('hora_fin','')}")
        fila(ws1, "MOTIVO",       nov.get("motivo",""))
        fila(ws1, "REMEDIO",      nov.get("remedio",""))
        fila(ws1, "TAREA PEND.",  nov.get("tarea_pendiente",""))
        fila(ws1, "COORDENADAS",  nov.get("coordenadas",""))
        ws1.append([])

    # ── Hoja 2: Checklist MPRIU ───────────────────────────────────────────────
    ws2 = wb.create_sheet("Checklists MPRIU")
    ws2.column_dimensions["A"].width = 50
    ws2.column_dimensions["B"].width = 10
    ws2.column_dimensions["C"].width = 10

    ws2.append(["CHECKLIST MPRIU — FOR FO 08"])
    ws2["A1"].font = Font(bold=True, color="FFFFFF")
    ws2["A1"].fill = PatternFill("solid", fgColor="1F4E79")
    ws2.append(["NOVEDAD", "CHECK", "CANTIDAD"])

    for motivo, vals in datos["mpriu"].get("novedades_check", {}).items():
        ws2.append([motivo, "SI" if vals.get("check") else "NO", vals.get("cantidad", 0)])

    ws2.append([])
    ws2.append(["OBSERVACIONES", datos["mpriu"].get("observaciones","")])

    # ── Hoja 3: MANGAS ────────────────────────────────────────────────────────
    if datos.get("mangas"):
        ws3 = wb.create_sheet("MANGAS")
        ws3.append(["NOMBRE", "DERIVACIÓN", "COORDENADAS", "OBSERVACIÓN"])
        ws3["A1"].font = Font(bold=True)
        for m in datos["mangas"]:
            ws3.append([m.get("nombre",""), m.get("derivacion","NO"), m.get("coordenadas",""), m.get("observacion","")])

    # ── Hoja 4: INVENTARIO HILOS ──────────────────────────────────────────────
    if datos["hilos"].get("filas"):
        ws4 = wb.create_sheet("INVENTARIO DE HILOS EN NODO")
        ws4.append([f"POSICIÓN ODF: {datos['hilos'].get('posicion_odf','')}"])
        ws4.append(["HILO", "DESCRIPCIÓN", "ESTADO"])
        for h in datos["hilos"]["filas"]:
            ws4.append([h.get("hilo_par",""), h.get("descripcion",""), h.get("estado","")])

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def datos_vacios():
    return {
        "recorrido": {
            "fecha": datetime.now().strftime("%d/%m/%Y"),
            "hora_inicio": "", "hora_fin": "",
            "nombre_ruta": "", "codigo_cuadrilla": "",
            "nodo_inicial": "", "nodo_final": "",
            "lider": "", "ayudante": "", "coordinador": "",
            "fotos_total": 0, "observaciones": "", "novedades": [],
        },
        "ciu":   {"vehiculo_placa": "", "distancia_ruta": ""},
        "mpriu": {"novedades_check": {}, "observaciones": ""},
        "mangas": [],
        "hilos":  {"posicion_odf": "", "filas": []},
    }


def novedad_vacia(numero):
    ahora = datetime.now()
    return {
        "numero": numero, "fecha": ahora.strftime("%d/%m/%Y"),
        "hora_inicio": ahora.strftime("%H:%M:%S"), "hora_fin": ahora.strftime("%H:%M:%S"),
        "motivo": "", "remedio": "", "tarea_pendiente": "", "coordenadas": "",
        "foto_antes": None, "foto_despues": None,
    }


def nombre_archivo(datos):
    ruta  = datos["recorrido"]["nombre_ruta"].split()[0].replace("/","-")
    fecha = datetime.now().strftime("%Y%m%d_%H%M")
    return f"FOR_FO_02_{ruta}_{fecha}.xlsx"


# ══════════════════════════════════════════════════════════════════════════════
#  BOT HANDLERS
# ══════════════════════════════════════════════════════════════════════════════

async def start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id in USUARIOS_AUTENTICADOS:
        return await menu_principal(update, ctx)
    await update.message.reply_text(
        "🔐 *RecorridosIA* — Acceso restringido\n\n"
        "Ingresa tu código de *6 dígitos* del autenticador:",
        parse_mode="Markdown", reply_markup=ReplyKeyboardRemove()
    )
    return ESPERANDO_TOTP


DOMINIO_PERMITIDO = os.getenv("DOMINIO_EMAIL", "telconet.ec")

async def verificar_totp(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    texto = update.message.text.strip().lower()
    email = ""
    codigo = ""
    for linea in texto.splitlines():
        if linea.startswith("email:"):
            email = linea.replace("email:", "").strip()
        elif linea.startswith("totp:"):
            codigo = linea.replace("totp:", "").strip()
    if not email or not codigo:
        await update.message.reply_text(
            "❌ Formato incorrecto. Usa exactamente:

"
            "email: tucorreo@telconet.ec
"
            "totp: 123456"
        )
        return ESPERANDO_TOTP
    if not email.endswith(DOMINIO_PERMITIDO):
        await update.message.reply_text(
            f"❌ Solo se permiten correos @{DOMINIO_PERMITIDO}

"
            "email: tucorreo@telconet.ec
"
            "totp: 123456"
        )
        return ESPERANDO_TOTP
    if verificar_codigo_totp(codigo):
        USUARIOS_AUTENTICADOS.add(update.effective_user.id)
        nombre = email.split("@")[0].upper()
        await update.message.reply_text(f"✅ *Acceso autorizado*
Bienvenido {nombre}", parse_mode="Markdown")
        return await menu_principal(update, ctx)
    await update.message.reply_text(
        "❌ Código incorrecto o expirado.

"
        "email: tucorreo@telconet.ec
"
        "totp: 123456"
    )
    return ESPERANDO_TOTP


async def menu_principal(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    teclado = [["🔍 Inspeccionar", "🗺 Nueva Ruta Base"], ["📋 Mis Rutas", "❓ Ayuda"]]
    await update.message.reply_text(
        "📡 *RecorridosIA* — Menú principal",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardMarkup(teclado, resize_keyboard=True)
    )
    return MENU_PRINCIPAL


async def inspeccionar(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in USUARIOS_AUTENTICADOS:
        return await start(update, ctx)
    ctx.user_data["datos"] = datos_vacios()
    ctx.user_data["novedad_actual"] = 0
    ctx.user_data["media_inspeccion"] = []
    await update.message.reply_text(
        "🔍 *Iniciando inspección*\n\n📝 ¿Nombre de la ruta?\n_Ejemplo: GOSSEAL-MACHACHI   TAREA: 157415066_",
        parse_mode="Markdown", reply_markup=ReplyKeyboardRemove()
    )
    return NOMBRE_RUTA


async def recv_nombre_ruta(update, ctx):
    ctx.user_data["datos"]["recorrido"]["nombre_ruta"] = update.message.text.upper()
    await update.message.reply_text("📝 ¿Código de cuadrilla?\n_Ejemplo: FO UIO INT 04_", parse_mode="Markdown")
    return CODIGO_CUADRILLA

async def recv_cuadrilla(update, ctx):
    ctx.user_data["datos"]["recorrido"]["codigo_cuadrilla"] = update.message.text.upper()
    await update.message.reply_text("📝 ¿Nodo inicial?\n_Ejemplo: GOSSEAL_", parse_mode="Markdown")
    return NODO_INICIAL

async def recv_nodo_inicial(update, ctx):
    ctx.user_data["datos"]["recorrido"]["nodo_inicial"] = update.message.text.upper()
    await update.message.reply_text("📝 ¿Nodo final?\n_Ejemplo: MACHACHI_", parse_mode="Markdown")
    return NODO_FINAL

async def recv_nodo_final(update, ctx):
    ctx.user_data["datos"]["recorrido"]["nodo_final"] = update.message.text.upper()
    await update.message.reply_text("👷 ¿Nombre del líder de cuadrilla?", parse_mode="Markdown")
    return LIDER

async def recv_lider(update, ctx):
    ctx.user_data["datos"]["recorrido"]["lider"] = update.message.text.upper()
    await update.message.reply_text("👷 ¿Nombre del ayudante técnico?", parse_mode="Markdown")
    return AYUDANTE

async def recv_ayudante(update, ctx):
    ctx.user_data["datos"]["recorrido"]["ayudante"] = update.message.text.upper()
    await update.message.reply_text("👷 ¿Nombre del coordinador de fibra óptica?", parse_mode="Markdown")
    return COORDINADOR

async def recv_coordinador(update, ctx):
    ctx.user_data["datos"]["recorrido"]["coordinador"] = update.message.text.upper()
    await update.message.reply_text("🚗 ¿Placa del vehículo?\n_Ejemplo: PCO3940_", parse_mode="Markdown")
    return PLACA

async def recv_placa(update, ctx):
    ctx.user_data["datos"]["ciu"]["vehiculo_placa"] = update.message.text.upper()
    await update.message.reply_text("📏 ¿Distancia de la ruta?\n_Ejemplo: 59KM_", parse_mode="Markdown")
    return DISTANCIA

async def recv_distancia(update, ctx):
    ctx.user_data["datos"]["ciu"]["distancia_ruta"] = update.message.text.upper()
    ctx.user_data["datos"]["recorrido"]["hora_inicio"] = datetime.now().strftime("%H:%M:%S")
    await update.message.reply_text(
        "📸 Envía las *fotos de la inspección*.\nCuando termines escribe: *LISTO*",
        parse_mode="Markdown"
    )
    return NOVEDADES_AUTO

async def recv_media(update, ctx):
    if "media_inspeccion" not in ctx.user_data:
        ctx.user_data["media_inspeccion"] = []
    if update.message.photo:
        foto = await update.message.photo[-1].get_file()
        ctx.user_data["media_inspeccion"].append(bytes(await foto.download_as_bytearray()))
        n = len(ctx.user_data["media_inspeccion"])
        await update.message.reply_text(f"📷 Foto {n} recibida. Envía más o escribe *LISTO*", parse_mode="Markdown")
    return NOVEDADES_AUTO

async def procesar_novedades(update, ctx):
    if update.message.text.upper() != "LISTO":
        return NOVEDADES_AUTO
    await update.message.reply_text("🤖 Analizando con IA... ⏳")
    datos    = ctx.user_data["datos"]
    media    = ctx.user_data.get("media_inspeccion", [])
    novedades = []
    for img in media:
        r = await analizar_imagen_gemini(img)
        if r:
            n = novedad_vacia(len(novedades)+1)
            n.update(r)
            novedades.append(n)
    if not novedades:
        n = novedad_vacia(1)
        n["motivo"]  = SIN_NOVEDAD_MOTIVO
        n["remedio"] = SIN_NOVEDAD_REMEDIO
        novedades.append(n)
    datos["recorrido"]["novedades"] = novedades
    for nov in novedades:
        m = nov["motivo"]
        if m != SIN_NOVEDAD_MOTIVO:
            datos["mpriu"]["novedades_check"][m] = {"check": True, "cantidad": datos["mpriu"]["novedades_check"].get(m,{}).get("cantidad",0)+1}
    cantidad = len(novedades)
    await update.message.reply_text(
        f"🔎 *{cantidad} novedad(es) detectada(s):*\n\n" +
        "\n".join([f"  {i+1}. {n['motivo']}" for i,n in enumerate(novedades)]) +
        "\n\n📝 ¿Tarea pendiente para la novedad #1?\n_Si no hay escribe: NINGUNA_",
        parse_mode="Markdown"
    )
    ctx.user_data["novedad_actual"] = 0
    return TAREA_PENDIENTE

async def recv_tarea(update, ctx):
    idx = ctx.user_data["novedad_actual"]
    if update.message.text.upper() != "NINGUNA":
        ctx.user_data["datos"]["recorrido"]["novedades"][idx]["tarea_pendiente"] = update.message.text.upper()
    await update.message.reply_text(f"📸 Foto *ANTES* mantenimiento novedad #{idx+1}\n_Sin foto escribe: SALTAR_", parse_mode="Markdown")
    return FOTO_ANTES

async def recv_foto_antes(update, ctx):
    idx = ctx.user_data["novedad_actual"]
    if update.message.photo:
        foto = await update.message.photo[-1].get_file()
        ctx.user_data["datos"]["recorrido"]["novedades"][idx]["foto_antes"] = bytes(await foto.download_as_bytearray())
    await update.message.reply_text(f"📸 Foto *DESPUÉS* mantenimiento novedad #{idx+1}\n_Sin foto escribe: SALTAR_", parse_mode="Markdown")
    return FOTO_DESPUES

async def recv_foto_despues(update, ctx):
    idx   = ctx.user_data["novedad_actual"]
    datos = ctx.user_data["datos"]
    if update.message.photo:
        foto = await update.message.photo[-1].get_file()
        datos["recorrido"]["novedades"][idx]["foto_despues"] = bytes(await foto.download_as_bytearray())
    ctx.user_data["novedad_actual"] += 1
    sig   = ctx.user_data["novedad_actual"]
    total = len(datos["recorrido"]["novedades"])
    if sig < total:
        await update.message.reply_text(f"📝 ¿Tarea pendiente novedad #{sig+1}?\n_Si no hay: NINGUNA_", parse_mode="Markdown")
        return TAREA_PENDIENTE
    await update.message.reply_text("📝 ¿*Observaciones generales*?\n_Si no hay: NINGUNA_", parse_mode="Markdown")
    return OBSERVACIONES

async def recv_observaciones(update, ctx):
    datos = ctx.user_data["datos"]
    if update.message.text.upper() != "NINGUNA":
        datos["recorrido"]["observaciones"] = update.message.text.upper()
        datos["mpriu"]["observaciones"]     = update.message.text.upper()
    datos["recorrido"]["hora_fin"] = datetime.now().strftime("%H:%M:%S")
    teclado = [["✅ SÍ, hubo cambio de mangas", "❌ No hubo cambio"]]
    await update.message.reply_text(
        "🔧 ¿Hubo *cambio de mangas*?", parse_mode="Markdown",
        reply_markup=ReplyKeyboardMarkup(teclado, resize_keyboard=True, one_time_keyboard=True)
    )
    return PREGUNTA_MANGAS

async def pregunta_mangas(update, ctx):
    if "SÍ" in update.message.text or "SI" in update.message.text.upper():
        await update.message.reply_text("🔧 Nombre de la manga:\n_Ejemplo: UIO-B-MAC/GOS-F1-DER-01_\nCuando termines: *FIN MANGAS*", parse_mode="Markdown", reply_markup=ReplyKeyboardRemove())
        return MANGA_NOMBRE
    teclado = [["✅ SÍ, hubo cambio en ODF", "❌ No hubo cambio"]]
    await update.message.reply_text("💡 ¿Hubo *cambio en ODF*?", parse_mode="Markdown", reply_markup=ReplyKeyboardMarkup(teclado, resize_keyboard=True, one_time_keyboard=True))
    return PREGUNTA_HILOS

async def recv_manga_nombre(update, ctx):
    if update.message.text.upper() == "FIN MANGAS":
        teclado = [["✅ SÍ, hubo cambio en ODF", "❌ No hubo cambio"]]
        await update.message.reply_text("💡 ¿Hubo *cambio en ODF*?", parse_mode="Markdown", reply_markup=ReplyKeyboardMarkup(teclado, resize_keyboard=True, one_time_keyboard=True))
        return PREGUNTA_HILOS
    ctx.user_data["manga_temp"] = {"nombre": update.message.text.upper(), "derivacion": "NO"}
    await update.message.reply_text("📍 Coordenadas:\n_Ejemplo: -0.477057,-78.579350_", parse_mode="Markdown")
    return MANGA_COORDS

async def recv_manga_coords(update, ctx):
    ctx.user_data["manga_temp"]["coordenadas"] = update.message.text
    await update.message.reply_text("📝 Observación:\n_Si no hay: NINGUNA_", parse_mode="Markdown")
    return MANGA_OBS

async def recv_manga_obs(update, ctx):
    manga = ctx.user_data.pop("manga_temp")
    manga["observacion"] = "" if update.message.text.upper() == "NINGUNA" else update.message.text
    ctx.user_data["datos"]["mangas"].append(manga)
    await update.message.reply_text("✅ Manga guardada. Siguiente nombre o *FIN MANGAS*:", parse_mode="Markdown")
    return MANGA_NOMBRE

async def pregunta_hilos(update, ctx):
    if "SÍ" in update.message.text or "SI" in update.message.text.upper():
        await update.message.reply_text("💡 ¿Posición del ODF?\n_Ejemplo: ODF #3_", parse_mode="Markdown", reply_markup=ReplyKeyboardRemove())
        return HILO_ODF
    return await enviar_excel(update, ctx)

async def recv_hilo_odf(update, ctx):
    ctx.user_data["datos"]["hilos"]["posicion_odf"] = update.message.text.upper()
    await update.message.reply_text("💡 Ingresa hilos:\n`HILO, DESCRIPCION, ESTADO`\n_Ejemplo: 1, TELCONET, OCUPADO_\nCuando termines: *FIN HILOS*", parse_mode="Markdown")
    return HILO_DATOS

async def recv_hilo_datos(update, ctx):
    if update.message.text.upper() == "FIN HILOS":
        return await enviar_excel(update, ctx)
    partes = update.message.text.split(",")
    if len(partes) >= 3:
        ctx.user_data["datos"]["hilos"]["filas"].append({"hilo_par": partes[0].strip(), "descripcion": partes[1].strip(), "estado": partes[2].strip().upper()})
    await update.message.reply_text("✅ Guardado. Siguiente o *FIN HILOS*:", parse_mode="Markdown")
    return HILO_DATOS

async def enviar_excel(update, ctx):
    await update.message.reply_text("⚙️ Generando informe *FOR FO 02*...", parse_mode="Markdown", reply_markup=ReplyKeyboardRemove())
    try:
        datos  = ctx.user_data["datos"]
        xl     = generar_excel(datos)
        nombre = nombre_archivo(datos)
        await update.message.reply_document(
            document=xl, filename=nombre,
            caption=f"✅ *FOR FO 02 generado*\n📍 {datos['recorrido']['nombre_ruta']}\n📋 Novedades: {len(datos['recorrido']['novedades'])}",
            parse_mode="Markdown"
        )
    except Exception as e:
        await update.message.reply_text(f"❌ Error: {e}")
    teclado = [["🔍 Inspeccionar", "🗺 Nueva Ruta Base"], ["📋 Mis Rutas"]]
    await update.message.reply_text("¿Qué deseas hacer?", reply_markup=ReplyKeyboardMarkup(teclado, resize_keyboard=True))
    return MENU_PRINCIPAL

async def cancelar(update, ctx):
    ctx.user_data.clear()
    await update.message.reply_text("❌ Cancelado.", reply_markup=ReplyKeyboardRemove())
    return ConversationHandler.END


# ══════════════════════════════════════════════════════════════════════════════
#  ARRANQUE
# ══════════════════════════════════════════════════════════════════════════════

def main():
