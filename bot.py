import os, io, hmac, struct, time, base64, hashlib, json, logging, threading
import httpx
from datetime import datetime
from http.server import HTTPServer, BaseHTTPRequestHandler
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from telegram import Update, ReplyKeyboardMarkup, ReplyKeyboardRemove
from telegram.ext import Application, CommandHandler, MessageHandler, ConversationHandler, ContextTypes, filters

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

BOT_TOKEN       = os.getenv("BOT_TOKEN")
GEMINI_API_KEY  = os.getenv("GEMINI_API_KEY")
MAPILLARY_TOKEN = os.getenv("MAPILLARY_TOKEN")
TOTP_SECRET     = os.getenv("TOTP_SECRET")
DOMINIO         = os.getenv("DOMINIO_EMAIL", "telconet.ec")
GEMINI_URL      = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=" + (GEMINI_API_KEY or "")

USUARIOS_AUTENTICADOS = set()

(ESPERANDO_TOTP, MENU_PRINCIPAL, NOMBRE_RUTA, CODIGO_CUADRILLA, NODO_INICIAL, NODO_FINAL,
 LIDER, AYUDANTE, COORDINADOR, PLACA, DISTANCIA, NOVEDADES_AUTO, TAREA_PENDIENTE,
 FOTO_ANTES, FOTO_DESPUES, OBSERVACIONES, PREGUNTA_MANGAS, PREGUNTA_HILOS,
 MANGA_NOMBRE, MANGA_COORDS, MANGA_OBS, HILO_ODF, HILO_DATOS) = range(23)

REMEDIOS = {
    "VEGETACION SOBRE FIBRA/MANGA.": "REALIZAR LA PODA O RETIRO DE VEGETACION QUE COMPROMETA LA INTEGRIDAD DEL CABLE.",
    "HERRAJES EN MAL ESTADO.": "REALIZAR EL REEMPLAZO INMEDIATO DEL HERRAJE AFECTADO.",
    "POSTES EN MAL ESTADO.": "DOCUMENTAR Y REPORTAR PARA GESTIONAR EL REEMPLAZO DEL POSTE.",
    "POSTES INCLINADOS.": "DOCUMENTAR Y REPORTAR PARA GESTIONAR EL APLOME DEL POSTE.",
    "MANGAS SUELTAS.": "ASEGURAR LA MANGA AL POSTE EN CONFIGURACION TIPO FIGURA 8.",
    "MANGAS ABIERTAS/DANADAS.": "REEMPLAZAR TAPAS Y SELLOS GARANTIZANDO EL CIERRE HERMETICO.",
    "CABLE LASTIMADO.": "DOCUMENTAR E INFORMAR PARA PROGRAMAR EL CAMBIO DEL TRAMO.",
    "CRUCES DE VIAS BAJOS.": "AJUSTAR LA ALTURA DEL CABLE A LA DISTANCIA REGLAMENTARIA.",
    "POZO SIN TAPA O EN MAL ESTADO.": "SOLICITAR TRABAJOS DE OBRA CIVIL PARA SU CORRECCION.",
    "FALTA DE HERRAJES.": "INSTALAR HERRAJES CONFORME A LA NORMATIVA TECNICA.",
    "VANOS POR RETEMPLAR.": "REALIZAR EL RETEMPLADO DEL CABLE PARA RESTABLECER LA TENSION.",
    "RESERVAS SUELTAS.": "REORGANIZAR Y ASEGURAR LA RESERVA EN FIGURA 8.",
}
SIN_NOVEDAD_MOTIVO  = "NO SE REGISTRAN NOVEDADES DURANTE LA INSPECCION."
SIN_NOVEDAD_REMEDIO = "EN ESTE PUNTO LA FIBRA SE ENCUENTRA SIN NOVEDAD."

def verificar_totp(codigo):
    if not TOTP_SECRET:
        return True
    try:
        secreto = base64.b32decode(TOTP_SECRET.upper().replace(" ",""), casefold=True)
        ahora = int(time.time()) // 30
        for delta in [0, -1, 1]:
            contador = struct.pack(">Q", ahora + delta)
            mac = hmac.new(secreto, contador, hashlib.sha1).digest()
            offset = mac[-1] & 0x0F
            c = struct.unpack(">I", mac[offset:offset+4])[0] & 0x7FFFFFFF
            if str(c % 1000000).zfill(6) == str(codigo).strip():
                return True
    except Exception:
        pass
    return False

async def analizar_imagen(img_bytes):
    img_b64 = base64.b64encode(img_bytes).decode()
    prompt = "Analiza esta imagen de inspeccion de fibra optica. Si hay problema responde JSON: {\"tiene_novedad\": true, \"motivo\": \"NOMBRE EN MAYUSCULAS\", \"coordenadas\": \"\"}. Si todo bien: {\"tiene_novedad\": false, \"motivo\": \"\", \"coordenadas\": \"\"}. Solo JSON."
    payload = {"contents": [{"parts": [{"text": prompt}, {"inline_data": {"mime_type": "image/jpeg", "data": img_b64}}]}]}
    try:
        async with httpx.AsyncClient(timeout=30) as client:
            resp = await client.post(GEMINI_URL, json=payload)
            texto = resp.json()["candidates"][0]["content"]["parts"][0]["text"]
            texto = texto.replace("```json","").replace("```","").strip()
            r = json.loads(texto)
            if not r.get("tiene_novedad"):
                return None
            motivo = r.get("motivo","").upper()
            return {"motivo": motivo, "remedio": REMEDIOS.get(motivo, "DOCUMENTAR Y REPORTAR AL COORDINADOR."), "coordenadas": r.get("coordenadas",""), "tarea_pendiente": "", "foto_antes": img_bytes, "foto_despues": None}
    except Exception as e:
        logger.error("Gemini error: " + str(e))
        return None

def generar_excel(datos):
    wb = Workbook()
    r  = datos["recorrido"]
    ws1 = wb.active
    ws1.title = "REPORTES_DE_RECORRIDOS"
    ws1.column_dimensions["A"].width = 45
    ws1.column_dimensions["B"].width = 30
    ws1.append(["REPORTE DE RECORRIDO DE RUTAS INTERURBANAS DE F.O. - FOR FO 02"])
    ws1["A1"].font = Font(bold=True, size=12, color="FFFFFF")
    ws1["A1"].fill = PatternFill("solid", fgColor="1F4E79")
    campos = [("FECHA", r["fecha"]), ("HORA INICIO", r["hora_inicio"]), ("HORA FIN", r["hora_fin"]),
              ("NOMBRE DE LA RUTA", r["nombre_ruta"]), ("CODIGO CUADRILLA", r["codigo_cuadrilla"]),
              ("NODO INICIAL", r["nodo_inicial"]), ("NODO FINAL", r["nodo_final"]),
              ("LIDER", r["lider"]), ("AYUDANTE", r["ayudante"]), ("COORDINADOR", r["coordinador"]),
              ("DISTANCIA", datos["ciu"].get("distancia_ruta","")), ("PLACA", datos["ciu"].get("vehiculo_placa","")),
              ("OBSERVACIONES", r["observaciones"])]
    for label, valor in campos:
        fila = ws1.max_row + 1
        ws1.cell(fila, 1, label).font = Font(bold=True)
        ws1.cell(fila, 2, valor)
    ws1.append([])
    fill_n = PatternFill("solid", fgColor="D6E4F0")
    for nov in r["novedades"]:
        ws1.append(["NOVEDAD #" + str(nov.get("numero",""))])
        ws1.cell(ws1.max_row, 1).fill = fill_n
        ws1.cell(ws1.max_row, 1).font = Font(bold=True)
        for label, key in [("MOTIVO", "motivo"), ("REMEDIO", "remedio"), ("TAREA PENDIENTE", "tarea_pendiente"), ("COORDENADAS", "coordenadas")]:
            fila = ws1.max_row + 1
            ws1.cell(fila, 1, label).font = Font(bold=True)
            ws1.cell(fila, 2, nov.get(key,""))
        ws1.append([])
    ws2 = wb.create_sheet("Checklists MPRIU")
    ws2.append(["NOVEDAD", "CHECK", "CANTIDAD"])
    for motivo, vals in datos["mpriu"].get("novedades_check", {}).items():
        ws2.append([motivo, "SI" if vals.get("check") else "NO", vals.get("cantidad", 0)])
    if datos.get("mangas"):
        ws3 = wb.create_sheet("MANGAS")
        ws3.append(["NOMBRE", "DERIVACION", "COORDENADAS", "OBSERVACION"])
        for m in datos["mangas"]:
            ws3.append([m.get("nombre",""), m.get("derivacion","NO"), m.get("coordenadas",""), m.get("observacion","")])
    if datos["hilos"].get("filas"):
        ws4 = wb.create_sheet("INVENTARIO DE HILOS EN NODO")
        ws4.append(["HILO", "DESCRIPCION", "ESTADO"])
        for h in datos["hilos"]["filas"]:
            ws4.append([h.get("hilo_par",""), h.get("descripcion",""), h.get("estado","")])
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()

def datos_vacios():
    return {
        "recorrido": {"fecha": datetime.now().strftime("%d/%m/%Y"), "hora_inicio": "", "hora_fin": "", "nombre_ruta": "", "codigo_cuadrilla": "", "nodo_inicial": "", "nodo_final": "", "lider": "", "ayudante": "", "coordinador": "", "fotos_total": 0, "observaciones": "", "novedades": []},
        "ciu": {"vehiculo_placa": "", "distancia_ruta": ""},
        "mpriu": {"novedades_check": {}, "observaciones": ""},
        "mangas": [], "hilos": {"posicion_odf": "", "filas": []},
    }

def novedad_vacia(numero):
    ahora = datetime.now()
    return {"numero": numero, "fecha": ahora.strftime("%d/%m/%Y"), "hora_inicio": ahora.strftime("%H:%M:%S"), "hora_fin": ahora.strftime("%H:%M:%S"), "motivo": "", "remedio": "", "tarea_pendiente": "", "coordenadas": "", "foto_antes": None, "foto_despues": None}

def nombre_archivo(datos):
    ruta = datos["recorrido"]["nombre_ruta"].split()[0].replace("/","-")
    return "FOR_FO_02_" + ruta + "_" + datetime.now().strftime("%Y%m%d_%H%M") + ".xlsx"

async def start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id in USUARIOS_AUTENTICADOS:
        return await menu_principal(update, ctx)
    await update.message.reply_text(
        "RecorridosIA - Acceso restringido\n\nIngresa tu correo y codigo de 6 digitos:\n\nemail: tucorreo@telconet.ec\ntotp: 123456",
        reply_markup=ReplyKeyboardRemove()
    )
    return ESPERANDO_TOTP

async def handler_totp(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    texto = update.message.text.strip().lower()
    email, codigo = "", ""
    for linea in texto.splitlines():
        if linea.startswith("email:"):
            email = linea.replace("email:", "").strip()
        elif linea.startswith("totp:"):
            codigo = linea.replace("totp:", "").strip()
    if not email or not codigo:
        await update.message.reply_text("Formato incorrecto. Usa:\nemail: tucorreo@telconet.ec\ntotp: 123456")
        return ESPERANDO_TOTP
    if not email.endswith(DOMINIO):
        await update.message.reply_text("Solo correos @" + DOMINIO + "\nemail: tucorreo@telconet.ec\ntotp: 123456")
        return ESPERANDO_TOTP
    if verificar_totp(codigo):
        USUARIOS_AUTENTICADOS.add(update.effective_user.id)
        nombre = email.split("@")[0].upper()
        await update.message.reply_text("Acceso autorizado. Bienvenido " + nombre)
        return await menu_principal(update, ctx)
    await update.message.reply_text("Codigo incorrecto.\nemail: tucorreo@telconet.ec\ntotp: 123456")
    return ESPERANDO_TOTP

async def menu_principal(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    teclado = [["Inspeccionar", "Nueva Ruta Base"], ["Mis Rutas", "Ayuda"]]
    await update.message.reply_text("RecorridosIA - Menu principal", reply_markup=ReplyKeyboardMarkup(teclado, resize_keyboard=True))
    return MENU_PRINCIPAL

async def inspeccionar(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in USUARIOS_AUTENTICADOS:
        return await start(update, ctx)
    ctx.user_data["datos"] = datos_vacios()
    ctx.user_data["novedad_actual"] = 0
    ctx.user_data["media_inspeccion"] = []
    await update.message.reply_text("Iniciando inspeccion\n\nNombre de la ruta?\nEjemplo: GOSSEAL-MACHACHI   TAREA: 157415066", reply_markup=ReplyKeyboardRemove())
    return NOMBRE_RUTA

async def recv_nombre_ruta(update, ctx):
    ctx.user_data["datos"]["recorrido"]["nombre_ruta"] = update.message.text.upper()
    await update.message.reply_text("Codigo de cuadrilla?\nEjemplo: FO UIO INT 04")
    return CODIGO_CUADRILLA

async def recv_cuadrilla(update, ctx):
    ctx.user_data["datos"]["recorrido"]["codigo_cuadrilla"] = update.message.text.upper()
    await update.message.reply_text("Nodo inicial?\nEjemplo: GOSSEAL")
    return NODO_INICIAL

async def recv_nodo_inicial(update, ctx):
    ctx.user_data["datos"]["recorrido"]["nodo_inicial"] = update.message.text.upper()
    await update.message.reply_text("Nodo final?\nEjemplo: MACHACHI")
    return NODO_FINAL

async def recv_nodo_final(update, ctx):
    ctx.user_data["datos"]["recorrido"]["nodo_final"] = update.message.text.upper()
    await update.message.reply_text("Nombre del lider de cuadrilla?")
    return LIDER

async def recv_lider(update, ctx):
    ctx.user_data["datos"]["recorrido"]["lider"] = update.message.text.upper()
    await update.message.reply_text("Nombre del ayudante tecnico?")
    return AYUDANTE

async def recv_ayudante(update, ctx):
    ctx.user_data["datos"]["recorrido"]["ayudante"] = update.message.text.upper()
    await update.message.reply_text("Nombre del coordinador de fibra optica?")
    return COORDINADOR

async def recv_coordinador(update, ctx):
    ctx.user_data["datos"]["recorrido"]["coordinador"] = update.message.text.upper()
    await update.message.reply_text("Placa del vehiculo?\nEjemplo: PCO3940")
    return PLACA

async def recv_placa(update, ctx):
    ctx.user_data["datos"]["ciu"]["vehiculo_placa"] = update.message.text.upper()
    await update.message.reply_text("Distancia de la ruta?\nEjemplo: 59KM")
    return DISTANCIA

async def recv_distancia(update, ctx):
    ctx.user_data["datos"]["ciu"]["distancia_ruta"] = update.message.text.upper()
    ctx.user_data["datos"]["recorrido"]["hora_inicio"] = datetime.now().strftime("%H:%M:%S")
    await update.message.reply_text("Envia las fotos de la inspeccion.\nCuando termines escribe: LISTO")
    return NOVEDADES_AUTO

async def recv_media(update, ctx):
    if "media_inspeccion" not in ctx.user_data:
        ctx.user_data["media_inspeccion"] = []
    if update.message.photo:
        foto = await update.message.photo[-1].get_file()
        ctx.user_data["media_inspeccion"].append(bytes(await foto.download_as_bytearray()))
        n = len(ctx.user_data["media_inspeccion"])
        await update.message.reply_text("Foto " + str(n) + " recibida. Envia mas o escribe LISTO")
    return NOVEDADES_AUTO

async def procesar_novedades(update, ctx):
    if update.message.text.upper() != "LISTO":
        return NOVEDADES_AUTO
    await update.message.reply_text("Analizando con IA...")
    datos = ctx.user_data["datos"]
    media = ctx.user_data.get("media_inspeccion", [])
    novedades = []
    for img in media:
        r = await analizar_imagen(img)
        if r:
            n = novedad_vacia(len(novedades)+1)
            n.update(r)
            novedades.append(n)
    if not novedades:
        n = novedad_vacia(1)
        n["motivo"] = SIN_NOVEDAD_MOTIVO
        n["remedio"] = SIN_NOVEDAD_REMEDIO
        novedades.append(n)
    datos["recorrido"]["novedades"] = novedades
    for nov in novedades:
        m = nov["motivo"]
        if m != SIN_NOVEDAD_MOTIVO:
            datos["mpriu"]["novedades_check"][m] = {"check": True, "cantidad": datos["mpriu"]["novedades_check"].get(m, {}).get("cantidad", 0)+1}
    msg = str(len(novedades)) + " novedad(es) detectada(s):\n\n"
    for i, n in enumerate(novedades):
        msg += str(i+1) + ". " + n["motivo"] + "\n"
    msg += "\nTarea pendiente para novedad #1?\nSi no hay escribe: NINGUNA"
    await update.message.reply_text(msg)
    ctx.user_data["novedad_actual"] = 0
    return TAREA_PENDIENTE

async def recv_tarea(update, ctx):
    idx = ctx.user_data["novedad_actual"]
    if update.message.text.upper() != "NINGUNA":
        ctx.user_data["datos"]["recorrido"]["novedades"][idx]["tarea_pendiente"] = update.message.text.upper()
    await update.message.reply_text("Foto ANTES mantenimiento novedad #" + str(idx+1) + "\nSin foto escribe: SALTAR")
    return FOTO_ANTES

async def recv_foto_antes(update, ctx):
    idx = ctx.user_data["novedad_actual"]
    if update.message.photo:
        foto = await update.message.photo[-1].get_file()
        ctx.user_data["datos"]["recorrido"]["novedades"][idx]["foto_antes"] = bytes(await foto.download_as_bytearray())
    await update.message.reply_text("Foto DESPUES mantenimiento novedad #" + str(idx+1) + "\nSin foto escribe: SALTAR")
    return FOTO_DESPUES

async def recv_foto_despues(update, ctx):
    idx = ctx.user_data["novedad_actual"]
    datos = ctx.user_data["datos"]
    if update.message.photo:
        foto = await update.message.photo[-1].get_file()
        datos["recorrido"]["novedades"][idx]["foto_despues"] = bytes(await foto.download_as_bytearray())
    ctx.user_data["novedad_actual"] += 1
    sig = ctx.user_data["novedad_actual"]
    total = len(datos["recorrido"]["novedades"])
    if sig < total:
        await update.message.reply_text("Tarea pendiente novedad #" + str(sig+1) + "?\nSi no hay: NINGUNA")
        return TAREA_PENDIENTE
    await update.message.reply_text("Observaciones generales?\nSi no hay: NINGUNA")
    return OBSERVACIONES

async def recv_observaciones(update, ctx):
    datos = ctx.user_data["datos"]
    if update.message.text.upper() != "NINGUNA":
        datos["recorrido"]["observaciones"] = update.message.text.upper()
    datos["recorrido"]["hora_fin"] = datetime.now().strftime("%H:%M:%S")
    teclado = [["SI, hubo cambio de mangas", "No hubo cambio"]]
    await update.message.reply_text("Hubo cambio de mangas?", reply_markup=ReplyKeyboardMarkup(teclado, resize_keyboard=True, one_time_keyboard=True))
    return PREGUNTA_MANGAS

async def pregunta_mangas(update, ctx):
    if "SI" in update.message.text.upper():
        await update.message.reply_text("Nombre de la manga:\nEjemplo: UIO-B-MAC/GOS-F1-DER-01\nCuando termines: FIN MANGAS", reply_markup=ReplyKeyboardRemove())
        return MANGA_NOMBRE
    teclado = [["SI, hubo cambio en ODF", "No hubo cambio"]]
    await update.message.reply_text("Hubo cambio en ODF?", reply_markup=ReplyKeyboardMarkup(teclado, resize_keyboard=True, one_time_keyboard=True))
    return PREGUNTA_HILOS

async def recv_manga_nombre(update, ctx):
    if update.message.text.upper() == "FIN MANGAS":
        teclado = [["SI, hubo cambio en ODF", "No hubo cambio"]]
        await update.message.reply_text("Hubo cambio en ODF?", reply_markup=ReplyKeyboardMarkup(teclado, resize_keyboard=True, one_time_keyboard=True))
        return PREGUNTA_HILOS
    ctx.user_data["manga_temp"] = {"nombre": update.message.text.upper(), "derivacion": "NO"}
    await update.message.reply_text("Coordenadas de la manga?\nEjemplo: -0.477057,-78.579350")
    return MANGA_COORDS

async def recv_manga_coords(update, ctx):
    ctx.user_data["manga_temp"]["coordenadas"] = update.message.text
    await update.message.reply_text("Observacion de la manga?\nSi no hay: NINGUNA")
    return MANGA_OBS

async def recv_manga_obs(update, ctx):
    manga = ctx.user_data.pop("manga_temp")
    manga["observacion"] = "" if update.message.text.upper() == "NINGUNA" else update.message.text
    ctx.user_data["datos"]["mangas"].append(manga)
    await update.message.reply_text("Manga guardada. Siguiente nombre o FIN MANGAS:")
    return MANGA_NOMBRE

async def pregunta_hilos(update, ctx):
    if "SI" in update.message.text.upper():
        await update.message.reply_text("Posicion del ODF?\nEjemplo: ODF #3", reply_markup=ReplyKeyboardRemove())
        return HILO_ODF
    return await enviar_excel(update, ctx)

async def recv_hilo_odf(update, ctx):
    ctx.user_data["datos"]["hilos"]["posicion_odf"] = update.message.text.upper()
    await update.message.reply_text("Ingresa hilos:\nHILO, DESCRIPCION, ESTADO\nEjemplo: 1, TELCONET, OCUPADO\nCuando termines: FIN HILOS")
    return HILO_DATOS

async def recv_hilo_datos(update, ctx):
    if update.message.text.upper() == "FIN HILOS":
        return await enviar_excel(update, ctx)
    partes = update.message.text.split(",")
    if len(partes) >= 3:
        ctx.user_data["datos"]["hilos"]["filas"].append({"hilo_par": partes[0].strip(), "descripcion": partes[1].strip(), "estado": partes[2].strip().upper()})
    await update.message.reply_text("Guardado. Siguiente o FIN HILOS:")
    return HILO_DATOS

async def enviar_excel(update, ctx):
    await update.message.reply_text("Generando informe FOR FO 02...", reply_markup=ReplyKeyboardRemove())
    try:
        datos = ctx.user_data["datos"]
        xl = generar_excel(datos)
        nombre = nombre_archivo(datos)
        await update.message.reply_document(document=xl, filename=nombre, caption="FOR FO 02 generado\nRuta: " + datos["recorrido"]["nombre_ruta"] + "\nNovedades: " + str(len(datos["recorrido"]["novedades"])))
    except Exception as e:
        await update.message.reply_text("Error: " + str(e))
    teclado = [["Inspeccionar", "Nueva Ruta Base"], ["Mis Rutas"]]
    await update.message.reply_text("Que deseas hacer?", reply_markup=ReplyKeyboardMarkup(teclado, resize_keyboard=True))
    return MENU_PRINCIPAL

async def cancelar(update, ctx):
    ctx.user_data.clear()
    await update.message.reply_text("Cancelado.", reply_markup=ReplyKeyboardRemove())
    return ConversationHandler.END

class PingHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b"RecorridosIA OK")
    def log_message(self, format, *args):
        pass

def start_web():
    port = int(os.getenv("PORT", 8080))
    server = HTTPServer(("0.0.0.0", port), PingHandler)
    logger.info("Servidor web en puerto " + str(port))
    server.serve_forever()

def build_app():
    app = Application.builder().token(BOT_TOKEN).build()
    conv = ConversationHandler(
        entry_points=[CommandHandler("start", start), CommandHandler("inspeccionar", inspeccionar), MessageHandler(filters.Regex("Inspeccionar"), inspeccionar)],
        states={
            ESPERANDO_TOTP:   [MessageHandler(filters.TEXT & ~filters.COMMAND, handler_totp)],
            MENU_PRINCIPAL:   [MessageHandler(filters.Regex("Inspeccionar"), inspeccionar), MessageHandler(filters.TEXT & ~filters.COMMAND, menu_principal)],
            NOMBRE_RUTA:      [MessageHandler(filters.TEXT & ~filters.COMMAND, recv_nombre_ruta)],
            CODIGO_CUADRILLA: [MessageHandler(filters.TEXT & ~filters.COMMAND, recv_cuadrilla)],
            NODO_INICIAL:     [MessageHandler(filters.TEXT & ~filters.COMMAND, recv_nodo_inicial)],
            NODO_FINAL:       [MessageHandler(filters.TEXT & ~filters.COMMAND, recv_nodo_final)],
            LIDER:            [MessageHandler(filters.TEXT & ~filters.COMMAND, recv_lider)],
            AYUDANTE:         [MessageHandler(filters.TEXT & ~filters.COMMAND, recv_ayudante)],
            COORDINADOR:      [MessageHandler(filters.TEXT & ~filters.COMMAND, recv_coordinador)],
            PLACA:            [MessageHandler(filters.TEXT & ~filters.COMMAND, recv_placa)],
            DISTANCIA:        [MessageHandler(filters.TEXT & ~filters.COMMAND, recv_distancia)],
            NOVEDADES_AUTO:   [MessageHandler(filters.PHOTO, recv_media), MessageHandler(filters.TEXT & ~filters.COMMAND, procesar_novedades)],
            TAREA_PENDIENTE:  [MessageHandler(filters.TEXT & ~filters.COMMAND, recv_tarea)],
            FOTO_ANTES:       [MessageHandler(filters.PHOTO, recv_foto_antes), MessageHandler(filters.TEXT & ~filters.COMMAND, recv_foto_antes)],
            FOTO_DESPUES:     [MessageHandler(filters.PHOTO, recv_foto_despues), MessageHandler(filters.TEXT & ~filters.COMMAND, recv_foto_despues)],
            OBSERVACIONES:    [MessageHandler(filters.TEXT & ~filters.COMMAND, recv_observaciones)],
            PREGUNTA_MANGAS:  [MessageHandler(filters.TEXT & ~filters.COMMAND, pregunta_mangas)],
            MANGA_NOMBRE:     [MessageHandler(filters.TEXT & ~filters.COMMAND, recv_manga_nombre)],
            MANGA_COORDS:     [MessageHandler(filters.TEXT & ~filters.COMMAND, recv_manga_coords)],
            MANGA_OBS:        [MessageHandler(filters.TEXT & ~filters.COMMAND, recv_manga_obs)],
            PREGUNTA_HILOS:   [MessageHandler(filters.TEXT & ~filters.COMMAND, pregunta_hilos)],
            HILO_ODF:         [MessageHandler(filters.TEXT & ~filters.COMMAND, recv_hilo_odf)],
            HILO_DATOS:       [MessageHandler(filters.TEXT & ~filters.COMMAND, recv_hilo_datos)],
        },
        fallbacks=[CommandHandler("cancelar", cancelar)],
        allow_reentry=True,
    )
    app.add_handler(conv)
    return app

if __name__ == "__main__":
    # Servidor web en hilo secundario
    t = threading.Thread(target=start_web, daemon=True)
    t.start()
    # Bot en hilo principal
    logger.info("RecorridosIA bot arrancando...")
    build_app().run_polling()
