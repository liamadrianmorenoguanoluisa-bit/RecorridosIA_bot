"""
RecorridosIA — Bot principal de Telegram
@RecorridosIA_bot

Variables de entorno (configurar en Render):
    BOT_TOKEN       = token del bot de Telegram
    GEMINI_API_KEY  = clave de API de Google Gemini AI
    MAPILLARY_TOKEN = token de acceso a Mapillary
    TOTP_SECRET     = clave secreta para acceso al bot (2FA)
"""

import os
import logging
from telegram import Update, ReplyKeyboardMarkup, ReplyKeyboardRemove
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    ConversationHandler, ContextTypes, filters
)
from reports.excel import (
    generar_excel, datos_vacios, novedad_vacia,
    nombre_archivo, verificar_codigo_totp
)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

BOT_TOKEN = os.getenv("BOT_TOKEN")

# ── Estados del ConversationHandler ──────────────────────────────────────────
(
    ESPERANDO_TOTP,
    MENU_PRINCIPAL,
    NOMBRE_RUTA, CODIGO_CUADRILLA, NODO_INICIAL, NODO_FINAL,
    LIDER, AYUDANTE, COORDINADOR, PLACA, DISTANCIA,
    NOVEDADES_AUTO, TAREA_PENDIENTE, FOTO_ANTES, FOTO_DESPUES,
    OBSERVACIONES, PREGUNTA_MANGAS, PREGUNTA_HILOS,
    MANGA_NOMBRE, MANGA_COORDS, MANGA_OBS,
    HILO_ODF, HILO_DATOS,
) = range(23)

# Usuarios autenticados en esta sesión
USUARIOS_AUTENTICADOS = set()


# ══════════════════════════════════════════════════════════════════════════════
#  AUTENTICACIÓN TOTP — PUERTA DE ENTRADA AL BOT
# ══════════════════════════════════════════════════════════════════════════════

async def start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id

    if user_id in USUARIOS_AUTENTICADOS:
        return await menu_principal(update, ctx)

    await update.message.reply_text(
        "🔐 *RecorridosIA* — Acceso restringido\n\n"
        "Ingresa el código de *6 dígitos* de tu autenticador\n"
        "_(Google Authenticator / Authy)_",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardRemove()
    )
    return ESPERANDO_TOTP


async def verificar_totp(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    codigo  = update.message.text.strip()

    if verificar_codigo_totp(codigo):
        USUARIOS_AUTENTICADOS.add(user_id)
        await update.message.reply_text(
            "✅ *Acceso autorizado*\n\n"
            "Bienvenido a RecorridosIA 👷",
            parse_mode="Markdown"
        )
        return await menu_principal(update, ctx)

    await update.message.reply_text(
        "❌ Código incorrecto o expirado.\n\n"
        "Intenta de nuevo con el código actual de tu autenticador:",
        parse_mode="Markdown"
    )
    return ESPERANDO_TOTP


# ══════════════════════════════════════════════════════════════════════════════
#  MENÚ PRINCIPAL
# ══════════════════════════════════════════════════════════════════════════════

async def menu_principal(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    teclado = [
        ["🗺 Nueva Ruta Base", "🔍 Inspeccionar"],
        ["📋 Mis Rutas",       "❓ Ayuda"]
    ]
    await update.message.reply_text(
        "📡 *RecorridosIA* — Menú principal\n\n"
        "¿Qué deseas hacer?",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardMarkup(teclado, resize_keyboard=True)
    )
    return MENU_PRINCIPAL


async def ayuda(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "📖 *Comandos disponibles:*\n\n"
        "/start          — Menú principal\n"
        "/inspeccionar   — Iniciar inspección de ruta\n"
        "/nueva_ruta     — Registrar nueva ruta base\n"
        "/mis_rutas      — Ver rutas registradas\n"
        "/cancelar       — Cancelar operación actual\n\n"
        "🔐 *Variables de entorno en Render:*\n"
        "`BOT_TOKEN` / `GEMINI_API_KEY` / `MAPILLARY_TOKEN` / `TOTP_SECRET`",
        parse_mode="Markdown"
    )


# ══════════════════════════════════════════════════════════════════════════════
#  FLUJO DE INSPECCIÓN
# ══════════════════════════════════════════════════════════════════════════════

async def inspeccionar(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if user_id not in USUARIOS_AUTENTICADOS:
        return await start(update, ctx)

    ctx.user_data["datos"] = datos_vacios()
    ctx.user_data["novedad_actual"] = 0
    await update.message.reply_text(
        "🔍 *Iniciando inspección*\n\n"
        "📝 ¿Cuál es el *nombre de la ruta*?\n"
        "_Ejemplo: GOSSEAL-MACHACHI   TAREA: 157415066_",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardRemove()
    )
    return NOMBRE_RUTA


async def recibir_nombre_ruta(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["datos"]["recorrido"]["nombre_ruta"] = update.message.text.upper()
    await update.message.reply_text(
        "✅ Ruta registrada.\n\n"
        "📝 ¿Cuál es el *código de cuadrilla*?\n"
        "_Ejemplo: FO UIO INT 04_",
        parse_mode="Markdown"
    )
    return CODIGO_CUADRILLA


async def recibir_cuadrilla(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["datos"]["recorrido"]["codigo_cuadrilla"] = update.message.text.upper()
    await update.message.reply_text(
        "📝 ¿Cuál es el *nodo inicial*?\n_Ejemplo: GOSSEAL_",
        parse_mode="Markdown"
    )
    return NODO_INICIAL


async def recibir_nodo_inicial(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["datos"]["recorrido"]["nodo_inicial"] = update.message.text.upper()
    await update.message.reply_text(
        "📝 ¿Cuál es el *nodo final*?\n_Ejemplo: MACHACHI_",
        parse_mode="Markdown"
    )
    return NODO_FINAL


async def recibir_nodo_final(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["datos"]["recorrido"]["nodo_final"] = update.message.text.upper()
    await update.message.reply_text("👷 ¿Nombre del *líder de cuadrilla*?", parse_mode="Markdown")
    return LIDER


async def recibir_lider(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["datos"]["recorrido"]["lider"] = update.message.text.upper()
    await update.message.reply_text("👷 ¿Nombre del *ayudante técnico*?", parse_mode="Markdown")
    return AYUDANTE


async def recibir_ayudante(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["datos"]["recorrido"]["ayudante"] = update.message.text.upper()
    await update.message.reply_text("👷 ¿Nombre del *coordinador de fibra óptica*?", parse_mode="Markdown")
    return COORDINADOR


async def recibir_coordinador(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["datos"]["recorrido"]["coordinador"] = update.message.text.upper()
    await update.message.reply_text(
        "🚗 ¿*Placa del vehículo*?\n_Ejemplo: PCO3940_",
        parse_mode="Markdown"
    )
    return PLACA


async def recibir_placa(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["datos"]["ciu"]["vehiculo_placa"] = update.message.text.upper()
    await update.message.reply_text(
        "📏 ¿*Distancia de la ruta*?\n_Ejemplo: 59KM_",
        parse_mode="Markdown"
    )
    return DISTANCIA


async def recibir_distancia(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["datos"]["ciu"]["distancia_ruta"] = update.message.text.upper()
    await update.message.reply_text(
        "📸 Perfecto. Ahora *envía las fotos de la inspección*.\n\n"
        "La IA analizará automáticamente las novedades.\n\n"
        "Cuando termines de enviar todo escribe: *LISTO*",
        parse_mode="Markdown"
    )
    return NOVEDADES_AUTO


async def recibir_media_inspeccion(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if "media_inspeccion" not in ctx.user_data:
        ctx.user_data["media_inspeccion"] = []

    if update.message.photo:
        foto = await update.message.photo[-1].get_file()
        foto_bytes = await foto.download_as_bytearray()
        ctx.user_data["media_inspeccion"].append(bytes(foto_bytes))
        await update.message.reply_text(
            f"📷 Foto recibida ({len(ctx.user_data['media_inspeccion'])}). "
            "Envía más o escribe *LISTO*",
            parse_mode="Markdown"
        )
    elif update.message.video or update.message.document:
        await update.message.reply_text(
            "🎥 Video recibido. Escribe *LISTO* para continuar.",
            parse_mode="Markdown"
        )
    return NOVEDADES_AUTO


async def procesar_novedades(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if update.message.text.upper() != "LISTO":
        return NOVEDADES_AUTO

    await update.message.reply_text("🤖 Analizando con IA... un momento ⏳")

    from vision.gemini import analizar_media
    media = ctx.user_data.get("media_inspeccion", [])
    novedades_ia = await analizar_media(media)

    datos = ctx.user_data["datos"]
    for i, nov in enumerate(novedades_ia):
        n = novedad_vacia(i + 1)
        n.update(nov)
        datos["recorrido"]["novedades"].append(n)

    _actualizar_mpriu(datos)

    cantidad = len(novedades_ia)
    if cantidad == 0:
        await update.message.reply_text(
            "✅ La IA no detectó novedades.\n\n"
            "📝 ¿Tienes alguna *observación general*?\n"
            "_Si no, escribe: NINGUNA_",
            parse_mode="Markdown"
        )
        return OBSERVACIONES

    await update.message.reply_text(
        f"🔎 La IA detectó *{cantidad} novedad(es)*:\n\n" +
        "\n".join([f"  {i+1}. {n['motivo']}" for i, n in enumerate(novedades_ia)]) +
        "\n\n📝 ¿Hay alguna *tarea pendiente* para la novedad #1?\n"
        "_Si no hay, escribe: NINGUNA_",
        parse_mode="Markdown"
    )
    ctx.user_data["novedad_actual"] = 0
    return TAREA_PENDIENTE


async def recibir_tarea_pendiente(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    idx   = ctx.user_data["novedad_actual"]
    datos = ctx.user_data["datos"]
    texto = update.message.text
    if texto.upper() != "NINGUNA":
        datos["recorrido"]["novedades"][idx]["tarea_pendiente"] = texto.upper()

    await update.message.reply_text(
        f"📸 Envía foto *ANTES* del mantenimiento — novedad #{idx+1}\n"
        "_Si no tienes, escribe: SALTAR_",
        parse_mode="Markdown"
    )
    return FOTO_ANTES


async def recibir_foto_antes(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    idx = ctx.user_data["novedad_actual"]
    if update.message.photo:
        foto = await update.message.photo[-1].get_file()
        ctx.user_data["datos"]["recorrido"]["novedades"][idx]["foto_antes"] = \
            bytes(await foto.download_as_bytearray())

    await update.message.reply_text(
        f"📸 Envía foto *DESPUÉS* del mantenimiento — novedad #{idx+1}\n"
        "_Si no tienes, escribe: SALTAR_",
        parse_mode="Markdown"
    )
    return FOTO_DESPUES


async def recibir_foto_despues(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    idx   = ctx.user_data["novedad_actual"]
    datos = ctx.user_data["datos"]

    if update.message.photo:
        foto = await update.message.photo[-1].get_file()
        datos["recorrido"]["novedades"][idx]["foto_despues"] = \
            bytes(await foto.download_as_bytearray())

    ctx.user_data["novedad_actual"] += 1
    siguiente = ctx.user_data["novedad_actual"]
    total     = len(datos["recorrido"]["novedades"])

    if siguiente < total:
        await update.message.reply_text(
            f"📝 ¿Tarea pendiente para la novedad #{siguiente+1}?\n"
            "_Si no hay, escribe: NINGUNA_",
            parse_mode="Markdown"
        )
        return TAREA_PENDIENTE

    await update.message.reply_text(
        "📝 ¿*Observaciones generales* del recorrido?\n"
        "_Si no hay, escribe: NINGUNA_",
        parse_mode="Markdown"
    )
    return OBSERVACIONES


async def recibir_observaciones(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    datos = ctx.user_data["datos"]
    texto = update.message.text
    if texto.upper() != "NINGUNA":
        datos["recorrido"]["observaciones"] = texto.upper()
        datos["mpriu"]["observaciones"]     = texto.upper()

    teclado = [["✅ SÍ, hubo cambio de mangas", "❌ No hubo cambio"]]
    await update.message.reply_text(
        "🔧 ¿Hubo *cambio o instalación de mangas* en esta inspección?",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardMarkup(teclado, resize_keyboard=True, one_time_keyboard=True)
    )
    return PREGUNTA_MANGAS


async def pregunta_mangas(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if "SÍ" in update.message.text or "SI" in update.message.text.upper():
        await update.message.reply_text(
            "🔧 *Ingresa el nombre de la manga:*\n"
            "_Ejemplo: UIO-B-MAC/GOS-F1-DER-01_\n\n"
            "Cuando termines escribe: *FIN MANGAS*",
            parse_mode="Markdown",
            reply_markup=ReplyKeyboardRemove()
        )
        return MANGA_NOMBRE

    teclado = [["✅ SÍ, hubo cambio en ODF", "❌ No hubo cambio"]]
    await update.message.reply_text(
        "💡 ¿Hubo *cambio de hilos en el ODF*?",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardMarkup(teclado, resize_keyboard=True, one_time_keyboard=True)
    )
    return PREGUNTA_HILOS


async def recibir_manga_nombre(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if update.message.text.upper() == "FIN MANGAS":
        teclado = [["✅ SÍ, hubo cambio en ODF", "❌ No hubo cambio"]]
        await update.message.reply_text(
            "💡 ¿Hubo *cambio de hilos en el ODF*?",
            parse_mode="Markdown",
            reply_markup=ReplyKeyboardMarkup(teclado, resize_keyboard=True, one_time_keyboard=True)
        )
        return PREGUNTA_HILOS

    ctx.user_data["manga_temp"] = {"nombre": update.message.text.upper(), "derivacion": "NO"}
    await update.message.reply_text(
        "📍 ¿*Coordenadas* de la manga?\n_Ejemplo: -0.477057,-78.579350_",
        parse_mode="Markdown"
    )
    return MANGA_COORDS


async def recibir_manga_coords(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["manga_temp"]["coordenadas"] = update.message.text
    await update.message.reply_text(
        "📝 ¿*Observación* de la manga?\n_Si no hay, escribe: NINGUNA_",
        parse_mode="Markdown"
    )
    return MANGA_OBS


async def recibir_manga_obs(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    manga = ctx.user_data.pop("manga_temp")
    texto = update.message.text
    manga["observacion"] = "" if texto.upper() == "NINGUNA" else texto
    ctx.user_data["datos"]["mangas"].append(manga)
    await update.message.reply_text(
        "✅ Manga guardada.\n\n"
        "🔧 Nombre de la siguiente manga o escribe *FIN MANGAS*:",
        parse_mode="Markdown"
    )
    return MANGA_NOMBRE


async def pregunta_hilos(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if "SÍ" in update.message.text or "SI" in update.message.text.upper():
        await update.message.reply_text(
            "💡 ¿*Posición del ODF*?\n_Ejemplo: ODF #3_",
            parse_mode="Markdown",
            reply_markup=ReplyKeyboardRemove()
        )
        return HILO_ODF

    return await _generar_y_enviar(update, ctx)


async def recibir_hilo_odf(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data["datos"]["hilos"]["posicion_odf"] = update.message.text.upper()
    await update.message.reply_text(
        "💡 Ingresa los hilos en formato:\n"
        "`HILO, DESCRIPCION, ESTADO`\n\n"
        "_Ejemplo: 1, TELCONET TRONCAL, OCUPADO_\n\n"
        "Uno por mensaje. Cuando termines escribe: *FIN HILOS*",
        parse_mode="Markdown"
    )
    return HILO_DATOS


async def recibir_hilo_datos(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if update.message.text.upper() == "FIN HILOS":
        return await _generar_y_enviar(update, ctx)

    partes = update.message.text.split(",")
    if len(partes) >= 3:
        ctx.user_data["datos"]["hilos"]["filas"].append({
            "hilo_par":    partes[0].strip(),
            "descripcion": partes[1].strip(),
            "estado":      partes[2].strip().upper(),
        })
    await update.message.reply_text(
        "✅ Hilo guardado. Envía el siguiente o escribe *FIN HILOS*",
        parse_mode="Markdown"
    )
    return HILO_DATOS


# ══════════════════════════════════════════════════════════════════════════════
#  GENERAR Y ENVIAR EXCEL
# ══════════════════════════════════════════════════════════════════════════════

async def _generar_y_enviar(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "⚙️ Generando informe *FOR FO 02*...",
        parse_mode="Markdown",
        reply_markup=ReplyKeyboardRemove()
    )
    try:
        datos          = ctx.user_data["datos"]
        excel_bytes    = generar_excel(datos)
        archivo_nombre = nombre_archivo(datos)

        await update.message.reply_document(
            document=excel_bytes,
            filename=archivo_nombre,
            caption=(
                f"✅ *Informe FOR FO 02 generado*\n\n"
                f"📍 Ruta: {datos['recorrido']['nombre_ruta']}\n"
                f"📋 Novedades: {len(datos['recorrido']['novedades'])}\n"
                f"📸 Fotos: {datos['recorrido']['fotos_total']}"
            ),
            parse_mode="Markdown"
        )
    except Exception as e:
        logger.error(f"Error generando Excel: {e}")
        await update.message.reply_text(f"❌ Error al generar el informe: {e}")

    teclado = [["🔍 Inspeccionar", "🗺 Nueva Ruta Base"], ["📋 Mis Rutas"]]
    await update.message.reply_text(
        "¿Qué deseas hacer ahora?",
        reply_markup=ReplyKeyboardMarkup(teclado, resize_keyboard=True)
    )
    return MENU_PRINCIPAL


# ══════════════════════════════════════════════════════════════════════════════
#  HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def _actualizar_mpriu(datos: dict):
    conteo = {}
    for nov in datos["recorrido"]["novedades"]:
        motivo = nov.get("motivo", "").upper()
        if motivo and motivo != "NO SE REGISTRAN NOVEDADES DURANTE LA INSPECCIÓN.":
            conteo[motivo] = conteo.get(motivo, 0) + 1
    for motivo, cantidad in conteo.items():
        datos["mpriu"]["novedades_check"][motivo] = {"check": True, "cantidad": cantidad}


async def cancelar(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data.clear()
    await update.message.reply_text(
        "❌ Operación cancelada.",
        reply_markup=ReplyKeyboardRemove()
    )
    return ConversationHandler.END


# ══════════════════════════════════════════════════════════════════════════════
#  ARRANQUE
# ══════════════════════════════════════════════════════════════════════════════

def main():
    app = Application.builder().token(BOT_TOKEN).build()

    conv = ConversationHandler(
        entry_points=[
            CommandHandler("start", start),
            CommandHandler("inspeccionar", inspeccionar),
            MessageHandler(filters.Regex("🔍 Inspeccionar"), inspeccionar),
        ],
        states={
            ESPERANDO_TOTP:   [MessageHandler(filters.TEXT & ~filters.COMMAND, verificar_totp)],
            MENU_PRINCIPAL:   [
                MessageHandler(filters.Regex("🔍 Inspeccionar"), inspeccionar),
                MessageHandler(filters.Regex("❓ Ayuda"), ayuda),
                MessageHandler(filters.TEXT & ~filters.COMMAND, menu_principal),
            ],
            NOMBRE_RUTA:      [MessageHandler(filters.TEXT & ~filters.COMMAND, recibir_nombre_ruta)],
            CODIGO_CUADRILLA: [MessageHandler(filters.TEXT & ~filters.COMMAND, recibir_cuadrilla)],
            NODO_INICIAL:     [MessageHandler(filters.TEXT & ~filters.COMMAND, recibir_nodo_inicial)],
            NODO_FINAL:       [MessageHandler(filters.TEXT & ~filters.COMMAND, recibir_nodo_final)],
            LIDER:            [MessageHandler(filters.TEXT & ~filters.COMMAND, recibir_lider)],
            AYUDANTE:         [MessageHandler(filters.TEXT & ~filters.COMMAND, recibir_ayudante)],
            COORDINADOR:      [MessageHandler(filters.TEXT & ~filters.COMMAND, recibir_coordinador)],
            PLACA:            [MessageHandler(filters.TEXT & ~filters.COMMAND, recibir_placa)],
            DISTANCIA:        [MessageHandler(filters.TEXT & ~filters.COMMAND, recibir_distancia)],
            NOVEDADES_AUTO:   [
                MessageHandler(filters.PHOTO | filters.VIDEO | filters.Document.ALL, recibir_media_inspeccion),
                MessageHandler(filters.TEXT & ~filters.COMMAND, procesar_novedades),
            ],
            TAREA_PENDIENTE:  [MessageHandler(filters.TEXT & ~filters.COMMAND, recibir_tarea_pendiente)],
            FOTO_ANTES:       [
                MessageHandler(filters.PHOTO, recibir_foto_antes),
                MessageHandler(filters.TEXT & ~filters.COMMAND, recibir_foto_antes),
            ],
            FOTO_DESPUES:     [
                MessageHandler(filters.PHOTO, recibir_foto_despues),
                MessageHandler(filters.TEXT & ~filters.COMMAND, recibir_foto_despues),
            ],
            OBSERVACIONES:    [MessageHandler(filters.TEXT & ~filters.COMMAND, recibir_observaciones)],
            PREGUNTA_MANGAS:  [MessageHandler(filters.TEXT & ~filters.COMMAND, pregunta_mangas)],
            MANGA_NOMBRE:     [MessageHandler(filters.TEXT & ~filters.COMMAND, recibir_manga_nombre)],
            MANGA_COORDS:     [MessageHandler(filters.TEXT & ~filters.COMMAND, recibir_manga_coords)],
            MANGA_OBS:        [MessageHandler(filters.TEXT & ~filters.COMMAND, recibir_manga_obs)],
            PREGUNTA_HILOS:   [MessageHandler(filters.TEXT & ~filters.COMMAND, pregunta_hilos)],
            HILO_ODF:         [MessageHandler(filters.TEXT & ~filters.COMMAND, recibir_hilo_odf)],
            HILO_DATOS:       [MessageHandler(filters.TEXT & ~filters.COMMAND, recibir_hilo_datos)],
        },
        fallbacks=[CommandHandler("cancelar", cancelar)],
        allow_reentry=True,
    )

    app.add_handler(conv)
    app.add_handler(CommandHandler("ayuda", ayuda))

    logger.info("🚀 RecorridosIA bot arrancando...")
    app.run_polling()


if __name__ == "__main__":
    main()
