import os, io, hmac, struct, time, base64, hashlib, json, logging, threading
import httpx
from datetime import datetime
from http.server import HTTPServer, BaseHTTPRequestHandler
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from telegram import Update, ReplyKeyboardMarkup, ReplyKeyboardRemove
from telegram import InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, MessageHandler, ConversationHandler, ContextTypes, filters, CallbackQueryHandler

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
 LIDER, AYUDANTE, COORDINADOR, PLACA, DISTANCIA,
 CIU_HERRAMIENTAS, CIU_EQUIPOS, CIU_MATERIALES,
 NOVEDADES_AUTO, TAREA_PENDIENTE,
 FOTO_ANTES, FOTO_DESPUES, OBSERVACIONES,
 MPRIU_CHECK,
 PREGUNTA_MANGAS, PREGUNTA_HILOS,
 MANGA_NOMBRE, MANGA_COORDS, MANGA_OBS, HILO_ODF, HILO_DATOS,
 NUEVA_RUTA_NOMBRE, NUEVA_RUTA_VIDEO,
 GENERAR_CONFIRMAR, GENERAR_NOMBRE_RUTA, GENERAR_CUADRILLA,
 GENERAR_NODO_INI, GENERAR_NODO_FIN, GENERAR_LIDER,
 GENERAR_AYUDANTE, GENERAR_COORDINADOR, GENERAR_PLACA,
 GENERAR_DISTANCIA, GENERAR_FECHA, GENERAR_HORA_INI,
 GENERAR_HORA_FIN, GENERAR_NOVEDADES,
 TAB_MENU, TAB_CIU_HERR, TAB_CIU_EQUI, TAB_CIU_MATE,
 TAB_MPRIU, TAB_REPORTES, TAB_NOVEDADES_IA,
 VIDEO_BASE_NOMBRE, VIDEO_BASE_UPLOAD) = range(52)

# Rutas guardadas en memoria
RUTAS_GUARDADAS = {}

# Herramientas CIU agrupadas para preguntar por grupos
HERRAMIENTAS_CIU = [
    ("Cinturon y Linea de Vida", "cinturon"),
    ("Casco", "casco"),
    ("Escalera de 28 pies", "escalera_28"),
    ("Conos reflectivos", "conos"),
    ("Juego de destornilladores", "destornilladores"),
    ("Martillo mediano", "martillo"),
    ("Estiletes", "estiletes"),
    ("Cortafrio", "cortafrio"),
    ("Juego de rachet", "rachet"),
    ("Pares de guantes aislantes", "guantes"),
    ("Tecle", "tecle"),
    ("Machete", "machete"),
    ("Cizalla", "cizalla"),
]
EQUIPOS_CIU = [
    ("Fusionadora", "fusionadora"),
    ("Cortadora de fibra", "cortadora"),
    ("OTDR con cargador", "otdr"),
    ("Llave Acsys", "acsys"),
    ("Inversor", "inversor"),
    ("Etiquetadora", "etiquetadora"),
]
MATERIALES_CIU = [
    ("Fibra 48h (500mt)", "fibra"),
    ("Mangas de 48h y/o 144h", "mangas_mat"),
    ("Rollo de cinta Eriband 3/4", "eriband"),
    ("Patchcord de fibra", "patchcord"),
    ("Adaptadores (Simplex-Duplex)", "adaptadores"),
    ("Paquetes de amarras", "amarras"),
]

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

NOVEDADES_MPRIU = [
    "HERRAJES EN MAL ESTADO.", "FALTA DE HERRAJES.", "POSTES EN MAL ESTADO.",
    "POSTE(S) CAMBIADO(S).", "POSTES POR INSTALAR.", "POSTE NUEVO INSTALADO - TN.",
    "POSTE NUEVO INSTALADO - EMPRESAS ELECTRICAS.", "POSTES INCLINADOS.",
    "RETENIDA(S) EN MAL ESTADO.", "RETENIDA(S) CORTADA(S).", "VANOS POR RETEMPLAR.",
    "MANGAS SUELTAS.", "MANGAS ABIERTAS/DANADAS.", "RESERVAS SUELTAS.",
    "CRUCES DE VIAS BAJOS.", "VEGETACION SOBRE FIBRA/MANGA.", "LOCALIZACION DE MANGA.",
    "DOCUMENTACION UNIFILAR DE HILOS.", "LINEA ELECTRICA EN MAL ESTADO.",
    "REGENERACION URBANA.", "AMPLIACION DE VIA.", "CABLE LASTIMADO.",
    "FIBRA INSTALADA INCORRECTAMENTE SOBRE MORDAZA.", "POZO SIN TAPA O EN MAL ESTADO.",
    "REPINTADO DE POZO.", "REPINTADO DE POSTE.", "ELEMENTOS SIN ETIQUETAS ACRILICAS.",
    "RIESGO DE DERRUMBE O DESLAVE.", "RIESGO DE INUNDACIONES.", "RIESGO DE INCENDIO.",
    "NO SE REGISTRAN NOVEDADES DURANTE LA INSPECCION.",
]

# Mapa de filas exactas del template Telconet
HERR_FILAS = {
    "Cinturon y Linea de Vida":14, "Casco":15, "Escalera de 24 pies":16,
    "Escalera de 28 pies":17, "Escalera de 32 pies":18, "Conos reflectivos":19,
    "Caja para herramientas":20, "Juego de destornilladores":21, "Martillo mediano":22,
    "Estiletes":23, "Cortafrio":24, "Alicate":25, "Llave francesa":26,
    "Juego de rachet":27, "Pares de guantes aislantes":28, "Tecle":29,
    "Machete":30, "Cizalla":31, "Pata de cabra":32, "Flejadora (Maquina Eriband)":33,
    "Extension con foco":34, "Motosierra":35, "Tijeras metalicas":36,
    "Arco de sierra":37, "Binoculares":38, "Parasol":39, "Remolque / Carrete para F.O.":40,
}
EQUI_FILAS = {
    "Fusionadora":42, "Cortadora de fibra":43, "Bobina de lanzamiento":44,
    "OTDR con cargador":45, "Llave Acsys":46, "GPS":47, "Inversor":48, "Etiquetadora":49,
}
MATE_FILAS = {
    "Fibra 48h (500mt)":51, "Mangas de 48h y/o 144h (2 minimo)":52,
    "Rollo de cinta Eriband 3/4":53, "Hebillas para cinta Eriband 3/4":54,
    "Hojas de sierra":55, "Patchcord de fibra":56, "Adaptadores (Simplex-Duplex)":57,
    "Paquetes de amarras":58, "Mesas plasticas":59, "Sillas plasticas":60,
    "Cuchillos":61, "Poleas":62, "Sogas de nylon medianas":63,
    "Sogas de nylon gruesas":64, "Repelente contra insectos":65,
    "Repelente contra abejas y avispas":66,
}

def _copiar_logos(wb_src, wb_dst):
    """Copia el logo Telconet de la plantilla original a cada hoja del nuevo workbook."""
    from openpyxl.drawing.image import Image as XLImg
    from openpyxl.drawing.spreadsheet_drawing import TwoCellAnchor, AnchorMarker
    import copy

    for sheet in wb_src.sheetnames:
        if sheet not in wb_dst.sheetnames:
            continue
        ws_src = wb_src[sheet]
        ws_dst = wb_dst[sheet]
        if ws_src._images:
            # Solo copiar el primer logo (indice 0)
            img_src = ws_src._images[0]
            logo_data = img_src._data()
            img_nuevo = XLImg(io.BytesIO(logo_data))
            # Copiar anchor
            anc_src = img_src.anchor
            fr = anc_src._from
            to = anc_src.to
            img_nuevo.anchor = img_src.anchor
            ws_dst.add_image(img_nuevo)

def _get_plantilla():
    """Carga la plantilla FOR FO 02 desde el mismo directorio del bot."""
    import os
    from openpyxl import load_workbook
    rutas = [
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "plantilla_FOR_FO_02.xlsx"),
        "plantilla_FOR_FO_02.xlsx",
        "/opt/render/project/src/plantilla_FOR_FO_02.xlsx",
    ]
    for ruta in rutas:
        if os.path.exists(ruta):
            return load_workbook(ruta)
    raise FileNotFoundError("No se encontro plantilla_FOR_FO_02.xlsx")

def generar_excel(datos):
    from openpyxl import load_workbook
    r   = datos["recorrido"]
    ciu = datos["ciu"]
    nch = datos["mpriu"].get("novedades_check", {})

    wb = _get_plantilla()

    # ══ HOJA 1: REPORTES_DE_RECORRIDOS ══════════════════════════════
    ws1 = wb["REPORTES_DE_RECORRIDOS"]
    ws1["B6"] = r["fecha"]
    ws1["C6"] = r["hora_inicio"]
    ws1["D6"] = r["hora_fin"]
    ws1["B7"] = r["nombre_ruta"]
    ws1["B8"] = r["codigo_cuadrilla"]
    ws1["B9"] = r["nodo_inicial"]

    # Novedades — cada novedad ocupa 6 filas desde fila 10
    fila_nov = 10
    for nov in r["novedades"]:
        ws1.cell(fila_nov,   2, nov.get("fecha",""))
        ws1.cell(fila_nov,   3, nov.get("hora_inicio",""))
        ws1.cell(fila_nov,   4, nov.get("hora_fin",""))
        ws1.cell(fila_nov+2, 2, nov.get("motivo",""))
        ws1.cell(fila_nov+3, 2, nov.get("remedio",""))
        ws1.cell(fila_nov+4, 2, nov.get("tarea_pendiente",""))
        ws1.cell(fila_nov+5, 2, nov.get("coordenadas",""))
        fila_nov += 6

    # Pie — buscar labels y llenar valores
    for row in ws1.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                v = cell.value.strip().upper()
                col_val = cell.column + 1
                if v == "NODO FINAL":
                    ws1.cell(cell.row, col_val, r["nodo_final"])
                elif "LIDER DE CUADRILLA" in v:
                    ws1.cell(cell.row, col_val, r["lider"])
                elif v == "AYUDANTE TECNICO":
                    ws1.cell(cell.row, col_val, r["ayudante"])
                elif "COORDINADOR FIBRA" in v:
                    ws1.cell(cell.row, col_val, r["coordinador"])
                elif "FOTOS ANEXAS" in v:
                    ws1.cell(cell.row, col_val, str(r["fotos_total"]))
                elif v == "OBSERVACIONES GENERALES":
                    ws1.cell(cell.row, col_val, r["observaciones"])

    # ══ HOJA 5: Checklist CIU ════════════════════════════════════════
    ws5 = wb["Checklist CIU"]
    ws5["C4"]  = r["fecha"]
    ws5["F4"]  = r["hora_inicio"]
    ws5["H4"]  = r["hora_fin"]
    ws5["C5"]  = r["nombre_ruta"]
    ws5["C6"]  = r["nodo_inicial"]
    ws5["C7"]  = r["nodo_final"]
    ws5["C8"]  = ciu.get("distancia_ruta","")
    ws5["C9"]  = r["lider"]
    ws5["C10"] = ciu.get("vehiculo_placa","")
    ws5["C11"] = r["coordinador"]

    for nombre, fila in HERR_FILAS.items():
        cant = ciu.get("herramientas",{}).get(nombre, {})
        if isinstance(cant, dict):
            cant = cant.get("cantidad", 0)
        ws5.cell(fila, 4, cant)
        ws5.cell(fila, 5, "BUEN ESTADO" if cant > 0 else "NINGUNA")

    for nombre, fila in EQUI_FILAS.items():
        cant = ciu.get("equipos",{}).get(nombre, {})
        if isinstance(cant, dict):
            cant = cant.get("cantidad", 0)
        ws5.cell(fila, 4, cant)
        ws5.cell(fila, 5, "BUEN ESTADO" if cant > 0 else "NINGUNA")

    for nombre, fila in MATE_FILAS.items():
        cant = ciu.get("materiales",{}).get(nombre, {})
        if isinstance(cant, dict):
            cant = cant.get("cantidad", 0)
        ws5.cell(fila, 4, cant)
        ws5.cell(fila, 5, "BUEN ESTADO" if cant > 0 else "NINGUNA")

    # ══ HOJA 6: Checklists MPRIU ═════════════════════════════════════
    ws6 = wb["Checklists MPRIU"]
    ws6["C4"]  = r["fecha"]
    ws6["F4"]  = r["hora_inicio"]
    ws6["H4"]  = r["hora_fin"]
    ws6["C5"]  = r["nombre_ruta"]
    ws6["C6"]  = r["nodo_inicial"]
    ws6["C7"]  = r["nodo_final"]
    ws6["C8"]  = ciu.get("distancia_ruta","")
    ws6["C9"]  = r["lider"]
    ws6["C10"] = ciu.get("vehiculo_placa","")
    ws6["C11"] = r["coordinador"]

    for row in ws6.iter_rows(min_row=13, max_row=ws6.max_row):
        for cell in row:
            if cell.column == 2 and cell.value and isinstance(cell.value, str):
                novedad_key = cell.value.strip().upper()
                for key, info in nch.items():
                    if key.upper() in novedad_key or novedad_key in key.upper():
                        ws6.cell(cell.row, 3, "SI" if info.get("check") else "NO")
                        ws6.cell(cell.row, 8, info.get("cantidad",0) if info.get("check") else 0)

    for row in ws6.iter_rows(min_row=ws6.max_row-5, max_row=ws6.max_row):
        for cell in row:
            if cell.value and "Observaciones" in str(cell.value):
                ws6.cell(cell.row, 3, r["observaciones"])

    # ══ HOJA 3: MANGAS ════════════════════════════════════════════════
    ws3 = wb["MANGAS"]
    mangas = datos.get("mangas",[])
    fila_m = 6
    for i in range(0, len(mangas), 2):
        m1 = mangas[i]
        m2 = mangas[i+1] if i+1 < len(mangas) else {}
        ws3.cell(fila_m,   3, m1.get("nombre",""))
        ws3.cell(fila_m,   5, m2.get("nombre",""))
        ws3.cell(fila_m+2, 3, m1.get("derivacion","NO"))
        ws3.cell(fila_m+2, 5, m2.get("derivacion",""))
        ws3.cell(fila_m+3, 3, m1.get("coordenadas",""))
        ws3.cell(fila_m+3, 5, m2.get("coordenadas",""))
        ws3.cell(fila_m+4, 3, m1.get("observacion",""))
        ws3.cell(fila_m+4, 5, m2.get("observacion",""))
        fila_m += 8

    # ══ HOJA 4: INVENTARIO DE HILOS EN NODO ══════════════════════════
    ws4 = wb["INVENTARIO DE HILOS EN NODO"]
    ws4["C4"] = r["nodo_final"]
    ws4["C5"] = datos["hilos"].get("posicion_odf","")
    hilos = datos["hilos"].get("filas",[])
    fila_h = 9
    for h in hilos:
        ws4.cell(fila_h, 3, h.get("descripcion",""))
        ws4.cell(fila_h, 4, h.get("estado",""))
        fila_h += 1

    # Agregar logos Telconet en todas las hojas
    try:
        wb_orig_logos = _get_plantilla()
        _copiar_logos(wb_orig_logos, wb)
    except Exception as e:
        logger.warning("No se pudo copiar logos: " + str(e))

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
    teclado = [
        ["🔍 Inspeccionar",    "🗺 Nueva Ruta Base"],
        ["📋 Generar Informe", "📁 Mis Rutas"],
        ["❓ Ayuda"]
    ]
    await update.message.reply_text(
        "📡 *RecorridosIA* — Menu principal",
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
    ctx.user_data["datos"]["ciu"]["herramientas"] = {}
    ctx.user_data["datos"]["ciu"]["equipos"] = {}
    ctx.user_data["datos"]["ciu"]["materiales"] = {}
    ctx.user_data["ciu_idx"] = 0

    msg = "CHECKLIST CIU - HERRAMIENTAS Y EPP\n\n"
    msg += "Indica cantidad de cada item (0 si no llevas):\n\n"
    for nombre, _ in HERRAMIENTAS_CIU:
        msg += "- " + nombre + ":\n"
    msg += "\nEnvia en formato:\n1, 2, 6, 1, 1, 2, 2, 2, 1, 2, 2, 2, 1"
    msg += "\n(un numero por cada item en orden)"
    await update.message.reply_text(msg, reply_markup=ReplyKeyboardRemove())
    return CIU_HERRAMIENTAS

async def recv_ciu_herramientas(update, ctx):
    datos = ctx.user_data["datos"]
    texto = update.message.text.strip()
    valores = [v.strip() for v in texto.replace(",", " ").split()]
    for i, (nombre, key) in enumerate(HERRAMIENTAS_CIU):
        cantidad = int(valores[i]) if i < len(valores) and valores[i].isdigit() else 0
        estado = "BUEN ESTADO" if cantidad > 0 else "NINGUNA"
        datos["ciu"]["herramientas"][nombre] = {"cantidad": cantidad, "obs": estado}

    msg = "CHECKLIST CIU - EQUIPOS ELECTRONICOS\n\n"
    msg += "Indica cantidad de cada equipo:\n\n"
    for nombre, _ in EQUIPOS_CIU:
        msg += "- " + nombre + ":\n"
    msg += "\nEnvia en formato:\n1, 2, 1, 1, 1, 1"
    await update.message.reply_text(msg)
    return CIU_EQUIPOS

async def recv_ciu_equipos(update, ctx):
    datos = ctx.user_data["datos"]
    texto = update.message.text.strip()
    valores = [v.strip() for v in texto.replace(",", " ").split()]
    for i, (nombre, key) in enumerate(EQUIPOS_CIU):
        cantidad = int(valores[i]) if i < len(valores) and valores[i].isdigit() else 0
        estado = "BUEN ESTADO" if cantidad > 0 else "NINGUNA"
        datos["ciu"]["equipos"][nombre] = {"cantidad": cantidad, "obs": estado}

    msg = "CHECKLIST CIU - MATERIALES E INSUMOS\n\n"
    msg += "Indica cantidad de cada material:\n\n"
    for nombre, _ in MATERIALES_CIU:
        msg += "- " + nombre + ":\n"
    msg += "\nEnvia en formato:\n335, 2, 1, 2, 10, 2"
    await update.message.reply_text(msg)
    return CIU_MATERIALES

async def recv_ciu_materiales(update, ctx):
    datos = ctx.user_data["datos"]
    texto = update.message.text.strip()
    valores = [v.strip() for v in texto.replace(",", " ").split()]
    for i, (nombre, key) in enumerate(MATERIALES_CIU):
        cantidad = int(valores[i]) if i < len(valores) and valores[i].isdigit() else 0
        estado = "BUEN ESTADO" if cantidad > 0 else "NINGUNA"
        datos["ciu"]["materiales"][nombre] = {"cantidad": cantidad, "obs": estado}

    await update.message.reply_text(
        "Checklist CIU completado\n\n"
        "Ahora envia las FOTOS de la inspeccion.\n"
        "Cuando termines escribe: LISTO"
    )
    return NOVEDADES_AUTO

async def recv_mpriu(update, ctx):
    datos = ctx.user_data["datos"]
    texto = update.message.text.strip().upper()
    if texto == "LISTO":
        return await recv_observaciones_mpriu(update, ctx)
    teclado = [["LISTO - terminar checklist"]]
    await update.message.reply_text(
        "Escribe LISTO para continuar o sigue marcando novedades.",
        reply_markup=ReplyKeyboardMarkup(teclado, resize_keyboard=True)
    )
    return MPRIU_CHECK

async def recv_observaciones_mpriu(update, ctx):
    datos = ctx.user_data["datos"]
    teclado = [["SI, hubo cambio de mangas", "No hubo cambio"]]
    await update.message.reply_text(
        "Checklist MPRIU completado\n\nHubo cambio de mangas?",
        reply_markup=ReplyKeyboardMarkup(teclado, resize_keyboard=True, one_time_keyboard=True)
    )
    return PREGUNTA_MANGAS

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
        datos["mpriu"]["observaciones"] = update.message.text.upper()
    datos["recorrido"]["hora_fin"] = datetime.now().strftime("%H:%M:%S")

    # Mostrar checklist MPRIU
    msg = "CHECKLIST MPRIU - Marca las novedades encontradas\n\n"
    msg += "La IA ya detecto estas novedades automaticamente:\n"
    novedades_ia = datos["mpriu"].get("novedades_check", {})
    for nov, vals in novedades_ia.items():
        msg += "SI - " + nov + " (cantidad: " + str(vals.get("cantidad",0)) + ")\n"
    msg += "\nSi hay novedades adicionales que la IA no detecto,\n"
    msg += "escribe el nombre exacto y cantidad. Ejemplo:\n"
    msg += "POSTES INCLINADOS: 2\n\n"
    msg += "Si no hay adicionales escribe: LISTO"
    teclado = [["LISTO - terminar checklist"]]
    await update.message.reply_text(msg, reply_markup=ReplyKeyboardMarkup(teclado, resize_keyboard=True))
    return MPRIU_CHECK

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

async def tab_menu(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Menu principal de pestanas con botones inline."""
    if update.effective_user.id not in USUARIOS_AUTENTICADOS:
        return await start(update, ctx)
    if "datos" not in ctx.user_data:
        ctx.user_data["datos"] = datos_vacios()

    datos = ctx.user_data["datos"]
    r = datos["recorrido"]

    ciu_ok   = bool(datos["ciu"].get("vehiculo_placa"))
    mpriu_ok = bool(datos["mpriu"].get("novedades_check"))
    rep_ok   = bool(r.get("nombre_ruta"))
    novedades = len(r.get("novedades", []))
    man_ok   = bool(datos.get("mangas"))
    hil_ok   = bool(datos["hilos"].get("filas"))

    def tick(ok): return "OK" if ok else "--"

    fotos_ok = bool(r.get("fotos_total", 0))
    teclado_inline = InlineKeyboardMarkup([
        [InlineKeyboardButton("Checklist CIU [" + tick(ciu_ok) + "]", callback_data="tab_1")],
        [InlineKeyboardButton("Checklists MPRIU [" + tick(mpriu_ok) + "]", callback_data="tab_2")],
        [InlineKeyboardButton("REPORTES_DE_RECORRIDOS [" + tick(rep_ok) + "]", callback_data="tab_reportes")],
        [InlineKeyboardButton("FOTOS_ANEXAS_AL_REPORTE [" + str(novedades) + " nov]", callback_data="tab_fotos")],
        [
            InlineKeyboardButton("Mangas [" + tick(man_ok) + "]", callback_data="tab_5"),
            InlineKeyboardButton("Hilos ODF [" + tick(hil_ok) + "]", callback_data="tab_6"),
        ],
        [InlineKeyboardButton("GENERAR EXCEL", callback_data="tab_generar")],
    ])

    msg = "INFORME FOR FO 02" + chr(10) + "Selecciona la pestana que quieres llenar:"

    if update.callback_query:
        await update.callback_query.answer()
        await update.callback_query.edit_message_text(msg, reply_markup=teclado_inline)
    else:
        await update.message.reply_text(msg, reply_markup=teclado_inline)
    return TAB_MENU

async def tab_callback(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Maneja los callbacks de los botones inline."""
    query = update.callback_query
    await query.answer()
    data = query.data

    if data == "tab_generar":
        await query.edit_message_text("Generando informe FOR FO 02...")
        return await enviar_excel(update, ctx)

    elif data == "tab_1":
        teclado = InlineKeyboardMarkup([
            [InlineKeyboardButton("Volver al menu", callback_data="tab_menu")]
        ])
        await query.edit_message_text(
            "CHECKLIST CIU - HERRAMIENTAS Y EPP" + chr(10) + chr(10) +
            "Indica cantidades separadas por coma:" + chr(10) +
            "Cinturon,Casco,Esc24,Esc28,Esc32,Conos,Caja," + chr(10) +
            "Destorn,Martillo,Estiletes,Cortafrio,Alicate,Llave," + chr(10) +
            "Rachet,Guantes,Tecle,Machete,Cizalla,Pata,Flejadora," + chr(10) +
            "Extension,Motosierra,Tijeras,Arco,Binoculares,Parasol,Remolque" + chr(10) + chr(10) +
            "Ejemplo: 2,2,0,2,0,6,0,1,1,2,2,0,0,1,2,2,2,1,0,0,0,0,0,0,0,0,0",
            reply_markup=teclado
        )
        ctx.user_data["tab_actual"] = "1"
        return TAB_CIU_HERR

    elif data == "tab_2":
        msg = "CHECKLIST MPRIU" + chr(10) + chr(10)
        msg += "Escribe los NUMEROS de las novedades (separados por coma):" + chr(10) + chr(10)
        for i, nov in enumerate(NOVEDADES_MPRIU, 1):
            msg += str(i) + ". " + nov + chr(10)
        msg += chr(10) + "Ejemplo: 16,18" + chr(10) + "Sin novedades: 31"
        teclado = InlineKeyboardMarkup([[InlineKeyboardButton("Volver al menu", callback_data="tab_menu")]])
        await query.edit_message_text(msg, reply_markup=teclado)
        ctx.user_data["tab_actual"] = "2"
        return TAB_MPRIU

    elif data == "tab_reportes":
        # REPORTES_DE_RECORRIDOS — hibrido
        teclado = InlineKeyboardMarkup([
            [
                InlineKeyboardButton("Manual (sin senial)", callback_data="rep_manual"),
                InlineKeyboardButton("Con IA (Gemini)", callback_data="rep_ia"),
            ],
            [InlineKeyboardButton("Volver al menu", callback_data="tab_menu")],
        ])
        await query.edit_message_text(
            "REPORTES_DE_RECORRIDOS" + chr(10) + chr(10) +
            "Elige como llenar esta pestana:",
            reply_markup=teclado
        )
        return TAB_MENU

    elif data == "tab_fotos":
        # FOTOS_ANEXAS_AL_REPORTE — hibrido
        teclado = InlineKeyboardMarkup([
            [
                InlineKeyboardButton("Manual (sin senial)", callback_data="rep_manual"),
                InlineKeyboardButton("Con IA (Gemini)", callback_data="rep_ia"),
            ],
            [InlineKeyboardButton("Volver al menu", callback_data="tab_menu")],
        ])
        await query.edit_message_text(
            "FOTOS_ANEXAS_AL_REPORTE" + chr(10) + chr(10) +
            "Elige como llenar esta pestana:" + chr(10) +
            "(Las fotos se agregan junto con los reportes de novedades)",
            reply_markup=teclado
        )
        return TAB_MENU

    elif data == "rep_manual":
        msg = (
            "REPORTES - Modo Manual" + chr(10) + chr(10) +
            "Ingresa los datos (uno por linea):" + chr(10) + chr(10) +
            "RUTA: GOSSEAL-MACHACHI   TAREA: 157415066" + chr(10) +
            "CUADRILLA: FO UIO INT 04" + chr(10) +
            "NODO_INI: GOSSEAL" + chr(10) +
            "NODO_FIN: MACHACHI" + chr(10) +
            "LIDER: RICHARD DAVID TAIPE COYAGO" + chr(10) +
            "AYUDANTE: JOSE LUIS ALLAICA CONDO" + chr(10) +
            "COORDINADOR: JUAN CARLOS YEPEZ ACAN" + chr(10) +
            "PLACA: PCO3940" + chr(10) +
            "DISTANCIA: 59KM" + chr(10) +
            "FECHA: HOY" + chr(10) +
            "HORA_INI: AHORA" + chr(10) +
            "HORA_FIN: AHORA" + chr(10) +
            "FOTOS: 6" + chr(10) +
            "OBS: texto o NINGUNA" + chr(10) + chr(10) +
            "Para novedades escribe:" + chr(10) +
            "NOV: VEGETACION SOBRE FIBRA/MANGA. | -0.477057,-78.579350" + chr(10) + chr(10) +
            "Cuando termines escribe: FIN"
        )
        teclado = InlineKeyboardMarkup([[InlineKeyboardButton("Volver al menu", callback_data="tab_menu")]])
        await query.edit_message_text(msg, reply_markup=teclado)
        ctx.user_data["tab_actual"] = "3"
        ctx.user_data["novedades_manuales"] = []
        return TAB_REPORTES

    elif data == "rep_ia":
        teclado = InlineKeyboardMarkup([[InlineKeyboardButton("Volver al menu", callback_data="tab_menu")]])
        await query.edit_message_text(
            "REPORTES + FOTOS con Gemini IA" + chr(10) + chr(10) +
            "Envia las fotos de la inspeccion." + chr(10) +
            "Gemini detectara automaticamente:" + chr(10) +
            "- Tipo de novedad (vegetacion, herrajes, etc)" + chr(10) +
            "- Remedio recomendado" + chr(10) +
            "- Coordenadas GPS del punto" + chr(10) + chr(10) +
            "Cuando termines de enviar fotos escribe: LISTO",
            reply_markup=teclado
        )
        ctx.user_data["tab_actual"] = "4"
        ctx.user_data["media_inspeccion"] = []
        return TAB_NOVEDADES_IA

    elif data == "tab_5":
        teclado = InlineKeyboardMarkup([[InlineKeyboardButton("Volver al menu", callback_data="tab_menu")]])
        await query.edit_message_text(
            "MANGAS" + chr(10) + chr(10) +
            "Ingresa cada manga en formato:" + chr(10) +
            "NOMBRE | DERIVACION | COORDENADAS | OBSERVACION" + chr(10) + chr(10) +
            "Ejemplo:" + chr(10) +
            "UIO-B-MAC/GOS-F1-DER-01 | NO | -0.477057,-78.579350 | NINGUNA" + chr(10) + chr(10) +
            "Cuando termines escribe: FIN MANGAS",
            reply_markup=teclado
        )
        ctx.user_data["tab_actual"] = "5"
        ctx.user_data["datos"]["mangas"] = []
        return MANGA_NOMBRE

    elif data == "tab_6":
        teclado = InlineKeyboardMarkup([[InlineKeyboardButton("Volver al menu", callback_data="tab_menu")]])
        await query.edit_message_text(
            "INVENTARIO DE HILOS EN NODO" + chr(10) + chr(10) +
            "Posicion del ODF:" + chr(10) +
            "Ejemplo: MAC-GOS-F01-R04-ODF02-48",
            reply_markup=teclado
        )
        ctx.user_data["tab_actual"] = "6"
        return HILO_ODF

    elif data == "tab_menu":
        return await tab_menu(update, ctx)

    return TAB_MENU

async def tab_selector(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Redirige texto al menu de pestanas."""
    txt = update.message.text.strip().upper()
    if txt == "GENERAR":
        return await enviar_excel(update, ctx)
    if txt == "CANCELAR":
        return await menu_principal(update, ctx)
    return await tab_menu(update, ctx)

async def tab_ciu_herr(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Procesa herramientas CIU."""
    valores = [v.strip() for v in update.message.text.replace(",", " ").split()]
    nombres = [
        "Cinturon y Linea de Vida","Casco","Escalera de 24 pies","Escalera de 28 pies",
        "Escalera de 32 pies","Conos reflectivos","Caja para herramientas",
        "Juego de destornilladores","Martillo mediano","Estiletes","Cortafrio",
        "Alicate","Llave francesa","Juego de rachet","Pares de guantes aislantes",
        "Tecle","Machete","Cizalla","Pata de cabra","Flejadora (Maquina Eriband)",
        "Extension con foco","Motosierra","Tijeras metalicas","Arco de sierra",
        "Binoculares","Parasol","Remolque / Carrete para F.O.",
    ]
    herr = {}
    for i, nombre in enumerate(nombres):
        cant = int(valores[i]) if i < len(valores) and valores[i].isdigit() else 0
        herr[nombre] = {"cantidad": cant, "obs": "BUEN ESTADO" if cant > 0 else "NINGUNA"}
    ctx.user_data["datos"]["ciu"]["herramientas"] = herr

    await update.message.reply_text(
        "Herramientas guardadas!\n\n"
        "EQUIPOS ELECTRONICOS - indica cantidades separadas por coma:\n\n"
        "Fusionadora,CortadoraFibra,BobinaLanz,OTDR,LlaveAcsys,GPS,Inversor,Etiquetadora\n\n"
        "Ejemplo: 1,2,0,1,1,0,1,1"
    )
    return TAB_CIU_EQUI

async def tab_ciu_equi(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Procesa equipos electronicos CIU."""
    valores = [v.strip() for v in update.message.text.replace(",", " ").split()]
    nombres = ["Fusionadora","Cortadora de fibra","Bobina de lanzamiento",
               "OTDR con cargador","Llave Acsys","GPS","Inversor","Etiquetadora"]
    equi = {}
    for i, nombre in enumerate(nombres):
        cant = int(valores[i]) if i < len(valores) and valores[i].isdigit() else 0
        equi[nombre] = {"cantidad": cant, "obs": "BUEN ESTADO" if cant > 0 else "NINGUNA"}
    ctx.user_data["datos"]["ciu"]["equipos"] = equi

    await update.message.reply_text(
        "Equipos guardados!\n\n"
        "MATERIALES E INSUMOS - indica cantidades separadas por coma:\n\n"
        "Fibra500m,Mangas,Eriband,Hebillas,HojasSierra,Patchcord,\n"
        "Adaptadores,Amarras,MesasP,SillasP,Cuchillos,Poleas,\n"
        "SogasMedianas,SogasGruesas,RepelInsect,RepelAbejas\n\n"
        "Ejemplo: 335,2,1,6,0,2,10,2,0,0,0,1,1,0,0,0"
    )
    return TAB_CIU_MATE

async def tab_ciu_mate(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Procesa materiales CIU."""
    valores = [v.strip() for v in update.message.text.replace(",", " ").split()]
    nombres = ["Fibra 48h (500mt)","Mangas de 48h y/o 144h (2 minimo)",
               "Rollo de cinta Eriband 3/4","Hebillas para cinta Eriband 3/4",
               "Hojas de sierra","Patchcord de fibra","Adaptadores (Simplex-Duplex)",
               "Paquetes de amarras","Mesas plasticas","Sillas plasticas",
               "Cuchillos","Poleas","Sogas de nylon medianas",
               "Sogas de nylon gruesas","Repelente contra insectos",
               "Repelente contra abejas y avispas"]
    mate = {}
    for i, nombre in enumerate(nombres):
        cant = int(valores[i]) if i < len(valores) and valores[i].isdigit() else 0
        mate[nombre] = {"cantidad": cant, "obs": "BUEN ESTADO" if cant > 0 else "NINGUNA"}
    ctx.user_data["datos"]["ciu"]["materiales"] = mate

    await update.message.reply_text(
        "Materiales guardados!\n\n"
        "Checklist CIU completo. Volviendo al menu de pestanas..."
    )
    return await tab_menu(update, ctx)

async def tab_mpriu(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Procesa checklist MPRIU por numeros."""
    txt = update.message.text.strip()
    numeros = [n.strip() for n in txt.replace(",", " ").split() if n.strip().isdigit()]
    nch = {}
    for num_str in numeros:
        idx = int(num_str) - 1
        if 0 <= idx < len(NOVEDADES_MPRIU):
            novedad = NOVEDADES_MPRIU[idx]
            nch[novedad] = {"check": True, "cantidad": 1}

    ctx.user_data["datos"]["mpriu"]["novedades_check"] = nch

    cant = len(nch)
    if cant == 1 and "NO SE REGISTRAN" in list(nch.keys())[0]:
        cant = 0

    await update.message.reply_text(
        str(cant) + " novedad(es) marcada(s) en MPRIU.\n\n"
        "Ahora ingresa la CANTIDAD exacta de cada novedad.\n"
        "Escribe en formato NUMERO:CANTIDAD (separados por coma)\n\n"
        "Ejemplo: 16:5,18:1\n"
        "Si las cantidades son correctas escribe: OK"
    )
    return TAB_MPRIU

async def tab_mpriu_cantidades(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Actualiza cantidades del MPRIU."""
    txt = update.message.text.strip().upper()
    if txt != "OK":
        pares = [p.strip() for p in txt.replace(",", " ").split()]
        nch = ctx.user_data["datos"]["mpriu"].get("novedades_check", {})
        for par in pares:
            if ":" in par:
                num_str, cant_str = par.split(":")
                if num_str.isdigit() and cant_str.isdigit():
                    idx = int(num_str) - 1
                    if 0 <= idx < len(NOVEDADES_MPRIU):
                        novedad = NOVEDADES_MPRIU[idx]
                        if novedad in nch:
                            nch[novedad]["cantidad"] = int(cant_str)
        ctx.user_data["datos"]["mpriu"]["novedades_check"] = nch

    await update.message.reply_text("Checklist MPRIU guardado!")
    return await tab_menu(update, ctx)

async def tab_reportes(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Procesa datos de REPORTES_DE_RECORRIDOS en modo manual."""
    lineas = update.message.text.strip().split("\n")
    datos = ctx.user_data["datos"]
    r = datos["recorrido"]

    for linea in lineas:
        if ":" not in linea:
            continue
        clave, valor = linea.split(":", 1)
        clave = clave.strip().upper()
        valor = valor.strip()

        if valor.upper() == "HOY":
            valor = datetime.now().strftime("%d/%m/%Y")
        elif valor.upper() == "AHORA":
            valor = datetime.now().strftime("%H:%M:%S")

        if clave == "RUTA":
            r["nombre_ruta"] = valor.upper()
        elif clave == "CUADRILLA":
            r["codigo_cuadrilla"] = valor.upper()
        elif clave == "NODO_INI":
            r["nodo_inicial"] = valor.upper()
        elif clave == "NODO_FIN":
            r["nodo_final"] = valor.upper()
        elif clave == "LIDER":
            r["lider"] = valor.upper()
        elif clave == "AYUDANTE":
            r["ayudante"] = valor.upper()
        elif clave == "COORDINADOR":
            r["coordinador"] = valor.upper()
        elif clave == "PLACA":
            datos["ciu"]["vehiculo_placa"] = valor.upper()
        elif clave == "DISTANCIA":
            datos["ciu"]["distancia_ruta"] = valor.upper()
        elif clave == "FECHA":
            r["fecha"] = valor
        elif clave == "HORA_INI":
            r["hora_inicio"] = valor
        elif clave == "HORA_FIN":
            r["hora_fin"] = valor
        elif clave == "FOTOS":
            r["fotos_total"] = int(valor) if valor.isdigit() else 0
        elif clave == "OBS":
            if valor.upper() != "NINGUNA":
                r["observaciones"] = valor.upper()

    # Si no hay novedades aun, agregar sin novedad
    if not r.get("novedades"):
        nov = novedad_vacia(1)
        nov["motivo"]  = "NO SE REGISTRAN NOVEDADES DURANTE LA INSPECCION."
        nov["remedio"] = "NO SE ENCUENTRAN NOVEDADES QUE SIGNIFIQUEN RIESGOS EN EL CABLE DE LA RED INTERURBANO."
        r["novedades"] = [nov]

    # Procesar novedades manuales si vienen en el mismo mensaje
    novedades_nuevas = []
    for linea in lineas:
        if linea.upper().startswith("NOV:"):
            partes = linea[4:].strip().split("|")
            motivo = partes[0].strip().upper()
            coords = partes[1].strip() if len(partes) > 1 else ""
            REMED = {
                "VEGETACION": "REALIZAR LA PODA O RETIRO DE VEGETACION.",
                "HERRAJES": "REALIZAR EL REEMPLAZO INMEDIATO DEL HERRAJE.",
                "POSTES": "DOCUMENTAR Y REPORTAR PARA GESTIONAR EL REEMPLAZO.",
                "MANGAS": "ASEGURAR LA MANGA EN CONFIGURACION TIPO FIGURA 8.",
                "CABLE": "DOCUMENTAR E INFORMAR PARA PROGRAMAR EL CAMBIO.",
                "DOCUMENTACION": "DOCUMENTAR O SOLICITAR PROGRAMACION; USAR SEGUIDOR DE SENAL.",
            }
            remedio = "DOCUMENTAR Y REPORTAR AL COORDINADOR."
            for key, rem in REMED.items():
                if key in motivo:
                    remedio = rem
                    break
            nov = novedad_vacia(len(novedades_nuevas)+1)
            nov["motivo"] = motivo
            nov["remedio"] = remedio
            nov["coordenadas"] = coords
            novedades_nuevas.append(nov)
            datos["mpriu"]["novedades_check"][motivo] = {"check": True, "cantidad": 1}
        elif linea.upper() == "FIN" and novedades_nuevas:
            break

    if novedades_nuevas:
        r["novedades"] = novedades_nuevas
    elif not r.get("novedades"):
        nov = novedad_vacia(1)
        nov["motivo"] = "NO SE REGISTRAN NOVEDADES DURANTE LA INSPECCION."
        nov["remedio"] = "NO SE ENCUENTRAN NOVEDADES QUE SIGNIFIQUEN RIESGOS."
        r["novedades"] = [nov]

    await update.message.reply_text(
        "Datos guardados!" + chr(10) +
        "Ruta: " + r.get("nombre_ruta","") + chr(10) +
        "Novedades: " + str(len(r.get("novedades",[])))
    )
    return await tab_menu(update, ctx)

async def tab_novedades_ia(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    """Recibe fotos y analiza con Gemini IA."""
    if "media_inspeccion" not in ctx.user_data:
        ctx.user_data["media_inspeccion"] = []

    if update.message.photo:
        foto = await update.message.photo[-1].get_file()
        ctx.user_data["media_inspeccion"].append(bytes(await foto.download_as_bytearray()))
        n = len(ctx.user_data["media_inspeccion"])
        await update.message.reply_text(
            "Foto " + str(n) + " recibida. Envia mas o escribe LISTO"
        )
        return TAB_NOVEDADES_IA

    if update.message.text and update.message.text.upper() == "LISTO":
        await update.message.reply_text("Analizando con Gemini IA...")
        media = ctx.user_data.get("media_inspeccion", [])
        novedades = []
        for img in media:
            r_ia = await analizar_imagen(img)
            if r_ia:
                n = novedad_vacia(len(novedades)+1)
                n.update(r_ia)
                novedades.append(n)
        if not novedades:
            n = novedad_vacia(1)
            n["motivo"]  = SIN_NOVEDAD_MOTIVO
            n["remedio"] = SIN_NOVEDAD_REMEDIO
            novedades = [n]
        ctx.user_data["datos"]["recorrido"]["novedades"] = novedades
        for nov in novedades:
            m = nov["motivo"]
            if m != SIN_NOVEDAD_MOTIVO:
                ctx.user_data["datos"]["mpriu"]["novedades_check"][m] = {"check": True, "cantidad": ctx.user_data["datos"]["mpriu"]["novedades_check"].get(m, {}).get("cantidad",0)+1}
        ctx.user_data["datos"]["recorrido"]["fotos_total"] = len(media)
        await update.message.reply_text(
            str(len(novedades)) + " novedad(es) detectadas por la IA!"
        )
        return await tab_menu(update, ctx)

    return TAB_NOVEDADES_IA

async def generar_informe(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in USUARIOS_AUTENTICADOS:
        return await start(update, ctx)
    if "datos" not in ctx.user_data:
        ctx.user_data["datos"] = datos_vacios()
    return await tab_menu(update, ctx)

async def generar_confirmar(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    txt = update.message.text
    if "Automatico" in txt:
        return await inspeccionar(update, ctx)
    elif "Manual" in txt:
        ctx.user_data["datos"] = datos_vacios()
        ctx.user_data["modo_manual"] = True
        ctx.user_data["novedades_manuales"] = []
        await update.message.reply_text(
            "Modo Manual\n\nVoy a pedirte los datos del recorrido.\n\nNombre de la ruta:\nEjemplo: GOSSEAL-MACHACHI   TAREA: 157415066",
            reply_markup=ReplyKeyboardRemove()
        )
        return GENERAR_NOMBRE_RUTA
    return await menu_principal(update, ctx)

async def gm_nombre_ruta(update, ctx):
    ctx.user_data["datos"]["recorrido"]["nombre_ruta"] = update.message.text.upper()
    await update.message.reply_text("Codigo de cuadrilla:\nEjemplo: FO UIO INT 04")
    return GENERAR_CUADRILLA

async def gm_cuadrilla(update, ctx):
    ctx.user_data["datos"]["recorrido"]["codigo_cuadrilla"] = update.message.text.upper()
    await update.message.reply_text("Nodo inicial:\nEjemplo: GOSSEAL")
    return GENERAR_NODO_INI

async def gm_nodo_ini(update, ctx):
    ctx.user_data["datos"]["recorrido"]["nodo_inicial"] = update.message.text.upper()
    await update.message.reply_text("Nodo final:\nEjemplo: MACHACHI")
    return GENERAR_NODO_FIN

async def gm_nodo_fin(update, ctx):
    ctx.user_data["datos"]["recorrido"]["nodo_final"] = update.message.text.upper()
    await update.message.reply_text("Lider de cuadrilla:\nEjemplo: RICHARD DAVID TAIPE COYAGO")
    return GENERAR_LIDER

async def gm_lider(update, ctx):
    ctx.user_data["datos"]["recorrido"]["lider"] = update.message.text.upper()
    await update.message.reply_text("Ayudante tecnico:\nEjemplo: JOSE LUIS ALLAICA CONDO")
    return GENERAR_AYUDANTE

async def gm_ayudante(update, ctx):
    ctx.user_data["datos"]["recorrido"]["ayudante"] = update.message.text.upper()
    await update.message.reply_text("Coordinador de fibra optica:\nEjemplo: JUAN CARLOS YEPEZ ACAN")
    return GENERAR_COORDINADOR

async def gm_coordinador(update, ctx):
    ctx.user_data["datos"]["recorrido"]["coordinador"] = update.message.text.upper()
    await update.message.reply_text("Placa del vehiculo:\nEjemplo: PCO3940")
    return GENERAR_PLACA

async def gm_placa(update, ctx):
    ctx.user_data["datos"]["ciu"]["vehiculo_placa"] = update.message.text.upper()
    await update.message.reply_text("Distancia de la ruta:\nEjemplo: 59KM")
    return GENERAR_DISTANCIA

async def gm_distancia(update, ctx):
    ctx.user_data["datos"]["ciu"]["distancia_ruta"] = update.message.text.upper()
    await update.message.reply_text(
        "Fecha del recorrido:\nEjemplo: 26/06/2026\nEscribe HOY para fecha actual"
    )
    return GENERAR_FECHA

async def gm_fecha(update, ctx):
    txt = update.message.text.upper()
    if txt == "HOY":
        txt = datetime.now().strftime("%d/%m/%Y")
    ctx.user_data["datos"]["recorrido"]["fecha"] = txt
    await update.message.reply_text(
        "Hora de inicio:\nEjemplo: 08:00\nEscribe AHORA para hora actual"
    )
    return GENERAR_HORA_INI

async def gm_hora_ini(update, ctx):
    txt = update.message.text.upper()
    if txt == "AHORA":
        txt = datetime.now().strftime("%H:%M:%S")
    ctx.user_data["datos"]["recorrido"]["hora_inicio"] = txt
    await update.message.reply_text(
        "Hora de fin:\nEjemplo: 10:30\nEscribe AHORA para hora actual"
    )
    return GENERAR_HORA_FIN

async def gm_hora_fin(update, ctx):
    txt = update.message.text.upper()
    if txt == "AHORA":
        txt = datetime.now().strftime("%H:%M:%S")
    ctx.user_data["datos"]["recorrido"]["hora_fin"] = txt
    await update.message.reply_text(
        "Novedades del recorrido:\n\n"
        "Escribe cada novedad en formato:\n"
        "MOTIVO | COORDENADAS\n\n"
        "Ejemplo:\n"
        "VEGETACION SOBRE FIBRA/MANGA. | -0.477057,-78.579350\n\n"
        "Cuando termines escribe: LISTO\n"
        "Si no hay novedades escribe: NINGUNA"
    )
    ctx.user_data["novedades_manuales"] = []
    return GENERAR_NOVEDADES

async def gm_novedades(update, ctx):
    txt = update.message.text.upper()
    if txt == "NINGUNA":
        nov = novedad_vacia(1)
        nov["motivo"]  = "NO SE REGISTRAN NOVEDADES DURANTE LA INSPECCION."
        nov["remedio"] = "NO SE ENCUENTRAN NOVEDADES QUE SIGNIFIQUEN RIESGOS EN EL CABLE DE LA RED INTERURBANO."
        ctx.user_data["datos"]["recorrido"]["novedades"] = [nov]
        return await _finalizar_manual(update, ctx)
    if txt == "LISTO":
        novedades = ctx.user_data.get("novedades_manuales", [])
        if not novedades:
            nov = novedad_vacia(1)
            nov["motivo"]  = "NO SE REGISTRAN NOVEDADES DURANTE LA INSPECCION."
            nov["remedio"] = "NO SE ENCUENTRAN NOVEDADES QUE SIGNIFIQUEN RIESGOS EN EL CABLE DE LA RED INTERURBANO."
            novedades = [nov]
        ctx.user_data["datos"]["recorrido"]["novedades"] = novedades
        return await _finalizar_manual(update, ctx)
    partes = update.message.text.split("|")
    motivo = partes[0].strip().upper()
    coords = partes[1].strip() if len(partes) > 1 else ""
    REMEDIOS_MAP = {
        "VEGETACION": "REALIZAR LA PODA O RETIRO DE VEGETACION QUE COMPROMETA LA INTEGRIDAD DEL CABLE.",
        "HERRAJES": "REALIZAR EL REEMPLAZO INMEDIATO DEL HERRAJE AFECTADO.",
        "POSTES": "DOCUMENTAR Y REPORTAR PARA GESTIONAR EL REEMPLAZO.",
        "MANGAS": "ASEGURAR LA MANGA AL POSTE EN CONFIGURACION TIPO FIGURA 8.",
        "CABLE": "DOCUMENTAR E INFORMAR PARA PROGRAMAR EL CAMBIO DEL TRAMO.",
        "DOCUMENTACION": "DOCUMENTAR O SOLICITAR PROGRAMACION DE TRABAJO; UTILIZAR SEGUIDOR DE SENAL.",
    }
    remedio = "DOCUMENTAR Y REPORTAR AL COORDINADOR."
    for key, rem in REMEDIOS_MAP.items():
        if key in motivo:
            remedio = rem
            break
    idx = len(ctx.user_data["novedades_manuales"]) + 1
    nov = novedad_vacia(idx)
    nov["motivo"]      = motivo
    nov["remedio"]     = remedio
    nov["coordenadas"] = coords
    ctx.user_data["novedades_manuales"].append(nov)
    ctx.user_data["datos"]["mpriu"]["novedades_check"][motivo] = {"check": True, "cantidad": 1}
    await update.message.reply_text(
        "Novedad #" + str(idx) + " guardada.\n\nEscribe otra o escribe LISTO"
    )
    return GENERAR_NOVEDADES

async def _finalizar_manual(update, ctx):
    teclado = [["SI, hubo cambio de mangas", "No hubo cambio"]]
    await update.message.reply_text(
        str(len(ctx.user_data["datos"]["recorrido"]["novedades"])) + " novedad(es) registrada(s)\n\nHubo cambio de mangas?",
        reply_markup=ReplyKeyboardMarkup(teclado, resize_keyboard=True, one_time_keyboard=True)
    )
    return PREGUNTA_MANGAS

async def nueva_ruta(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in USUARIOS_AUTENTICADOS:
        return await start(update, ctx)
    await update.message.reply_text(
        "Nueva Ruta Base\n\n"
        "Este comando registra el video base de una ruta.\n\n"
        "Nombre de la nueva ruta?\n"
        "Ejemplo: GOSSEAL-MACHACHI",
        reply_markup=ReplyKeyboardRemove()
    )
    return NUEVA_RUTA_NOMBRE

async def recv_nueva_ruta_nombre(update, ctx):
    ctx.user_data["nueva_ruta_nombre"] = update.message.text.upper()
    await update.message.reply_text(
        "Ruta: " + ctx.user_data["nueva_ruta_nombre"] + "\n\n"
        "Ahora graba el video con tu Insta360 y subelo a Mapillary.\n\n"
        "Cuando termines pega el LINK de Mapillary aqui.\n"
        "Ejemplo: https://www.mapillary.com/app/user/xxx?pKey=xxx\n\n"
        "O si quieres subir el video directo, envialo aqui y el bot lo procesara."
    )
    return NUEVA_RUTA_VIDEO

async def recv_nueva_ruta_video(update, ctx):
    nombre = ctx.user_data.get("nueva_ruta_nombre", "SIN NOMBRE")
    
    if update.message.text and update.message.text.startswith("http"):
        # Guarda el link de Mapillary
        RUTAS_GUARDADAS[nombre] = {
            "nombre": nombre,
            "mapillary_link": update.message.text.strip(),
            "tipo": "mapillary",
            "fecha": datetime.now().strftime("%d/%m/%Y %H:%M")
        }
        await update.message.reply_text(
            "Ruta base guardada\n\n"
            "Nombre: " + nombre + "\n"
            "Link: " + update.message.text.strip() + "\n\n"
            "Ya puedes usar esta ruta en futuras inspecciones."
        )
    elif update.message.video or update.message.document:
        # Guarda referencia del video enviado
        RUTAS_GUARDADAS[nombre] = {
            "nombre": nombre,
            "tipo": "video_telegram",
            "fecha": datetime.now().strftime("%d/%m/%Y %H:%M")
        }
        await update.message.reply_text(
            "Video recibido\n\n"
            "Ruta base guardada: " + nombre + "\n"
            "Fecha: " + datetime.now().strftime("%d/%m/%Y %H:%M") + "\n\n"
            "Ya puedes usar esta ruta en futuras inspecciones."
        )
    else:
        await update.message.reply_text(
            "Envia el link de Mapillary o el video de la ruta.\n"
            "Ejemplo link: https://www.mapillary.com/..."
        )
        return NUEVA_RUTA_VIDEO

    teclado = [["Inspeccionar", "Nueva Ruta Base"], ["Mis Rutas", "Ayuda"]]
    await update.message.reply_text(
        "Que deseas hacer?",
        reply_markup=ReplyKeyboardMarkup(teclado, resize_keyboard=True)
    )
    return MENU_PRINCIPAL

async def mis_rutas(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in USUARIOS_AUTENTICADOS:
        return await start(update, ctx)
    if not RUTAS_GUARDADAS:
        await update.message.reply_text(
            "No tienes rutas base guardadas todavia.\n\n"
            "Usa Nueva Ruta Base para registrar tu primera ruta."
        )
        return MENU_PRINCIPAL
    
    msg = "Rutas base registradas:\n\n"
    for i, (nombre, info) in enumerate(RUTAS_GUARDADAS.items()):
        msg += str(i+1) + ". " + nombre + "\n"
        msg += "   Fecha: " + info.get("fecha","") + "\n"
        msg += "   Tipo: " + info.get("tipo","") + "\n"
        if info.get("mapillary_link"):
            msg += "   Link: " + info["mapillary_link"][:50] + "...\n"
        msg += "\n"
    
    await update.message.reply_text(msg)
    return MENU_PRINCIPAL

async def ayuda(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "RecorridosIA - Ayuda\n\n"
        "COMANDOS DISPONIBLES:\n\n"
        "Inspeccionar\n"
        "  Inicia una inspeccion de ruta.\n"
        "  La IA analiza fotos y genera el informe\n"
        "  FOR FO 02 automaticamente.\n\n"
        "Nueva Ruta Base\n"
        "  Registra el video base de una ruta.\n"
        "  Graba con Insta360, sube a Mapillary\n"
        "  y pega el link aqui.\n\n"
        "Mis Rutas\n"
        "  Lista todas las rutas base guardadas.\n\n"
        "Ayuda\n"
        "  Muestra este mensaje.\n\n"
        "VARIABLES DE ENTORNO EN RENDER:\n"
        "BOT_TOKEN / GEMINI_API_KEY\n"
        "MAPILLARY_TOKEN / TOTP_SECRET\n"
        "DOMINIO_EMAIL / RENDER_EXTERNAL_URL"
    )
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

def ping_render():
    """Hace ping cada 4 minutos a si mismo para no dormirse."""
    import urllib.request
    while True:
        time.sleep(720)
        try:
            url = os.getenv("RENDER_EXTERNAL_URL", "")
            if url:
                urllib.request.urlopen(url, timeout=10)
                logger.info("Ping OK - bot despierto")
        except Exception as e:
            logger.warning("Ping error: " + str(e))

def start_web():
    port = int(os.getenv("PORT", 8080))
    server = HTTPServer(("0.0.0.0", port), PingHandler)
    logger.info("Servidor web en puerto " + str(port))
    # Ping en hilo separado
    t = threading.Thread(target=ping_render, daemon=True)
    t.start()
    server.serve_forever()

def build_app():
    app = Application.builder().token(BOT_TOKEN).build()
    conv = ConversationHandler(
        entry_points=[
            CommandHandler("start", start),
            CommandHandler("inspeccionar", inspeccionar),
            CommandHandler("nueva_ruta", nueva_ruta),
            CommandHandler("generar", generar_informe),
            MessageHandler(filters.Regex("📋 Generar Informe"), generar_informe),
            CommandHandler("mis_rutas", mis_rutas),
            CommandHandler("ayuda", ayuda),
            MessageHandler(filters.Regex("Inspeccionar"), inspeccionar),
            MessageHandler(filters.Regex("Nueva Ruta Base"), nueva_ruta),
            MessageHandler(filters.Regex("Mis Rutas"), mis_rutas),
            MessageHandler(filters.Regex("Ayuda"), ayuda),
        ],
        states={
            ESPERANDO_TOTP:   [MessageHandler(filters.TEXT & ~filters.COMMAND, handler_totp)],
            MENU_PRINCIPAL:   [
                MessageHandler(filters.Regex("Inspeccionar"), inspeccionar),
                MessageHandler(filters.Regex("Nueva Ruta Base"), nueva_ruta),
                MessageHandler(filters.Regex("Generar Informe"), generar_informe),
            MessageHandler(filters.Regex("Generar"), generar_informe),
                MessageHandler(filters.Regex("Mis Rutas"), mis_rutas),
                MessageHandler(filters.Regex("Ayuda"), ayuda),
                MessageHandler(filters.TEXT & ~filters.COMMAND, menu_principal),
            ],
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
            NUEVA_RUTA_NOMBRE: [MessageHandler(filters.TEXT & ~filters.COMMAND, recv_nueva_ruta_nombre)],
            NUEVA_RUTA_VIDEO:  [MessageHandler(filters.TEXT | filters.VIDEO | filters.Document.ALL & ~filters.COMMAND, recv_nueva_ruta_video)],
            GENERAR_CONFIRMAR:   [MessageHandler(filters.TEXT & ~filters.COMMAND, generar_confirmar)],
            TAB_MENU:            [MessageHandler(filters.TEXT & ~filters.COMMAND, tab_selector)],
            TAB_CIU_HERR:        [MessageHandler(filters.TEXT & ~filters.COMMAND, tab_ciu_herr)],
            TAB_CIU_EQUI:        [MessageHandler(filters.TEXT & ~filters.COMMAND, tab_ciu_equi)],
            TAB_CIU_MATE:        [MessageHandler(filters.TEXT & ~filters.COMMAND, tab_ciu_mate)],
            TAB_MPRIU:           [MessageHandler(filters.TEXT & ~filters.COMMAND, tab_mpriu), MessageHandler(filters.TEXT & ~filters.COMMAND, tab_mpriu_cantidades)],
            TAB_REPORTES:        [MessageHandler(filters.TEXT & ~filters.COMMAND, tab_reportes)],
            TAB_NOVEDADES_IA:    [MessageHandler(filters.PHOTO, tab_novedades_ia), MessageHandler(filters.TEXT & ~filters.COMMAND, tab_novedades_ia)],
            GENERAR_NOMBRE_RUTA: [MessageHandler(filters.TEXT & ~filters.COMMAND, gm_nombre_ruta)],
            GENERAR_CUADRILLA:   [MessageHandler(filters.TEXT & ~filters.COMMAND, gm_cuadrilla)],
            GENERAR_NODO_INI:    [MessageHandler(filters.TEXT & ~filters.COMMAND, gm_nodo_ini)],
            GENERAR_NODO_FIN:    [MessageHandler(filters.TEXT & ~filters.COMMAND, gm_nodo_fin)],
            GENERAR_LIDER:       [MessageHandler(filters.TEXT & ~filters.COMMAND, gm_lider)],
            GENERAR_AYUDANTE:    [MessageHandler(filters.TEXT & ~filters.COMMAND, gm_ayudante)],
            GENERAR_COORDINADOR: [MessageHandler(filters.TEXT & ~filters.COMMAND, gm_coordinador)],
            GENERAR_PLACA:       [MessageHandler(filters.TEXT & ~filters.COMMAND, gm_placa)],
            GENERAR_DISTANCIA:   [MessageHandler(filters.TEXT & ~filters.COMMAND, gm_distancia)],
            GENERAR_FECHA:       [MessageHandler(filters.TEXT & ~filters.COMMAND, gm_fecha)],
            GENERAR_HORA_INI:    [MessageHandler(filters.TEXT & ~filters.COMMAND, gm_hora_ini)],
            GENERAR_HORA_FIN:    [MessageHandler(filters.TEXT & ~filters.COMMAND, gm_hora_fin)],
            GENERAR_NOVEDADES:   [MessageHandler(filters.TEXT & ~filters.COMMAND, gm_novedades)],
            CIU_HERRAMIENTAS: [MessageHandler(filters.TEXT & ~filters.COMMAND, recv_ciu_herramientas)],
            CIU_EQUIPOS:      [MessageHandler(filters.TEXT & ~filters.COMMAND, recv_ciu_equipos)],
            CIU_MATERIALES:   [MessageHandler(filters.TEXT & ~filters.COMMAND, recv_ciu_materiales)],
            OBSERVACIONES:    [MessageHandler(filters.TEXT & ~filters.COMMAND, recv_observaciones)],
            MPRIU_CHECK:      [MessageHandler(filters.TEXT & ~filters.COMMAND, recv_mpriu)],
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
    app.add_handler(CallbackQueryHandler(tab_callback))
    return app

async def run_bot():
    import asyncio
    app = build_app()
    await app.initialize()
    await app.start()
    await app.updater.start_polling()
    logger.info("RecorridosIA bot arrancando...")
    while True:
        await asyncio.sleep(1)

def bot_thread():
    import asyncio
    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    loop.run_until_complete(run_bot())

if __name__ == "__main__":
    # Bot en hilo secundario con su propio event loop
    t = threading.Thread(target=bot_thread, daemon=True)
    t.start()
    logger.info("RecorridosIA bot arrancando...")
    # Servidor web en hilo PRINCIPAL — Render no mata el proceso
    start_web()
