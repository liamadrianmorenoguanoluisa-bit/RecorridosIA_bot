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
    """Carga la plantilla FOR FO 02 — si no existe crea un workbook basico."""
    import os
    from openpyxl import load_workbook, Workbook
    rutas = [
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "plantilla_FOR_FO_02.xlsx"),
        "plantilla_FOR_FO_02.xlsx",
        "/opt/render/project/src/plantilla_FOR_FO_02.xlsx",
    ]
    for ruta in rutas:
        if os.path.exists(ruta):
            logger.info("Plantilla encontrada: " + ruta)
            return load_workbook(ruta)

    # Fallback: crear workbook con las hojas necesarias
    logger.warning("Plantilla no encontrada, usando workbook basico")
    wb = Workbook()
    wb.active.title = "REPORTES_DE_RECORRIDOS"
    for nombre in ["FOTOS_ANEXAS_AL_REPORTE","MANGAS","INVENTARIO DE HILOS EN NODO","Checklist CIU","Checklists MPRIU"]:
        wb.create_sheet(nombre)
    return wb

def generar_excel(datos):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    r   = datos["recorrido"]
    ciu = datos["ciu"]
    nch = datos["mpriu"].get("novedades_check", {})

    wb  = Workbook()

    # ── HOJA 1: REPORTES_DE_RECORRIDOS ──────────────────────────────
    ws1 = wb.active
    ws1.title = "REPORTES_DE_RECORRIDOS"
    ws1.column_dimensions["A"].width = 41
    ws1.column_dimensions["B"].width = 35
    ws1.column_dimensions["C"].width = 32
    ws1.column_dimensions["D"].width = 30

    AZUL = "0000FF"; GRIS = "969696"; GRIS2 = "D9D9D9"; GRIS3 = "C0C0C0"
    AZUL2 = "0070C0"; VERDE = "00B050"; ROJO = "FF0000"; BLANCO = "FFFFFF"

    def cel(ws, f, c, v, bold=False, bg=None, halign="left", color="000000", merge_end=None):
        cell = ws.cell(f, c, v)
        cell.font = Font(bold=bold, name="Calibri", size=11, color=color)
        cell.alignment = Alignment(horizontal=halign, vertical="center", wrap_text=True)
        if bg: cell.fill = PatternFill("solid", fgColor=bg)
        if merge_end: ws.merge_cells(start_row=f, start_column=c, end_row=merge_end[0], end_column=merge_end[1])
        return cell

    # Titulo
    ws1.row_dimensions[2].height = 57
    cel(ws1,2,2,"REPORTE DE RECORRIDOS DE MANTENIMIENTO PREVENTIVO PARA RUTAS INTERURBANAS",bold=True,bg=AZUL,color=BLANCO,halign="center",merge_end=(2,3))
    cel(ws1,2,4,"Codigo: FOR FO 02 Version: 3 (28/05/2021)",bold=True)
    ws1.row_dimensions[4].height = 24
    cel(ws1,4,1,"REPORTE DE RECORRIDO DE RUTAS INTERURBANAS DE F. O.",bold=True,bg=AZUL,color=BLANCO,halign="center",merge_end=(4,4))

    # Fecha/hora
    ws1.row_dimensions[5].height = 38; ws1.row_dimensions[6].height = 38
    ws1.merge_cells(start_row=5,start_column=1,end_row=6,end_column=1)
    cel(ws1,5,1,"FECHA Y HORA DEL RECORRIDO",bold=True,bg=GRIS,halign="center")
    cel(ws1,5,2,"FECHA",bold=True); cel(ws1,5,3,"HORA INICIO",bold=True); cel(ws1,5,4,"HORA FIN",bold=True)
    cel(ws1,6,2,r.get("fecha","")); cel(ws1,6,3,r.get("hora_inicio","")); cel(ws1,6,4,r.get("hora_fin",""))

    # Datos generales
    for i,(label,valor) in enumerate([
        ("NOMBRE DE LA RUTA",   r.get("nombre_ruta","")),
        ("CODIGO DE CUADRILLA", r.get("codigo_cuadrilla","")),
        ("NODO INICIAL",        r.get("nodo_inicial","")),
    ]):
        f = 7+i; ws1.row_dimensions[f].height = 38
        cel(ws1,f,1,label,bold=True,bg=GRIS2)
        cel(ws1,f,2,valor,bold=True,merge_end=(f,4))

    # Novedades
    fila = 10
    for nov in r.get("novedades", []):
        num = str(nov.get("numero",""))
        ws1.row_dimensions[fila].height = 38
        ws1.merge_cells(start_row=fila,start_column=1,end_row=fila+1,end_column=1)
        cel(ws1,fila,1,"FECHA Y HORA NOVEDAD # "+num,bold=True,bg=GRIS,halign="center")
        cel(ws1,fila,2,"FECHA",bold=True); cel(ws1,fila,3,"HORA INICIO",bold=True); cel(ws1,fila,4,"HORA FIN",bold=True)
        fila += 1
        ws1.row_dimensions[fila].height = 38
        cel(ws1,fila,2,nov.get("fecha","")); cel(ws1,fila,3,nov.get("hora_inicio","")); cel(ws1,fila,4,nov.get("hora_fin",""))
        fila += 1
        for label,key in [
            ("MOTIVO APARENTE DE LA NOVEDAD","motivo"),
            ("REMEDIO DEFINITIVO A LA NOVEDAD","remedio"),
            ("TAREA PENDIENTE","tarea_pendiente"),
            ("COORDENADAS","coordenadas"),
        ]:
            ws1.row_dimensions[fila].height = 42
            cel(ws1,fila,1,label,bold=True,bg=GRIS3)
            cel(ws1,fila,2,nov.get(key,""),merge_end=(fila,4))
            fila += 1

    # Pie
    for label,valor in [
        ("NODO FINAL",r.get("nodo_final","")),
        ("LIDER DE CUADRILLA QUE ELABORA INFORME",r.get("lider","")),
        ("AYUDANTE TECNICO",r.get("ayudante","")),
        ("COORDINADOR FIBRA OPTICA",r.get("coordinador","")),
        ("FOTOS ANEXAS AL REPORTE",str(r.get("fotos_total",0))),
        ("OBSERVACIONES GENERALES",r.get("observaciones","")),
    ]:
        ws1.row_dimensions[fila].height = 38
        cel(ws1,fila,1,label,bold=True,bg=GRIS2)
        cel(ws1,fila,2,valor,merge_end=(fila,4))
        fila += 1

    # ── HOJA 2: FOTOS_ANEXAS_AL_REPORTE ─────────────────────────────
    ws2 = wb.create_sheet("FOTOS_ANEXAS_AL_REPORTE")
    ws2.column_dimensions["A"].width = 3; ws2.column_dimensions["B"].width = 20
    ws2.column_dimensions["C"].width = 59; ws2.column_dimensions["D"].width = 24
    cel(ws2,2,3,"REPORTE DE RECORRIDOS DE MANTENIMIENTO PREVENTIVO PARA RUTAS INTERURBANAS",bold=True,bg=AZUL,color=BLANCO,halign="center",merge_end=(2,4))
    cel(ws2,2,5,"Codigo: FOR FO 02",bold=True)
    cel(ws2,4,2,"FOTOS DE LAS ACCIONES CORRECTIVAS",bold=True,bg=AZUL,color=BLANCO,halign="center",merge_end=(4,5))
    cel(ws2,7,2,"NODO INICIO RECORRIDO",bold=True); cel(ws2,7,3,"FOTO",bold=True,bg=AZUL,color=BLANCO,halign="center"); cel(ws2,7,4,"NOMBRE DEL NODO",bold=True,bg=AZUL,color=BLANCO,halign="center",merge_end=(7,5))
    ws2.row_dimensions[8].height = 315
    cel(ws2,8,3,"[FOTO NODO INICIO]"); cel(ws2,8,4,r.get("nodo_inicial",""),bold=True,merge_end=(8,5))
    f2 = 10
    for nov in r.get("novedades",[]):
        cel(ws2,f2,2,"NOVEDAD # "+str(nov.get("numero","")),bold=True)
        cel(ws2,f2,3,"ANTES DEL MANTENIMIENTO",bold=True,bg=AZUL,color=BLANCO,halign="center")
        cel(ws2,f2,4,"DESPUES DEL MANTENIMIENTO",bold=True,bg=AZUL,color=BLANCO,halign="center",merge_end=(f2,5))
        f2 += 1; ws2.row_dimensions[f2].height = 315
        cel(ws2,f2,3,"[FOTO ANTES]"); cel(ws2,f2,4,"[FOTO DESPUES]",merge_end=(f2,5)); f2 += 2
    cel(ws2,f2,2,"NODO FINAL DEL RECORRIDO",bold=True)
    cel(ws2,f2,3,"FOTO",bold=True,bg=AZUL,color=BLANCO,halign="center")
    cel(ws2,f2,4,r.get("nodo_final",""),bold=True,merge_end=(f2,5))

    # ── HOJA 3: MANGAS ───────────────────────────────────────────────
    ws3 = wb.create_sheet("MANGAS")
    ws3.column_dimensions["B"].width = 25; ws3.column_dimensions["C"].width = 35
    ws3.column_dimensions["D"].width = 25; ws3.column_dimensions["E"].width = 35
    cel(ws3,2,3,"REPORTE DE RECORRIDOS DE MANTENIMIENTO PREVENTIVO PARA RUTAS INTERURBANAS",bold=True,bg=AZUL,color=BLANCO,halign="center",merge_end=(2,4))
    cel(ws3,4,2,"FOTOS DE LAS MANGAS DESDE EL NODO A AL B",bold=True,bg="00133A",color=BLANCO,halign="center",merge_end=(4,5))
    mangas = datos.get("mangas",[])
    f3 = 6
    if not mangas:
        cel(ws3,f3,2,"SIN CAMBIO DE MANGAS EN ESTE RECORRIDO")
    else:
        for i in range(0,len(mangas),2):
            m1 = mangas[i]; m2 = mangas[i+1] if i+1<len(mangas) else {}
            ws3.row_dimensions[f3].height = 315
            cel(ws3,f3,3,"[FOTO MANGA]"); cel(ws3,f3,5,"[FOTO MANGA]"); f3+=1
            for label,k in [("NOMBRE:","nombre"),("DERIVACION:","derivacion"),("COORDENADAS:","coordenadas"),("OBSERVACION:","observacion")]:
                ws3.row_dimensions[f3].height = 21
                cel(ws3,f3,2,label,bold=True,bg="1F4E79",color=BLANCO)
                cel(ws3,f3,3,m1.get(k,""))
                cel(ws3,f3,4,label,bold=True,bg="1F4E79",color=BLANCO)
                cel(ws3,f3,5,m2.get(k,"")); f3+=1

    # ── HOJA 4: INVENTARIO DE HILOS EN NODO ─────────────────────────
    ws4 = wb.create_sheet("INVENTARIO DE HILOS EN NODO")
    ws4.column_dimensions["A"].width = 10; ws4.column_dimensions["B"].width = 14
    ws4.column_dimensions["C"].width = 45; ws4.column_dimensions["D"].width = 10
    cel(ws4,2,3,"REPORTE DE RECORRIDOS DE MANTENIMIENTO PREVENTIVO PARA RUTAS INTERURBANAS",bold=True,bg=AZUL,color=BLANCO,halign="center",merge_end=(2,6))
    cel(ws4,3,1,"NODO: ",bold=True); cel(ws4,3,3,r.get("nodo_final",""))
    cel(ws4,4,1,"NOMBRE ODF DE RUTA:",bold=True); cel(ws4,4,3,datos["hilos"].get("posicion_odf",""))
    for col,txt in [(2,"PAR"),(3,"HILO"),(4,"NOMENCLATURA"),(5,"RACK #")]:
        cel(ws4,6,col,txt,bold=True,bg=AZUL2,color=BLANCO,halign="center")
    f4 = 7
    hilos = datos["hilos"].get("filas",[])
    if not hilos:
        cel(ws4,f4,2,"SIN CAMBIOS EN ODF EN ESTE RECORRIDO")
    else:
        for h in hilos:
            cel(ws4,f4,2,h.get("hilo_par","")); cel(ws4,f4,4,h.get("descripcion","")); cel(ws4,f4,5,h.get("estado","")); f4+=1

    # ── HOJA 5: Checklist CIU ────────────────────────────────────────
    ws5 = wb.create_sheet("Checklist CIU")
    for col,w in [("A",9),("B",26),("C",11),("D",21),("E",11),("F",14),("G",11),("H",14)]:
        ws5.column_dimensions[col].width = w
    cel(ws5,2,2,"CHECKLIST CUADRILLA INTERURBANA",bold=True,bg=AZUL,color=BLANCO,halign="center",merge_end=(2,7))
    cel(ws5,2,8,"Codigo: FOR FO 05 Version: 3",bold=True)
    cel(ws5,3,2,"Fecha del Recorrido",bold=True); cel(ws5,3,3,r.get("fecha",""))
    cel(ws5,3,5,"Hora Inicio",bold=True); cel(ws5,3,6,r.get("hora_inicio",""))
    cel(ws5,3,7,"Hora Fin",bold=True); cel(ws5,3,8,r.get("hora_fin",""))
    f5 = 4
    for label,valor in [
        ("Nombre de Ruta",r.get("nombre_ruta","")),("Nodo Inicio",r.get("nodo_inicial","")),
        ("Nodo Final",r.get("nodo_final","")),("Distancia",ciu.get("distancia_ruta","")),
        ("Lider de Cuadrilla",r.get("lider","")),("Vehiculo Placa",ciu.get("vehiculo_placa","")),
        ("Coordinador",r.get("coordinador","")),
    ]:
        cel(ws5,f5,2,label,bold=True); cel(ws5,f5,3,valor,merge_end=(f5,8)); f5+=1

    ciu_h = ciu.get("herramientas",{}); ciu_e = ciu.get("equipos",{}); ciu_m = ciu.get("materiales",{})
    HERR = ["Cinturon y Linea de Vida","Casco","Escalera de 24 pies","Escalera de 28 pies","Escalera de 32 pies","Conos reflectivos","Caja para herramientas","Juego de destornilladores","Martillo mediano","Estiletes","Cortafrio","Alicate","Llave francesa","Juego de rachet","Pares de guantes aislantes","Tecle","Machete","Cizalla","Pata de cabra","Flejadora (Maquina Eriband)","Extension con foco","Motosierra","Tijeras metalicas","Arco de sierra","Binoculares","Parasol","Remolque / Carrete para F.O."]
    EQUI = ["Fusionadora","Cortadora de fibra","Bobina de lanzamiento","OTDR con cargador","Llave Acsys","GPS","Inversor","Etiquetadora"]
    MATE = ["Fibra 48h (500mt)","Mangas de 48h y/o 144h (2 minimo)","Rollo de cinta Eriband 3/4","Hebillas para cinta Eriband 3/4","Hojas de sierra","Patchcord de fibra","Adaptadores (Simplex-Duplex)","Paquetes de amarras","Mesas plasticas","Sillas plasticas","Cuchillos","Poleas","Sogas de nylon medianas","Sogas de nylon gruesas","Repelente contra insectos","Repelente contra abejas y avispas"]

    for sec,items,data_s in [("HERRAMIENTAS Y EPP",HERR,ciu_h),("EQUIPOS ELECTRONICOS",EQUI,ciu_e),("MATERIALES E INSUMOS",MATE,ciu_m)]:
        cel(ws5,f5,2,sec,bold=True,bg=AZUL2,color=BLANCO,halign="center",merge_end=(f5,3))
        cel(ws5,f5,4,"CANTIDAD",bold=True,bg=AZUL2,color=BLANCO,halign="center")
        cel(ws5,f5,5,"OBSERVACIONES",bold=True,bg=AZUL2,color=BLANCO,halign="center",merge_end=(f5,8)); f5+=1
        for nombre in items:
            info = data_s.get(nombre,{})
            cant = info.get("cantidad",0) if isinstance(info,dict) else int(info or 0)
            obs  = info.get("obs","NINGUNA") if isinstance(info,dict) else ("BUEN ESTADO" if cant>0 else "NINGUNA")
            cel(ws5,f5,2,nombre)
            cel(ws5,f5,4,cant,halign="center")
            bg_o = VERDE if obs=="BUEN ESTADO" else (ROJO if obs=="MAL ESTADO" else "808080")
            cel(ws5,f5,5,obs,bg=bg_o,color=BLANCO,halign="center",merge_end=(f5,8)); f5+=1

    # ── HOJA 6: Checklists MPRIU ─────────────────────────────────────
    ws6 = wb.create_sheet("Checklists MPRIU")
    for col,w in [("A",9),("B",26),("C",11),("D",32),("E",11),("F",14),("G",11),("H",14)]:
        ws6.column_dimensions[col].width = w
    cel(ws6,2,2,"CHECKLIST DE RECORRIDO DE MANTENIMIENTO PREVENTIVO DE RUTAS INTERURBANAS",bold=True,bg=AZUL,color=BLANCO,halign="center",merge_end=(2,7))
    cel(ws6,2,8,"Codigo: FOR FO 08 Version: 02",bold=True)
    cel(ws6,3,2,"Fecha del Recorrido",bold=True); cel(ws6,3,3,r.get("fecha",""))
    cel(ws6,3,5,"Hora Inicio",bold=True); cel(ws6,3,6,r.get("hora_inicio",""))
    cel(ws6,3,7,"Hora Fin",bold=True); cel(ws6,3,8,r.get("hora_fin",""))
    f6 = 4
    for label,valor in [
        ("Nombre de Ruta",r.get("nombre_ruta","")),("Nodo Inicio",r.get("nodo_inicial","")),
        ("Nodo Final",r.get("nodo_final","")),("Distancia",ciu.get("distancia_ruta","")),
        ("Lider de Cuadrilla",r.get("lider","")),("Vehiculo Placa",ciu.get("vehiculo_placa","")),
        ("Coordinador",r.get("coordinador","")),
    ]:
        cel(ws6,f6,2,label,bold=True); cel(ws6,f6,3,valor,merge_end=(f6,8)); f6+=1

    f6 += 1
    cel(ws6,f6,2,"NOVEDAD",bold=True,bg=AZUL2,color=BLANCO,halign="center")
    cel(ws6,f6,3,"CHECK",bold=True,bg=AZUL2,color=BLANCO,halign="center")
    cel(ws6,f6,4,"SOLUCION",bold=True,bg=AZUL2,color=BLANCO,halign="center",merge_end=(f6,7))
    cel(ws6,f6,8,"CANTIDAD",bold=True,bg=AZUL2,color=BLANCO,halign="center"); f6+=1

    SOLUCIONES = {
        "HERRAJES EN MAL ESTADO.":"REALIZAR EL REEMPLAZO INMEDIATO DEL HERRAJE AFECTADO.",
        "FALTA DE HERRAJES.":"INSTALAR LOS HERRAJES CONFORME A LA NORMATIVA TECNICA.",
        "POSTES EN MAL ESTADO.":"DOCUMENTAR Y REPORTAR PARA GESTIONAR EL REEMPLAZO DEL POSTE.",
        "POSTES INCLINADOS.":"DOCUMENTAR Y REPORTAR PARA GESTIONAR EL APLOME DEL POSTE.",
        "VANOS POR RETEMPLAR.":"REALIZAR EL RETEMPLADO DEL CABLE PARA RESTABLECER LA TENSION.",
        "MANGAS SUELTAS.":"ASEGURAR LA MANGA AL POSTE EN CONFIGURACION TIPO FIGURA 8.",
        "MANGAS ABIERTAS/DANADAS.":"REEMPLAZAR TAPAS Y SELLOS GARANTIZANDO EL CIERRE HERMETICO.",
        "RESERVAS SUELTAS.":"REORGANIZAR Y ASEGURAR LA RESERVA EN FIGURA 8.",
        "CRUCES DE VIAS BAJOS.":"AJUSTAR LA ALTURA DEL CABLE A LA DISTANCIA REGLAMENTARIA.",
        "VEGETACION SOBRE FIBRA/MANGA.":"REALIZAR LA PODA O RETIRO DE VEGETACION QUE COMPROMETA LA INTEGRIDAD DEL CABLE.",
        "DOCUMENTACION UNIFILAR DE HILOS.":"DOCUMENTAR O SOLICITAR PROGRAMACION DE TRABAJO; UTILIZAR SEGUIDOR DE SENAL.",
        "LINEA ELECTRICA EN MAL ESTADO.":"DOCUMENTAR Y SOLICITAR REPORTE AL AREA DE REGULATORIO.",
        "CABLE LASTIMADO.":"DOCUMENTAR E INFORMAR PARA PROGRAMAR EL CAMBIO DEL TRAMO.",
        "POZO SIN TAPA O EN MAL ESTADO.":"SOLICITAR TRABAJOS DE OBRA CIVIL PARA SU CORRECCION.",
        "ELEMENTOS SIN ETIQUETAS ACRILICAS.":"VERIFICAR Y COLOCAR ETIQUETA ACRILICA CON EL CODIGO DE RUTA.",
        "RIESGO DE DERRUMBE O DESLAVE.":"DOCUMENTAR Y SOLICITAR REUBICACION DEL RECORRIDO DEL CABLE.",
        "RIESGO DE INUNDACIONES.":"DOCUMENTAR Y SOLICITAR REUBICACION DEL RECORRIDO DEL CABLE.",
        "RIESGO DE INCENDIO.":"DOCUMENTAR Y SOLICITAR REUBICACION DEL RECORRIDO DEL CABLE.",
        "NO SE REGISTRAN NOVEDADES DURANTE LA INSPECCION.":"NO SE ENCUENTRAN NOVEDADES QUE SIGNIFIQUEN RIESGOS EN EL CABLE DE LA RED INTERURBANO.",
    }

    for novedad in NOVEDADES_MPRIU:
        ws6.row_dimensions[f6].height = 43
        info  = nch.get(novedad,{})
        tiene = info.get("check",False)
        cant  = info.get("cantidad",0)
        chk   = "SI" if tiene else "NO"
        sol   = SOLUCIONES.get(novedad,"DOCUMENTAR Y REPORTAR AL COORDINADOR.")
        cel(ws6,f6,2,novedad)
        cel(ws6,f6,3,chk,bold=True,bg=(VERDE if tiene else ROJO),color=BLANCO,halign="center")
        cel(ws6,f6,4,sol,merge_end=(f6,7))
        cel(ws6,f6,8,cant if tiene else 0,halign="center"); f6+=1

    ws6.row_dimensions[f6].height = 60
    cel(ws6,f6,2,"Observaciones:",bold=True)
    cel(ws6,f6,3,r.get("observaciones",""),merge_end=(f6,8))

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


