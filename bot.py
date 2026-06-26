import os, io, hmac, struct, time, base64, hashlib, json, logging, threading
import httpx
from datetime import datetime
from http.server import HTTPServer, BaseHTTPRequestHandler
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
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
 LIDER, AYUDANTE, COORDINADOR, PLACA, DISTANCIA,
 CIU_HERRAMIENTAS, CIU_EQUIPOS, CIU_MATERIALES,
 NOVEDADES_AUTO, TAREA_PENDIENTE,
 FOTO_ANTES, FOTO_DESPUES, OBSERVACIONES,
 MPRIU_CHECK,
 PREGUNTA_MANGAS, PREGUNTA_HILOS,
 MANGA_NOMBRE, MANGA_COORDS, MANGA_OBS, HILO_ODF, HILO_DATOS) = range(27)

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

SOLUCIONES_MPRIU = {
    "HERRAJES EN MAL ESTADO.": "REALIZAR EL REEMPLAZO INMEDIATO DEL HERRAJE AFECTADO, GARANTIZANDO LA CORRECTA SUJECION DEL CABLE Y LA ESTABILIDAD MECANICA DEL TENDIDO.",
    "FALTA DE HERRAJES.": "INSTALAR LOS HERRAJES CONFORME A LA NORMATIVA TECNICA, ASEGURANDO LA CORRECTA FIJACION DEL CABLE AL POSTE.",
    "POSTES EN MAL ESTADO.": "DOCUMENTAR MEDIANTE REGISTRO FOTOGRAFICO Y COORDENADAS, Y REPORTAR PARA GESTIONAR EL REEMPLAZO DEL POSTE CON LA ENTIDAD RESPONSABLE.",
    "POSTE(S) CAMBIADO(S).": "INSTALAR LOS HERRAJES NECESARIOS Y ASEGURAR CORRECTAMENTE EL CABLE AL NUEVO POSTE. DOCUMENTAR EL CAMBIO PARA ACTUALIZACION DE INVENTARIO.",
    "POSTES POR INSTALAR.": "DOCUMENTAR LA UBICACION EXACTA Y REPORTAR PARA LA COORDINACION E INSTALACION DEL NUEVO POSTE REQUERIDO.",
    "POSTE NUEVO INSTALADO - TN.": "DOCUMENTAR, ETIQUETAR CON CODIGO DE IDENTIFICACION Y APLICAR PINTURA DE SENALIZACION CONFORME A ESTANDARES OPERATIVOS.",
    "POSTE NUEVO INSTALADO - EMPRESAS ELECTRICAS.": "DOCUMENTAR, COLOCAR ETIQUETA ACRILICA Y ASEGURAR EL CABLE DE FIBRA OPTICA CONFORME A LA NORMATIVA TECNICA VIGENTE.",
    "POSTES INCLINADOS.": "DOCUMENTAR MEDIANTE REGISTRO FOTOGRAFICO Y COORDENADAS, Y REPORTAR PARA GESTIONAR EL APLOME DEL POSTE CON EL CONTRATISTA.",
    "RETENIDA(S) EN MAL ESTADO.": "DOCUMENTAR MEDIANTE REGISTRO FOTOGRAFICO Y COORDENADAS, Y REPORTAR PARA GESTIONAR LA CORRECCION CON EL CONTRATISTA.",
    "RETENIDA(S) CORTADA(S).": "DOCUMENTAR MEDIANTE REGISTRO FOTOGRAFICO Y COORDENADAS, Y REPORTAR PARA GESTIONAR LA CORRECCION CON EL CONTRATISTA.",
    "VANOS POR RETEMPLAR.": "REALIZAR EL RETEMPLADO DEL CABLE PARA RESTABLECER LA TENSION ADECUADA Y EVITAR RIESGOS DE DANO O CAIDA.",
    "MANGAS SUELTAS.": "ASEGURAR LA MANGA AL POSTE EN CONFIGURACION TIPO FIGURA 8, CONFORME AL ESTANDAR.",
    "MANGAS ABIERTAS/DANADAS.": "REEMPLAZAR TAPAS Y SELLOS, GARANTIZANDO EL CIERRE HERMETICO Y LA PROTECCION DEL EMPALME CONTRA AGENTES EXTERNOS.",
    "RESERVAS SUELTAS.": "REORGANIZAR Y ASEGURAR LA RESERVA EN FIGURA 8 CONFORME A LO ESTABLECIDO.",
    "CRUCES DE VIAS BAJOS.": "AJUSTAR LA ALTURA DEL CABLE ELEVANDOLO A LA DISTANCIA REGLAMENTARIA O REPORTAR PARA LA IMPLEMENTACION DE UNA SOLUCION ESTRUCTURAL.",
    "VEGETACION SOBRE FIBRA/MANGA.": "REALIZAR LA PODA O RETIRO DE VEGETACION QUE COMPROMETA LA INTEGRIDAD O SEGURIDAD DEL CABLE. EN CASO DE REQUERIR PERMISOS, DOCUMENTAR LA NOVEDAD.",
    "LOCALIZACION DE MANGA.": "DOCUMENTAR LA UBICACION MEDIANTE COORDENADAS GPS Y REGISTRO FOTOGRAFICO PARA ACTUALIZACION DE INVENTARIO.",
    "DOCUMENTACION UNIFILAR DE HILOS.": "DOCUMENTAR O SOLICITAR LA PROGRAMACION DE TRABAJO PARA OBTENER LA INFORMACION; UTILIZAR UN SEGUIDOR DE SENAL.",
    "LINEA ELECTRICA EN MAL ESTADO.": "DOCUMENTAR EL RIESGO Y SOLICITAR AL COORDINADOR EL REPORTE AL AREA DE REGULATORIO.",
    "REGENERACION URBANA.": "ESTABLECER CONTACTO CON EL CONSORCIO, DOCUMENTAR LA AFECTACION Y COORDINAR LAS MEDIDAS DE MITIGACION.",
    "AMPLIACION DE VIA.": "DOCUMENTAR, REGISTRAR EL CONTACTO DEL RESPONSABLE DE LA OBRA Y COORDINAR MEDIDAS DE MITIGACION CON EL COORDINADOR DE FO.",
    "CABLE LASTIMADO.": "DOCUMENTAR E INFORMAR PARA PROGRAMAR EL CAMBIO DEL TRAMO DE CABLE.",
    "FIBRA INSTALADA INCORRECTAMENTE SOBRE MORDAZA.": "CORREGIR LA INSTALACION SEPARANDO ADECUADAMENTE EL CABLE DE FIBRA DEL MENSAJERO CONFORME A LA NORMATIVA TECNICA.",
    "POZO SIN TAPA O EN MAL ESTADO.": "SOLICITAR LA EJECUCION DE TRABAJOS DE OBRA CIVIL PARA SU INSTALACION O CORRECCION.",
    "REPINTADO DE POZO.": "REALIZAR EL PINTADO DEL POZO TELCONET CON EL CODIGO ASIGNADO POR GIS.",
    "REPINTADO DE POSTE.": "REALIZAR EL PINTADO DEL POSTE TELCONET CON EL CODIGO ASIGNADO POR GIS.",
    "ELEMENTOS SIN ETIQUETAS ACRILICAS.": "VERIFICAR, COLOCAR ETIQUETA ACRILICA Y ETIQUETAR CON EL CODIGO DE RUTA.",
    "RIESGO DE DERRUMBE O DESLAVE.": "DOCUMENTAR EL RIESGO Y SOLICITAR AL COORDINADOR LA REUBICACION DEL RECORRIDO DEL CABLE.",
    "RIESGO DE INUNDACIONES.": "DOCUMENTAR EL RIESGO Y SOLICITAR AL COORDINADOR LA REUBICACION DEL RECORRIDO DEL CABLE.",
    "RIESGO DE INCENDIO.": "DOCUMENTAR EL RIESGO Y SOLICITAR AL COORDINADOR LA REUBICACION DEL RECORRIDO DEL CABLE.",
    "NO SE REGISTRAN NOVEDADES DURANTE LA INSPECCION.": "NO SE ENCUENTRAN NOVEDADES QUE SIGNIFIQUEN RIESGOS EN EL CABLE DE LA RED INTERURBANO.",
}

def _estilo_header(ws, fila, col, texto, fg="1F4E79"):
    c = ws.cell(fila, col, texto)
    c.font = Font(bold=True, color="FFFFFF", name="Arial", size=9)
    c.fill = PatternFill("solid", fgColor=fg)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    return c

def _estilo_label(ws, fila, col, texto):
    c = ws.cell(fila, col, texto)
    c.font = Font(bold=True, name="Arial", size=9)
    c.alignment = Alignment(wrap_text=True, vertical="center")
    return c

def _estilo_valor(ws, fila, col, texto):
    c = ws.cell(fila, col, texto)
    c.font = Font(name="Arial", size=9)
    c.alignment = Alignment(wrap_text=True, vertical="center")
    return c

def generar_excel(datos):
    from openpyxl.styles import Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    r  = datos["recorrido"]

    borde = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )

    # ══════════════════════════════════════════════════════
    # HOJA 1: REPORTES_DE_RECORRIDOS
    # ══════════════════════════════════════════════════════
    ws1 = wb.active
    ws1.title = "REPORTES_DE_RECORRIDOS"
    ws1.column_dimensions["A"].width = 50
    ws1.column_dimensions["B"].width = 20
    ws1.column_dimensions["C"].width = 20
    ws1.column_dimensions["D"].width = 20

    # Titulo principal
    ws1.merge_cells("A1:D1")
    c = ws1["A1"]
    c.value = "REPORTE DE RECORRIDOS DE MANTENIMIENTO PREVENTIVO PARA RUTAS INTERURBANAS"
    c.font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
    c.fill = PatternFill("solid", fgColor="1F4E79")
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws1.row_dimensions[1].height = 30

    ws1.merge_cells("A2:D2")
    ws1["A2"].value = "Codigo: FOR FO 02 Version: 3 (28/05/2021)"
    ws1["A2"].font = Font(bold=True, name="Arial", size=9)
    ws1["A2"].alignment = Alignment(horizontal="right")

    # Subtitulo
    ws1.merge_cells("A3:D3")
    ws1["A3"].value = "REPORTE DE RECORRIDO DE RUTAS INTERURBANAS DE F. O."
    ws1["A3"].font = Font(bold=True, color="FFFFFF", name="Arial", size=10)
    ws1["A3"].fill = PatternFill("solid", fgColor="2E75B6")
    ws1["A3"].alignment = Alignment(horizontal="center", vertical="center")

    # Fecha y hora
    _estilo_label(ws1, 4, 1, "FECHA Y HORA DEL RECORRIDO")
    _estilo_label(ws1, 4, 2, "FECHA")
    _estilo_label(ws1, 4, 3, "HORA INICIO")
    _estilo_label(ws1, 4, 4, "HORA FIN")
    _estilo_valor(ws1, 5, 2, r["fecha"])
    _estilo_valor(ws1, 5, 3, r["hora_inicio"])
    _estilo_valor(ws1, 5, 4, r["hora_fin"])

    # Datos principales
    datos_principales = [
        ("NOMBRE DE LA RUTA", r["nombre_ruta"]),
        ("CODIGO DE CUADRILLA", r["codigo_cuadrilla"]),
        ("NODO INICIAL", r["nodo_inicial"]),
    ]
    fila = 6
    for label, valor in datos_principales:
        ws1.merge_cells(f"A{fila}:A{fila}")
        _estilo_label(ws1, fila, 1, label)
        ws1.merge_cells(f"B{fila}:D{fila}")
        _estilo_valor(ws1, fila, 2, valor)
        fila += 1

    # Novedades
    for nov in r["novedades"]:
        num = nov.get("numero", "")
        ws1.merge_cells(f"A{fila}:D{fila}")
        c = ws1.cell(fila, 1, "FECHA Y HORA NOVEDAD # " + str(num))
        c.font = Font(bold=True, color="FFFFFF", name="Arial", size=9)
        c.fill = PatternFill("solid", fgColor="2E75B6")
        c.alignment = Alignment(horizontal="center")
        fila += 1

        _estilo_label(ws1, fila, 2, "FECHA")
        _estilo_label(ws1, fila, 3, "HORA INICIO")
        _estilo_label(ws1, fila, 4, "HORA FIN")
        fila += 1
        _estilo_valor(ws1, fila, 2, nov.get("fecha",""))
        _estilo_valor(ws1, fila, 3, nov.get("hora_inicio",""))
        _estilo_valor(ws1, fila, 4, nov.get("hora_fin",""))
        fila += 1

        for label, key in [
            ("MOTIVO APARENTE DE LA NOVEDAD", "motivo"),
            ("REMEDIO DEFINITIVO A LA NOVEDAD", "remedio"),
            ("TAREA PENDIENTE (por regulatorio/obra civil, contratista)", "tarea_pendiente"),
            ("COORDENADAS SITIO DE LA NOVEDAD (Grados decimales)", "coordenadas"),
        ]:
            ws1.merge_cells(f"A{fila}:A{fila}")
            _estilo_label(ws1, fila, 1, label)
            ws1.merge_cells(f"B{fila}:D{fila}")
            _estilo_valor(ws1, fila, 2, nov.get(key,""))
            fila += 1

    # Pie del reporte
    pie = [
        ("NODO FINAL", r["nodo_final"]),
        ("LIDER DE CUADRILLA QUE ELABORA INFORME", r["lider"]),
        ("AYUDANTE TECNICO", r["ayudante"]),
        ("COORDINADOR FIBRA OPTICA", r["coordinador"]),
        ("FOTOS ANEXAS AL REPORTE (INDIQUE CUANTAS)", str(r["fotos_total"])),
        ("OBSERVACIONES GENERALES", r["observaciones"]),
    ]
    for label, valor in pie:
        _estilo_label(ws1, fila, 1, label)
        ws1.merge_cells(f"B{fila}:D{fila}")
        _estilo_valor(ws1, fila, 2, valor)
        fila += 1

    # ══════════════════════════════════════════════════════
    # HOJA 2: FOTOS_ANEXAS_AL_REPORTE
    # ══════════════════════════════════════════════════════
    ws2 = wb.create_sheet("FOTOS_ANEXAS_AL_REPORTE")
    ws2.column_dimensions["B"].width = 25
    ws2.column_dimensions["C"].width = 25
    ws2.column_dimensions["D"].width = 25

    ws2.merge_cells("B1:D1")
    ws2["B1"].value = "REPORTE DE RECORRIDOS DE MANTENIMIENTO PREVENTIVO PARA RUTAS INTERURBANAS"
    ws2["B1"].font = Font(bold=True, color="FFFFFF", name="Arial")
    ws2["B1"].fill = PatternFill("solid", fgColor="1F4E79")
    ws2["B1"].alignment = Alignment(horizontal="center")

    ws2.merge_cells("B2:D2")
    ws2["B2"].value = "FOTOS DE LAS ACCIONES CORRECTIVAS"
    ws2["B2"].font = Font(bold=True, name="Arial")
    ws2["B2"].alignment = Alignment(horizontal="center")

    # Nodo inicio
    ws2["B3"].value = "NODO INICIO RECORRIDO"
    ws2["C3"].value = "FOTO"
    ws2["D3"].value = r["nodo_inicial"]
    ws2["B3"].font = Font(bold=True, name="Arial", size=9)

    fila2 = 4
    for nov in r["novedades"]:
        num = nov.get("numero","")
        ws2.merge_cells(f"B{fila2}:D{fila2}")
        c = ws2.cell(fila2, 2, "NOVEDAD # " + str(num))
        c.font = Font(bold=True, color="FFFFFF", name="Arial", size=9)
        c.fill = PatternFill("solid", fgColor="2E75B6")
        c.alignment = Alignment(horizontal="center")
        fila2 += 1

        ws2.cell(fila2, 2, "ANTES DEL MANTENIMIENTO").font = Font(bold=True, name="Arial", size=9)
        ws2.cell(fila2, 4, "DESPUES DEL MANTENIMIENTO").font = Font(bold=True, name="Arial", size=9)
        fila2 += 1

        ws2.row_dimensions[fila2].height = 90
        ws2.cell(fila2, 2, "[FOTO ANTES]").font = Font(name="Arial", size=9, color="808080")
        ws2.cell(fila2, 4, "[FOTO DESPUES]").font = Font(name="Arial", size=9, color="808080")
        fila2 += 2

    ws2.cell(fila2, 2, "NODO FINAL DEL RECORRIDO").font = Font(bold=True, name="Arial", size=9)
    ws2.cell(fila2, 3, "FOTO").font = Font(bold=True, name="Arial", size=9)
    ws2.cell(fila2, 4, r["nodo_final"]).font = Font(name="Arial", size=9)

    # ══════════════════════════════════════════════════════
    # HOJA 3: Checklist CIU
    # ══════════════════════════════════════════════════════
    ws_ciu = wb.create_sheet("Checklist CIU")
    ws_ciu.column_dimensions["B"].width = 40
    ws_ciu.column_dimensions["C"].width = 12
    ws_ciu.column_dimensions["D"].width = 20

    ws_ciu.merge_cells("B1:D1")
    ws_ciu["B1"].value = "CHECKLIST CUADRILLA INTERURBANA - Codigo: FOR FO 05"
    ws_ciu["B1"].font = Font(bold=True, color="FFFFFF", name="Arial")
    ws_ciu["B1"].fill = PatternFill("solid", fgColor="1F4E79")
    ws_ciu["B1"].alignment = Alignment(horizontal="center")

    datos_ciu = [
        ("Fecha del Recorrido", r["fecha"], "Hora Inicio", r["hora_inicio"], "Hora Fin", r["hora_fin"]),
    ]
    ws_ciu["B2"].value = "Fecha del Recorrido"
    ws_ciu["C2"].value = r["fecha"]
    ws_ciu["D2"].value = "Hora Inicio: " + r["hora_inicio"] + "  Hora Fin: " + r["hora_fin"]

    info_ciu = [
        ("Nombre de Ruta", r["nombre_ruta"]),
        ("Nodo Inicio", r["nodo_inicial"]),
        ("Nodo Final", r["nodo_final"]),
        ("Distancia de la Ruta", datos["ciu"].get("distancia_ruta","")),
        ("Lider de Cuadrilla", r["lider"]),
        ("Vehiculo Placa", datos["ciu"].get("vehiculo_placa","")),
        ("Coordinador Fibra Optica", r["coordinador"]),
    ]
    fila_c = 3
    for label, valor in info_ciu:
        ws_ciu.cell(fila_c, 2, label).font = Font(bold=True, name="Arial", size=9)
        ws_ciu.merge_cells(f"C{fila_c}:D{fila_c}")
        ws_ciu.cell(fila_c, 3, valor).font = Font(name="Arial", size=9)
        fila_c += 1

    ws_ciu.cell(fila_c, 2, "HERRAMIENTAS Y EPP").font = Font(bold=True, color="FFFFFF", name="Arial", size=9)
    ws_ciu.cell(fila_c, 2).fill = PatternFill("solid", fgColor="2E75B6")
    ws_ciu.cell(fila_c, 3, "CANTIDAD").font = Font(bold=True, name="Arial", size=9)
    ws_ciu.cell(fila_c, 4, "OBSERVACIONES").font = Font(bold=True, name="Arial", size=9)
    fila_c += 1

    herramientas = [
        "Cinturon y Linea de Vida", "Casco", "Escalera de 24 pies", "Escalera de 28 pies",
        "Escalera de 32 pies", "Conos reflectivos", "Juego de destornilladores",
        "Martillo mediano", "Estiletes", "Cortafrio", "Alicate", "Llave francesa",
        "Juego de rachet", "Pares de guantes aislantes", "Tecle", "Machete", "Cizalla",
        "Fusionadora", "Cortadora de fibra", "OTDR con cargador", "Llave Acsys",
        "Inversor", "Etiquetadora",
    ]
    for h in herramientas:
        ws_ciu.cell(fila_c, 2, h).font = Font(name="Arial", size=9)
        ws_ciu.cell(fila_c, 3, 0).font = Font(name="Arial", size=9)
        ws_ciu.cell(fila_c, 4, "NINGUNA").font = Font(name="Arial", size=9)
        fila_c += 1

    # ══════════════════════════════════════════════════════
    # HOJA 4: Checklists MPRIU
    # ══════════════════════════════════════════════════════
    ws_mp = wb.create_sheet("Checklists MPRIU")
    ws_mp.column_dimensions["B"].width = 45
    ws_mp.column_dimensions["C"].width = 8
    ws_mp.column_dimensions["D"].width = 60
    ws_mp.column_dimensions["E"].width = 12

    ws_mp.merge_cells("B1:E1")
    ws_mp["B1"].value = "CHECKLIST DE RECORRIDO DE MANTENIMIENTO PREVENTIVO DE RUTAS INTERURBANAS - Codigo: FOR FO 08"
    ws_mp["B1"].font = Font(bold=True, color="FFFFFF", name="Arial")
    ws_mp["B1"].fill = PatternFill("solid", fgColor="1F4E79")
    ws_mp["B1"].alignment = Alignment(horizontal="center", wrap_text=True)

    info_mp = [
        ("Fecha del Recorrido", r["fecha"] + "  Hora Inicio: " + r["hora_inicio"] + "  Hora Fin: " + r["hora_fin"]),
        ("Nombre de Ruta", r["nombre_ruta"]),
        ("Nodo Inicio", r["nodo_inicial"]),
        ("Nodo Final", r["nodo_final"]),
        ("Distancia de la Ruta", datos["ciu"].get("distancia_ruta","")),
        ("Lider de Cuadrilla", r["lider"]),
        ("Vehiculo Placa", datos["ciu"].get("vehiculo_placa","")),
        ("Coordinador Fibra Optica", r["coordinador"]),
    ]
    fila_m = 2
    for label, valor in info_mp:
        ws_mp.cell(fila_m, 2, label).font = Font(bold=True, name="Arial", size=9)
        ws_mp.merge_cells(f"C{fila_m}:E{fila_m}")
        ws_mp.cell(fila_m, 3, valor).font = Font(name="Arial", size=9)
        fila_m += 1

    # Encabezados tabla novedades
    for col, texto in [(2,"NOVEDAD"),(3,"CHECK"),(4,"SOLUCION"),(5,"CANTIDAD")]:
        c = ws_mp.cell(fila_m, col, texto)
        c.font = Font(bold=True, color="FFFFFF", name="Arial", size=9)
        c.fill = PatternFill("solid", fgColor="2E75B6")
        c.alignment = Alignment(horizontal="center")
    fila_m += 1

    novedades_check = datos["mpriu"].get("novedades_check", {})
    for novedad in NOVEDADES_MPRIU:
        check_info = novedades_check.get(novedad, {})
        tiene = check_info.get("check", False)
        cantidad = check_info.get("cantidad", 0)
        check_str = "SI" if tiene else "NO"
        solucion = SOLUCIONES_MPRIU.get(novedad, "")

        c_nov = ws_mp.cell(fila_m, 2, novedad)
        c_nov.font = Font(name="Arial", size=9)
        c_nov.alignment = Alignment(wrap_text=True)

        c_chk = ws_mp.cell(fila_m, 3, check_str)
        c_chk.font = Font(bold=True, name="Arial", size=9, color="FF0000" if tiene else "000000")
        c_chk.alignment = Alignment(horizontal="center")

        c_sol = ws_mp.cell(fila_m, 4, solucion)
        c_sol.font = Font(name="Arial", size=9)
        c_sol.alignment = Alignment(wrap_text=True)

        c_cant = ws_mp.cell(fila_m, 5, cantidad)
        c_cant.font = Font(name="Arial", size=9)
        c_cant.alignment = Alignment(horizontal="center")

        ws_mp.row_dimensions[fila_m].height = 30
        fila_m += 1

    ws_mp.cell(fila_m, 2, "Observaciones:").font = Font(bold=True, name="Arial", size=9)
    ws_mp.merge_cells(f"C{fila_m}:E{fila_m}")
    ws_mp.cell(fila_m, 3, r["observaciones"]).font = Font(name="Arial", size=9)

    # ══════════════════════════════════════════════════════
    # HOJA 5: MANGAS (solo si hay)
    # ══════════════════════════════════════════════════════
    if datos.get("mangas"):
        ws_mg = wb.create_sheet("MANGAS")
        ws_mg.column_dimensions["B"].width = 30
        ws_mg.column_dimensions["C"].width = 20
        ws_mg.column_dimensions["D"].width = 30
        ws_mg.column_dimensions["E"].width = 20

        ws_mg.merge_cells("B1:E1")
        ws_mg["B1"].value = "FOTOS DE LAS MANGAS DESDE EL NODO A AL B"
        ws_mg["B1"].font = Font(bold=True, color="FFFFFF", name="Arial")
        ws_mg["B1"].fill = PatternFill("solid", fgColor="1F4E79")
        ws_mg["B1"].alignment = Alignment(horizontal="center")

        fila_mg = 2
        mangas = datos["mangas"]
        for i in range(0, len(mangas), 2):
            m1 = mangas[i]
            m2 = mangas[i+1] if i+1 < len(mangas) else {}
            for label, k1, k2 in [
                ("NOMBRE:", "nombre", "nombre"),
                ("DERIVACION:", "derivacion", "derivacion"),
                ("COORDENADAS:", "coordenadas", "coordenadas"),
                ("OBSERVACION:", "observacion", "observacion"),
            ]:
                ws_mg.cell(fila_mg, 2, label).font = Font(bold=True, name="Arial", size=9)
                ws_mg.cell(fila_mg, 3, m1.get(k1,"")).font = Font(name="Arial", size=9)
                ws_mg.cell(fila_mg, 4, label).font = Font(bold=True, name="Arial", size=9)
                ws_mg.cell(fila_mg, 5, m2.get(k2,"")).font = Font(name="Arial", size=9)
                fila_mg += 1
            fila_mg += 1

    # ══════════════════════════════════════════════════════
    # HOJA 6: INVENTARIO DE HILOS (solo si hay)
    # ══════════════════════════════════════════════════════
    if datos["hilos"].get("filas"):
        ws_h = wb.create_sheet("INVENTARIO DE HILOS EN NODO")
        ws_h.column_dimensions["B"].width = 8
        ws_h.column_dimensions["C"].width = 8
        ws_h.column_dimensions["D"].width = 30

        ws_h.merge_cells("B1:D1")
        ws_h["B1"].value = "INVENTARIO DE HILOS EN NODO - FOR FO 02"
        ws_h["B1"].font = Font(bold=True, color="FFFFFF", name="Arial")
        ws_h["B1"].fill = PatternFill("solid", fgColor="1F4E79")
        ws_h["B1"].alignment = Alignment(horizontal="center")

        ws_h.cell(2, 2, "POSICION ODF:").font = Font(bold=True, name="Arial", size=9)
        ws_h.cell(2, 3, datos["hilos"].get("posicion_odf","")).font = Font(name="Arial", size=9)

        for col, txt in [(2,"PAR"),(3,"HILO"),(4,"NOMENCLATURA"),(5,"ESTADO")]:
            c = ws_h.cell(3, col, txt)
            c.font = Font(bold=True, color="FFFFFF", name="Arial", size=9)
            c.fill = PatternFill("solid", fgColor="2E75B6")
            c.alignment = Alignment(horizontal="center")

        fila_h = 4
        for h in datos["hilos"]["filas"]:
            ws_h.cell(fila_h, 2, h.get("hilo_par","")).font = Font(name="Arial", size=9)
            ws_h.cell(fila_h, 3, "").font = Font(name="Arial", size=9)
            ws_h.cell(fila_h, 4, h.get("descripcion","")).font = Font(name="Arial", size=9)
            ws_h.cell(fila_h, 5, h.get("estado","")).font = Font(name="Arial", size=9)
            fila_h += 1

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
