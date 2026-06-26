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

HERRAMIENTAS_EPP = [
    "Cinturon y Linea de Vida", "Casco", "Escalera de 24 pies", "Escalera de 28 pies",
    "Escalera de 32 pies", "Conos reflectivos", "Caja para herramientas",
    "Juego de destornilladores", "Martillo mediano", "Estiletes", "Cortafrio",
    "Alicate", "Llave francesa", "Juego de rachet", "Pares de guantes aislantes",
    "Tecle", "Machete", "Cizalla", "Pata de cabra", "Flejadora (Maquina Eriband)",
    "Extension con foco", "Motosierra", "Tijeras metalicas", "Arco de sierra",
    "Binoculares", "Parasol", "Remolque / Carrete para F.O.",
]
EQUIPOS_ELECTRONICOS = [
    "Fusionadora", "Cortadora de fibra", "Bobina de lanzamiento",
    "OTDR con cargador", "Llave Acsys", "GPS", "Inversor", "Etiquetadora",
]
MATERIALES_INSUMOS = [
    "Fibra 48h (500mt)", "Mangas de 48h y/o 144h (2 minimo)",
    "Rollo de cinta Eriband 3/4", "Hebillas para cinta Eriband 3/4",
    "Hojas de sierra", "Patchcord de fibra", "Adaptadores (Simplex-Duplex)",
    "Paquetes de amarras", "Mesas plasticas", "Sillas plasticas",
    "Cuchillos", "Poleas", "Sogas de nylon medianas", "Sogas de nylon gruesas",
    "Repelente contra insectos", "Repelente contra abejas y avispas",
]

# Colores exactos del original Telconet
AZUL      = "0000FF"   # titulos principales
GRIS_OSC  = "969696"   # FECHA Y HORA header
GRIS_MED  = "D9D9D9"   # labels datos generales
GRIS_CLAR = "C0C0C0"   # labels de novedad
AZUL_MANG = "1F4E79"   # mangas header
AZUL_CIU  = "0070C0"   # CIU/MPRIU headers
VERDE     = "00B050"
ROJO      = "FF0000"
GRIS_LEY  = "808080"
TEXTO     = "333333"
BLANCO    = "FFFFFF"

def cel(ws, fila, col, valor, bold=False, size=11, color=TEXTO, bg=None, halign="left", wrap=True, merge_to=None):
    from openpyxl.styles import Alignment as Al
    c = ws.cell(fila, col, valor)
    c.font = Font(bold=bold, size=size, color=color, name="Calibri")
    if bg:
        c.fill = PatternFill("solid", fgColor=bg)
    c.alignment = Al(horizontal=halign, vertical="center", wrap_text=wrap)
    if merge_to:
        ws.merge_cells(start_row=fila, start_column=col, end_row=merge_to[0], end_column=merge_to[1])
    return c

def generar_excel(datos):
    from openpyxl.styles import Alignment as Al
    wb = Workbook()
    r   = datos["recorrido"]
    ciu = datos["ciu"]
    nch = datos["mpriu"].get("novedades_check", {})

    # ══ HOJA 1: REPORTES_DE_RECORRIDOS ══════════════════════════════
    ws1 = wb.active
    ws1.title = "REPORTES_DE_RECORRIDOS"
    ws1.column_dimensions["A"].width = 41.11
    ws1.column_dimensions["B"].width = 35.56
    ws1.column_dimensions["C"].width = 32.34
    ws1.column_dimensions["D"].width = 30.56

    # Fila 2 — titulo
    ws1.row_dimensions[2].height = 57
    cel(ws1,2,2,"REPORTE DE RECORRIDOS DE MANTENIMIENTO PREVENTIVO PARA RUTAS INTERURBANAS",
        bold=True,color=BLANCO,bg=AZUL,halign="center",merge_to=(2,3))
    cel(ws1,2,4,"Codigo: FOR FO 02\nVersion: 3 (28/05/2021)",bold=True,size=11)

    # Fila 4 — subtitulo
    ws1.row_dimensions[4].height = 24
    cel(ws1,4,1,"REPORTE DE RECORRIDO DE RUTAS INTERURBANAS DE F. O.",
        bold=True,color=BLANCO,bg=AZUL,halign="center",merge_to=(4,4))

    # Fila 5-6 — fecha/hora
    ws1.row_dimensions[5].height = 38
    ws1.row_dimensions[6].height = 38
    ws1.merge_cells(start_row=5,start_column=1,end_row=6,end_column=1)
    cel(ws1,5,1,"FECHA Y HORA DEL RECORRIDO",bold=True,bg=GRIS_OSC,halign="center",color=TEXTO)
    cel(ws1,5,2,"FECHA",bold=True,color=TEXTO)
    cel(ws1,5,3,"HORA INICIO",bold=True,color=TEXTO)
    cel(ws1,5,4,"HORA FIN",bold=True,color=TEXTO)
    cel(ws1,6,2,r["fecha"],color=TEXTO)
    cel(ws1,6,3,r["hora_inicio"],color=TEXTO)
    cel(ws1,6,4,r["hora_fin"],color=TEXTO)

    # Filas 7-9 — datos generales
    for i,(label,valor) in enumerate([
        ("NOMBRE DE LA RUTA",    r["nombre_ruta"]),
        ("CODIGO DE CUADRILLA",  r["codigo_cuadrilla"]),
        ("NODO INICIAL",         r["nodo_inicial"]),
    ]):
        f = 7+i
        ws1.row_dimensions[f].height = 38
        cel(ws1,f,1,label,bold=True,bg=GRIS_MED,color=TEXTO)
        cel(ws1,f,2,valor,bold=True,color=TEXTO,merge_to=(f,4))

    # Novedades
    fila = 10
    for nov in r["novedades"]:
        num = str(nov.get("numero",""))
        ws1.row_dimensions[fila].height = 38
        ws1.merge_cells(start_row=fila,start_column=1,end_row=fila+1,end_column=1)
        cel(ws1,fila,1,"FECHA Y HORA NOVEDAD # "+num,bold=True,bg=GRIS_OSC,color=TEXTO,halign="center")
        cel(ws1,fila,2,"FECHA",bold=True,color=TEXTO)
        cel(ws1,fila,3,"HORA INICIO",bold=True,color=TEXTO)
        cel(ws1,fila,4,"HORA FIN",bold=True,color=TEXTO)
        fila += 1
        ws1.row_dimensions[fila].height = 38
        cel(ws1,fila,2,nov.get("fecha",""),color=TEXTO)
        cel(ws1,fila,3,nov.get("hora_inicio",""),color=TEXTO)
        cel(ws1,fila,4,nov.get("hora_fin",""),color=TEXTO)
        fila += 1
        for label,key in [
            ("MOTIVO APARENTE DE LA NOVEDAD","motivo"),
            ("REMEDIO DEFINITIVO A LA NOVEDAD","remedio"),
            ("TAREA PENDIENTE (por regulatorio/obra civil, contratista)","tarea_pendiente"),
            ("COORDENADAS SITIO DE LA NOVEDAD (Grados decimales)","coordenadas"),
        ]:
            ws1.row_dimensions[fila].height = 42
            cel(ws1,fila,1,label,bold=True,bg=GRIS_CLAR,color=TEXTO)
            cel(ws1,fila,2,nov.get(key,""),color=TEXTO,merge_to=(fila,4))
            fila += 1

    # Pie
    for label,valor in [
        ("NODO FINAL",r["nodo_final"]),
        ("LIDER DE CUADRILLA QUE ELABORA INFORME",r["lider"]),
        ("AYUDANTE TECNICO",r["ayudante"]),
        ("COORDINADOR FIBRA OPTICA",r["coordinador"]),
        ("FOTOS ANEXAS AL REPORTE (INDIQUE CUANTAS)",str(r["fotos_total"])),
        ("OBSERVACIONES GENERALES",r["observaciones"]),
    ]:
        ws1.row_dimensions[fila].height = 38
        cel(ws1,fila,1,label,bold=True,bg=GRIS_MED,color=TEXTO)
        cel(ws1,fila,2,valor,color=TEXTO,merge_to=(fila,4))
        fila += 1

    # ══ HOJA 2: FOTOS_ANEXAS_AL_REPORTE ═════════════════════════════
    ws2 = wb.create_sheet("FOTOS_ANEXAS_AL_REPORTE")
    ws2.column_dimensions["A"].width = 2.67
    ws2.column_dimensions["B"].width = 19.67
    ws2.column_dimensions["C"].width = 58.78
    ws2.column_dimensions["D"].width = 23.78
    ws2.column_dimensions["E"].width = 35.77

    ws2.row_dimensions[2].height = 57
    cel(ws2,2,3,"REPORTE DE RECORRIDOS DE MANTENIMIENTO PREVENTIVO PARA RUTAS INTERURBANAS",
        bold=True,color=BLANCO,bg=AZUL,halign="center",merge_to=(2,4))
    cel(ws2,2,5,"Codigo: FOR FO 02\nVersion: 3 (28/05/2021)",bold=True)

    cel(ws2,4,2,"FOTOS DE LAS ACCIONES CORRECTIVAS",bold=True,color=BLANCO,bg=AZUL,halign="center",merge_to=(4,5))
    ws2.row_dimensions[4].height = 30

    cel(ws2,7,2,"NODO INICIO RECORRIDO",bold=True,color=TEXTO)
    cel(ws2,7,3,"FOTO",bold=True,color=BLANCO,bg=AZUL,halign="center")
    cel(ws2,7,4,"NOMBRE DEL NODO",bold=True,color=BLANCO,bg=AZUL,halign="center")
    cel(ws2,8,4,r["nodo_inicial"],bold=True,size=24,color=TEXTO)

    f2 = 10
    for nov in r["novedades"]:
        num = str(nov.get("numero",""))
        cel(ws2,f2,2,"NOVEDAD # "+num,bold=True,color=TEXTO)
        cel(ws2,f2,3,"ANTES DEL MANTENIMIENTO",bold=True,color=BLANCO,bg=AZUL,halign="center")
        cel(ws2,f2,4,"DESPUES DEL MANTENIMIENTO",bold=True,color=BLANCO,bg=AZUL,halign="center")
        f2 += 1
        ws2.row_dimensions[f2].height = 315
        ws2.cell(f2,3).value = "[FOTO ANTES]"
        ws2.cell(f2,3).font = Font(name="Calibri",size=9,color=GRIS_LEY,italic=True)
        ws2.cell(f2,4).value = "[FOTO DESPUES]"
        ws2.cell(f2,4).font = Font(name="Calibri",size=9,color=GRIS_LEY,italic=True)
        f2 += 2

    cel(ws2,f2,2,"NODO FINAL DEL RECORRIDO",bold=True,color=TEXTO)
    cel(ws2,f2,3,"FOTO",bold=True,color=BLANCO,bg=AZUL,halign="center")
    cel(ws2,f2,4,r["nodo_final"],bold=True,size=24,color=TEXTO)

    # ══ HOJA 3: MANGAS ═══════════════════════════════════════════════
    ws3 = wb.create_sheet("MANGAS")
    ws3.column_dimensions["A"].width = 2.67
    ws3.column_dimensions["B"].width = 24.78
    ws3.column_dimensions["C"].width = 34.78
    ws3.column_dimensions["D"].width = 24.78
    ws3.column_dimensions["E"].width = 34.78

    ws3.row_dimensions[2].height = 57
    cel(ws3,2,3,"REPORTE DE RECORRIDOS DE MANTENIMIENTO PREVENTIVO PARA RUTAS INTERURBANAS",
        bold=True,color=BLANCO,bg=AZUL,halign="center",merge_to=(2,4))
    cel(ws3,2,5,"Codigo: FOR FO 02\nVersion: 3 (28/05/2021)",bold=True)
    cel(ws3,4,2,"FOTOS DE LAS MANGAS DESDE EL NODO A AL B",
        bold=True,color=BLANCO,bg="00133A",halign="center",merge_to=(4,5))
    ws3.row_dimensions[4].height = 39

    f3 = 6
    mangas = datos.get("mangas",[])
    if not mangas:
        cel(ws3,f3,2,"SIN CAMBIO DE MANGAS EN ESTE RECORRIDO",color=GRIS_LEY)
    else:
        for i in range(0,len(mangas),2):
            m1 = mangas[i]
            m2 = mangas[i+1] if i+1<len(mangas) else {}
            for label,k in [("NOMBRE:","nombre"),("DERIVACION:","derivacion"),
                             ("COORDENADAS:","coordenadas"),("OBSERVACION:","observacion")]:
                ws3.row_dimensions[f3].height = 21
                cel(ws3,f3,2,label,bold=True,color=BLANCO,bg=AZUL_MANG)
                cel(ws3,f3,3,m1.get(k,""),color=TEXTO)
                cel(ws3,f3,4,label,bold=True,color=BLANCO,bg=AZUL_MANG)
                cel(ws3,f3,5,m2.get(k,""),color=TEXTO)
                f3 += 1
            ws3.row_dimensions[f3].height = 315
            ws3.cell(f3,3).value = "[FOTO MANGA]"
            ws3.cell(f3,3).font = Font(name="Calibri",size=9,color=GRIS_LEY,italic=True)
            ws3.cell(f3,5).value = "[FOTO MANGA]"
            ws3.cell(f3,5).font = Font(name="Calibri",size=9,color=GRIS_LEY,italic=True)
            f3 += 2

    # ══ HOJA 4: INVENTARIO DE HILOS EN NODO ═════════════════════════
    ws4 = wb.create_sheet("INVENTARIO DE HILOS EN NODO")
    ws4.column_dimensions["A"].width = 9.56
    ws4.column_dimensions["B"].width = 14.11
    ws4.column_dimensions["C"].width = 45.44
    ws4.column_dimensions["D"].width = 9.11
    ws4.column_dimensions["E"].width = 24.67
    ws4.column_dimensions["F"].width = 23.67

    ws4.row_dimensions[2].height = 57
    cel(ws4,2,3,"REPORTE DE RECORRIDOS DE MANTENIMIENTO PREVENTIVO PARA RUTAS INTERURBANAS",
        bold=True,color=BLANCO,bg=AZUL,halign="center",merge_to=(2,6))
    cel(ws4,2,7,"Codigo: FOR FO 02\nVersion: 3 (28/05/2021)",bold=True)

    cel(ws4,3,1,"NODO: ",bold=True,color=TEXTO)
    cel(ws4,3,3,r["nodo_final"],color=TEXTO)
    cel(ws4,4,1,"NOMBRE ODF DE RUTA:",bold=True,color=TEXTO)
    cel(ws4,4,3,datos["hilos"].get("posicion_odf",""),color=TEXTO)

    ws4.row_dimensions[6].height = 27
    for col,txt in [(2,"PAR"),(3,"HILO"),(4,"NOMENCLATURA"),(5,"RACK #")]:
        cel(ws4,6,col,txt,bold=True,color=BLANCO,bg=AZUL_CIU,halign="center")

    f4 = 7
    hilos = datos["hilos"].get("filas",[])
    if not hilos:
        cel(ws4,f4,2,"SIN CAMBIOS EN ODF EN ESTE RECORRIDO",color=GRIS_LEY)
    else:
        for h in hilos:
            ws4.row_dimensions[f4].height = 21
            cel(ws4,f4,2,h.get("hilo_par",""),color=TEXTO)
            cel(ws4,f4,4,h.get("descripcion",""),color=TEXTO)
            cel(ws4,f4,5,h.get("estado",""),color=TEXTO)
            f4 += 1

    # ══ HOJA 5: Checklist CIU ════════════════════════════════════════
    ws5 = wb.create_sheet("Checklist CIU")
    ws5.column_dimensions["A"].width = 8.67
    ws5.column_dimensions["B"].width = 25.67
    ws5.column_dimensions["C"].width = 10.66
    ws5.column_dimensions["D"].width = 20.66
    ws5.column_dimensions["E"].width = 10.66
    ws5.column_dimensions["F"].width = 13.67
    ws5.column_dimensions["G"].width = 10.66
    ws5.column_dimensions["H"].width = 13.67

    ws5.row_dimensions[2].height = 45.75
    cel(ws5,2,2,"CHECKLIST CUADRILLA INTERURBANA",
        bold=True,color=BLANCO,bg=AZUL,halign="center",merge_to=(2,7))
    cel(ws5,2,8,"Codigo: FOR FO 05\nVersion: 3 (26/06/2025)",bold=True)

    cel(ws5,3,2,"Fecha del Recorrido",bold=True,color=TEXTO)
    cel(ws5,3,3,r["fecha"],color=TEXTO)
    cel(ws5,3,5,"Hora Inicio",bold=True,color=TEXTO)
    cel(ws5,3,6,r["hora_inicio"],color=TEXTO)
    cel(ws5,3,7,"Hora Fin",bold=True,color=TEXTO)
    cel(ws5,3,8,r["hora_fin"],color=TEXTO)

    f5 = 4
    for label,valor in [
        ("Nombre de Ruta",r["nombre_ruta"]),
        ("Nodo Inicio",r["nodo_inicial"]),
        ("Nodo Final",r["nodo_final"]),
        ("Distancia de la Ruta",ciu.get("distancia_ruta","")),
        ("Lider de Cuadrilla",r["lider"]),
        ("Vehiculo Placa",ciu.get("vehiculo_placa","")),
        ("Coordinador Fibra Optica",r["coordinador"]),
    ]:
        ws5.row_dimensions[f5].height = 21
        cel(ws5,f5,2,label,bold=True,color=TEXTO)
        cel(ws5,f5,3,valor,color=TEXTO,merge_to=(f5,8))
        f5 += 1

    ciu_h = ciu.get("herramientas",{})
    ciu_e = ciu.get("equipos",{})
    ciu_m = ciu.get("materiales",{})

    for sec_nom,items,data_s in [
        ("HERRAMIENTAS Y EPP",HERRAMIENTAS_EPP,ciu_h),
        ("EQUIPOS ELECTRONICOS",EQUIPOS_ELECTRONICOS,ciu_e),
        ("MATERIALES E INSUMOS",MATERIALES_INSUMOS,ciu_m),
    ]:
        ws5.row_dimensions[f5].height = 21
        cel(ws5,f5,2,sec_nom,bold=True,color=BLANCO,bg=AZUL_CIU,halign="center",merge_to=(f5,3))
        cel(ws5,f5,4,"CANTIDAD",bold=True,color=BLANCO,bg=AZUL_CIU,halign="center")
        cel(ws5,f5,5,"OBSERVACIONES",bold=True,color=BLANCO,bg=AZUL_CIU,halign="center",merge_to=(f5,8))
        f5 += 1
        for nombre in items:
            ws5.row_dimensions[f5].height = 21
            cel(ws5,f5,2,nombre,color=TEXTO)
            info = data_s.get(nombre,{})
            cant = info.get("cantidad",0)
            obs  = info.get("obs","NINGUNA")
            cel(ws5,f5,4,cant,color=TEXTO,halign="center")
            bg_obs = VERDE if obs=="BUEN ESTADO" else (ROJO if obs=="MAL ESTADO" else GRIS_LEY)
            cel(ws5,f5,5,obs,color=BLANCO,bg=bg_obs,halign="center",merge_to=(f5,8))
            f5 += 1

    f5 += 1
    for txt,bg in [("BUEN ESTADO",VERDE),("MAL ESTADO",ROJO),("NINGUNA",GRIS_LEY)]:
        cel(ws5,f5,2,txt,bold=True,color=BLANCO,bg=bg,halign="center")
        f5 += 1

    # ══ HOJA 6: Checklists MPRIU ════════════════════════════════════
    ws6 = wb.create_sheet("Checklists MPRIU")
    ws6.column_dimensions["A"].width = 8.67
    ws6.column_dimensions["B"].width = 25.67
    ws6.column_dimensions["C"].width = 10.66
    ws6.column_dimensions["D"].width = 31.67
    ws6.column_dimensions["E"].width = 10.66
    ws6.column_dimensions["F"].width = 13.67
    ws6.column_dimensions["G"].width = 10.66
    ws6.column_dimensions["H"].width = 13.67

    ws6.row_dimensions[2].height = 45.75
    cel(ws6,2,2,"CHECKLIST DE RECORRIDO DE MANTENIMIENTO PREVENTIVO DE RUTAS INTERURBANAS",
        bold=True,color=BLANCO,bg=AZUL,halign="center",merge_to=(2,7))
    cel(ws6,2,8,"Codigo: FOR FO 08\nVersion: 02 (28/05/2021)",bold=True)

    cel(ws6,3,2,"Fecha del Recorrido",bold=True,color=TEXTO)
    cel(ws6,3,3,r["fecha"],color=TEXTO)
    cel(ws6,3,5,"Hora Inicio",bold=True,color=TEXTO)
    cel(ws6,3,6,r["hora_inicio"],color=TEXTO)
    cel(ws6,3,7,"Hora Fin",bold=True,color=TEXTO)
    cel(ws6,3,8,r["hora_fin"],color=TEXTO)

    f6 = 4
    for label,valor in [
        ("Nombre de Ruta",r["nombre_ruta"]),
        ("Nodo Inicio",r["nodo_inicial"]),
        ("Nodo Final",r["nodo_final"]),
        ("Distancia de la Ruta",ciu.get("distancia_ruta","")),
        ("Lider de Cuadrilla",r["lider"]),
        ("Vehiculo Placa",ciu.get("vehiculo_placa","")),
        ("Coordinador Fibra Optica",r["coordinador"]),
    ]:
        ws6.row_dimensions[f6].height = 30
        cel(ws6,f6,2,label,bold=True,color=TEXTO)
        cel(ws6,f6,3,valor,color=TEXTO,merge_to=(f6,8))
        f6 += 1

    ws6.row_dimensions[f6].height = 30
    cel(ws6,f6,2,"NOVEDAD",bold=True,color=BLANCO,bg=AZUL_CIU,halign="center")
    cel(ws6,f6,3,"CHECK",bold=True,color=BLANCO,bg=AZUL_CIU,halign="center")
    cel(ws6,f6,4,"SOLUCION",bold=True,color=BLANCO,bg=AZUL_CIU,halign="center",merge_to=(f6,7))
    cel(ws6,f6,8,"CANTIDAD",bold=True,color=BLANCO,bg=AZUL_CIU,halign="center")
    f6 += 1

    for novedad in NOVEDADES_MPRIU:
        ws6.row_dimensions[f6].height = 43.5
        info = nch.get(novedad,{})
        tiene = info.get("check",False)
        cant  = info.get("cantidad",0)
        chk   = "SI" if tiene else "NO"
        sol   = SOLUCIONES_MPRIU.get(novedad,"")
        cel(ws6,f6,2,novedad,color=TEXTO)
        bg_chk = VERDE if tiene else ROJO
        cel(ws6,f6,3,chk,bold=True,color=BLANCO,bg=bg_chk,halign="center")
        cel(ws6,f6,4,sol,color=TEXTO,merge_to=(f6,7))
        cel(ws6,f6,8,cant if tiene else 0,color=TEXTO,halign="center")
        f6 += 1

    ws6.row_dimensions[f6].height = 60
    cel(ws6,f6,2,"Observaciones:",bold=True,color=TEXTO)
    cel(ws6,f6,3,r["observaciones"],color=TEXTO,merge_to=(f6,8))

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
