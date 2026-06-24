"""
RecorridosIA — Análisis de imágenes con Gemini AI
"""

import os
import base64
import json
import httpx
from datetime import datetime

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
GEMINI_URL     = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={GEMINI_API_KEY}"

NOVEDADES_TELCONET = [
    "HERRAJES EN MAL ESTADO.",
    "FALTA DE HERRAJES.",
    "POSTES EN MAL ESTADO.",
    "POSTES INCLINADOS.",
    "RETENCIÓN(S) EN MAL ESTADO.",
    "VANOS POR RETEMPLAR.",
    "MANGAS SUELTAS.",
    "MANGAS ABIERTAS/DAÑADAS.",
    "RESERVAS SUELTAS.",
    "CRUCES DE VÍAS BAJOS.",
    "VEGETACIÓN SOBRE FIBRA/MANGA.",
    "DOCUMENTACIÓN UNIFILAR DE HILOS.",
    "LÍNEA ELÉCTRICA EN MAL ESTADO.",
    "AMPLIACIÓN DE VÍA.",
    "CABLE LASTIMADO.",
    "POZO SIN TAPA O EN MAL ESTADO.",
    "ELEMENTOS SIN ETIQUETAS ACRÍLICAS.",
    "RIESGO DE DERRUMBE O DESLAVE.",
    "RIESGO DE INUNDACIONES.",
    "RIESGO DE INCENDIO.",
]

REMEDIOS = {
    "VEGETACIÓN SOBRE FIBRA/MANGA.": "REALIZAR LA PODA O RETIRO DE VEGETACIÓN QUE COMPROMETA LA INTEGRIDAD O SEGURIDAD DEL CABLE. EN CASO DE REQUERIR PERMISOS, DOCUMENTAR LA NOVEDAD.",
    "HERRAJES EN MAL ESTADO.": "REALIZAR EL REEMPLAZO INMEDIATO DEL HERRAJE AFECTADO, GARANTIZANDO LA CORRECTA SUJECIÓN DEL CABLE.",
    "POSTES EN MAL ESTADO.": "DOCUMENTAR MEDIANTE REGISTRO FOTOGRÁFICO Y COORDENADAS, Y REPORTAR PARA GESTIONAR EL REEMPLAZO DEL POSTE.",
    "POSTES INCLINADOS.": "DOCUMENTAR MEDIANTE REGISTRO FOTOGRÁFICO Y COORDENADAS, Y REPORTAR PARA GESTIONAR EL APLOME DEL POSTE.",
    "MANGAS SUELTAS.": "ASEGURAR LA MANGA AL POSTE EN CONFIGURACIÓN TIPO 'FIGURA 8', CONFORME AL ESTÁNDAR.",
    "MANGAS ABIERTAS/DAÑADAS.": "REEMPLAZAR TAPAS Y SELLOS, GARANTIZANDO EL CIERRE HERMÉTICO Y LA PROTECCIÓN DEL EMPALME.",
    "CABLE LASTIMADO.": "DOCUMENTAR E INFORMAR PARA PROGRAMAR EL CAMBIO DEL TRAMO DE CABLE.",
    "DOCUMENTACIÓN UNIFILAR DE HILOS.": "DOCUMENTAR O SOLICITAR LA PROGRAMACIÓN DE TRABAJO PARA OBTENER LA INFORMACIÓN; UTILIZAR UN SEGUIDOR DE SEÑAL.",
    "CRUCES DE VÍAS BAJOS.": "AJUSTAR LA ALTURA DEL CABLE ELEVÁNDOLO A LA DISTANCIA REGLAMENTARIA.",
    "POZO SIN TAPA O EN MAL ESTADO.": "SOLICITAR LA EJECUCIÓN DE TRABAJOS DE OBRA CIVIL PARA SU INSTALACIÓN O CORRECCIÓN.",
    "RIESGO DE DERRUMBE O DESLAVE.": "DOCUMENTAR EL RIESGO Y SOLICITAR AL COORDINADOR LA REUBICACIÓN DEL RECORRIDO DEL CABLE.",
    "RIESGO DE INUNDACIONES.": "DOCUMENTAR EL RIESGO Y SOLICITAR AL COORDINADOR LA REUBICACIÓN DEL RECORRIDO DEL CABLE.",
    "RIESGO DE INCENDIO.": "DOCUMENTAR EL RIESGO Y SOLICITAR AL COORDINADOR LA REUBICACIÓN DEL RECORRIDO DEL CABLE.",
    "LÍNEA ELÉCTRICA EN MAL ESTADO.": "DOCUMENTAR EL RIESGO Y SOLICITAR AL COORDINADOR EL REPORTE AL ÁREA DE REGULATORIO.",
    "AMPLIACIÓN DE VÍA.": "DOCUMENTAR, REGISTRAR EL CONTACTO DEL RESPONSABLE DE LA OBRA Y COORDINAR MEDIDAS DE MITIGACIÓN.",
    "ELEMENTOS SIN ETIQUETAS ACRÍLICAS.": "VERIFICAR, COLOCAR ETIQUETA ACRÍLICA Y ETIQUETAR CON EL CÓDIGO DE RUTA.",
    "FALTA DE HERRAJES.": "INSTALAR LOS HERRAJES CONFORME A LA NORMATIVA TÉCNICA, ASEGURANDO LA CORRECTA FIJACIÓN DEL CABLE AL POSTE.",
    "VANOS POR RETEMPLAR.": "REALIZAR EL RETEMPLADO DEL CABLE PARA RESTABLECER LA TENSIÓN ADECUADA.",
    "RESERVAS SUELTAS.": "REORGANIZAR Y ASEGURAR LA RESERVA EN 'FIGURA 8' CONFORME A LO ESTABLECIDO.",
}

SIN_NOVEDAD_MOTIVO  = "NO SE REGISTRAN NOVEDADES DURANTE LA INSPECCIÓN."
SIN_NOVEDAD_REMEDIO = "EN ESTE PUNTO LA FIBRA SE ENCUENTRA SIN NOVEDAD."


async def analizar_media(imagenes: list) -> list:
    """
    Recibe lista de bytes de imágenes.
    Devuelve lista de dicts con novedades detectadas.
    Si no hay imágenes o no detecta nada, retorna novedad vacía.
    """
    if not imagenes:
        return [_sin_novedad()]

    novedades = []
    ahora = datetime.now()

    for i, img_bytes in enumerate(imagenes):
        resultado = await _analizar_imagen(img_bytes, i + 1)
        if resultado:
            resultado["fecha"]       = ahora.strftime("%d/%m/%Y")
            resultado["hora_inicio"] = ahora.strftime("%H:%M:%S")
            resultado["hora_fin"]    = ahora.strftime("%H:%M:%S")
            novedades.append(resultado)

    return novedades if novedades else [_sin_novedad()]


async def _analizar_imagen(img_bytes: bytes, numero: int) -> dict | None:
    """Envía una imagen a Gemini y retorna la novedad detectada."""
    img_b64 = base64.b64encode(img_bytes).decode("utf-8")

    prompt = f"""Eres un experto en inspección de rutas de fibra óptica de Telconet Ecuador.
Analiza esta imagen de una inspección de ruta interurbana.

Clasifica ÚNICAMENTE si detectas alguna de estas novedades:
{chr(10).join(f'- {n}' for n in NOVEDADES_TELCONET)}

Responde SOLO con JSON válido, sin texto adicional:
{{
  "tiene_novedad": true/false,
  "motivo": "NOMBRE EXACTO DE LA NOVEDAD DE LA LISTA O vacío",
  "coordenadas": "si se ven coordenadas en la imagen o vacío"
}}

Si la imagen muestra infraestructura en buen estado, responde tiene_novedad: false."""

    payload = {
        "contents": [{
            "parts": [
                {"text": prompt},
                {"inline_data": {"mime_type": "image/jpeg", "data": img_b64}}
            ]
        }]
    }

    try:
        async with httpx.AsyncClient(timeout=30) as client:
            resp = await client.post(GEMINI_URL, json=payload)
            resp.raise_for_status()
            data = resp.json()
            texto = data["candidates"][0]["content"]["parts"][0]["text"]
            texto = texto.replace("```json", "").replace("```", "").strip()
            resultado = json.loads(texto)

            if not resultado.get("tiene_novedad"):
                return None

            motivo = resultado.get("motivo", "").upper()
            return {
                "motivo":      motivo,
                "remedio":     REMEDIOS.get(motivo, "DOCUMENTAR Y REPORTAR AL COORDINADOR."),
                "coordenadas": resultado.get("coordenadas", ""),
                "tarea_pendiente": "",
                "foto_antes":  img_bytes,
                "foto_despues": None,
            }
    except Exception as e:
        print(f"Error Gemini imagen {numero}: {e}")
        return None


def _sin_novedad() -> dict:
    ahora = datetime.now()
    return {
        "motivo":           SIN_NOVEDAD_MOTIVO,
        "remedio":          SIN_NOVEDAD_REMEDIO,
        "coordenadas":      "",
        "tarea_pendiente":  "",
        "fecha":            ahora.strftime("%d/%m/%Y"),
        "hora_inicio":      ahora.strftime("%H:%M:%S"),
        "hora_fin":         ahora.strftime("%H:%M:%S"),
        "foto_antes":       None,
        "foto_despues":     None,
    }
