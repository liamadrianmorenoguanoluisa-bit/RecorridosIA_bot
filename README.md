# RecorridosIA 🗺️

Bot de Telegram para inspección inteligente de rutas de fibra óptica Telconet.
Genera automáticamente el informe **FOR FO 02** en Excel.

## Estructura del proyecto

```
RecorridosIA/
├── main.py                        ← punto de entrada (Render ejecuta esto)
├── requirements.txt
├── README.md
├── bot/
│   └── main.py                   ← lógica del bot de Telegram
├── vision/
│   └── gemini.py                 ← análisis de imágenes con Gemini AI
└── reports/
    ├── excel.py                  ← generador del FOR FO 02
    └── plantilla_FOR_FO_02.xlsx  ← plantilla base (subir manualmente)
```

## Variables de entorno — configurar en Render

| Variable | Descripción |
|----------|-------------|
| `BOT_TOKEN` | Token del bot @RecorridosIA_bot (BotFather) |
| `GEMINI_API_KEY` | Clave de API de Google Gemini AI |
| `MAPILLARY_TOKEN` | Token de acceso a Mapillary |
| `TOTP_SECRET` | Clave secreta 2FA para acceso al bot |

## Cómo generar el TOTP_SECRET

1. Instala en tu PC: `pip install pyotp`
2. Ejecuta:
```python
import pyotp
print(pyotp.random_base32())
# Ejemplo: JBSWY3DPEHPK3PXP
```
3. Copia ese valor como `TOTP_SECRET` en Render
4. Abre **Google Authenticator** → Agregar cuenta → Clave manual
5. Nombre: `RecorridosIA` / Clave: el valor generado

## Deploy en Render

1. Sube este repositorio a GitHub
2. Render → **New Web Service** → conectar repo
3. Configurar:
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `python main.py`
4. Agregar las 4 variables de entorno
5. Click en **Deploy** 🚀

## Cómo usar el bot

```
1. Abre @RecorridosIA_bot en Telegram
2. Escribe /start
3. Ingresa el código de 6 dígitos de Google Authenticator ← TOTP
4. Acceso autorizado ✅
5. Selecciona: Inspeccionar
6. El bot pide los datos del recorrido uno a uno
7. Envías las fotos → la IA detecta novedades automáticamente
8. El bot genera el Excel FOR FO 02 y te lo envía por Telegram ✅
```
