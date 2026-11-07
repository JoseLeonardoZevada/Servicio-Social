"""Configuración global: API key de Gemini, modelo y constantes del sistema."""
import os

import google.generativeai as genai


# La API key se lee de la variable de entorno configurada en Render
genai.configure(api_key=os.environ["GEMINI_API_KEY"])
MODEL_NAME = os.environ.get("GEMINI_MODEL", "gemini-3.5-flash-lite")

MAX_INTENTOS = 3      # cuántas veces el Jefe deja reintentar a un agente
TOPE_IMAGENES = 12    # tope absoluto de imágenes por presentación (el cupo real escala con las diapositivas)
TITULO_FUENTES = "Fuentes y créditos"
# Wikimedia pide identificar a quien usa su API. Opcional: define CONTACTO en Render.
USER_AGENT = "PDF2SlidesAgents/1.0 (proyecto de servicio social; uso educativo) " + os.environ.get("CONTACTO", "")
