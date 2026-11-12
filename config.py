"""Configuración global: API key de Gemini, modelo y constantes del sistema."""
import os

import google.generativeai as genai


# La API key se lee de la variable de entorno configurada en Render
genai.configure(api_key=os.environ["GEMINI_API_KEY"])
MODEL_NAME = os.environ.get("GEMINI_MODEL", "gemini-3.5-flash-lite")

MAX_INTENTOS = 3      # cuántas veces el Jefe deja reintentar a un agente
TITULO_FUENTES = "Fuentes y créditos"
# Wikimedia pide identificar a quien usa su API. Opcional: define CONTACTO en Render.
USER_AGENT = "PDF2SlidesAgents/1.0 (proyecto de servicio social; uso educativo) " + os.environ.get("CONTACTO", "")

# ---- Tamaño de la presentación: lo dictan los temas y los caracteres de cada tema ----
CHARS_POR_DIAPOSITIVA = 450   # cuántos caracteres de contenido de un tema caben en una diapositiva
MAX_DIAP_POR_TEMA = 4         # tope de diapositivas para un solo tema
MAX_DIAPOSITIVAS = 60         # tope absoluto de diapositivas de contenido
COBERTURA_MINIMA = 0.95       # fracción mínima de temas del Recopilador que el plan debe cubrir

# ---- Imágenes ----
PROPORCION_IMAGENES = 0.6     # mínimo de diapositivas de contenido que llevan imagen

# ---- Textos largos ----
LIMITE_TEXTO = 150_000          # caracteres que el Recopilador lee de una vez (si hay más, lo parte en bloques)
LIMITE_TEXTO_REVISION = 80_000  # caracteres del documento que ve el Jefe al revisar
