"""Utilidades generales: llamada al LLM, JSON, texto y registro de revisiones del Jefe."""
import json
import re
from html import unescape

import google.generativeai as genai

from config import MODEL_NAME, TOPE_IMAGENES


def llamar_llm(system_instruction, prompt):
    """Una llamada a Gemini. 'prompt' puede ser texto o una lista (texto + imágenes)."""
    model = genai.GenerativeModel(MODEL_NAME, system_instruction=system_instruction)
    return model.generate_content(prompt).text.strip()


def parsear_json(raw):
    """Convierte la respuesta del modelo en JSON, quitando ``` si los puso."""
    raw = re.sub(r"^```(?:json)?", "", raw.strip())
    raw = re.sub(r"```$", "", raw).strip()
    return json.loads(raw)


def num_palabras(texto):
    return len(str(texto).split())


def con_feedback(prompt, feedback):
    """Si el Jefe rechazó un intento anterior, se lo decimos al agente."""
    if not feedback:
        return prompt
    return (
        prompt
        + "\n\nIMPORTANTE: el Jefe rechazó tu intento anterior por estos motivos:\n"
        + feedback
        + "\nCorrígelo en este nuevo intento."
    )


def _max_imagenes(n_slides):
    """Cupo de imágenes según el tamaño de la presentación."""
    return max(3, min(TOPE_IMAGENES, n_slides // 3))


def _chk(lista, regla, ok, detalle="", bloquea=True):
    """Registra una revisión del Jefe. Si 'bloquea' es False, solo avisa (⚠️)."""
    lista.append({"regla": regla, "ok": bool(ok), "detalle": "" if ok else (detalle or regla),
                  "bloquea": bloquea})


def _limpiar_html(s):
    return re.sub(r"\s+", " ", unescape(re.sub(r"<[^>]+>", "", s or ""))).strip()


def _recortar_texto(texto, limite=15000):
    """Si el PDF es largo, conserva el inicio y el final (ahí suele estar la bibliografía)."""
    if len(texto) <= limite:
        return texto
    return texto[:11500] + "\n[... texto omitido ...]\n" + texto[-3500:]


def _referencia_en_texto(ref, texto_min):
    """Guardia anti-invención: casi todas las palabras de la referencia deben estar en el PDF."""
    palabras = re.findall(r"\w{4,}", ref.lower())
    if not palabras:
        return False
    return sum(1 for w in palabras if w in texto_min) / len(palabras) >= 0.7
