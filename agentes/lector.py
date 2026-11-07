"""Agente Lector: lee el PDF con Gemini (respaldo: pdfplumber)."""
import os
import re

import pdfplumber

from agentes.base import Agente
from utils import llamar_llm


SYS_LECTOR = """
Eres un transcriptor de documentos. Copias el texto de un PDF de forma fiel, en orden,
sin resumir, sin corregir y sin inventar nada. Incluye títulos, tablas (como texto) y
la bibliografía. Si una página es una imagen escaneada, haz OCR. Si hay una figura
importante, descríbela en una línea entre corchetes: [Figura: ...].
Responde solo con el texto transcrito.
"""


def _extraer_plumber(ruta):
    paginas = []
    with pdfplumber.open(ruta) as pdf:
        for p in pdf.pages:
            paginas.append(p.extract_text() or "")
    return paginas


def _solapamiento(a, b):
    """Fracción de las palabras (5+ letras) de 'a' que también aparecen en 'b'."""
    pa = set(re.findall(r"\w{5,}", a.lower()))
    if not pa:
        return 1.0
    pb = set(re.findall(r"\w{5,}", b.lower()))
    return len(pa & pb) / len(pa)


class Lector(Agente):
    nombre = "Lector"
    rol = "Lee el PDF con Gemini (también escaneados) y usa pdfplumber como respaldo"

    def ejecutar(self, entrada, feedback=None):
        paginas = _extraer_plumber(entrada)                    # respaldo, siempre se calcula
        t_plumber = "\n".join(t for t in paginas if t)
        vacias = sum(1 for t in paginas if not t)

        fuente, t_ia, nota = "pdfplumber", "", ""
        if feedback:
            nota = "El Jefe rechazó la lectura anterior, así que usé el respaldo (pdfplumber)."
        else:
            try:
                if os.path.getsize(entrada) > 15_000_000:
                    raise ValueError("el PDF es muy pesado para enviarlo a Gemini")
                with open(entrada, "rb") as f:
                    datos = f.read()
                t_ia = llamar_llm(SYS_LECTOR, ["Transcribe todo el texto de este PDF.",
                                                {"mime_type": "application/pdf", "data": datos}])
                fuente = "Gemini"
            except Exception as e:
                nota = f"Gemini no pudo leer el PDF ({type(e).__name__}: {e}); usé pdfplumber."

        texto = t_ia if fuente == "Gemini" else t_plumber
        solap = None
        if fuente == "Gemini" and len(t_plumber.strip()) >= 200:
            solap = _solapamiento(t_plumber, t_ia)

        lineas = [f"Fuente usada: {fuente}"]
        if nota:
            lineas.append(nota)
        if solap is not None:
            lineas.append(f"Coincidencia de palabras con pdfplumber: {round(solap * 100)}%")
        if fuente == "Gemini":
            lineas.append(f"Inicio de la transcripción: {t_ia[:400].strip()!r}...")
        lineas.append("")
        lineas.append("Lectura de respaldo (pdfplumber) por página:")
        for i, t in enumerate(paginas[:12], 1):
            muestra = t[:200].replace("\n", " ").strip()
            lineas.append(f"Pág. {i} ({len(t)} caracteres): {muestra}..." if t else f"Pág. {i}: sin texto extraíble")
        if len(paginas) > 12:
            lineas.append(f"... y {len(paginas) - 12} páginas más")

        return {
            "salida": {"texto": texto, "paginas": len(paginas), "fuente": fuente,
                       "texto_plumber": t_plumber, "solapamiento": solap},
            "recibio": f"PDF: {os.path.basename(entrada)}",
            "entrego": f"{len(paginas)} páginas, {len(texto)} caracteres de texto (leído con {fuente})",
            "razonamiento": (
                "Leí el PDF con pdfplumber (respaldo) y se lo mandé a Gemini para una transcripción más "
                "completa, que sirve también con PDFs escaneados o con tablas. "
                + (nota or "Usé la transcripción de Gemini porque coincide con la lectura de pdfplumber.")
            ),
            "detalle": "\n".join(lineas),
            "notas": f"{vacias} páginas sin texto extraíble para pdfplumber",
        }
