"""Agente Seccionador: decide cuántas diapositivas hay y qué contenido lleva cada una."""
import json

from agentes.base import Agente
from utils import con_feedback, llamar_llm, parsear_json


SYS_SECCIONADOR = """
Eres un asistente experto en diseño de presentaciones. Decides cuántas
diapositivas se necesitan, qué información va en cada una y en qué formato,
pensando en explicarlas en voz alta frente a una audiencia.

Reglas que debes seguir SIEMPRE:
- Responde ÚNICAMENTE con JSON válido. Nada de texto extra ni ``` .
- Español neutro, tono claro y profesional.
- NO incluyas portada ni diapositiva de fuentes, y NO elijas imágenes: otros agentes se encargan.
- Cada "title" máximo 8 palabras.
- Cada diapositiva tiene un "tipo":
    * "parrafo": un texto corrido de 25 a 55 palabras, redactado con fluidez
      (nada de listas). Deja "bullets" vacío.
    * "mixto": un "parrafo" corto (10 a 30 palabras) más 2 a 5 bullets.
    * "bullets": 3 a 8 bullets. Úsalo cuando el contenido sea naturalmente una lista
      (ingredientes, pasos, requisitos). En el resto de los casos prefiere texto corrido.
  Como máximo el 60% de las diapositivas pueden ser solo "bullets".
- Cada bullet máximo 15 palabras, frase corta. No repitas ideas entre diapositivas.
- La primera diapositiva presenta el tema general (introducción).
- La última diapositiva es un resumen o conclusión.
- Todos los temas del resumen deben quedar cubiertos. En "temas" pon los números
  (campo "numero" del resumen) de los temas que cubre cada diapositiva.
- Explica tu decisión en "razonamiento": por qué elegiste ese número de diapositivas.
"""

EJEMPLO_SLIDES = {
    "razonamiento": "Usé 2 diapositivas porque el texto es corto: una explicativa y una de cierre.",
    "slides": [
        {
            "title": "¿Qué es el reciclaje?",
            "temas": [1],
            "tipo": "parrafo",
            "parrafo": "El reciclaje convierte materiales de desecho en nuevos productos. Así se "
                       "reduce el consumo de materias primas y se disminuye la contaminación del aire "
                       "y del agua, además de ahorrar energía en la producción.",
            "bullets": [],
        },
        {
            "title": "Conclusión",
            "temas": [1, 2],
            "tipo": "mixto",
            "parrafo": "Reciclar bien empieza en casa y beneficia a todos.",
            "bullets": ["Separar los materiales desde el hogar", "Reducir el impacto ambiental"],
        },
    ],
}


class Seccionador(Agente):
    nombre = "Seccionador"
    rol = "Decide cuántas diapositivas hay y qué formato y contenido lleva cada una (IA)"

    def __init__(self, tope):
        self.tope = tope

    def ejecutar(self, entrada, feedback=None):
        resumen = {k: v for k, v in entrada.items() if k not in ("referencias", "secciones")}
        resumen["temas"] = [{"numero": i, "tema": sec.get("tema"), "puntos": sec.get("puntos")}
                            for i, sec in enumerate(entrada.get("secciones", []), 1)]
        prompt = (
            "Ejemplo del formato esperado (solo la estructura, no el contenido):\n"
            + json.dumps(EJEMPLO_SLIDES, ensure_ascii=False)
            + "\n\nAhora, con base en este resumen del documento, arma las diapositivas de contenido. "
            f"Decide TÚ cuántas hacen falta (mínimo 3, máximo {self.tope}): si el documento es una "
            "colección de unidades (recetas, casos, capítulos, secciones), usa una diapositiva por unidad "
            "(o más si la unidad es larga); si no, las que pida el contenido. Todos los temas deben quedar "
            'cubiertos. Devuelve SOLO el JSON {"razonamiento": "...", "slides": [{"title": "...", '
            '"temas": [1], "tipo": "...", "parrafo": "...", "bullets": ["..."]}]}.\n\n'
            "RESUMEN:\n" + json.dumps(resumen, ensure_ascii=False)
        )
        data = parsear_json(llamar_llm(SYS_SECCIONADOR, con_feedback(prompt, feedback)))
        slides = data["slides"]

        detalle = []
        for i, sl in enumerate(slides, 1):
            temas = ", ".join(str(t) for t in (sl.get("temas") or []))
            detalle.append(f"Diapositiva {i} [{sl.get('tipo', '?')}]: {sl.get('title', '?')}   (cubre temas: {temas or '-'})")
            if sl.get("parrafo"):
                detalle.append(f"   Texto: {sl['parrafo']}")
            for b in sl.get("bullets", []):
                detalle.append(f"   - {b}")

        return {
            "salida": {
                "slides": slides,
                "titulo": entrada.get("titulo_documento", "Presentación"),
                "subtitulo": entrada.get("subtitulo", ""),
                "tema_general": entrada.get("tema_general", ""),
                "razonamiento": data.get("razonamiento", ""),
            },
            "recibio": f"{len(entrada.get('secciones', []))} temas del Recopilador (tope: {self.tope} diapositivas)",
            "entrego": f"Plan de {len(slides)} diapositivas de contenido",
            "razonamiento": data.get("razonamiento", ""),
            "detalle": "\n".join(detalle),
            "notas": " / ".join(sl.get("title", "?") for sl in slides),
        }
