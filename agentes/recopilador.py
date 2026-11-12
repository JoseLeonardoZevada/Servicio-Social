"""Agente Recopilador: resume el texto, lo agrupa en temas y recoge referencias."""
import json
import math

from utils import Agente, con_feedback, llamar_llm, parsear_json
from config import LIMITE_TEXTO


SYS_RECOPILADOR = """
Eres un analista que lee documentos y extrae su información clave.
Responde ÚNICAMENTE con JSON válido, sin texto extra ni marcado de código.
Escribe en español neutro. No inventes información que no esté en el texto.
Nunca inventes referencias bibliográficas: solo copia las que aparezcan en el documento.
"""


def _fmt_secciones(data):
    partes = [
        f"Título: {data.get('titulo_documento', '')}",
        f"Subtítulo: {data.get('subtitulo', '')}",
        f"Tema general: {data.get('tema_general', '')}",
        "",
    ]
    for i, s in enumerate(data.get("secciones", []), 1):
        partes.append(f"Tema {i}: {s.get('tema', '?')}")
        for pt in s.get("puntos", []):
            partes.append(f"   - {pt}")
    partes.append("")
    refs = data.get("referencias", [])
    partes.append(f"Referencias citadas en el documento ({len(refs)}):")
    partes += [f"   - {r}" for r in refs] or ["   (ninguna)"]
    return "\n".join(partes)


def _trocear(texto, limite):
    """Parte un texto muy largo en bloques de hasta 'limite' caracteres, cortando en saltos de línea."""
    if len(texto) <= limite:
        return [texto]
    bloques, actual = [], ""
    for linea in texto.split("\n"):
        if actual and len(actual) + len(linea) + 1 > limite:
            bloques.append(actual)
            actual = ""
        while len(linea) > limite:
            bloques.append(linea[:limite])
            linea = linea[limite:]
        actual += linea + "\n"
    if actual.strip():
        bloques.append(actual)
    return bloques


class Recopilador(Agente):
    nombre = "Recopilador"
    rol = "Lee TODO el texto, lo agrupa en temas (tantos como pida el documento) y recoge las referencias que cita (IA)"

    def __init__(self, rango_temas=(4, 8)):
        self.rango = rango_temas

    def _prompt(self, texto, minimo, maximo, parte=None, paginas="?"):
        esquema = {
            "titulo_documento": "título corto del documento (máx. 10 palabras)",
            "subtitulo": "frase que resume de qué trata (máx. 15 palabras)",
            "tema_general": "...",
            "razonamiento": "1 o 2 frases: cómo agrupaste el contenido y por qué",
            "secciones": [{"tema": "...", "puntos": ["...", "..."]}],
            "referencias": ["referencia bibliográfica tal como aparece en el documento"],
        }
        donde = (f"Este es el bloque {parte[0]} de {parte[1]} del documento; extrae los temas de ESTE bloque. "
                 if parte else "")
        return (
            donde + "Lee el documento COMPLETO, de principio a fin (no te quedes solo con el inicio y el "
            "final), y extrae su información clave agrupada por temas:\n"
            "- Tú decides cuántos temas tiene realmente el documento: uno por cada sección o idea principal "
            f"distinta, en el mismo orden en que aparecen. (Tiene unas {paginas} páginas; como referencia, uno así "
            f"suele dar entre {minimo} y {maximo} temas, pero puede ser distinto.) Si trata muchos asuntos "
            "distintos, usa muchos temas; no los mezcles para quedarte con menos.\n"
            "- Cada tema lleva de 3 a 8 puntos. Cada punto es una frase completa (15 a 35 palabras) con el "
            "dato concreto que da el documento: definiciones, cifras, nombres, causas, pasos o ejemplos. "
            "Evita las frases vagas.\n"
            "- En 'referencias' copia (máximo 8) las fuentes bibliográficas o autores citados que "
            "aparezcan en el documento. Si el documento no cita ninguna, devuelve una lista vacía. "
            "NO inventes ninguna.\n"
            "Estructura exacta: " + json.dumps(esquema, ensure_ascii=False) + "\n\n"
            'DOCUMENTO:\n"""\n' + texto + '\n"""'
        )

    def ejecutar(self, entrada, feedback=None):
        texto = entrada["texto"]
        bloques = _trocear(texto, LIMITE_TEXTO)
        mn, mx = self.rango
        partes = []
        for j, bloque in enumerate(bloques, 1):
            frac = len(bloque) / max(1, len(texto))
            mn_b, mx_b = (mn, mx) if len(bloques) == 1 else (max(2, math.ceil(mn * frac)), max(3, math.ceil(mx * frac)))
            prompt = self._prompt(bloque, mn_b, mx_b, (j, len(bloques)) if len(bloques) > 1 else None,
                                  entrada.get("paginas", "?"))
            partes.append(parsear_json(llamar_llm(SYS_RECOPILADOR, con_feedback(prompt, feedback))))

        data = dict(partes[0])
        if len(partes) > 1:                      # PDF muy largo: se juntan los temas de todos los bloques
            data["secciones"] = [s for d in partes for s in d.get("secciones", [])]
            refs = []
            for d in partes:
                for r in d.get("referencias", []) or []:
                    if r not in refs:
                        refs.append(r)
            data["referencias"] = refs[:8]
            data["razonamiento"] = " ".join(str(d.get("razonamiento", "")) for d in partes).strip()
        data.setdefault("referencias", [])
        n_puntos = sum(len(s.get("puntos", [])) for s in data.get("secciones", []))
        return {
            "salida": data,
            "recibio": (f"{len(texto)} caracteres de texto completos"
                        + (f" (leídos en {len(bloques)} bloques)" if len(bloques) > 1 else "")
                        + f"; referencia: entre {mn} y {mx} temas (la IA decide)"),
            "entrego": (f"{len(data.get('secciones', []))} temas, {n_puntos} puntos clave, "
                        f"{len(data['referencias'])} referencias"),
            "razonamiento": data.get("razonamiento", ""),
            "detalle": _fmt_secciones(data),
            "notas": f"Tema general: {data.get('tema_general', '?')}",
        }
