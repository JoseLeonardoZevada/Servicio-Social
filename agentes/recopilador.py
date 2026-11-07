"""Agente Recopilador: resume el texto, lo agrupa en temas y recoge referencias."""
import json

from agentes.base import Agente
from utils import _recortar_texto, con_feedback, llamar_llm, parsear_json


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


class Recopilador(Agente):
    nombre = "Recopilador"
    rol = "Resume el texto, lo agrupa en temas y recoge las referencias que cita el documento (IA)"

    def ejecutar(self, entrada, feedback=None):
        esquema = {
            "titulo_documento": "título corto del documento (máx. 10 palabras)",
            "subtitulo": "frase que resume de qué trata (máx. 15 palabras)",
            "tema_general": "...",
            "razonamiento": "1 o 2 frases: cómo agrupaste el contenido y por qué",
            "secciones": [{"tema": "...", "puntos": ["...", "..."]}],
            "referencias": ["referencia bibliográfica tal como aparece en el documento"],
        }
        prompt = (
            "Lee el siguiente documento y extrae su información clave, agrupada "
            "por temas (entre 4 y 8 temas, cada uno con 2 a 5 puntos).\n"
            "En 'referencias' copia (máximo 8) las fuentes bibliográficas o autores citados que "
            "aparezcan en el documento. Si el documento no cita ninguna, devuelve una lista vacía. "
            "NO inventes ninguna.\n"
            "Estructura exacta: " + json.dumps(esquema, ensure_ascii=False) + "\n\n"
            'DOCUMENTO:\n"""\n' + _recortar_texto(entrada["texto"]) + '\n"""'
        )
        data = parsear_json(llamar_llm(SYS_RECOPILADOR, con_feedback(prompt, feedback)))
        data.setdefault("referencias", [])
        n_puntos = sum(len(s.get("puntos", [])) for s in data.get("secciones", []))
        return {
            "salida": data,
            "recibio": f"{len(entrada['texto'])} caracteres de texto",
            "entrego": (f"{len(data.get('secciones', []))} temas, {n_puntos} puntos clave, "
                        f"{len(data['referencias'])} referencias"),
            "razonamiento": data.get("razonamiento", ""),
            "detalle": _fmt_secciones(data),
            "notas": f"Tema general: {data.get('tema_general', '?')}",
        }
