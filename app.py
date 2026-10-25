import os
import json
import re
import time
import tempfile
from html import escape

import gradio as gr
import pdfplumber
import google.generativeai as genai
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.dml import MSO_FILL
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn

# La API key se lee de la variable de entorno configurada en Render
genai.configure(api_key=os.environ["GEMINI_API_KEY"])
MODEL_NAME = os.environ.get("GEMINI_MODEL", "gemini-3.5-flash-lite")

MAX_INTENTOS = 3  # cuántas veces el Jefe deja reintentar a un agente


# ----------------------------------------------------------------------
# Utilidades
# ----------------------------------------------------------------------
def llamar_llm(system_instruction, prompt):
    """Una llamada a Gemini con la instrucción de sistema del agente."""
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


def _chk(lista, regla, ok, detalle=""):
    """Registra una revisión del Jefe: qué regla, si pasó y por qué falló."""
    lista.append({"regla": regla, "ok": bool(ok), "detalle": "" if ok else (detalle or regla)})


# ----------------------------------------------------------------------
# Agentes (cada uno hace UNA cosa y devuelve qué recibió/entregó/razonó)
# ----------------------------------------------------------------------
class Agente:
    nombre = "Agente"
    rol = ""

    def ejecutar(self, entrada, feedback=None):
        """Debe devolver {"salida", "recibio", "entrego", "razonamiento", "detalle", "notas"}"""
        raise NotImplementedError


class Lector(Agente):
    nombre = "Lector"
    rol = "Extrae el texto del PDF página por página (sin IA)"

    def ejecutar(self, entrada, feedback=None):
        paginas = []
        with pdfplumber.open(entrada) as pdf:
            for p in pdf.pages:
                paginas.append(p.extract_text() or "")
        texto = "\n".join(t for t in paginas if t)
        vacias = sum(1 for t in paginas if not t)

        lineas = []
        for i, t in enumerate(paginas[:12], 1):
            if t:
                muestra = t[:300].replace("\n", " ").strip()
                lineas.append(f"Pág. {i} ({len(t)} caracteres): {muestra}...")
            else:
                lineas.append(f"Pág. {i}: sin texto extraíble")
        if len(paginas) > 12:
            lineas.append(f"... y {len(paginas) - 12} páginas más")

        return {
            "salida": {"texto": texto, "paginas": len(paginas)},
            "recibio": f"PDF: {os.path.basename(entrada)}",
            "entrego": f"{len(paginas)} páginas, {len(texto)} caracteres de texto",
            "razonamiento": (
                "Abrí el PDF con pdfplumber, leí cada página en orden y uní el texto. "
                "No usé IA: es extracción directa. Si hay páginas sin texto, suelen ser "
                "imágenes escaneadas."
            ),
            "detalle": "\n".join(lineas),
            "notas": f"{vacias} páginas sin texto extraíble",
        }


SYS_RECOPILADOR = """
Eres un analista que lee documentos y extrae su información clave.
Responde ÚNICAMENTE con JSON válido, sin texto extra ni marcado de código.
Escribe en español neutro. No inventes información que no esté en el texto.
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
    return "\n".join(partes)


class Recopilador(Agente):
    nombre = "Recopilador"
    rol = "Resume el texto y lo agrupa en temas con puntos clave (IA)"

    def ejecutar(self, entrada, feedback=None):
        esquema = {
            "titulo_documento": "título corto del documento (máx. 10 palabras)",
            "subtitulo": "frase que resume de qué trata (máx. 15 palabras)",
            "tema_general": "...",
            "razonamiento": "1 o 2 frases: cómo agrupaste el contenido y por qué",
            "secciones": [{"tema": "...", "puntos": ["...", "..."]}],
        }
        prompt = (
            "Lee el siguiente documento y extrae su información clave, agrupada "
            "por temas (entre 4 y 8 temas, cada uno con 2 a 5 puntos).\n"
            "Estructura exacta: " + json.dumps(esquema, ensure_ascii=False) + "\n\n"
            'DOCUMENTO:\n"""\n' + entrada["texto"][:15000] + '\n"""'
        )
        data = parsear_json(llamar_llm(SYS_RECOPILADOR, con_feedback(prompt, feedback)))
        n_puntos = sum(len(s.get("puntos", [])) for s in data.get("secciones", []))
        return {
            "salida": data,
            "recibio": f"{len(entrada['texto'])} caracteres de texto",
            "entrego": f"{len(data.get('secciones', []))} temas, {n_puntos} puntos clave",
            "razonamiento": data.get("razonamiento", ""),
            "detalle": _fmt_secciones(data),
            "notas": f"Tema general: {data.get('tema_general', '?')}",
        }


SYS_SECCIONADOR = """
Eres un asistente experto en diseño de presentaciones. Decides cuántas
diapositivas se necesitan y qué información va en cada una, pensando en
explicarlas en voz alta frente a una audiencia.

Reglas que debes seguir SIEMPRE:
- Responde ÚNICAMENTE con JSON válido. Nada de texto extra ni ``` .
- Español neutro, tono claro y profesional.
- NO incluyas portada: otro agente la crea. Tú solo planeas el contenido.
- Cada "title" máximo 8 palabras.
- Cada "bullet" máximo 15 palabras, frase corta, nunca un párrafo.
- No repitas la misma idea en dos diapositivas.
- La primera diapositiva presenta el tema general (introducción).
- La última diapositiva es un resumen o conclusión.
- Explica tu decisión en "razonamiento" y, en cada diapositiva, indica en
  "origen" de qué tema(s) del resumen salió su contenido.
"""

EJEMPLO_SLIDES = {
    "razonamiento": "Usé 2 diapositivas porque el texto es corto: una de contenido y una de cierre.",
    "slides": [
        {
            "title": "¿Qué es el reciclaje?",
            "origen": "Tema 1: Definición",
            "bullets": [
                "Convierte materiales de desecho en nuevos productos",
                "Reduce el consumo de materias primas frescas",
                "Disminuye la contaminación del aire y el agua",
            ],
        },
        {
            "title": "Conclusión",
            "origen": "Temas 1 y 2: síntesis",
            "bullets": [
                "El reciclaje reduce el impacto ambiental",
                "La separación desde el hogar es clave",
                "Todos podemos participar",
            ],
        },
    ],
}


class Seccionador(Agente):
    nombre = "Seccionador"
    rol = "Decide cuántas diapositivas hay y qué contenido lleva cada una (IA)"

    def ejecutar(self, entrada, feedback=None):
        prompt = (
            "Ejemplo del formato esperado (solo la estructura, no el contenido):\n"
            + json.dumps(EJEMPLO_SLIDES, ensure_ascii=False)
            + "\n\nAhora, con base en este resumen estructurado del documento, "
            "arma entre 5 y 10 diapositivas de contenido, cada una con entre 3 y 5 bullets. "
            'Devuelve SOLO el JSON {"razonamiento": "...", "slides": [{"title": "...", '
            '"origen": "...", "bullets": ["..."]}]}.\n\n'
            "RESUMEN:\n" + json.dumps(entrada, ensure_ascii=False)
        )
        data = parsear_json(llamar_llm(SYS_SECCIONADOR, con_feedback(prompt, feedback)))
        slides = data["slides"]

        detalle = []
        for i, s in enumerate(slides, 1):
            detalle.append(f"Diapositiva {i}: {s.get('title', '?')}   [origen: {s.get('origen', '-')}]")
            for b in s.get("bullets", []):
                detalle.append(f"   - {b}")

        return {
            "salida": {
                "slides": slides,
                "titulo": entrada.get("titulo_documento", "Presentación"),
                "subtitulo": entrada.get("subtitulo", ""),
                "tema_general": entrada.get("tema_general", ""),
                "razonamiento": data.get("razonamiento", ""),
            },
            "recibio": f"{len(entrada.get('secciones', []))} temas del Recopilador",
            "entrego": f"Plan de {len(slides)} diapositivas de contenido",
            "razonamiento": data.get("razonamiento", ""),
            "detalle": "\n".join(detalle),
            "notas": " / ".join(s.get("title", "?") for s in slides),
        }


class GeneradorPPTX(Agente):
    nombre = "Generador PPTX"
    rol = "Construye el archivo .pptx (portada + contenido) a partir del plan (sin IA)"

    def __init__(self, ruta_salida):
        self.ruta_salida = ruta_salida

    def ejecutar(self, entrada, feedback=None):
        prs = Presentation()
        prs.slide_width = Inches(13.333)  # formato 16:9
        prs.slide_height = Inches(7.5)

        # Portada (layout 0 = diapositiva de título)
        portada = prs.slides.add_slide(prs.slide_layouts[0])
        portada.shapes.title.text = entrada["titulo"]
        portada.placeholders[1].text = entrada["subtitulo"] or entrada.get("tema_general", "")

        detalle = [f"Diapositiva 1 (portada): {entrada['titulo']}"]
        for i, info in enumerate(entrada["slides"], 2):
            slide = prs.slides.add_slide(prs.slide_layouts[1])
            slide.shapes.title.text = info.get("title", "Sin titulo")
            tf = slide.placeholders[1].text_frame
            tf.clear()
            for j, bullet in enumerate(info.get("bullets", [])):
                if j == 0:
                    tf.text = bullet
                else:
                    tf.add_paragraph().text = bullet
            detalle.append(f"Diapositiva {i}: {info.get('title', '')} ({len(info.get('bullets', []))} bullets)")
        prs.save(self.ruta_salida)

        total = len(entrada["slides"]) + 1
        return {
            "salida": {
                "ruta": self.ruta_salida,
                "n_slides": total,
                "titulo": entrada["titulo"],
                "tema": entrada.get("tema_general", ""),
            },
            "recibio": f"Plan de {len(entrada['slides'])} diapositivas + título y subtítulo",
            "entrego": f"Archivo {os.path.basename(self.ruta_salida)} con {total} diapositivas (portada incluida)",
            "razonamiento": (
                "Creé primero la portada con el título y subtítulo del documento, luego una "
                "diapositiva por cada elemento del plan, en el mismo orden. Formato 16:9. "
                "Aún sin estilo: eso lo hace el Decorador."
            ),
            "detalle": "\n".join(detalle),
            "notas": "",
        }


PALETAS = {
    "oceano": {"portada": (0x0B, 0x3C, 0x5D), "fondo": (0xF7, 0xFA, 0xFC), "titulo": (0x0B, 0x3C, 0x5D),
               "texto": (0x2B, 0x35, 0x40), "acento": (0x32, 0x82, 0xB8), "acento2": (0x7F, 0xC4, 0xE8),
               "tinte": (0xDC, 0xEC, 0xF6)},
    "bosque": {"portada": (0x1B, 0x4D, 0x2E), "fondo": (0xF6, 0xFA, 0xF6), "titulo": (0x1B, 0x4D, 0x2E),
               "texto": (0x2A, 0x33, 0x2C), "acento": (0x4C, 0xAF, 0x50), "acento2": (0xA5, 0xD6, 0xA7),
               "tinte": (0xE0, 0xF0, 0xE0)},
    "atardecer": {"portada": (0x7A, 0x24, 0x0E), "fondo": (0xFF, 0xF8, 0xF1), "titulo": (0x7A, 0x24, 0x0E),
                  "texto": (0x3A, 0x2E, 0x2A), "acento": (0xE6, 0x7E, 0x22), "acento2": (0xF5, 0xB7, 0x6B),
                  "tinte": (0xFB, 0xE5, 0xD0)},
    "grafito": {"portada": (0x22, 0x22, 0x2B), "fondo": (0xF5, 0xF5, 0xF7), "titulo": (0x22, 0x22, 0x2B),
                "texto": (0x3C, 0x3C, 0x44), "acento": (0x6C, 0x5C, 0xE7), "acento2": (0xB1, 0xA8, 0xF5),
                "tinte": (0xE6, 0xE3, 0xFA)},
}
BLANCO = (0xFF, 0xFF, 0xFF)

SYS_DECORADOR = """
Eres un diseñador de presentaciones. Eliges la paleta de colores que mejor
va con el tema del documento. Responde ÚNICAMENTE con JSON válido.
"""


def _rgb(t):
    return RGBColor(*t)


def _forma(slide, tipo, x, y, w, h, color, atras=True):
    """Dibuja una forma de color sin borde ni sombra; opcionalmente detrás del texto."""
    f = slide.shapes.add_shape(tipo, int(x), int(y), int(w), int(h))
    f.fill.solid()
    f.fill.fore_color.rgb = _rgb(color)
    f.line.fill.background()
    f.shadow.inherit = False
    estilo = f._element.find(qn("p:style"))  # quita sombras/bordes heredados del tema
    if estilo is not None:
        f._element.remove(estilo)
    if atras:
        arbol = slide.shapes._spTree
        arbol.remove(f._element)
        arbol.insert(2, f._element)
    return f


def _colocar(shape, x, y, w, h):
    shape.left, shape.top, shape.width, shape.height = int(x), int(y), int(w), int(h)


def _estilo(shape, tam, color, negrita=False, alinear=PP_ALIGN.LEFT,
            anclaje=MSO_ANCHOR.MIDDLE, espacio=0):
    tf = shape.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anclaje
    for par in tf.paragraphs:
        par.alignment = alinear
        if espacio:
            par.space_after = Pt(espacio)
        for run in par.runs:
            run.font.size = Pt(tam)
            run.font.bold = negrita
            run.font.color.rgb = _rgb(color)


class Decorador(Agente):
    nombre = "Decorador"
    rol = "Elige paleta (IA) y aplica diseño: portada, colores, tipografía y formas (código)"

    def ejecutar(self, entrada, feedback=None):
        ruta, tema = entrada["ruta"], entrada["tema"]
        nombre_paleta, motivo = "oceano", "paleta por defecto"
        try:
            prompt = (
                f"Tema del documento: {tema}\n"
                f"Paletas disponibles: {list(PALETAS)}\n"
                'Responde {"paleta": "<nombre>", "motivo": "<una frase corta>"}'
            )
            d = parsear_json(llamar_llm(SYS_DECORADOR, prompt))
            if d.get("paleta") in PALETAS:
                nombre_paleta, motivo = d["paleta"], d.get("motivo", "")
        except Exception as e:  # si la IA falla, no tiramos todo el proceso
            motivo = f"la IA falló ({e}); usé la paleta por defecto"

        p = PALETAS[nombre_paleta]
        prs = Presentation(ruta)
        W, H = prs.slide_width, prs.slide_height
        for i, slide in enumerate(prs.slides):
            slide.background.fill.solid()
            if i == 0:
                self._portada(slide, p, W, H)
            else:
                self._contenido(slide, p, W, H, i)
        prs.save(ruta)

        detalle = (
            f"Paleta elegida: {nombre_paleta}\n"
            "Portada: fondo oscuro, dos círculos decorativos, título 44 pt en blanco, "
            "línea de acento y subtítulo.\n"
            f"Diapositivas de contenido ({len(prs.slides) - 1}): fondo claro, barra lateral de acento, "
            "círculo suave en la esquina, título 34 pt alineado a la izquierda con subrayado, "
            "bullets 24 pt con espaciado y número de diapositiva en un círculo."
        )
        return {
            "salida": {"ruta": ruta, "n_slides": len(prs.slides), "paleta": nombre_paleta},
            "recibio": f"Archivo básico de {len(prs.slides)} diapositivas",
            "entrego": f"Mismo archivo con diseño aplicado (paleta '{nombre_paleta}')",
            "razonamiento": motivo,
            "detalle": detalle,
            "notas": "",
        }

    def _portada(self, slide, p, W, H):
        slide.background.fill.fore_color.rgb = _rgb(p["portada"])
        _forma(slide, MSO_SHAPE.OVAL, W - Inches(4.6), -Inches(1.6), Inches(6.4), Inches(6.4), p["acento"])
        _forma(slide, MSO_SHAPE.OVAL, W - Inches(3.0), H - Inches(3.0), Inches(4.4), Inches(4.4), p["acento2"])
        _forma(slide, MSO_SHAPE.RECTANGLE, Inches(0.9), Inches(4.3), Inches(1.6), Inches(0.1), p["acento2"])

        titulo = slide.shapes.title
        _colocar(titulo, Inches(0.9), Inches(1.6), Inches(7.4), Inches(2.6))
        _estilo(titulo, 44, BLANCO, negrita=True, anclaje=MSO_ANCHOR.BOTTOM)

        sub = slide.placeholders[1]
        _colocar(sub, Inches(0.9), Inches(4.6), Inches(7.4), Inches(1.6))
        _estilo(sub, 22, p["acento2"], anclaje=MSO_ANCHOR.TOP)

    def _contenido(self, slide, p, W, H, numero):
        slide.background.fill.fore_color.rgb = _rgb(p["fondo"])
        _forma(slide, MSO_SHAPE.OVAL, W - Inches(2.6), -Inches(1.6), Inches(4.0), Inches(4.0), p["tinte"])
        _forma(slide, MSO_SHAPE.RECTANGLE, 0, 0, Inches(0.3), H, p["acento"])
        _forma(slide, MSO_SHAPE.RECTANGLE, Inches(0.9), Inches(1.7), Inches(1.4), Inches(0.08), p["acento"])

        titulo = slide.shapes.title
        _colocar(titulo, Inches(0.9), Inches(0.45), W - Inches(3.0), Inches(1.2))
        _estilo(titulo, 34, p["titulo"], negrita=True)

        cuerpo = slide.placeholders[1]
        _colocar(cuerpo, Inches(0.9), Inches(2.0), W - Inches(1.8), H - Inches(3.0))
        _estilo(cuerpo, 24, p["texto"], anclaje=MSO_ANCHOR.TOP, espacio=12)

        circulo = _forma(slide, MSO_SHAPE.OVAL, W - Inches(1.1), H - Inches(0.95),
                         Inches(0.6), Inches(0.6), p["acento"], atras=False)
        tf = circulo.text_frame
        tf.text = str(numero)
        par = tf.paragraphs[0]
        par.alignment = PP_ALIGN.CENTER
        run = par.runs[0]
        run.font.size = Pt(14)
        run.font.bold = True
        run.font.color.rgb = _rgb(BLANCO)


# ----------------------------------------------------------------------
# El Jefe: coordina, valida, manda a corregir y, si hace falta, corrige él
# ----------------------------------------------------------------------
class ErrorPipeline(Exception):
    pass


class Jefe:
    nombre = "Jefe"
    rol = "Supervisa y corrige a los agentes"

    def __init__(self):
        self.log = []
        self.t_inicio = time.time()

    # ---- bitácora ----
    def _anotar(self, agente, rol, intento, estado, segundos, res=None, revisiones=None, notas=""):
        res = res or {}
        self.log.append(
            {
                "paso": len(self.log) + 1,
                "agente": agente,
                "rol": rol,
                "intento": intento,
                "estado": estado,
                "recibio": res.get("recibio", "-"),
                "entrego": res.get("entrego", "-"),
                "razonamiento": res.get("razonamiento", ""),
                "detalle": res.get("detalle", ""),
                "notas": notas or res.get("notas", ""),
                "revisiones": revisiones or [],
                "segundos": round(segundos, 1),
            }
        )

    # ---- reglas de supervisión: devuelven lista de revisiones ----
    def revisar(self, nombre, s):
        c = []
        if nombre == "Lector":
            _chk(c, "Texto extraído suficiente (al menos 200 caracteres)",
                 len(s["texto"].strip()) >= 200, "Texto casi vacío (¿PDF escaneado como imagen?)")
        elif nombre == "Recopilador":
            secciones = s.get("secciones", [])
            _chk(c, "Trae título y subtítulo del documento",
                 s.get("titulo_documento") and s.get("subtitulo"), "Falta 'titulo_documento' o 'subtitulo'")
            _chk(c, "Título de máximo 12 palabras",
                 num_palabras(s.get("titulo_documento", "")) <= 12, "El título del documento es muy largo")
            _chk(c, "Trae tema general", s.get("tema_general"), "Falta 'tema_general'")
            _chk(c, "Explica su razonamiento", s.get("razonamiento"), "Falta el campo 'razonamiento'")
            _chk(c, "Al menos 3 temas", len(secciones) >= 3,
                 f"Hay {len(secciones)} temas; extrae al menos 3")
            sin_puntos = [x.get("tema", "?") for x in secciones if not x.get("puntos")]
            _chk(c, "Todos los temas tienen puntos", not sin_puntos,
                 f"Temas sin puntos: {sin_puntos}")
        elif nombre == "Seccionador":
            c += self._checks_plan(s["slides"], s.get("razonamiento"))
        elif nombre == "Generador PPTX":
            existe = os.path.exists(s["ruta"])
            _chk(c, "El archivo .pptx existe", existe, "No se creó el archivo")
            if existe:
                prs = Presentation(s["ruta"])
                _chk(c, "El número de diapositivas coincide con el plan (portada + contenido)",
                     len(prs.slides) == s["n_slides"], "El número de diapositivas no coincide con el plan")
                portada = prs.slides[0].shapes.title.text if len(prs.slides) else ""
                _chk(c, "La primera diapositiva es la portada con el título",
                     portada == s["titulo"], "La portada no tiene el título esperado")
        elif nombre == "Decorador":
            prs = Presentation(s["ruta"])
            _chk(c, "No cambió el número de diapositivas", len(prs.slides) == s["n_slides"],
                 "Cambió el número de diapositivas")
            sin_fondo = [i for i, sl in enumerate(prs.slides, 1) if sl.background.fill.type != MSO_FILL.SOLID]
            _chk(c, "Todas las diapositivas tienen fondo aplicado", not sin_fondo,
                 f"Sin fondo en las diapositivas: {sin_fondo}")
            sin_formas = [i for i, sl in enumerate(prs.slides, 1) if len(sl.shapes) <= len(sl.placeholders)]
            _chk(c, "Todas las diapositivas tienen elementos decorativos", not sin_formas,
                 f"Sin decoración en las diapositivas: {sin_formas}")
        return c

    def _checks_plan(self, slides, razonamiento):
        c = []
        _chk(c, "Entre 5 y 10 diapositivas de contenido", 5 <= len(slides) <= 10,
             f"Debe haber entre 5 y 10 diapositivas (hay {len(slides)})")
        malos = [i for i, s in enumerate(slides, 1) if not s.get("title") or num_palabras(s["title"]) > 8]
        _chk(c, "Títulos de máximo 8 palabras", not malos,
             f"Títulos vacíos o de más de 8 palabras en las diapositivas {malos}")
        malos = [i for i, s in enumerate(slides, 1) if not 2 <= len(s.get("bullets", [])) <= 5]
        _chk(c, "Entre 2 y 5 bullets por diapositiva", not malos,
             f"Cantidad de bullets fuera de rango en las diapositivas {malos}")
        largos = [i for i, s in enumerate(slides, 1)
                  if any(num_palabras(b) > 15 for b in s.get("bullets", []))]
        _chk(c, "Bullets de máximo 15 palabras", not largos,
             f"Bullets de más de 15 palabras en las diapositivas {largos}")
        vistos, repetidos = set(), []
        for i, s in enumerate(slides, 1):
            for b in s.get("bullets", []):
                clave = str(b).strip().lower()
                if clave in vistos:
                    repetidos.append(i)
                vistos.add(clave)
        _chk(c, "Sin bullets repetidos", not repetidos,
             f"Bullets repetidos en las diapositivas {sorted(set(repetidos))}")
        sin_origen = [i for i, s in enumerate(slides, 1) if not s.get("origen")]
        _chk(c, "Cada diapositiva indica de qué tema sale", not sin_origen,
             f"Falta 'origen' en las diapositivas {sin_origen}")
        _chk(c, "Explica su razonamiento", razonamiento, "Falta el campo 'razonamiento'")
        return c

    # ---- corrección directa del Jefe (cuando el agente ya no puede) ----
    def _corregir_plan(self, salida):
        cambios, vistos, nuevos = [], set(), []
        slides = salida["slides"]
        if len(slides) > 10:
            cambios.append("recortó el plan a 10 diapositivas")
        for i, s in enumerate(slides[:10], 1):
            t = " ".join(str(s.get("title", "Sin título")).split()[:8])
            bullets = []
            for b in s.get("bullets", []):
                palabras = str(b).split()
                if len(palabras) > 15:
                    b = " ".join(palabras[:15]) + "…"
                    cambios.append(f"recortó un bullet de la diapositiva {i}")
                clave = b.strip().lower()
                if clave in vistos:
                    cambios.append(f"quitó un bullet repetido en la diapositiva {i}")
                    continue
                vistos.add(clave)
                bullets.append(b)
            if len(bullets) > 5:
                bullets = bullets[:5]
                cambios.append(f"dejó solo 5 bullets en la diapositiva {i}")
            if not s.get("origen"):
                cambios.append(f"marcó el origen faltante de la diapositiva {i}")
            nuevos.append({"title": t, "bullets": bullets, "origen": s.get("origen") or "(no indicado)"})
        razon = salida.get("razonamiento")
        if not razon:
            cambios.append("completó el razonamiento faltante")
        return {**salida, "slides": nuevos, "razonamiento": razon or "(el agente no explicó su decisión)"}, cambios

    # ---- ciclo de supervisión de un agente ----
    def _supervisar(self, agente, entrada):
        feedback, res = None, None
        for intento in range(1, MAX_INTENTOS + 1):
            t0 = time.time()
            res, revisiones = None, []
            try:
                res = agente.ejecutar(entrada, feedback)
                revisiones = self.revisar(agente.nombre, res["salida"])
            except Exception as e:
                _chk(revisiones, "La ejecución termina sin errores", False, f"{type(e).__name__}: {e}")
            dt = time.time() - t0
            problemas = [r["detalle"] for r in revisiones if not r["ok"]]

            if not problemas:
                self._anotar(agente.nombre, agente.rol, intento, "✅ Aprobado", dt, res, revisiones)
                return res["salida"]

            self._anotar(agente.nombre, agente.rol, intento, "❌ Rechazado por el Jefe", dt, res, revisiones,
                         notas="Motivos: " + "; ".join(problemas))
            feedback = "- " + "\n- ".join(problemas)

        # Se acabaron los intentos: el Jefe interviene si puede
        if agente.nombre == "Seccionador" and res:
            nuevo, cambios = self._corregir_plan(res["salida"])
            revisiones = self._checks_plan(nuevo["slides"], nuevo["razonamiento"])
            if all(r["ok"] for r in revisiones):
                self._anotar(
                    self.nombre, self.rol, "-", "🛠️ Corrigió él mismo", 0,
                    {
                        "recibio": f"Plan rechazado {MAX_INTENTOS} veces",
                        "entrego": f"Plan de {len(nuevo['slides'])} diapositivas corregido",
                        "razonamiento": "El Seccionador no logró cumplir las reglas tras varios intentos, "
                                        "así que apliqué los ajustes mínimos necesarios yo mismo.",
                        "detalle": "\n".join(sorted(set(cambios))) or "ajustes menores",
                    },
                    revisiones,
                )
                return nuevo

        self._anotar(self.nombre, self.rol, "-", "🛑 Detuvo el proceso", 0,
                     {"recibio": f"{agente.nombre} falló {MAX_INTENTOS} veces",
                      "razonamiento": "No pude corregirlo yo mismo, así que detengo el proceso para no "
                                      "entregar una presentación incorrecta."},
                     notas=feedback or "")
        raise ErrorPipeline(f"{agente.nombre} no logró entregar un resultado válido")

    # ---- el flujo completo ----
    def coordinar(self, pdf_path, ruta_pptx):
        lectura = self._supervisar(Lector(), pdf_path)
        resumen = self._supervisar(Recopilador(), lectura)
        plan = self._supervisar(Seccionador(), resumen)
        base = self._supervisar(GeneradorPPTX(ruta_pptx), {**plan})
        final = self._supervisar(Decorador(), base)
        return final["ruta"]

    # ---- reporte para ti (el "jefe de jefes") ----
    def reporte_html(self, exito):
        e = escape
        rechazos = sum(1 for x in self.log if "Rechazado" in x["estado"])
        correcciones = sum(1 for x in self.log if "Corrigió" in x["estado"])
        total = round(time.time() - self.t_inicio, 1)
        caja = "border:1px solid rgba(128,128,128,.4);border-radius:8px;margin:8px 0;padding:8px 12px;"
        celda = "border:1px solid rgba(128,128,128,.35);padding:6px 8px;text-align:left;"

        h = ["<div>"]
        h.append("<h3>Informe final del Jefe</h3><ul>")
        h.append(f"<li>Resultado: {'✅ presentación generada' if exito else '🛑 el proceso se detuvo'}</li>")
        h.append(f"<li>Rechazos a agentes: {rechazos}</li>")
        h.append(f"<li>Correcciones hechas por el Jefe: {correcciones}</li>")
        h.append(f"<li>Tiempo total: {total} s</li></ul>")

        h.append("<h3>Resumen de la corrida</h3>")
        h.append('<table style="border-collapse:collapse;width:100%;"><tr>'
                 + "".join(f'<th style="{celda}">{t}</th>' for t in ["#", "Agente", "Intento", "Estado", "Seg"])
                 + "</tr>")
        for x in self.log:
            h.append("<tr>" + "".join(
                f'<td style="{celda}">{e(str(v))}</td>'
                for v in [x["paso"], x["agente"], x["intento"], x["estado"], x["segundos"]]
            ) + "</tr>")
        h.append("</table>")

        h.append("<h3>Detalle por agente (clic para abrir)</h3>")
        for x in self.log:
            abierto = " open" if any(k in x["estado"] for k in ("Rechazado", "Corrigió", "Detuvo")) else ""
            h.append(f'<details{abierto} style="{caja}"><summary style="cursor:pointer;font-weight:600;">'
                     f'#{x["paso"]} {e(x["agente"])} — {e(x["estado"])} '
                     f'(intento {e(str(x["intento"]))}, {x["segundos"]} s)</summary>')
            h.append(f'<p><b>Qué hace:</b> {e(x["rol"])}</p>')
            h.append(f'<p><b>Recibió:</b> {e(x["recibio"])}</p>')
            h.append(f'<p><b>Entregó:</b> {e(x["entrego"])}</p>')
            if x["razonamiento"]:
                h.append(f'<p><b>Razonamiento:</b> {e(x["razonamiento"])}</p>')
            if x["notas"]:
                h.append(f'<p><b>Notas:</b> {e(x["notas"])}</p>')
            if x["detalle"]:
                h.append('<p><b>Detalle del trabajo:</b></p>'
                         f'<pre style="white-space:pre-wrap;margin:0;padding:8px;'
                         f'border:1px solid rgba(128,128,128,.35);border-radius:6px;">{e(x["detalle"])}</pre>')
            if x["revisiones"]:
                h.append("<p><b>Revisiones del Jefe:</b></p><ul>")
                for r in x["revisiones"]:
                    marca = "✅" if r["ok"] else "❌"
                    extra = f" — {e(r['detalle'])}" if not r["ok"] else ""
                    h.append(f"<li>{marca} {e(r['regla'])}{extra}</li>")
                h.append("</ul>")
            h.append("</details>")
        h.append("</div>")
        return "".join(h)


# ----------------------------------------------------------------------
# Interfaz
# ----------------------------------------------------------------------
def process_pdf(pdf_file):
    if pdf_file is None:
        raise gr.Error("Por favor sube un archivo PDF.")

    carpeta = tempfile.mkdtemp()
    ruta_pptx = os.path.join(carpeta, "presentacion_generada.pptx")
    ruta_log = os.path.join(carpeta, "bitacora.json")

    jefe = Jefe()
    try:
        jefe.coordinar(pdf_file, ruta_pptx)
        exito = True
    except ErrorPipeline:
        exito = False

    with open(ruta_log, "w", encoding="utf-8") as f:
        json.dump(jefe.log, f, ensure_ascii=False, indent=2)

    return (ruta_pptx if exito else None), jefe.reporte_html(exito), ruta_log


demo = gr.Interface(
    fn=process_pdf,
    inputs=gr.File(label="Sube tu PDF", file_types=[".pdf"]),
    outputs=[
        gr.File(label="Descarga tu presentación (.pptx)"),
        gr.HTML(label="Informe del Jefe y bitácora de cada agente"),
        gr.File(label="Bitácora completa (.json)"),
    ],
    title="Sistema multiagente: PDF a Diapositivas",
    description=(
        "Sube un PDF. Cinco agentes trabajan en cadena (Lector, Recopilador, "
        "Seccionador, Generador, Decorador) supervisados por un Jefe que valida, "
        "corrige y te entrega un informe de quién hizo qué y por qué."
    ),
    flagging_mode="never",
)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    demo.launch(server_name="0.0.0.0", server_port=port)
