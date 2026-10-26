import os
import json
import math
import re
import time
import tempfile
from html import escape, unescape

import requests
import gradio as gr
import pdfplumber
import google.generativeai as genai
from PIL import Image
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.enum.dml import MSO_FILL
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
import io

# La API key se lee de la variable de entorno configurada en Render
genai.configure(api_key=os.environ["GEMINI_API_KEY"])
MODEL_NAME = os.environ.get("GEMINI_MODEL", "gemini-3.5-flash-lite")

MAX_INTENTOS = 3      # cuántas veces el Jefe deja reintentar a un agente
TOPE_IMAGENES = 12    # tope absoluto de imágenes por presentación (el cupo real escala con las diapositivas)
TITULO_FUENTES = "Fuentes y créditos"
# Wikimedia pide identificar a quien usa su API. Opcional: define CONTACTO en Render.
USER_AGENT = "PDF2SlidesAgents/1.0 (proyecto de servicio social; uso educativo) " + os.environ.get("CONTACTO", "")


# ----------------------------------------------------------------------
# Utilidades
# ----------------------------------------------------------------------
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


# ---- estimación de si un texto cabe en su caja (sin renderizar) ----
def _altura_in(textos, ancho_in, tam, espacio_pt, sangria_in):
    cpl = max(1.0, (ancho_in - sangria_in) * 144 / tam / 1.1)  # caracteres por línea (aprox., conservador)
    lineas = sum(max(1, math.ceil(len(t) / cpl)) for t in textos)
    return lineas * tam * 1.2 / 72 + len(textos) * espacio_pt / 72


def _mejor_tam(textos, ancho_emu, alto_emu, tams, espacio_pt=0, sangria_in=0.4):
    ancho_in = ancho_emu / 914400 - 0.2
    alto_in = alto_emu / 914400 - 0.1
    for t in tams:
        if _altura_in(textos, ancho_in, t, espacio_pt, sangria_in) <= alto_in * 0.95:
            return t
    return tams[-1]


def _cabe(shape, sangria_in=0.4):
    pars = shape.text_frame.paragraphs
    textos = [p.text for p in pars]
    tam = next((r.font.size.pt for p in pars for r in p.runs if r.font.size), 18)
    esp = pars[0].space_after.pt if pars and pars[0].space_after is not None else 0
    return _altura_in(textos, shape.width / 914400 - 0.2, tam, esp, sangria_in) <= shape.height / 914400 - 0.1


# ---- color y contraste (para validar el diseño propuesto por la IA) ----
def _hex_a_rgb(h):
    h = str(h).strip().lstrip("#")
    if not re.fullmatch(r"[0-9a-fA-F]{6}", h):
        raise ValueError(f"color inválido: {h!r}")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def _a_hex(t):
    return "#%02X%02X%02X" % tuple(t)


def _luminancia(rgb):
    def canal(v):
        v = v / 255
        return v / 12.92 if v <= 0.03928 else ((v + 0.055) / 1.055) ** 2.4
    r, g, b = (canal(x) for x in rgb)
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def _contraste(a, b):
    l1, l2 = sorted((_luminancia(a), _luminancia(b)), reverse=True)
    return (l1 + 0.05) / (l2 + 0.05)


def _problemas_contraste(pal):
    blanco = (0xFF, 0xFF, 0xFF)
    reglas = [("texto", pal["texto"], pal["fondo"], 4.5), ("título", pal["titulo"], pal["fondo"], 4.5),
              ("texto sobre tinte", pal["texto"], pal["tinte"], 4.5),
              ("título sobre tinte", pal["titulo"], pal["tinte"], 4.5),
              ("blanco sobre portada", blanco, pal["portada"], 4.5),
              ("acento2 sobre portada", pal["acento2"], pal["portada"], 3.0),
              ("blanco sobre acento", blanco, pal["acento"], 3.0)]
    out = []
    for nombre, a, b, minimo in reglas:
        r = _contraste(a, b)
        if r < minimo:
            out.append(f"contraste {nombre} = {r:.1f} (mínimo {minimo})")
    return out


# ----------------------------------------------------------------------
# Acceso a Wikimedia Commons (imágenes con licencia libre y autor)
# ----------------------------------------------------------------------
def _licencia_libre(lic):
    return bool(re.match(r"^(cc[\s-]|cc0|public domain|pd[\s-]|pd$|gfdl)", lic.strip().lower()))


def buscar_commons(consulta, limite=8):
    """Busca imágenes en Wikimedia Commons y devuelve solo las que traen autor y licencia libre."""
    r = requests.get(
        "https://commons.wikimedia.org/w/api.php",
        params={
            "action": "query", "generator": "search", "gsrsearch": consulta,
            "gsrnamespace": 6, "gsrlimit": limite, "prop": "imageinfo",
            "iiprop": "url|extmetadata|mime|size", "iiurlwidth": 800, "format": "json",
        },
        headers={"User-Agent": USER_AGENT}, timeout=20,
    )
    r.raise_for_status()
    paginas = r.json().get("query", {}).get("pages", {}).values()
    resultado = []
    for pg in sorted(paginas, key=lambda x: x.get("index", 0)):
        info = (pg.get("imageinfo") or [{}])[0]
        meta = info.get("extmetadata", {})
        if info.get("mime") not in ("image/jpeg", "image/png"):
            continue
        if info.get("width", 0) < 600 or info.get("height", 0) < 400:
            continue
        licencia = _limpiar_html(meta.get("LicenseShortName", {}).get("value", ""))
        autor = _limpiar_html(meta.get("Artist", {}).get("value", ""))
        if not licencia or not autor or not _licencia_libre(licencia):
            continue
        resultado.append({
            "titulo": pg.get("title", "").replace("File:", ""),
            "thumb": info.get("thumburl") or info.get("url"),
            "pagina": info.get("descriptionurl", ""),
            "autor": autor,
            "licencia": licencia,
            "descripcion": _limpiar_html(meta.get("ImageDescription", {}).get("value", ""))[:200],
        })
    return resultado


def descargar_imagen(url):
    r = requests.get(url, headers={"User-Agent": USER_AGENT}, timeout=25)
    r.raise_for_status()
    mime = r.headers.get("Content-Type", "").split(";")[0]
    if mime not in ("image/jpeg", "image/png"):
        raise ValueError(f"tipo de imagen no soportado: {mime}")
    if len(r.content) > 6_000_000:
        raise ValueError("imagen demasiado pesada")
    return r.content, mime


# ----------------------------------------------------------------------
# Agentes (cada uno hace UNA cosa y devuelve qué recibió/entregó/razonó)
# ----------------------------------------------------------------------
class Agente:
    nombre = "Agente"
    rol = ""

    def ejecutar(self, entrada, feedback=None):
        """Debe devolver {"salida", "recibio", "entrego", "razonamiento", "detalle", "notas"}"""
        raise NotImplementedError


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


SYS_ILUSTRADOR_PLAN = """
Eres un editor gráfico que planea las imágenes de una presentación. Eliges solo las
diapositivas donde una imagen real aporta valor, y escribes búsquedas precisas.
Responde ÚNICAMENTE con JSON válido.
"""

SYS_ILUSTRADOR = """
Eres un editor gráfico muy estricto. Eliges imágenes para diapositivas, pero solo
si muestran de forma clara y directa el tema. Si hay duda, no eliges ninguna.
Responde ÚNICAMENTE con JSON válido.
"""


class Ilustrador(Agente):
    nombre = "Ilustrador"
    rol = ("Decide qué diapositivas llevan imagen y qué buscar; la trae de Wikimedia Commons (licencia libre), "
           "verifica con IA (viendo la imagen) que ilustre el tema y guarda autor y licencia")

    def __init__(self, carpeta):
        self.carpeta = carpeta

    def _planear(self, slides, cap, feedback):
        resumen = [{"diapositiva": i, "titulo": sl.get("title", ""),
                    "texto": (sl.get("parrafo") or " ".join(sl.get("bullets", [])))[:160]}
                   for i, sl in enumerate(slides, 1)]
        prompt = (
            f"Estas son las diapositivas de una presentación ({len(slides)} en total). Elige hasta {cap} "
            "diapositivas que se beneficien de una imagen concreta y real (un objeto, lugar, escena, "
            "alimento, diagrama...), no de conceptos abstractos. No pongas imagen a todas: elige solo "
            "donde la imagen ilustre claramente el contenido. Para cada una escribe una búsqueda corta "
            "EN INGLÉS (2 a 5 palabras) para Wikimedia Commons.\n"
            'Responde {"razonamiento": "<por qué elegiste esas diapositivas>", '
            '"imagenes": [{"diapositiva": N, "consulta": "..."}]}\n\n'
            "DIAPOSITIVAS:\n" + json.dumps(resumen, ensure_ascii=False)
        )
        d = parsear_json(llamar_llm(SYS_ILUSTRADOR_PLAN, con_feedback(prompt, feedback)))
        plan = [{"diapositiva": x.get("diapositiva"), "consulta": str(x.get("consulta", "")).strip()}
                for x in d.get("imagenes", []) if isinstance(x, dict)]
        return str(d.get("razonamiento", "")), plan

    def _elegir(self, slide, candidatas):
        texto = slide.get("parrafo") or " ".join(slide.get("bullets", []))
        partes = [
            f"Diapositiva: «{slide.get('title', '')}». Contenido: {texto}\n"
            "Estas imágenes vienen de Wikimedia Commons. Elige la que muestre de forma clara y "
            "directa el tema de la diapositiva. Descarta las que solo se relacionen vagamente, "
            "logotipos, capturas con texto ilegible o imágenes que no se entiendan. Si NINGUNA "
            'cumple, devuelve null. Responde {"elegida": <número de imagen o null>, '
            '"razon": "<una frase: qué se ve y por qué se relaciona con el tema>"}'
        ]
        for k, c in enumerate(candidatas):
            partes.append(f"Imagen {k} (archivo: {c['titulo']}):")
            partes.append({"mime_type": c["mime"], "data": c["bytes"]})
        d = parsear_json(llamar_llm(SYS_ILUSTRADOR, partes))
        k = d.get("elegida")
        if isinstance(k, int) and 0 <= k < len(candidatas):
            return candidatas[k], d.get("razon", "")
        return None, d.get("razon", "ninguna imagen ilustraba el tema con claridad")

    def ejecutar(self, entrada, feedback=None):
        slides = [dict(sl) for sl in entrada["slides"]]
        for sl in slides:
            sl.pop("imagen", None)
        cap = _max_imagenes(len(slides))
        razon_plan, plan = self._planear(slides, cap, feedback)

        lineas = [f"Cupo de imágenes: hasta {cap} (para {len(slides)} diapositivas)",
                  f"Por qué eligió esas diapositivas: {razon_plan}",
                  "Plan: " + (" · ".join(f"diap. {p['diapositiva']} → «{p['consulta']}»" for p in plan) or "(vacío)"),
                  ""]
        validos = [p for p in plan if isinstance(p["diapositiva"], int)
                   and 1 <= p["diapositiva"] <= len(slides) and p["consulta"]][:cap]
        errores, elegidas = [], 0

        for p in validos:
            i = p["diapositiva"] - 1
            sl = slides[i]
            etiqueta = f"Diap. {i + 1} «{sl.get('title', '')}» (búsqueda: {p['consulta']})"
            try:
                cands = buscar_commons(p["consulta"])
                bajadas = []
                for c in cands:
                    if len(bajadas) >= 3:
                        break
                    try:
                        datos, mime = descargar_imagen(c["thumb"])
                        bajadas.append({**c, "bytes": datos, "mime": mime})
                    except Exception:
                        continue
                if not bajadas:
                    lineas.append(f"{etiqueta}: sin candidatas utilizables ({len(cands)} con licencia y autor)")
                    continue
                elegida, razon = self._elegir(sl, bajadas)
                if not elegida:
                    lineas.append(f"{etiqueta}: ninguna sirvió → {razon}")
                    continue
                ext = ".png" if elegida["mime"] == "image/png" else ".jpg"
                ruta = os.path.join(self.carpeta, f"imagen_{i + 1}{ext}")
                with open(ruta, "wb") as f:
                    f.write(elegida["bytes"])
                ancho, alto = Image.open(io.BytesIO(elegida["bytes"])).size
                sl["imagen"] = {
                    "ruta": ruta, "titulo": elegida["titulo"], "autor": elegida["autor"],
                    "licencia": elegida["licencia"], "pagina": elegida["pagina"],
                    "razon": razon, "ancho": ancho, "alto": alto,
                }
                elegidas += 1
                lineas.append(f"{etiqueta}: ✔ {elegida['titulo']} — {elegida['autor']}, "
                              f"{elegida['licencia']} → {razon}")
            except Exception as e:
                errores.append(str(e))
                lineas.append(f"{etiqueta}: error ({type(e).__name__}: {e})")

        if validos and len(errores) == len(validos):
            raise RuntimeError("No se pudo consultar Wikimedia Commons: " + errores[0])

        return {
            "salida": {**entrada, "slides": slides, "plan_imagenes": plan, "max_imagenes": cap,
                       "razonamiento_imagenes": razon_plan},
            "recibio": f"Plan de {len(slides)} diapositivas (sin imágenes)",
            "entrego": f"{elegidas} imágenes aceptadas de {len(validos)} búsquedas",
            "razonamiento": (
                "Primero leí el plan y elegí las diapositivas donde una imagen real aporta valor, con su "
                "búsqueda. Luego consulté Wikimedia Commons (solo imágenes con autor y licencia libre), "
                "descargué hasta 3 candidatas por búsqueda y le pedí a la IA que las viera y eligiera solo una "
                "que muestre claramente el tema; si ninguna lo hacía, esa diapositiva queda sin imagen. "
                f"Selección: {razon_plan}"
            ),
            "detalle": "\n".join(lineas),
            "notas": "",
        }


def _sin_vineta(par):
    """Quita la viñeta y la sangría de un párrafo (para texto corrido)."""
    pPr = par._p.get_or_add_pPr()
    pPr.set("marL", "0")
    pPr.set("indent", "0")
    if pPr.find(qn("a:buNone")) is None:
        pPr.append(pPr.makeelement(qn("a:buNone"), {}))


def _escribir(tf, items):
    """items = [(texto, con_vineta)]"""
    tf.clear()
    for k, (texto, vineta) in enumerate(items):
        par = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
        par.text = texto
        if not vineta:
            _sin_vineta(par)


def _corto(texto, n):
    return texto if len(texto) <= n else texto[: n - 1].rstrip() + "…"


class GeneradorPPTX(Agente):
    nombre = "Generador PPTX"
    rol = "Construye el archivo .pptx (portada, contenido, imágenes y fuentes) a partir del plan (sin IA)"

    def __init__(self, ruta_salida):
        self.ruta_salida = ruta_salida

    def ejecutar(self, entrada, feedback=None):
        prs = Presentation()
        prs.slide_width = Inches(13.333)  # formato 16:9
        prs.slide_height = Inches(7.5)

        portada = prs.slides.add_slide(prs.slide_layouts[0])
        portada.shapes.title.text = entrada["titulo"]
        portada.placeholders[1].text = entrada["subtitulo"] or entrada.get("tema_general", "")

        detalle = [f"Diapositiva 1 (portada): {entrada['titulo']}"]
        creditos, n_imgs = [], 0
        for i, info in enumerate(entrada["slides"], 1):
            slide = prs.slides.add_slide(prs.slide_layouts[1])
            slide.shapes.title.text = info.get("title", "Sin titulo")
            tipo = info.get("tipo", "bullets")
            items = []
            if tipo in ("parrafo", "mixto") and info.get("parrafo"):
                items.append((info["parrafo"], False))
            if tipo in ("bullets", "mixto"):
                items += [(b, True) for b in info.get("bullets", [])]
            _escribir(slide.placeholders[1].text_frame, items)

            marca = ""
            img = info.get("imagen")
            if img:
                slide.shapes.add_picture(img["ruta"], Inches(8), Inches(2), width=Inches(4))
                cap = slide.shapes.add_textbox(Inches(8), Inches(5.5), Inches(4), Inches(0.5))
                cap.name = "Credito imagen"
                cap.text_frame.text = (f"Imagen: {_corto(img['autor'], 45)} · {img['licencia']} · "
                                       "Wikimedia Commons")
                n_imgs += 1
                creditos.append((i, img))
                marca = " + imagen"
            detalle.append(f"Diapositiva {i + 1} [{tipo}]: {info.get('title', '')}{marca}")

        # Diapositiva final: fuentes de información y créditos de imágenes
        fuentes = prs.slides.add_slide(prs.slide_layouts[1])
        fuentes.shapes.title.text = TITULO_FUENTES
        items = [("Fuentes de información", False),
                 (f"Documento base: {entrada['titulo']} ({entrada.get('documento', 'PDF')})", False)]
        refs = entrada.get("referencias", [])[:6]
        if refs:
            items += [(f"Citado en el documento: {_corto(r, 150)}", False) for r in refs]
        else:
            items.append(("El documento base no cita referencias bibliográficas.", False))
        items.append(("Créditos de imágenes", False))
        if creditos:
            for n, img in creditos:
                titulo_archivo = _corto(img["titulo"], 60)
                items.append((_corto(f"Diap. {n}: «{titulo_archivo}» — {img['autor']}, {img['licencia']}. "
                                     f"{img['pagina']}", 220), False))
        else:
            items.append(("Esta presentación no incluye imágenes de terceros.", False))
        _escribir(fuentes.placeholders[1].text_frame, items)
        detalle.append(f"Diapositiva {len(entrada['slides']) + 2} (fuentes): {len(refs)} referencias, "
                       f"{len(creditos)} créditos de imagen")

        prs.save(self.ruta_salida)
        total = len(entrada["slides"]) + 2
        return {
            "salida": {
                "ruta": self.ruta_salida, "n_slides": total, "n_imagenes": n_imgs,
                "titulo": entrada["titulo"], "tema": entrada.get("tema_general", ""),
            },
            "recibio": f"Plan de {len(entrada['slides'])} diapositivas, {n_imgs} imágenes y referencias",
            "entrego": f"Archivo {os.path.basename(self.ruta_salida)} con {total} diapositivas "
                       "(portada + contenido + fuentes)",
            "razonamiento": (
                "Creé la portada, luego una diapositiva por elemento del plan: texto corrido sin viñetas "
                "en las de tipo 'parrafo', viñetas en las de tipo 'bullets', y ambos en 'mixto'. "
                "Inserté cada imagen con su crédito y cerré con la diapositiva de fuentes, armada solo con "
                "datos verificados (documento, referencias presentes en el PDF y créditos de las imágenes). "
                "Aún sin estilo: eso lo hace el Decorador."
            ),
            "detalle": "\n".join(detalle),
            "notas": "",
        }


PALETAS = {  # diseños de respaldo y ejemplo de referencia para la IA
    "oceano": {"portada": (0x0B, 0x3C, 0x5D), "fondo": (0xF7, 0xFA, 0xFC), "titulo": (0x0B, 0x3C, 0x5D),
               "texto": (0x2B, 0x35, 0x40), "acento": (0x32, 0x82, 0xB8), "acento2": (0x7F, 0xC4, 0xE8),
               "tinte": (0xDC, 0xEC, 0xF6)},
}
BLANCO = (0xFF, 0xFF, 0xFF)
ENCABEZADOS_FUENTES = ("Fuentes de información", "Créditos de imágenes")
CLAVES_COLOR = ("portada", "fondo", "titulo", "texto", "acento", "acento2", "tinte")
ESTILOS_FORMAS = ("circulos", "bloques", "minimal")
BARRAS = ("izquierda", "superior")
FUENTES_OK = ("Calibri", "Arial", "Georgia", "Trebuchet MS", "Cambria")

DISENO_RESPALDO = {
    "paleta": PALETAS["oceano"], "estilo_formas": "circulos", "barra": "izquierda",
    "alternar_fondos": False, "fuente_titulo": "Calibri", "fuente_cuerpo": "Calibri",
    "motivo": "diseño de respaldo 'oceano'",
}

SYS_DECORADOR = """
Eres un diseñador de presentaciones. Propones un diseño completo y coherente con el tema
del documento. Responde ÚNICAMENTE con JSON válido.
"""


def _validar_diseno(d):
    """Valida la propuesta de la IA. Devuelve (diseno, problemas)."""
    prob, pal = [], {}
    paleta = d.get("paleta") if isinstance(d.get("paleta"), dict) else {}
    for k in CLAVES_COLOR:
        try:
            pal[k] = _hex_a_rgb(paleta.get(k, ""))
        except Exception:
            prob.append(f"color '{k}' inválido o ausente (usa #RRGGBB)")
    estilo, barra = d.get("estilo_formas"), d.get("barra")
    ft, fc = d.get("fuente_titulo"), d.get("fuente_cuerpo")
    if estilo not in ESTILOS_FORMAS:
        prob.append(f"'estilo_formas' debe ser uno de {list(ESTILOS_FORMAS)}")
    if barra not in BARRAS:
        prob.append(f"'barra' debe ser uno de {list(BARRAS)}")
    if ft not in FUENTES_OK or fc not in FUENTES_OK:
        prob.append(f"las fuentes deben ser de {list(FUENTES_OK)}")
    if not prob:
        prob += _problemas_contraste(pal)
    diseno = {"paleta": pal, "estilo_formas": estilo, "barra": barra,
              "alternar_fondos": bool(d.get("alternar_fondos")), "fuente_titulo": ft,
              "fuente_cuerpo": fc, "motivo": str(d.get("motivo", ""))}
    return diseno, prob


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
            anclaje=MSO_ANCHOR.MIDDLE, espacio=0, cursiva=False, fuente=None):
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
            run.font.italic = cursiva
            run.font.color.rgb = _rgb(color)
            if fuente:
                run.font.name = fuente


def _imagenes(slide):
    return [sh for sh in slide.shapes if sh.shape_type == MSO_SHAPE_TYPE.PICTURE]


class Decorador(Agente):
    nombre = "Decorador"
    rol = "La IA propone el diseño (colores, formas, tipografía) y el código lo aplica y lo acomoda"

    def ejecutar(self, entrada, feedback=None):
        ruta, tema = entrada["ruta"], entrada["tema"]
        ejemplo = {"paleta": {k: _a_hex(v) for k, v in PALETAS["oceano"].items()},
                   "estilo_formas": "circulos", "barra": "izquierda", "alternar_fondos": False,
                   "fuente_titulo": "Calibri", "fuente_cuerpo": "Calibri", "motivo": "una frase corta"}
        prompt = (
            f"Tema del documento: {tema}\nTítulo: {entrada.get('titulo', '')}\n"
            "Propón el diseño completo de la presentación (16:9). Reglas:\n"
            "- 'portada': color OSCURO (llevará texto blanco encima).\n"
            "- 'fondo' y 'tinte': colores CLAROS.\n"
            "- 'titulo' y 'texto': colores oscuros con buen contraste sobre 'fondo' y 'tinte'.\n"
            "- 'acento': color vivo de brillo medio (llevará texto blanco encima).\n"
            "- 'acento2': color más claro, visible sobre 'portada'.\n"
            "- 'estilo_formas': 'circulos' (suave), 'bloques' (geométrico, formal) o 'minimal' (sobrio).\n"
            "- 'barra': 'izquierda' o 'superior'.\n"
            "- 'alternar_fondos': true para alternar el fondo entre diapositivas y dar ritmo.\n"
            f"- 'fuente_titulo' y 'fuente_cuerpo': una de {list(FUENTES_OK)}.\n"
            "Elige colores y estilo acordes al tema. Colores en formato #RRGGBB.\n"
            "Ejemplo de formato (no lo copies, inspírate): " + json.dumps(ejemplo, ensure_ascii=False)
        )
        origen, problemas = "IA", []
        try:
            propuesta = parsear_json(llamar_llm(SYS_DECORADOR, con_feedback(prompt, feedback)))
            diseno, problemas = _validar_diseno(propuesta)
        except Exception as e:
            diseno, problemas = None, [f"{type(e).__name__}: {e}"]
        if problemas:
            origen = "respaldo"
            motivo = ("La propuesta de la IA se descartó (" + "; ".join(problemas)
                      + "); usé el diseño de respaldo 'oceano'.")
            diseno = dict(DISENO_RESPALDO)
        else:
            motivo = diseno["motivo"] or "diseño propuesto por la IA"

        prs = Presentation(ruta)
        W, H = prs.slide_width, prs.slide_height
        total = len(prs.slides)
        tamanos = []
        for i, slide in enumerate(prs.slides):
            slide.background.fill.solid()
            if i == 0:
                self._portada(slide, diseno, W, H)
            else:
                tam = self._contenido(slide, diseno, W, H, i, es_fuentes=(i == total - 1))
                tamanos.append(f"Diap. {i + 1}: texto a {tam} pt" + (" (con imagen)" if _imagenes(slide) else ""))
        prs.save(ruta)

        pal = diseno["paleta"]
        detalle = (
            f"Origen del diseño: {origen}\n"
            "Paleta: " + ", ".join(f"{k} {_a_hex(v)}" for k, v in pal.items()) + "\n"
            f"Formas: {diseno['estilo_formas']} · Barra: {diseno['barra']} · "
            f"Fondos alternados: {'sí' if diseno['alternar_fondos'] else 'no'}\n"
            f"Tipografía: títulos {diseno['fuente_titulo']}, cuerpo {diseno['fuente_cuerpo']}\n"
            "Las imágenes van a la derecha con marco de color y su crédito debajo; el tamaño de "
            "letra se ajusta para que el texto quepa; la última diapositiva usa letra pequeña.\n"
            + "\n".join(tamanos)
        )
        return {
            "salida": {"ruta": ruta, "n_slides": total, "diseno": diseno,
                       "origen_diseno": origen, "problemas_ia": problemas},
            "recibio": f"Archivo básico de {total} diapositivas y el tema del documento",
            "entrego": f"Mismo archivo con diseño aplicado (diseño {'propuesto por la IA' if origen == 'IA' else 'de respaldo'})",
            "razonamiento": motivo,
            "detalle": detalle,
            "notas": "",
        }

    def _portada(self, slide, d, W, H):
        p, estilo = d["paleta"], d["estilo_formas"]
        slide.background.fill.fore_color.rgb = _rgb(p["portada"])
        if estilo == "circulos":
            _forma(slide, MSO_SHAPE.OVAL, W - Inches(4.6), -Inches(1.6), Inches(6.4), Inches(6.4), p["acento"])
            _forma(slide, MSO_SHAPE.OVAL, W - Inches(3.0), H - Inches(3.0), Inches(4.4), Inches(4.4), p["acento2"])
        elif estilo == "bloques":
            _forma(slide, MSO_SHAPE.RECTANGLE, W - Inches(4.2), 0, Inches(4.2), H, p["acento"])
            _forma(slide, MSO_SHAPE.RECTANGLE, W - Inches(4.2), H - Inches(1.6), Inches(4.2), Inches(1.6), p["acento2"])
        else:
            _forma(slide, MSO_SHAPE.RECTANGLE, 0, H - Inches(0.5), W, Inches(0.5), p["acento"])
        _forma(slide, MSO_SHAPE.RECTANGLE, Inches(0.9), Inches(4.3), Inches(1.6), Inches(0.1), p["acento2"])

        titulo = slide.shapes.title
        _colocar(titulo, Inches(0.9), Inches(1.6), Inches(7.4), Inches(2.6))
        tam_t = _mejor_tam([titulo.text], titulo.width, titulo.height, [44, 40, 36, 32], 0, 0)
        _estilo(titulo, tam_t, BLANCO, negrita=True, anclaje=MSO_ANCHOR.BOTTOM, fuente=d["fuente_titulo"])

        sub = slide.placeholders[1]
        _colocar(sub, Inches(0.9), Inches(4.6), Inches(7.4), Inches(1.6))
        _estilo(sub, 22, p["acento2"], anclaje=MSO_ANCHOR.TOP, fuente=d["fuente_cuerpo"])

    def _contenido(self, slide, d, W, H, numero, es_fuentes=False):
        p, estilo = d["paleta"], d["estilo_formas"]
        alterna = d["alternar_fondos"] and numero % 2 == 0 and not es_fuentes
        fondo = p["tinte"] if alterna else p["fondo"]
        color_forma = p["fondo"] if alterna else p["tinte"]
        slide.background.fill.fore_color.rgb = _rgb(fondo)

        if estilo == "circulos":
            _forma(slide, MSO_SHAPE.OVAL, W - Inches(2.6), -Inches(1.6), Inches(4.0), Inches(4.0), color_forma)
        elif estilo == "bloques":
            _forma(slide, MSO_SHAPE.RECTANGLE, W - Inches(2.4), 0, Inches(2.4), Inches(1.2), color_forma)
        if d["barra"] == "izquierda":
            _forma(slide, MSO_SHAPE.RECTANGLE, 0, 0, Inches(0.3), H, p["acento"])
        else:
            _forma(slide, MSO_SHAPE.RECTANGLE, 0, 0, W, Inches(0.25), p["acento"])
        _forma(slide, MSO_SHAPE.RECTANGLE, Inches(0.9), Inches(1.7), Inches(1.4), Inches(0.08), p["acento"])

        titulo = slide.shapes.title
        _colocar(titulo, Inches(0.9), Inches(0.45), W - Inches(3.0), Inches(1.2))
        tam_t = _mejor_tam([titulo.text], titulo.width, titulo.height, [34, 30, 28, 24], 0, 0)
        _estilo(titulo, tam_t, p["titulo"], negrita=True, fuente=d["fuente_titulo"])

        cuerpo = slide.placeholders[1]
        pics = _imagenes(slide)
        top_cuerpo, alto_cuerpo = Inches(2.0), H - Inches(3.0)

        if pics and not es_fuentes:
            pic = pics[0]
            img_w, img_h = Inches(4.9), Inches(3.5)
            img_x = W - Inches(0.9) - img_w
            pw, ph = pic.image.size
            esc = min(img_w / pw, img_h / ph)
            nw, nh = int(pw * esc), int(ph * esc)
            _colocar(pic, img_x + (img_w - nw) / 2, top_cuerpo + (img_h - nh) / 2, nw, nh)
            _forma(slide, MSO_SHAPE.RECTANGLE, pic.left + Inches(0.12), pic.top + Inches(0.12),
                   nw, nh, p["acento2"])
            for sh in slide.shapes:
                if sh.name == "Credito imagen":
                    _colocar(sh, img_x, pic.top + nh + Inches(0.22), img_w, Inches(0.7))
                    _estilo(sh, 10, p["texto"], anclaje=MSO_ANCHOR.TOP, cursiva=True, fuente=d["fuente_cuerpo"])
            ancho_cuerpo = img_x - Inches(0.9) - Inches(0.4)
        else:
            ancho_cuerpo = W - Inches(1.8)
        _colocar(cuerpo, Inches(0.9), top_cuerpo, ancho_cuerpo, alto_cuerpo)

        pars = cuerpo.text_frame.paragraphs
        textos = [x.text for x in pars]
        if es_fuentes:
            espacio, tams = 4, [18, 16, 14, 13, 12, 11, 10]
        elif len(pars) == 1 and not pics:
            espacio, tams = 12, [28, 26, 24, 22, 20, 18]   # texto corrido a ancho completo
        else:
            espacio, tams = 12, [24, 22, 20, 18, 16, 14]
        tam = _mejor_tam(textos, cuerpo.width, cuerpo.height, tams, espacio)
        solo_texto = len(pars) == 1 and not pics and not es_fuentes
        _estilo(cuerpo, tam, p["texto"], anclaje=MSO_ANCHOR.MIDDLE if solo_texto else MSO_ANCHOR.TOP,
                espacio=espacio, fuente=d["fuente_cuerpo"])

        if es_fuentes:
            for par in pars:
                _sin_vineta(par)
                if par.text in ENCABEZADOS_FUENTES:
                    for run in par.runs:
                        run.font.bold = True
                        run.font.size = Pt(tam + 3)
                        run.font.color.rgb = _rgb(p["acento"])

        circulo = _forma(slide, MSO_SHAPE.OVAL, W - Inches(1.1), H - Inches(0.95),
                         Inches(0.6), Inches(0.6), p["acento"], atras=False)
        tf = circulo.text_frame
        tf.text = str(numero)
        par = tf.paragraphs[0]
        par.alignment = PP_ALIGN.CENTER
        run = par.runs[0]
        run.font.size = Pt(14)
        run.font.bold = True
        run.font.name = d["fuente_cuerpo"]
        run.font.color.rgb = _rgb(BLANCO)
        return tam


# ----------------------------------------------------------------------
# El Jefe: coordina, valida, manda a corregir y, si hace falta, corrige él
# ----------------------------------------------------------------------
class ErrorPipeline(Exception):
    pass


SYS_JEFE = """
Eres el Jefe supervisor de un equipo de agentes que convierte documentos en presentaciones.
Revisas el trabajo con criterio y sin ser quisquilloso: señalas errores reales, no gustos.
Responde ÚNICAMENTE con JSON válido.
"""


class Jefe:
    nombre = "Jefe"
    rol = "Supervisa y corrige a los agentes"

    def __init__(self):
        self.log = []
        self.t_inicio = time.time()
        self.texto_pdf = ""
        self.texto_original = ""
        self.tope = 10          # tope de diapositivas (escala con las páginas del PDF)
        self.n_temas = 0        # temas que entregó el Recopilador
        self.secciones = []

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
            if s.get("solapamiento") is not None:
                _chk(c, "La lectura de Gemini coincide con la de pdfplumber (al menos 60% de palabras)",
                     s["solapamiento"] >= 0.6,
                     f"La transcripción de Gemini solo coincide en {round(s['solapamiento'] * 100)}% con "
                     "pdfplumber; puede estar incompleta o alterada")
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
            _chk(c, "Todos los temas tienen puntos", not sin_puntos, f"Temas sin puntos: {sin_puntos}")
            refs = s.get("referencias", [])
            _chk(c, "'referencias' es una lista de máximo 8", isinstance(refs, list) and len(refs) <= 8,
                 "'referencias' debe ser una lista de máximo 8 elementos")
            if isinstance(refs, list):
                inventadas = [r for r in refs if not _referencia_en_texto(str(r), self.texto_pdf)]
                _chk(c, "Todas las referencias aparecen realmente en el PDF (anti-invención)",
                     not inventadas,
                     f"Estas referencias no aparecen en el documento, quítalas: {inventadas}")
        elif nombre == "Seccionador":
            c += self._checks_plan(s["slides"], s.get("razonamiento"))
            if all(r["ok"] or not r["bloquea"] for r in c):  # la IA revisa solo si pasó las reglas fijas
                try:
                    r = self._revision_ia_plan(s)
                    _chk(c, "Revisión con IA: fiel al documento, sin ideas repetidas y coherente",
                         not r["problemas_graves"], "La revisión con IA encontró: " + " | ".join(r["problemas_graves"]))
                    if r["sugerencias"]:
                        _chk(c, "Sugerencias de la IA (no bloquean)", False, " | ".join(r["sugerencias"]), bloquea=False)
                    if r["veredicto"]:
                        _chk(c, "Veredicto de la IA: " + r["veredicto"], True)
                except Exception as e:
                    _chk(c, "Revisión con IA", False, f"No se pudo ejecutar ({type(e).__name__}: {e})", bloquea=False)
        elif nombre == "Ilustrador":
            n_sl, cap, plan = len(s["slides"]), s["max_imagenes"], s.get("plan_imagenes", [])
            _chk(c, f"Planea entre 1 y {cap} imágenes (según el tamaño de la presentación)",
                 1 <= len(plan) <= cap, f"El plan tiene {len(plan)} imágenes; debe tener entre 1 y {cap}")
            fuera = [p.get("diapositiva") for p in plan
                     if not (isinstance(p.get("diapositiva"), int) and 1 <= p["diapositiva"] <= n_sl)]
            _chk(c, "Cada imagen planeada apunta a una diapositiva que existe", not fuera,
                 f"Diapositivas inexistentes en el plan de imágenes: {fuera}")
            nums = [p.get("diapositiva") for p in plan]
            _chk(c, "No planea dos imágenes para la misma diapositiva", len(nums) == len(set(nums)),
                 "Hay diapositivas repetidas en el plan de imágenes")
            malas_q = [p.get("diapositiva") for p in plan
                       if not p.get("consulta") or num_palabras(p["consulta"]) > 6]
            _chk(c, "Las búsquedas son cortas (máximo 6 palabras) y no están vacías", not malas_q,
                 f"Búsquedas vacías o muy largas en las diapositivas {malas_q}")
            _chk(c, "Explica por qué eligió esas diapositivas", s.get("razonamiento_imagenes"),
                 "Falta el razonamiento de la selección de imágenes")
            imgs = [(i, x["imagen"]) for i, x in enumerate(s["slides"], 1) if x.get("imagen")]
            _chk(c, f"Máximo {cap} imágenes", len(imgs) <= cap, f"Hay {len(imgs)} imágenes")
            malas, sin_cred, sin_razon = [], [], []
            for i, img in imgs:
                try:
                    with Image.open(img["ruta"]) as im:
                        im.verify()
                    if img["ancho"] < 400:
                        malas.append(i)
                except Exception:
                    malas.append(i)
                if not (img.get("autor") and img.get("licencia") and img.get("pagina")):
                    sin_cred.append(i)
                if not img.get("razon"):
                    sin_razon.append(i)
            _chk(c, "Cada imagen es un archivo legible de al menos 400 px de ancho", not malas,
                 f"Imágenes ilegibles o muy pequeñas en las diapositivas {malas}")
            _chk(c, "Cada imagen trae autor, licencia y enlace de origen", not sin_cred,
                 f"Faltan datos de crédito en las diapositivas {sin_cred}")
            _chk(c, "Cada imagen trae la justificación de por qué ilustra el tema", not sin_razon,
                 f"Falta la justificación en las diapositivas {sin_razon}")
            titulos = [img["titulo"] for _, img in imgs]
            _chk(c, "No hay imágenes repetidas", len(titulos) == len(set(titulos)), "Imagen repetida")
        elif nombre == "Generador PPTX":
            existe = os.path.exists(s["ruta"])
            _chk(c, "El archivo .pptx existe", existe, "No se creó el archivo")
            if existe:
                prs = Presentation(s["ruta"])
                _chk(c, "El número de diapositivas coincide con el plan (portada + contenido + fuentes)",
                     len(prs.slides) == s["n_slides"], "El número de diapositivas no coincide con el plan")
                portada = prs.slides[0].shapes.title.text if len(prs.slides) else ""
                _chk(c, "La primera diapositiva es la portada con el título",
                     portada == s["titulo"], "La portada no tiene el título esperado")
                ultima = prs.slides[-1].shapes.title.text if len(prs.slides) else ""
                _chk(c, f"La última diapositiva es '{TITULO_FUENTES}'", ultima == TITULO_FUENTES,
                     "Falta la diapositiva de fuentes al final")
                n_pics = sum(len(_imagenes(sl)) for sl in prs.slides)
                _chk(c, "Cada imagen del plan quedó insertada en su diapositiva",
                     n_pics == s["n_imagenes"], f"Hay {n_pics} imágenes y el plan tenía {s['n_imagenes']}")
        elif nombre == "Decorador":
            prs = Presentation(s["ruta"])
            W, H = prs.slide_width, prs.slide_height
            prob_c = _problemas_contraste(s["diseno"]["paleta"])
            _chk(c, "Los colores tienen contraste suficiente para leerse", not prob_c,
                 "Contraste bajo: " + "; ".join(prob_c))
            _chk(c, "El diseño lo propuso la IA (si no, se usó el diseño de respaldo)",
                 s.get("origen_diseno") == "IA",
                 "Se usó el diseño de respaldo: " + "; ".join(s.get("problemas_ia", [])), bloquea=False)
            _chk(c, "No cambió el número de diapositivas", len(prs.slides) == s["n_slides"],
                 "Cambió el número de diapositivas")
            sin_fondo = [i for i, sl in enumerate(prs.slides, 1) if sl.background.fill.type != MSO_FILL.SOLID]
            _chk(c, "Todas las diapositivas tienen fondo aplicado", not sin_fondo,
                 f"Sin fondo en las diapositivas: {sin_fondo}")
            sin_formas = [i for i, sl in enumerate(prs.slides, 1) if len(sl.shapes) <= len(sl.placeholders)]
            _chk(c, "Todas las diapositivas tienen elementos decorativos", not sin_formas,
                 f"Sin decoración en las diapositivas: {sin_formas}")
            fuera, solapa = [], []
            for i, sl in enumerate(prs.slides, 1):
                for pic in _imagenes(sl):
                    if pic.left < 0 or pic.left + pic.width > W or pic.top + pic.height + Inches(0.8) > H:
                        fuera.append(i)
                    cuerpo = sl.placeholders[1]
                    if cuerpo.left + cuerpo.width > pic.left:
                        solapa.append(i)
            _chk(c, "Las imágenes y sus créditos caben dentro de la diapositiva", not fuera,
                 f"Imagen fuera de los márgenes en las diapositivas {fuera}")
            _chk(c, "Las imágenes no se encimen con el texto", not solapa,
                 f"La imagen se encima con el texto en las diapositivas {solapa}")
            desborda = []
            for i, sl in enumerate(prs.slides, 1):
                if i == 1:
                    continue
                if not _cabe(sl.placeholders[1]) or not _cabe(sl.shapes.title, sangria_in=0):
                    desborda.append(i)
            _chk(c, "El texto cabe en su caja (estimación, conviene revisarlo en PowerPoint)",
                 not desborda, f"Posible texto desbordado en las diapositivas {desborda}", bloquea=False)
        return c

    def _checks_plan(self, slides, razonamiento):
        c = []
        n = len(slides)
        _chk(c, f"Entre 3 y {self.tope} diapositivas de contenido (tope según las páginas del PDF)",
             3 <= n <= self.tope, f"Hay {n} diapositivas; debe haber entre 3 y {self.tope}")
        malos = [i for i, sl in enumerate(slides, 1) if not sl.get("title") or num_palabras(sl["title"]) > 8]
        _chk(c, "Títulos de máximo 8 palabras", not malos,
             f"Títulos vacíos o de más de 8 palabras en las diapositivas {malos}")

        tipos_mal = [i for i, sl in enumerate(slides, 1) if sl.get("tipo") not in ("parrafo", "mixto", "bullets")]
        _chk(c, "Cada diapositiva tiene un 'tipo' válido (parrafo, mixto o bullets)", not tipos_mal,
             f"'tipo' inválido o ausente en las diapositivas {tipos_mal}")

        incoherentes = []
        for i, sl in enumerate(slides, 1):
            t, par, bl = sl.get("tipo"), sl.get("parrafo", "") or "", sl.get("bullets", []) or []
            if t == "parrafo" and not 25 <= num_palabras(par) <= 55:
                incoherentes.append(f"{i} (párrafo de {num_palabras(par)} palabras; debe tener 25 a 55)")
            elif t == "mixto" and not (10 <= num_palabras(par) <= 30 and 2 <= len(bl) <= 5):
                incoherentes.append(f"{i} (mixto: párrafo de 10 a 30 palabras y 2 a 5 bullets)")
            elif t == "bullets" and not 3 <= len(bl) <= 8:
                incoherentes.append(f"{i} (bullets: debe tener 3 a 8)")
        _chk(c, "El contenido de cada diapositiva coincide con su tipo y sus límites", not incoherentes,
             "Revisa las diapositivas: " + "; ".join(incoherentes))

        solo_bullets = sum(1 for sl in slides if sl.get("tipo") == "bullets")
        permitido = math.ceil(0.6 * n)
        _chk(c, "No más del 60% de las diapositivas son solo bullets", solo_bullets <= permitido,
             f"{solo_bullets} de {n} diapositivas son solo bullets (máximo {permitido}); usa más texto corrido")

        largos = [i for i, sl in enumerate(slides, 1)
                  if any(num_palabras(b) > 15 for b in (sl.get("bullets") or []))]
        _chk(c, "Bullets de máximo 15 palabras", not largos,
             f"Bullets de más de 15 palabras en las diapositivas {largos}")
        vistos, repetidos = set(), []
        for i, sl in enumerate(slides, 1):
            for b in (sl.get("bullets") or []):
                clave = str(b).strip().lower()
                if clave in vistos:
                    repetidos.append(i)
                vistos.add(clave)
        _chk(c, "Sin bullets repetidos", not repetidos,
             f"Bullets repetidos en las diapositivas {sorted(set(repetidos))}")

        def temas_validos(sl):
            t = sl.get("temas")
            return (isinstance(t, list) and t
                    and all(isinstance(x, int) and not isinstance(x, bool) and 1 <= x <= self.n_temas for x in t))
        temas_mal = [i for i, sl in enumerate(slides, 1) if not temas_validos(sl)]
        _chk(c, "Cada diapositiva indica qué temas del Recopilador cubre (números válidos)", not temas_mal,
             f"Falta 'temas' o tiene números inválidos (1 a {self.n_temas}) en las diapositivas {temas_mal}")
        cubiertos = {x for sl in slides if temas_validos(sl) for x in sl["temas"]}
        faltan = sorted(set(range(1, self.n_temas + 1)) - cubiertos)
        _chk(c, "Cobertura: todos los temas del Recopilador aparecen en alguna diapositiva", not faltan,
             "Ninguna diapositiva cubre los temas " + ", ".join(
                 f"{k} ({self.secciones[k - 1].get('tema', '?')})" for k in faltan))
        _chk(c, "Explica su razonamiento", razonamiento, "Falta el campo 'razonamiento'")
        return c

    # ---- revisiones con IA del Jefe ----
    def _revision_ia_plan(self, salida):
        plan = [{k: sl.get(k) for k in ("title", "tipo", "parrafo", "bullets")} for sl in salida["slides"]]
        prompt = (
            "Compara el PLAN de diapositivas con el DOCUMENTO original y revisa:\n"
            "1. Fidelidad: ¿alguna afirmación del plan NO aparece en el documento o lo contradice?\n"
            "2. Repetición: ¿hay ideas repetidas entre diapositivas?\n"
            "3. Coherencia: ¿el orden es lógico (introducción, desarrollo, conclusión) y todo es del "
            "tema del documento?\n"
            "Pon en 'problemas_graves' SOLO lo que sea claramente un error (afirmación inventada, "
            "contradicción, idea repetida, diapositiva fuera de tema), indicando el número de diapositiva. "
            "Los asuntos de estilo o mejoras opcionales van en 'sugerencias'. Si no hay problemas graves, "
            "deja la lista vacía.\n"
            'Responde {"problemas_graves": ["diapositiva N: ..."], "sugerencias": ["..."], '
            '"veredicto": "<una frase>"}\n\n'
            "DOCUMENTO:\n" + _recortar_texto(self.texto_original) + "\n\nPLAN:\n"
            + json.dumps(plan, ensure_ascii=False)
        )
        d = parsear_json(llamar_llm(SYS_JEFE, prompt))
        return {"problemas_graves": [str(x) for x in d.get("problemas_graves", [])],
                "sugerencias": [str(x) for x in d.get("sugerencias", [])],
                "veredicto": str(d.get("veredicto", ""))}

    def _revision_final(self, ruta):
        t0 = time.time()
        try:
            prs = Presentation(ruta)
            textos = []
            for i, sl in enumerate(prs.slides, 1):
                partes = [sh.text_frame.text for sh in sl.shapes
                          if sh.has_text_frame and len(sh.text_frame.text.strip()) > 2]
                textos.append(f"Diapositiva {i}: " + " | ".join(partes))
            prompt = (
                "Esta es la presentación FINAL generada a partir del DOCUMENTO. Califica del 1 al 10 su "
                "calidad (claridad, coherencia, fidelidad al documento, equilibrio entre texto e "
                "imágenes). Indica fortalezas y problemas. Si hay problemas, di qué agente debería "
                "revisarse (Lector, Recopilador, Seccionador, Ilustrador, Generador PPTX o Decorador).\n"
                'Responde {"calificacion": <1-10>, "fortalezas": ["..."], "problemas": ["..."], '
                '"agente_a_revisar": "<nombre o null>", "veredicto": "<una frase>"}\n\n'
                "DOCUMENTO:\n" + _recortar_texto(self.texto_original, 8000)
                + "\n\nPRESENTACIÓN:\n" + "\n".join(textos)
            )
            d = parsear_json(llamar_llm(SYS_JEFE, prompt))
            detalle = ["Fortalezas:"] + [f"   + {x}" for x in d.get("fortalezas", [])]
            detalle += ["Problemas:"] + ([f"   - {x}" for x in d.get("problemas", [])] or ["   (ninguno)"])
            detalle.append(f"Agente a revisar: {d.get('agente_a_revisar') or 'ninguno'}")
            self._anotar(self.nombre, "Revisión final de la presentación con IA", "-", "🧐 Revisión final",
                         time.time() - t0,
                         {"recibio": f"Presentación final de {len(prs.slides)} diapositivas",
                          "entrego": f"Calificación {d.get('calificacion', '?')}/10",
                          "razonamiento": str(d.get("veredicto", "")), "detalle": "\n".join(detalle)})
        except Exception as e:
            self._anotar(self.nombre, "Revisión final de la presentación con IA", "-",
                         "⚠️ Revisión final no disponible", time.time() - t0,
                         {"recibio": "-", "entrego": "-",
                          "razonamiento": f"No se pudo ejecutar la revisión con IA: {type(e).__name__}: {e}"})

    # ---- corrección directa del Jefe (cuando el agente ya no puede) ----
    def _corregir_plan(self, salida):
        cambios, vistos, nuevos = [], set(), []
        slides = salida["slides"]
        if len(slides) > self.tope:
            cambios.append(f"recortó el plan a {self.tope} diapositivas")
        for i, sl in enumerate(slides[:self.tope], 1):
            t = " ".join(str(sl.get("title", "Sin título")).split()[:8])
            par = (sl.get("parrafo") or "").strip()
            if num_palabras(par) > 55:
                par = " ".join(par.split()[:55]) + "…"
                cambios.append(f"recortó el párrafo de la diapositiva {i}")
            bullets = []
            for b in (sl.get("bullets") or []):
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
            if len(bullets) > 8:
                bullets = bullets[:8]
                cambios.append(f"dejó solo 8 bullets en la diapositiva {i}")
            tipo = "mixto" if par and bullets else ("parrafo" if par else "bullets")
            if tipo != sl.get("tipo"):
                cambios.append(f"ajustó el tipo de la diapositiva {i} a '{tipo}'")
            temas = [x for x in (sl.get("temas") or [])
                     if isinstance(x, int) and not isinstance(x, bool) and 1 <= x <= self.n_temas]
            nuevos.append({"title": t, "tipo": tipo, "parrafo": par, "bullets": bullets, "temas": temas})

        # Cobertura: por cada tema que nadie cubrió, el Jefe agrega una diapositiva con los puntos del Recopilador
        cubiertos = {x for sl in nuevos for x in sl["temas"]}
        for k in range(1, self.n_temas + 1):
            if k in cubiertos or len(nuevos) >= self.tope:
                continue
            sec = self.secciones[k - 1]
            puntos = []
            for pt in (sec.get("puntos") or [])[:8]:
                palabras = str(pt).split()
                puntos.append(" ".join(palabras[:15]) + ("…" if len(palabras) > 15 else ""))
            nueva = {"title": " ".join(str(sec.get("tema", f"Tema {k}")).split()[:8]), "tipo": "bullets",
                     "parrafo": "", "bullets": puntos, "temas": [k]}
            nuevos.insert(max(0, len(nuevos) - 1), nueva)   # antes de la conclusión
            cambios.append(f"agregó una diapositiva para el tema {k} («{sec.get('tema', '?')}») que nadie cubría")
        razon = salida.get("razonamiento")
        if not razon:
            cambios.append("completó el razonamiento faltante")
        return {**salida, "slides": nuevos, "razonamiento": razon or "(el agente no explicó su decisión)"}, cambios

    # ---- ciclo de supervisión de un agente ----
    def _supervisar(self, agente, entrada, opcional=False, respaldo=None):
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
            problemas = [r["detalle"] for r in revisiones if not r["ok"] and r["bloquea"]]

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
            if all(r["ok"] or not r["bloquea"] for r in revisiones):
                self._anotar(
                    self.nombre, self.rol, "-", "🛠️ Corrigió él mismo", 0,
                    {
                        "recibio": f"Plan rechazado {MAX_INTENTOS} veces",
                        "entrego": f"Plan de {len(nuevo['slides'])} diapositivas corregido",
                        "razonamiento": "El Seccionador no logró cumplir las reglas tras varios intentos, "
                                        "así que apliqué los ajustes mínimos necesarios yo mismo.",
                        "detalle": "\n".join(sorted(set(cambios))) or
                                   ("Sin cambios de formato. Se aceptó con las observaciones pendientes:\n"
                                    + (feedback or "")),
                    },
                    revisiones,
                )
                return nuevo

        if opcional and respaldo is not None:
            self._anotar(self.nombre, self.rol, "-", "⚠️ Siguió sin este paso", 0,
                         {"recibio": f"{agente.nombre} falló {MAX_INTENTOS} veces",
                          "entrego": "Presentación sin imágenes",
                          "razonamiento": "Las imágenes son un extra, no un requisito para tener una "
                                          "presentación útil, así que continué sin ellas en vez de detener todo."},
                         notas=feedback or "")
            return respaldo

        self._anotar(self.nombre, self.rol, "-", "🛑 Detuvo el proceso", 0,
                     {"recibio": f"{agente.nombre} falló {MAX_INTENTOS} veces",
                      "razonamiento": "No pude corregirlo yo mismo, así que detengo el proceso para no "
                                      "entregar una presentación incorrecta."},
                     notas=feedback or "")
        raise ErrorPipeline(f"{agente.nombre} no logró entregar un resultado válido")

    # ---- el flujo completo ----
    def coordinar(self, pdf_path, ruta_pptx, carpeta):
        lectura = self._supervisar(Lector(), pdf_path)
        self.texto_original = lectura["texto"]
        self.tope = min(60, max(10, int(lectura.get("paginas", 1))))   # 1 diapositiva por página, entre 10 y 60
        self.texto_pdf = (lectura["texto"] + "\n" + lectura.get("texto_plumber", "")).lower()
        resumen = self._supervisar(Recopilador(), lectura)
        self.secciones = resumen.get("secciones", [])
        self.n_temas = len(self.secciones)
        plan = self._supervisar(Seccionador(self.tope), resumen)
        plan = {**plan, "documento": os.path.basename(pdf_path), "referencias": resumen.get("referencias", [])}
        respaldo = {**plan, "slides": [{k: v for k, v in s.items() if k != "imagen"} for s in plan["slides"]]}
        plan = self._supervisar(Ilustrador(carpeta), plan, opcional=True, respaldo=respaldo)
        base = self._supervisar(GeneradorPPTX(ruta_pptx), plan)
        final = self._supervisar(Decorador(), base)
        self._revision_final(final["ruta"])
        return final["ruta"]

    # ---- reporte para ti (el "jefe de jefes") ----
    def reporte_html(self, exito):
        e = escape
        rechazos = sum(1 for x in self.log if "Rechazado" in x["estado"])
        correcciones = sum(1 for x in self.log if "Corrigió" in x["estado"])
        omitidos = sum(1 for x in self.log if "Siguió" in x["estado"])
        total = round(time.time() - self.t_inicio, 1)
        caja = "border:1px solid rgba(128,128,128,.4);border-radius:8px;margin:8px 0;padding:8px 12px;"
        celda = "border:1px solid rgba(128,128,128,.35);padding:6px 8px;text-align:left;"

        h = ['<div style="max-height:75vh;overflow-y:auto;padding-right:10px;">']
        h.append("<h3>Informe final del Jefe</h3><ul>")
        h.append(f"<li>Resultado: {'✅ presentación generada' if exito else '🛑 el proceso se detuvo'}</li>")
        h.append(f"<li>Rechazos a agentes: {rechazos}</li>")
        h.append(f"<li>Correcciones hechas por el Jefe: {correcciones}</li>")
        h.append(f"<li>Pasos opcionales omitidos: {omitidos}</li>")
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
            abierto = " open" if any(k in x["estado"] for k in ("Rechazado", "Corrigió", "Detuvo", "Siguió", "Revisión final")) else ""
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
                    marca = "✅" if r["ok"] else ("❌" if r["bloquea"] else "⚠️")
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
        jefe.coordinar(pdf_file, ruta_pptx, carpeta)
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
        "Sube un PDF. Seis agentes trabajan en cadena (Lector, Recopilador, Seccionador, "
        "Ilustrador, Generador, Decorador) supervisados por un Jefe que valida con reglas y con IA, "
        "corrige y te entrega un informe de quién hizo qué y por qué. Las imágenes vienen de "
        "Wikimedia Commons con su crédito."
    ),
    flagging_mode="never",
)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    demo.launch(server_name="0.0.0.0", server_port=port)
