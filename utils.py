"""Utilidades compartidas por todos los agentes: IA (Gemini), JSON y texto, diseño, Wikimedia Commons,
ilustraciones con IA, funciones de python-pptx y cálculos del plan. Cada sección tiene su encabezado."""
import io
import json
import math
import re
import time
from html import unescape
from urllib.parse import quote

import google.generativeai as genai
import requests
from PIL import Image
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Pt

from config import (CHARS_POR_DIAPOSITIVA, LIMITE_TEXTO_REVISION, MAX_DIAP_POR_TEMA, MAX_DIAPOSITIVAS,
                    MODEL_NAME, PROPORCION_IMAGENES, USER_AGENT)


# ==============================================================================
# Base de los agentes y errores del pipeline
# ==============================================================================


class Agente:
    nombre = "Agente"
    rol = ""

    def ejecutar(self, entrada, feedback=None):
        """Debe devolver {"salida", "recibio", "entrego", "razonamiento", "detalle", "notas"}"""
        raise NotImplementedError


class ErrorPipeline(Exception):
    pass


# ==============================================================================
# IA (Gemini), JSON y texto
# ==============================================================================


def llamar_llm(system_instruction, prompt):
    """Una llamada a Gemini. 'prompt' puede ser texto o una lista (texto + imágenes).
    Si Gemini responde que hay demasiadas peticiones o está saturado, espera y reintenta."""
    model = genai.GenerativeModel(MODEL_NAME, system_instruction=system_instruction)
    esperas = (6, 20, 45)
    for k in range(len(esperas) + 1):
        try:
            return model.generate_content(prompt).text.strip()
        except Exception as e:
            descripcion = f"{type(e).__name__} {e}"
            pasajero = any(t in descripcion for t in ("429", "503", "ResourceExhausted", "ServiceUnavailable",
                                                      "TooManyRequests", "DeadlineExceeded"))
            if not pasajero or k == len(esperas):
                raise
            time.sleep(esperas[k])


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


def _chk(lista, regla, ok, detalle="", bloquea=True):
    """Registra una revisión del Jefe. Si 'bloquea' es False, solo avisa (⚠️)."""
    lista.append({"regla": regla, "ok": bool(ok), "detalle": "" if ok else (detalle or regla),
                  "bloquea": bloquea})


def _limpiar_html(s):
    return re.sub(r"\s+", " ", unescape(re.sub(r"<[^>]+>", "", s or ""))).strip()


def _recortar_texto(texto, limite=LIMITE_TEXTO_REVISION):
    """Si el texto pasa del límite, conserva el inicio y el final (ahí suele estar la bibliografía)."""
    if len(texto) <= limite:
        return texto
    ini = int(limite * 0.77)
    return texto[:ini] + "\n[... texto omitido ...]\n" + texto[-(limite - ini):]


def _referencia_en_texto(ref, texto_min):
    """Guardia anti-invención: casi todas las palabras de la referencia deben estar en el PDF."""
    palabras = re.findall(r"\w{4,}", ref.lower())
    if not palabras:
        return False
    return sum(1 for w in palabras if w in texto_min) / len(palabras) >= 0.7


# ==============================================================================
# Diseño: ¿cabe el texto en su caja? y contraste de colores
# ==============================================================================


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


# ==============================================================================
# Wikimedia Commons (imágenes con licencia libre y autor)
# ==============================================================================


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


# ==============================================================================
# Ilustraciones generadas con IA (último recurso; Pollinations, gratis y sin clave)
# ==============================================================================


_URL = "https://image.pollinations.ai/prompt/"
_ESTILO = ", flat vector illustration, simple composition, clean light background, no text, no letters, no watermark"
_ESPERA_ENTRE_PETICIONES = 16     # segundos (el plan anónimo permite ~1 petición cada 15 s)
_ultima = 0.0


def generar_imagen_ia(prompt, ancho=1024, alto=768, intentos=3):
    """Devuelve (bytes_jpeg, 'image/jpeg'). Lanza una excepción si no pudo generarla."""
    global _ultima
    texto = quote((prompt.strip()[:300]) + _ESTILO)
    url = f"{_URL}{texto}"
    ultimo_error = None
    for k in range(intentos):
        falta = _ESPERA_ENTRE_PETICIONES - (time.time() - _ultima)
        if falta > 0:
            time.sleep(falta)
        _ultima = time.time()
        try:
            r = requests.get(url, params={"width": ancho, "height": alto, "nologo": "true"},
                             headers={"User-Agent": USER_AGENT}, timeout=120)
            if r.status_code in (429, 500, 502, 503, 504):
                ultimo_error = RuntimeError(f"HTTP {r.status_code}")
                continue
            r.raise_for_status()
            if not r.headers.get("Content-Type", "").startswith("image/"):
                raise ValueError("la respuesta no es una imagen")
            im = Image.open(io.BytesIO(r.content))
            im.load()
            salida = io.BytesIO()
            im.convert("RGB").save(salida, "JPEG", quality=90)
            return salida.getvalue(), "image/jpeg"
        except Exception as e:
            ultimo_error = e
    raise RuntimeError(f"no se pudo generar la imagen ({type(ultimo_error).__name__}: {ultimo_error})")


# ==============================================================================
# Funciones para dibujar y escribir en diapositivas (python-pptx)
# ==============================================================================


def _sin_vineta(par):
    """Quita la viñeta y la sangría de un párrafo (para texto corrido)."""
    pPr = par._p.get_or_add_pPr()
    pPr.set("marL", "0")
    pPr.set("indent", "0")
    if pPr.find(qn("a:buNone")) is None:
        pPr.append(pPr.makeelement(qn("a:buNone"), {}))


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


# ==============================================================================
# Plan: temas, diapositivas e imágenes por diapositiva
# ==============================================================================


def rango_de_temas(paginas):
    """Cuántos temas debe sacar el Recopilador según el tamaño del PDF (20 páginas -> de 6 a 20)."""
    minimo = max(3, paginas // 3)
    maximo = min(40, max(8, paginas))
    return minimo, max(minimo, maximo)


def caracteres_tema(sec):
    """Caracteres de contenido de un tema (su nombre y todos sus puntos)."""
    return len(str(sec.get("tema", ""))) + sum(len(str(p)) for p in (sec.get("puntos") or []))


def presupuesto_diapositivas(secciones):
    """Cada tema recibe las diapositivas que pida la cantidad de caracteres de su contenido."""
    filas = []
    for i, sec in enumerate(secciones, 1):
        chars = caracteres_tema(sec)
        m = min(MAX_DIAP_POR_TEMA, max(1, math.floor(chars / CHARS_POR_DIAPOSITIVA + 0.5)))
        filas.append({"numero": i, "tema": str(sec.get("tema", f"Tema {i}")), "caracteres": chars,
                      "diapositivas": m})
    # Si se pasa del tope absoluto (sin contar introducción y conclusión), se recortan los temas más grandes
    limite = max(len(filas), MAX_DIAPOSITIVAS - 2)
    while sum(f["diapositivas"] for f in filas) > limite:
        mayor = max(filas, key=lambda f: f["diapositivas"])
        if mayor["diapositivas"] <= 1:
            break
        mayor["diapositivas"] -= 1
    return filas


def total_objetivo(presupuesto):
    """Diapositivas de contenido que corresponden: las de los temas + introducción + conclusión."""
    return sum(f["diapositivas"] for f in presupuesto) + 2


def rango_diapositivas(presupuesto):
    """Límites del veto del Jefe: la IA decide cuántas diapositivas, pero no menos de 70% ni más de 150%
    de lo que sugiere la fórmula de los caracteres de cada tema."""
    obj = total_objetivo(presupuesto)
    return max(3, math.floor(obj * 0.7)), min(MAX_DIAPOSITIVAS, math.ceil(obj * 1.5))


def min_imagenes(n):
    """Mínimo de diapositivas con imagen (60% de las de contenido, redondeado hacia arriba)."""
    return math.ceil(round(PROPORCION_IMAGENES * n, 6))


def asegurar_imagenes(slides):
    """Deja 'lleva_imagen' en cada diapositiva: al menos 60% con imagen y nunca dos seguidas sin ella.
    Respeta las que ya estaban marcadas y solo agrega marcas. Devuelve los números de diapositiva
    (desde 1) que tuvo que marcar."""
    n = len(slides)
    marcas = [str(sl.get("lleva_imagen")).strip().lower() in ("true", "1", "si", "sí") for sl in slides]
    cambios = []
    antes = list(marcas)
    # 1) alternar: empezar con imagen y no dejar dos diapositivas seguidas sin ella
    if n >= 2 and not marcas[0] and not marcas[1]:
        marcas[0] = True
    for i in range(1, n):
        if not marcas[i] and not marcas[i - 1]:
            marcas[i] = True
    # 2) cuota mínima: marcar las que falten, repartidas a lo largo de la presentación
    falta = min_imagenes(n) - sum(marcas)
    if falta > 0:
        libres = [i for i, m in enumerate(marcas) if not m]
        for j in range(min(falta, len(libres))):
            marcas[libres[int((j + 0.5) * len(libres) / falta)]] = True
    for i, (sl, m) in enumerate(zip(slides, marcas)):
        sl["lleva_imagen"] = m
        if m and not antes[i]:
            cambios.append(i + 1)
    return cambios
