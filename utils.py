"""Utilidades compartidas por todos los agentes: IA (Gemini), JSON y texto, diseño, Wikimedia Commons,
ilustraciones con IA, funciones de python-pptx, cálculos del plan y motor RAG ligero en memoria."""

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
# RAG ligero en memoria (por fragmentos y páginas)
# ==============================================================================

def _tokens_rag(texto):
    """Palabras normalizadas de 5+ letras recortadas a su raíz de 6 caracteres."""
    return {w[:6] for w in re.findall(r"[a-záéíóúñü0-9]{5,}", str(texto).lower())}


def crear_indice_rag(paginas_texto, texto_completo="", max_chars=900):
    """Divide el PDF en fragmentos (chunks) con número de página y tokens para búsqueda rápida."""
    indice = []
    fuentes = [(i, t) for i, t in enumerate(paginas_texto or [], 1) if t and t.strip()]
    if not fuentes and texto_completo:
        partes = [texto_completo[i:i + max_chars] for i in range(0, len(texto_completo), max_chars)]
        fuentes = [(i, p) for i, p in enumerate(partes, 1)]

    for num_pag, txt in fuentes:
        parrafos = [p.strip() for p in re.split(r"\n{2,}", txt.strip()) if p.strip()]
        actual = ""
        for p in parrafos:
            if actual and len(actual) + len(p) + 1 > max_chars:
                indice.append({
                    "id": len(indice) + 1,
                    "pagina": num_pag,
                    "texto": actual.strip(),
                    "tokens": list(_tokens_rag(actual)),
                })
                actual = ""
            if len(p) > max_chars:
                for k in range(0, len(p), max_chars):
                    sub = p[k:k + max_chars].strip()
                    if sub:
                        indice.append({
                            "id": len(indice) + 1,
                            "pagina": num_pag,
                            "texto": sub,
                            "tokens": list(_tokens_rag(sub)),
                        })
            else:
                actual = f"{actual}\n{p}".strip() if actual else p
        if actual:
            indice.append({
                "id": len(indice) + 1,
                "pagina": num_pag,
                "texto": actual.strip(),
                "tokens": list(_tokens_rag(actual)),
            })
    return indice


def recuperar_contexto_rag(indice_rag, consulta, top_k=3):
    """Recupera los 'top_k' fragmentos más relevantes del PDF para un tema y sus puntos."""
    if not indice_rag:
        return []
    q_tokens = _tokens_rag(consulta)
    if not q_tokens:
        return [f"[Pág. {ch['pagina']}] {ch['texto']}" for ch in indice_rag[:top_k]]

    calificados = []
    for ch in indice_rag:
        c_tokens = set(ch.get("tokens") or _tokens_rag(ch.get("texto", "")))
        if not c_tokens:
            continue
        inter = len(q_tokens & c_tokens)
        if inter > 0:
            score = inter / (math.sqrt(len(q_tokens)) * math.sqrt(len(c_tokens)))
            calificados.append((score, ch["id"], ch))

    calificados.sort(key=lambda x: (-x[0], x[1]))
    elegidos = [x[2] for x in calificados[:top_k]]
    if not elegidos:
        elegidos = indice_rag[:min(top_k, len(indice_rag))]
    return [f"[Pág. {ch['pagina']}] {ch['texto']}" for ch in elegidos]


# ==============================================================================
# Diseño: ¿cabe el texto en su caja? y contraste de colores (con auto-ajuste)
# ==============================================================================

def _altura_in(textos, ancho_in, tam, espacio_pt, sangria_in):
    cpl = max(1.0, (ancho_in - sangria_in) * 144 / tam / 1.1)
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


def _mezclar_rgb(color, meta, factor):
    return tuple(max(0, min(255, int(round(c + (m - c) * factor)))) for c, m in zip(color, meta))


def _forzar_contraste(color, referencia, minimo, hacia_oscuro):
    """Acerca gradualmente 'color' hacia negro o blanco hasta alcanzar el contraste 'minimo'."""
    meta = (0, 0, 0) if hacia_oscuro else (255, 255, 255)
    actual = color
    for paso in range(1, 26):
        if _contraste(actual, referencia) >= minimo:
            return actual
        actual = _mezclar_rgb(color, meta, paso / 25.0)
    return actual


def ajustar_contraste_paleta(pal):
    """Corrige suavemente los colores de la paleta propuesta por la IA si alguno tiene bajo contraste,
    conservando su tono y evitando descartar el diseño de la IA."""
    p = dict(pal)
    blanco = (255, 255, 255)
    ajustes = []

    # 1. Asegurar que fondo y tinte sean claros
    for k in ("fondo", "tinte"):
        if _luminancia(p[k]) < 0.65:
            nuevo = _mezclar_rgb(p[k], blanco, 0.75)
            ajustes.append(f"{k} aclarado ({_a_hex(p[k])}→{_a_hex(nuevo)})")
            p[k] = nuevo

    # 2. Portada oscura (para texto blanco encima >= 4.5)
    if _contraste(blanco, p["portada"]) < 4.5:
        nuevo = _forzar_contraste(p["portada"], blanco, 4.6, hacia_oscuro=True)
        ajustes.append(f"portada oscurecida ({_a_hex(p['portada'])}→{_a_hex(nuevo)})")
        p["portada"] = nuevo

    # 3. Acento con contraste suficiente para texto blanco (>= 3.0)
    if _contraste(blanco, p["acento"]) < 3.0:
        nuevo = _forzar_contraste(p["acento"], blanco, 3.1, hacia_oscuro=True)
        ajustes.append(f"acento ajustado ({_a_hex(p['acento'])}→{_a_hex(nuevo)})")
        p["acento"] = nuevo

    # 4. Acento2 visible sobre portada (>= 3.0)
    if _contraste(p["acento2"], p["portada"]) < 3.0:
        nuevo = _forzar_contraste(p["acento2"], p["portada"], 3.1, hacia_oscuro=False)
        ajustes.append(f"acento2 aclarado ({_a_hex(p['acento2'])}→{_a_hex(nuevo)})")
        p["acento2"] = nuevo

    # 5. Texto y título legibles sobre fondo y tinte (>= 4.5)
    for k in ("texto", "titulo"):
        for base in ("fondo", "tinte"):
            if _contraste(p[k], p[base]) < 4.5:
                nuevo = _forzar_contraste(p[k], p[base], 4.6, hacia_oscuro=True)
                if nuevo != p[k]:
                    ajustes.append(f"{k} oscurecido sobre {base} ({_a_hex(p[k])}→{_a_hex(nuevo)})")
                    p[k] = nuevo

    return p, ajustes


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
        "[https://commons.wikimedia.org/w/api.php](https://commons.wikimedia.org/w/api.php)",
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

_URL = "[https://image.pollinations.ai/prompt/](https://image.pollinations.ai/prompt/)"
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
    estilo = f._element.find(qn("p:style"))
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
    """Respeta en qué diapositivas decidió la IA poner imagen. Solo interviene si hay 3 o más
    diapositivas seguidas sin imagen o si el total global queda por debajo del 60%."""
    n = len(slides)
    marcas = [str(sl.get("lleva_imagen")).strip().lower() in ("true", "1", "si", "sí") for sl in slides]
    cambios = []
    antes = list(marcas)

    # 1) Evitar bloques largos sin apoyo visual (3 seguidas sin imagen -> marca la de en medio)
    for i in range(2, n):
        if not marcas[i] and not marcas[i - 1] and not marcas[i - 2]:
            marcas[i - 1] = True

    # 2) Cuota mínima global (60%): completar solo si la IA marcó menos del mínimo
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