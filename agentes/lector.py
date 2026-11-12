"""Agente Lector: lee el PDF con Gemini (respaldo: pdfplumber) y saca las imágenes que trae el PDF."""
import os
import re

import pdfplumber
from PIL import ImageStat

from utils import Agente, llamar_llm, parsear_json


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


SYS_LECTOR_IMG = """
Eres un editor gráfico. Describes imágenes extraídas de un PDF y decides si sirven para ilustrar
una presentación. Responde ÚNICAMENTE con JSON válido.
"""


def _describir_imagenes(imagenes):
    """La IA mira las imágenes (de 4 en 4), las describe y descarta logotipos, adornos e imágenes vacías.
    Si no puede dar su dictamen sobre una imagen, la conserva."""
    utiles, descartadas = [], 0
    for ini in range(0, len(imagenes), 4):
        lote = imagenes[ini:ini + 4]
        partes = ["Estas imágenes salieron de un PDF. Para cada una devuelve 'descripcion' (una frase en español "
                  "de qué muestra), 'tipo' (foto, diagrama, grafica, tabla u otro) y 'util' (true si sirve para "
                  "ilustrar contenido; false si es un logotipo, adorno, fondo, imagen casi vacía o ilegible). "
                  'Responde {"imagenes": [{"id": "...", "descripcion": "...", "tipo": "...", "util": true}]}']
        for im in lote:
            partes.append(f"Imagen {im['id']} (página {im['pagina']}):")
            with open(im["ruta"], "rb") as f:
                partes.append({"mime_type": "image/jpeg", "data": f.read()})
        try:
            info = {x.get("id"): x for x in parsear_json(llamar_llm(SYS_LECTOR_IMG, partes)).get("imagenes", [])
                    if isinstance(x, dict)}
        except Exception:
            info = {}
        for im in lote:
            x = info.get(im["id"])
            if x is not None and x.get("util") is False:
                descartadas += 1
                continue
            im["descripcion"] = str((x or {}).get("descripcion", ""))
            im["tipo"] = str((x or {}).get("tipo", "otro"))
            utiles.append(im)
    return utiles, descartadas


MAX_IMAGENES_PDF = 40       # tope de imágenes que se extraen de un PDF
MAX_IMG_POR_PAGINA = 3      # y por página (las más grandes)


def _huella(im):
    """Huella visual de 16x16 para detectar imágenes repetidas (logos, encabezados en cada página)."""
    g = im.convert("L").resize((16, 16))
    px = list(g.getdata())
    media = sum(px) / len(px)
    return "".join("1" if v > media else "0" for v in px)


def _extraer_imagenes_pdf(ruta, carpeta, paginas_texto):
    """Recorta de cada página las imágenes que trae el PDF (fotos, diagramas) y las guarda en 'carpeta'.
    Descarta iconos, páginas escaneadas completas, imágenes casi lisas y repetidas."""
    imagenes, huellas = [], set()
    with pdfplumber.open(ruta) as pdf:
        for num, page in enumerate(pdf.pages, 1):
            if len(imagenes) >= MAX_IMAGENES_PDF:
                break
            area_pagina = float(page.width * page.height)
            cajas = []
            for im in page.images:
                x0, x1 = max(im["x0"], 0), min(im["x1"], page.width)
                top, bottom = max(im["top"], 0), min(im["bottom"], page.height)
                w, h = x1 - x0, bottom - top
                if w < 100 or h < 80:
                    continue                                   # iconos y viñetas
                if not 0.03 <= (w * h) / area_pagina <= 0.85:
                    continue                                   # demasiado pequeña o página escaneada completa
                if not 0.2 <= w / h <= 5:
                    continue                                   # franjas o líneas
                cajas.append((w * h, (x0, top, x1, bottom)))
            for _, caja in sorted(cajas, reverse=True)[:MAX_IMG_POR_PAGINA]:
                try:
                    pil = page.crop(caja).to_image(resolution=150).original.convert("RGB")
                    if pil.width < 250 or pil.height < 150:
                        continue
                    if ImageStat.Stat(pil.convert("L")).stddev[0] < 8:
                        continue                               # imagen casi lisa
                    h16 = _huella(pil)
                    if h16 in huellas:
                        continue
                    huellas.add(h16)
                    n = len(imagenes) + 1
                    ruta_img = os.path.join(carpeta, f"pdf_img_{n}.jpg")
                    pil.save(ruta_img, "JPEG", quality=88)
                    imagenes.append({"id": f"pdf-{n}", "ruta": ruta_img, "pagina": num,
                                     "ancho": pil.width, "alto": pil.height,
                                     "contexto": (paginas_texto[num - 1] or "")[:3000]})
                except Exception:
                    continue
    return imagenes


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

    def __init__(self, carpeta):
        self.carpeta = carpeta
        self._imagenes = None          # se extraen una sola vez, aunque el Jefe pida reintentar
        self._descartadas_ia = 0

    def ejecutar(self, entrada, feedback=None):
        paginas = _extraer_plumber(entrada)                    # respaldo, siempre se calcula
        nota_img = ""
        if self._imagenes is None:
            try:
                self._imagenes = _extraer_imagenes_pdf(entrada, self.carpeta, paginas)
                self._descartadas_ia = 0
                if self._imagenes:
                    self._imagenes, self._descartadas_ia = _describir_imagenes(self._imagenes)
            except Exception as e:
                self._imagenes = []
                nota_img = f"No se pudieron extraer imágenes del PDF ({type(e).__name__}: {e})."
        imagenes = self._imagenes
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
        lineas.append("")
        if imagenes:
            lineas.append(f"Imágenes útiles del PDF ({len(imagenes)}; la IA descartó {self._descartadas_ia} por ser "
                          "logotipos, adornos o estar vacías), disponibles para el Ilustrador:")
            for im in imagenes:
                lineas.append(f"   {im['id']}: página {im['pagina']}, {im.get('tipo', '?')} — {im.get('descripcion', '')}")
        else:
            lineas.append(nota_img or "El PDF no trae imágenes aprovechables (o son páginas escaneadas).")

        return {
            "salida": {"texto": texto, "paginas": len(paginas), "fuente": fuente,
                       "texto_plumber": t_plumber, "solapamiento": solap, "imagenes_pdf": imagenes},
            "recibio": f"PDF: {os.path.basename(entrada)}",
            "entrego": (f"{len(paginas)} páginas, {len(texto)} caracteres de texto (leído con {fuente}) "
                        f"y {len(imagenes)} imágenes del PDF"),
            "razonamiento": (
                "Leí el PDF con pdfplumber (respaldo) y se lo mandé a Gemini para una transcripción más "
                "completa, que sirve también con PDFs escaneados o con tablas. "
                + (nota or "Usé la transcripción de Gemini porque coincide con la lectura de pdfplumber.")
            ),
            "detalle": "\n".join(lineas),
            "notas": f"{vacias} páginas sin texto extraíble para pdfplumber",
        }
