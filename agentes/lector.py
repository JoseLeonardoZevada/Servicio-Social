"""Agente Lector: extrae texto rapido con pdfplumber (usa Gemini OCR si es escaneado),
extrae imagenes cuidando los 512 MB de RAM de Render y construye el indice RAG."""

import gc
import os
import re

import pdfplumber
from PIL import ImageStat

from utils import Agente, crear_indice_rag, llamar_llm, parsear_json

SYS_LECTOR = """
Eres un transcriptor de documentos. Copias el texto de un PDF de forma fiel, en orden,
sin resumir, sin corregir y sin inventar nada. Incluye titulos, tablas (como texto) y
la bibliografia. Si una pagina es una imagen escaneada, haz OCR. Si hay una figura
importante, describela en una linea entre corchetes: [Figura: ...].
Responde solo con el texto transcrito.
"""


def _extraer_plumber(ruta):
    paginas = []
    with pdfplumber.open(ruta) as pdf:
        for p in pdf.pages:
            paginas.append(p.extract_text() or "")
    return paginas


SYS_LECTOR_IMG = """
Eres un editor grafico. Describes imagenes extraidas de un PDF y decides si sirven para ilustrar
una presentacion. Responde UNICAMENTE con JSON valido.
"""


def _describir_imagenes(imagenes):
    """La IA mira las imagenes (de 4 en 4), las describe y descarta logotipos, adornos e imagenes vacias."""
    utiles, descartadas = [], 0
    for ini in range(0, len(imagenes), 4):
        lote = imagenes[ini:ini + 4]
        partes = ["Estas imagenes salieron de un PDF. Para cada una devuelve 'descripcion' (una frase en espanol "
                  "de que muestra), 'tipo' (foto, diagrama, grafica, tabla u otro) y 'util' (true si sirve para "
                  "ilustrar contenido; false si es un logotipo, adorno, fondo, imagen casi vacia o ilegible). "
                  'Responde {"imagenes": [{"id": "...", "descripcion": "...", "tipo": "...", "util": true}]}']
        for im in lote:
            partes.append(f"Imagen {im['id']} (pagina {im['pagina']}):")
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


MAX_IMAGENES_PDF = 16       # reducido para cuidar RAM y llamadas a Gemini en Render Free
MAX_IMG_POR_PAGINA = 2


def _huella(im):
    """Huella visual de 16x16 para detectar imagenes repetidas (logos, encabezados en cada pagina)."""
    g = im.convert("L").resize((16, 16))
    px = list(g.getdata())
    media = sum(px) / len(px)
    return "".join("1" if v > media else "0" for v in px)


def _extraer_imagenes_pdf(ruta, carpeta, paginas_texto):
    """Recorta de cada pagina las imagenes que trae el PDF cuidando la memoria RAM de Render (100 DPI)."""
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
                    continue
                if not 0.03 <= (w * h) / area_pagina <= 0.85:
                    continue
                if not 0.2 <= w / h <= 5:
                    continue
                cajas.append((w * h, (x0, top, x1, bottom)))
            for _, caja in sorted(cajas, reverse=True)[:MAX_IMG_POR_PAGINA]:
                try:
                    pil = page.crop(caja).to_image(resolution=100).original.convert("RGB")
                    if pil.width < 220 or pil.height < 140:
                        continue
                    if ImageStat.Stat(pil.convert("L")).stddev[0] < 8:
                        continue
                    h16 = _huella(pil)
                    if h16 in huellas:
                        continue
                    huellas.add(h16)
                    n = len(imagenes) + 1
                    ruta_img = os.path.join(carpeta, f"pdf_img_{n}.jpg")
                    pil.save(ruta_img, "JPEG", quality=85)
                    ancho, alto = pil.width, pil.height
                    pil.close()
                    imagenes.append({"id": f"pdf-{n}", "ruta": ruta_img, "pagina": num,
                                     "ancho": ancho, "alto": alto,
                                     "contexto": (paginas_texto[num - 1] or "")[:2000]})
                except Exception:
                    continue
            page.flush_cache()
    gc.collect()
    return imagenes


def _solapamiento(a, b):
    """Fraccion de las palabras (5+ letras) de 'a' que tambien aparecen en 'b'."""
    pa = set(re.findall(r"\w{5,}", a.lower()))
    if not pa:
        return 1.0
    pb = set(re.findall(r"\w{5,}", b.lower()))
    return len(pa & pb) / len(pa)


class Lector(Agente):
    nombre = "Lector"
    rol = "Lee el PDF con pdfplumber (o Gemini si viene escaneado), extrae imagenes e indexa fragmentos para RAG"

    def __init__(self, carpeta):
        self.carpeta = carpeta
        self._imagenes = None
        self._descartadas_ia = 0

    def ejecutar(self, entrada, feedback=None):
        print(f"[LECTOR] Extrayendo texto de {os.path.basename(entrada)}...", flush=True)
        paginas = _extraer_plumber(entrada)
        t_plumber = "\n\n".join(t for t in paginas if t)
        vacias = sum(1 for t in paginas if not t or len(t.strip()) < 40)

        nota_img = ""
        if self._imagenes is None:
            try:
                print("[LECTOR] Buscando imagenes dentro del PDF...", flush=True)
                self._imagenes = _extraer_imagenes_pdf(entrada, self.carpeta, paginas)
                self._descartadas_ia = 0
                if self._imagenes:
                    print(f"[LECTOR] Describiendo {len(self._imagenes)} imagenes con IA...", flush=True)
                    self._imagenes, self._descartadas_ia = _describir_imagenes(self._imagenes)
            except Exception as e:
                self._imagenes = []
                nota_img = f"No se pudieron extraer imagenes del PDF ({type(e).__name__}: {e})."
        imagenes = self._imagenes

        fuente, t_ia, nota = "pdfplumber", "", ""
        # Si pdfplumber ya extrajo texto solido en la mayoria de las paginas, lo usamos directo sin gastar 2 min en OCR
        if len(t_plumber.strip()) >= 400 and vacias <= len(paginas) * 0.4 and not feedback:
            nota = "El PDF tiene texto digital nativo completo; se uso pdfplumber directo e indexacion RAG por pagina."
        elif feedback:
            nota = "El Jefe rechazo la lectura anterior, asi que use el respaldo (pdfplumber)."
        else:
            try:
                print("[LECTOR] PDF escaneado o con poco texto; transcribiendo con Gemini...", flush=True)
                if os.path.getsize(entrada) > 12_000_000:
                    raise ValueError("el PDF es muy pesado para enviarlo a Gemini")
                with open(entrada, "rb") as f:
                    datos = f.read()
                t_ia = llamar_llm(SYS_LECTOR, ["Transcribe todo el texto de este PDF.",
                                               {"mime_type": "application/pdf", "data": datos}])
                fuente = "Gemini"
            except Exception as e:
                nota = f"Gemini no pudo transcribir el PDF ({type(e).__name__}: {e}); use pdfplumber."

        texto = t_ia if fuente == "Gemini" else t_plumber
        solap = None
        if fuente == "Gemini" and len(t_plumber.strip()) >= 200:
            solap = _solapamiento(t_plumber, t_ia)

        indice_rag = crear_indice_rag(paginas, texto)
        print(f"[LECTOR] Listo: {len(paginas)} pags, {len(indice_rag)} chunks RAG, {len(imagenes)} imgs.", flush=True)

        lineas = [f"Fuente usada: {fuente}"]
        if nota:
            lineas.append(nota)
        if solap is not None:
            lineas.append(f"Coincidencia de palabras con pdfplumber: {round(solap * 100)}%")
        lineas.append(f"Indice RAG creado en memoria: {len(indice_rag)} fragmentos por pagina listos para consulta.")
        lineas.append("")
        lineas.append("Lectura por pagina (pdfplumber):")
        for i, t in enumerate(paginas[:12], 1):
            muestra = t[:200].replace("\n", " ").strip()
            lineas.append(f"Pag. {i} ({len(t)} caracteres): {muestra}..." if t else f"Pag. {i}: sin texto extraible")
        if len(paginas) > 12:
            lineas.append(f"... y {len(paginas) - 12} paginas mas")
        lineas.append("")
        if imagenes:
            lineas.append(f"Imagenes utiles del PDF ({len(imagenes)}; la IA descarto {self._descartadas_ia} por ser "
                          "logotipos, adornos o estar vacias), disponibles para el Ilustrador:")
            for im in imagenes:
                lineas.append(f"   {im['id']}: pagina {im['pagina']}, {im.get('tipo', '?')} — {im.get('descripcion', '')}")
        else:
            lineas.append(nota_img or "El PDF no trae imagenes aprovechables (o son paginas escaneadas).")

        return {
            "salida": {"texto": texto, "paginas": len(paginas), "fuente": fuente,
                       "texto_plumber": t_plumber, "solapamiento": solap, "imagenes_pdf": imagenes,
                       "indice_rag": indice_rag},
            "recibio": f"PDF: {os.path.basename(entrada)}",
            "entrego": (f"{len(paginas)} paginas, {len(texto)} caracteres de texto (leido con {fuente}), "
                        f"{len(indice_rag)} fragmentos RAG y {len(imagenes)} imagenes del PDF"),
            "razonamiento": (
                f"Extraje el texto del PDF ({fuente}) y construi un indice RAG de {len(indice_rag)} fragmentos "
                "para recuperar informacion precisa por tema. " + nota
            ),
            "detalle": "\n".join(lineas),
            "notas": f"{vacias} paginas con poco texto",
        }