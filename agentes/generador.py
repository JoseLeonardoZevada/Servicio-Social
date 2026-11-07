"""Agente Generador PPTX: construye el archivo .pptx a partir del plan (sin IA)."""
import os

from pptx import Presentation
from pptx.util import Inches

from agentes.base import Agente
from config import TITULO_FUENTES
from pptx_utils import _sin_vineta


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
