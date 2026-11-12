"""Agente Generador PPTX: construye el archivo .pptx a partir del plan; la IA escribe las notas del orador."""
import json
import os

from pptx import Presentation
from pptx.util import Inches

from utils import Agente, llamar_llm, parsear_json, _sin_vineta
from config import TITULO_FUENTES


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


MAX_CREDITOS_POR_DIAPOSITIVA = 7

SYS_NOTAS = """
Eres un coach de presentaciones. Escribes el guion que el presentador dirá en voz alta en cada diapositiva.
Usas SOLO información del documento (los puntos de los temas); no inventas datos ni cifras.
Responde ÚNICAMENTE con JSON válido.
"""


def _lista(nums):
    nums = [str(n) for n in nums]
    return nums[0] if len(nums) == 1 else ", ".join(nums[:-1]) + " y " + nums[-1]


def _pie_imagen(img):
    """Texto que va debajo de cada imagen, según de dónde salió."""
    fuente = img.get("fuente", "web")
    if fuente == "pdf":
        return f"Imagen: documento base, pág. {img.get('pagina_pdf', '?')}"
    if fuente == "ia":
        return "Ilustración generada con IA (solo ilustrativa)"
    return f"Imagen: {_corto(img['autor'], 45)} · {img['licencia']} · Wikimedia Commons"


class GeneradorPPTX(Agente):
    nombre = "Generador PPTX"
    rol = ("Construye el archivo .pptx (portada, contenido, imágenes y fuentes) y la IA escribe las notas del "
           "orador de cada diapositiva")

    def __init__(self, ruta_salida):
        self.ruta_salida = ruta_salida

    def _notas(self, entrada):
        """La IA escribe, de 8 en 8 diapositivas, lo que el presentador dirá en voz alta (notas del orador)."""
        slides, secciones, notas = entrada["slides"], entrada.get("secciones", []) or [], {}
        for ini in range(0, len(slides), 8):
            bloque = []
            for k in range(ini, min(ini + 8, len(slides))):
                sl = slides[k]
                temas = [t for t in (sl.get("temas") or []) if isinstance(t, int) and 1 <= t <= len(secciones)]
                puntos = [pt for t in temas[:2] for pt in (secciones[t - 1].get("puntos") or [])][:8]
                bloque.append({"diapositiva": k + 1, "titulo": sl.get("title", ""),
                               "texto": " ".join([sl.get("parrafo") or ""] + list(sl.get("bullets") or [])).strip(),
                               "puntos_del_tema": puntos})
            prompt = ("Escribe el guion del presentador de cada diapositiva: 2 a 4 frases (40 a 90 palabras) en "
                      "español, con tono natural para decirlas en voz alta, que amplíen la diapositiva con los "
                      "datos de 'puntos_del_tema' sin repetirla palabra por palabra y sin inventar nada. "
                      'Responde {"notas": [{"diapositiva": N, "nota": "..."}]}\n\nDIAPOSITIVAS:\n'
                      + json.dumps(bloque, ensure_ascii=False))
            try:
                for x in parsear_json(llamar_llm(SYS_NOTAS, prompt)).get("notas", []):
                    if isinstance(x, dict) and isinstance(x.get("diapositiva"), int) and str(x.get("nota", "")).strip():
                        notas[x["diapositiva"]] = str(x["nota"]).strip()
            except Exception:
                continue
        return notas

    def ejecutar(self, entrada, feedback=None):
        notas = self._notas(entrada)
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
            if notas.get(i):
                slide.notes_slide.notes_text_frame.text = notas[i]

            marca = ""
            img = info.get("imagen")
            if img:
                slide.shapes.add_picture(img["ruta"], Inches(8), Inches(2), width=Inches(4))
                cap = slide.shapes.add_textbox(Inches(8), Inches(5.5), Inches(4), Inches(0.5))
                cap.name = "Credito imagen"
                cap.text_frame.text = _pie_imagen(img)
                n_imgs += 1
                creditos.append((i, img))
                marca = " + imagen"
            detalle.append(f"Diapositiva {i + 1} [{tipo}]: {info.get('title', '')}{marca}")

        # Diapositivas finales: fuentes de información y créditos de imágenes (pueden ser varias)
        items = [("Fuentes de información", False),
                 (f"Documento base: {entrada['titulo']} ({entrada.get('documento', 'PDF')})", False)]
        refs = entrada.get("referencias", [])[:6]
        if refs:
            items += [(f"Citado en el documento: {_corto(r, 150)}", False) for r in refs]
        else:
            items.append(("El documento base no cita referencias bibliográficas.", False))
        paginas_finales = []
        if not creditos:
            items += [("Créditos de imágenes", False), ("Esta presentación no incluye imágenes de terceros.", False)]
        paginas_finales.append((TITULO_FUENTES, items))

        if creditos:
            lineas = []
            for n, img in creditos:
                if img.get("fuente", "web") == "web":
                    titulo_archivo = _corto(img["titulo"], 60)
                    lineas.append(_corto(f"Diap. {n}: «{titulo_archivo}» — {img['autor']}, {img['licencia']}. "
                                         f"{img['pagina']}", 220))
            pdf = [(n, img) for n, img in creditos if img.get("fuente") == "pdf"]
            ia = [n for n, img in creditos if img.get("fuente") == "ia"]
            if pdf:
                lineas.append(f"Diap. {_lista([n for n, _ in pdf])}: imágenes tomadas del documento base "
                              f"(págs. {_lista(sorted({im['pagina_pdf'] for _, im in pdf}))}).")
            if ia:
                lineas.append(f"Diap. {_lista(ia)}: ilustraciones generadas con IA, solo ilustrativas.")
            trozos = [lineas[k:k + MAX_CREDITOS_POR_DIAPOSITIVA]
                      for k in range(0, len(lineas), MAX_CREDITOS_POR_DIAPOSITIVA)]
            for k, trozo in enumerate(trozos, 1):
                titulo = f"{TITULO_FUENTES}: imágenes" + (f" ({k}/{len(trozos)})" if len(trozos) > 1 else "")
                paginas_finales.append((titulo, [("Créditos de imágenes", False)] + [(t, False) for t in trozo]))

        for titulo, its in paginas_finales:
            fs = prs.slides.add_slide(prs.slide_layouts[1])
            fs.shapes.title.text = titulo
            _escribir(fs.placeholders[1].text_frame, its)
        n_final = len(entrada["slides"]) + 1
        for k, (titulo, _) in enumerate(paginas_finales, 1):
            detalle.append(f"Diapositiva {n_final + k} (fuentes) «{titulo}»: "
                           + (f"{len(refs)} referencias" if k == 1 else "créditos de imágenes"))

        prs.save(self.ruta_salida)
        total = len(entrada["slides"]) + 1 + len(paginas_finales)
        return {
            "salida": {
                "ruta": self.ruta_salida, "n_slides": total, "n_imagenes": n_imgs,
                "titulo": entrada["titulo"], "tema": entrada.get("tema_general", ""),
                "n_notas": len(notas), "n_contenido": len(entrada["slides"]),
            },
            "recibio": f"Plan de {len(entrada['slides'])} diapositivas, {n_imgs} imágenes y referencias",
            "entrego": f"Archivo {os.path.basename(self.ruta_salida)} con {total} diapositivas "
                       f"(portada + contenido + fuentes y créditos) y {len(notas)} notas del orador",
            "razonamiento": (
                "Creé la portada, luego una diapositiva por elemento del plan: texto corrido sin viñetas "
                "en las de tipo 'parrafo', viñetas en las de tipo 'bullets', y ambos en 'mixto'. "
                "Inserté cada imagen con su crédito y cerré con la diapositiva de fuentes, armada solo con "
                "datos verificados (documento, referencias presentes en el PDF y créditos de las imágenes). "
                "La IA escribió las notas del orador de cada diapositiva con los puntos del tema. "
                "Aún sin estilo: eso lo hace el Decorador."
            ),
            "detalle": "\n".join(detalle),
            "notas": "",
        }
