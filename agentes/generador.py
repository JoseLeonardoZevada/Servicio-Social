"""Agente Generador PPTX: construye el archivo .pptx a partir del plan; la IA escribe las notas del orador
en simultaneo (paralelo) por bloques."""

from concurrent.futures import ThreadPoolExecutor, as_completed
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
Eres un coach de presentaciones. Escribes el guion que el presentador dira en voz alta en cada diapositiva.
Usas SOLO informacion del documento (los puntos de los temas); no inventas datos ni cifras.
Responde UNICAMENTE con JSON valido.
"""


def _lista(nums):
    nums = [str(n) for n in nums]
    return nums[0] if len(nums) == 1 else ", ".join(nums[:-1]) + " y " + nums[-1]


def _pie_imagen(img):
    """Texto que va debajo de cada imagen, segun de donde salio."""
    fuente = img.get("fuente", "web")
    if fuente == "pdf":
        return f"Imagen: documento base, pag. {img.get('pagina_pdf', '?')}"
    if fuente == "ia":
        return "Ilustracion generada con IA (solo ilustrativa)"
    return f"Imagen: {_corto(img.get('autor', 'Autor'), 45)} · {img.get('licencia', 'CC')} · Wikimedia Commons"


class GeneradorPPTX(Agente):
    nombre = "Generador PPTX"
    rol = ("Construye el archivo .pptx (portada, contenido, imagenes y fuentes) y la IA escribe las notas del "
           "orador de cada diapositiva en simultaneo")

    def __init__(self, ruta_salida):
        self.ruta_salida = ruta_salida

    def _pedir_bloque_notas(self, bloque):
        prompt = (
            "Escribe las notas del orador (35 a 75 palabras en espanol natural para decir en voz alta) "
            "de cada diapositiva de este bloque, explicando sus ideas sin limitarte a leer las vinetas.\n"
            'Responde {"notas": [{"diapositiva": N, "guion": "..."}]}\n\n'
            "DIAPOSITIVAS:\n" + json.dumps(bloque, ensure_ascii=False)
        )
        try:
            d = parsear_json(llamar_llm(SYS_NOTAS, prompt))
            return {x["diapositiva"]: str(x["guion"]).strip()
                    for x in d.get("notas", [])
                    if isinstance(x, dict) and isinstance(x.get("diapositiva"), int) and str(x.get("guion", "")).strip()}
        except Exception:
            return {}

    def _notas(self, entrada):
        """Si el Jefe ya precalculo las notas en paralelo con el Ilustrador, las reutiliza; si no, las pide en paralelo."""
        if isinstance(entrada.get("notas_precalculadas"), dict) and entrada["notas_precalculadas"]:
            return entrada["notas_precalculadas"]
        slides, secciones, notas = entrada["slides"], entrada.get("secciones", []) or [], {}
        bloques = []
        for ini in range(0, len(slides), 8):
            bloque = []
            for k in range(ini, min(ini + 8, len(slides))):
                sl = slides[k]
                pts = []
                for num_t in sl.get("temas") or []:
                    if 1 <= num_t <= len(secciones):
                        pts.extend(secciones[num_t - 1].get("puntos") or [])
                bloque.append({
                    "diapositiva": k + 1,
                    "titulo": sl.get("title", ""),
                    "parrafo": sl.get("parrafo", ""),
                    "bullets": sl.get("bullets", []),
                    "puntos_tema": pts[:6],
                })
            bloques.append(bloque)

        if len(bloques) == 1:
            notas.update(self._pedir_bloque_notas(bloques[0]))
        elif bloques:
            with ThreadPoolExecutor(max_workers=min(3, len(bloques))) as pool:
                for res in pool.map(self._pedir_bloque_notas, bloques):
                    notas.update(res)
        return notas

    def ejecutar(self, entrada, feedback=None):
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)

        # 1) Portada
        s_port = prs.slides.add_slide(prs.slide_layouts[0])
        s_port.shapes.title.text = entrada.get("titulo", "Presentación")
        s_port.placeholders[1].text = entrada.get("subtitulo", "")

        notas = self._notas(entrada)
        n_imgs = 0
        creditos_web, pdf, ia = [], [], []
        detalle = [f"Diapositiva 1 (portada): {entrada.get('titulo', 'Presentación')}"]

        # 2) Diapositivas de contenido
        for idx, sl in enumerate(entrada["slides"], 1):
            slide = prs.slides.add_slide(prs.slide_layouts[1])
            slide.shapes.title.text = sl.get("title", f"Diapositiva {idx}")
            items = []
            t = sl.get("tipo", "bullets")
            par = (sl.get("parrafo") or "").strip()
            bullets = [str(b).strip() for b in (sl.get("bullets") or []) if str(b).strip()]
            if t == "parrafo":
                items.append((par or " ".join(bullets), False))
            elif t == "mixto":
                if par:
                    items.append((par, False))
                items.extend((b, True) for b in bullets)
            else:
                items.extend((b, True) for b in (bullets or [par]))
            _escribir(slide.placeholders[1].text_frame, items)

            img = sl.get("imagen")
            if img and os.path.exists(img.get("ruta", "")):
                pic = slide.shapes.add_picture(img["ruta"], Inches(7.5), Inches(2.0), width=Inches(4.5))
                caja = slide.shapes.add_textbox(Inches(7.5), Inches(6.2), Inches(4.5), Inches(0.6))
                caja.name = "Credito imagen"
                caja.text_frame.text = _pie_imagen(img)
                n_imgs += 1
                num_real = idx + 1
                if img.get("fuente") == "pdf":
                    pdf.append((num_real, img))
                elif img.get("fuente") == "ia":
                    ia.append(num_real)
                else:
                    creditos_web.append((num_real, img))

            guion = notas.get(idx, "")
            if guion:
                slide.notes_slide.notes_text_frame.text = guion
            detalle.append(f"Diapositiva {idx + 1} [{t}]: {sl.get('title', '')}")

        # 3) Diapositiva(s) de fuentes y creditos
        refs = [str(r).strip() for r in (entrada.get("referencias") or []) if str(r).strip()]
        lineas = [("Fuentes de información", False)]
        lineas.append((f"Documento base: {entrada.get('documento', 'PDF proporcionado')}", True))
        for r in refs:
            lineas.append((r, True))
        if creditos_web or pdf or ia:
            lineas.append(("Créditos de imágenes", False))
            for num_real, im in creditos_web:
                lineas.append((f"Diap. {num_real}: «{_corto(im.get('titulo', ''), 45)}» — "
                               f"{_corto(im.get('autor', ''), 35)} ({im.get('licencia', '')}) · {im.get('pagina', '')}", True))
            if pdf:
                lineas.append((f"Diap. {_lista([n for n, _ in pdf])}: imágenes tomadas del documento base "
                               f"(págs. {_lista(sorted({im['pagina_pdf'] for _, im in pdf}))}).", True))
            if ia:
                lineas.append((f"Diap. {_lista(ia)}: ilustraciones generadas con IA, solo ilustrativas.", True))

        trozos = [lineas[k:k + MAX_CREDITOS_POR_DIAPOSITIVA]
                  for k in range(0, len(lineas), MAX_CREDITOS_POR_DIAPOSITIVA)] or [[("Fuentes de información", False)]]
        for k, trozo in enumerate(trozos, 1):
            s_f = prs.slides.add_slide(prs.slide_layouts[1])
            sufijo = f" ({k}/{len(trozos)})" if len(trozos) > 1 else ""
            s_f.shapes.title.text = f"{TITULO_FUENTES}{sufijo}"
            _escribir(s_f.placeholders[1].text_frame, trozo)
            detalle.append(f"Diapositiva {len(prs.slides)} (fuentes) «{s_f.shapes.title.text}»: {len(refs)} referencias")

        prs.save(self.ruta_salida)
        total = len(prs.slides)
        return {
            "salida": {
                "ruta": self.ruta_salida,
                "n_slides": total,
                "n_imagenes": n_imgs,
                "n_notas": len(notas),
                "n_contenido": len(entrada["slides"]),
                "titulo": entrada.get("titulo", "Presentación"),
                "tema": entrada.get("tema_general", ""),
            },
            "recibio": f"Plan de {len(entrada['slides'])} diapositivas, {n_imgs} imágenes y referencias",
            "entrego": f"Archivo {os.path.basename(self.ruta_salida)} con {total} diapositivas "
                       f"(portada + contenido + fuentes y créditos) y {len(notas)} notas del orador",
            "razonamiento": (
                "Creé la portada, luego una diapositiva por elemento del plan: texto corrido sin viñetas "
                "en las de tipo 'parrafo', viñetas en las de tipo 'bullets', y ambos en 'mixto'. "
                "Inserté cada imagen con su crédito y cerré con la diapositiva de fuentes, armada solo con "
                "datos verificados (documento, referencias presentes en el PDF y créditos de las imágenes). "
                "La IA escribió las notas del orador en simultáneo. Aún sin estilo: eso lo hace el Decorador."
            ),
            "detalle": "\n".join(detalle),
            "notas": "",
        }
