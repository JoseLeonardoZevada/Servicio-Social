"""Agente Decorador: la IA propone el diseño y el código lo aplica y lo acomoda."""
import json

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Pt, Inches

from utils import (_a_hex, Agente, _colocar, con_feedback, _estilo, _forma, _hex_a_rgb, _imagenes,
                   llamar_llm, _mejor_tam, parsear_json, _problemas_contraste, _rgb, _sin_vineta)
from config import TITULO_FUENTES


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


SYS_DECORADOR_LAYOUT = """
Eres un diseñador de presentaciones con criterio visual. Decides cómo se distribuye cada diapositiva
para que la presentación tenga ritmo y se lea bien. Responde ÚNICAMENTE con JSON válido.
"""

DISTRIBUCIONES = ("imagen_derecha", "imagen_izquierda", "imagen_grande_derecha", "imagen_grande_izquierda")
FONDOS = ("claro", "tinte")
MAX_PALABRAS_IMAGEN_GRANDE = 60


class Decorador(Agente):
    nombre = "Decorador"
    rol = "La IA propone el diseño (colores, formas, tipografía) y el código lo aplica y lo acomoda"

    def _proponer_layouts(self, prs, feedback):
        """La IA elige para cada diapositiva de contenido el lado y el tamaño de la imagen y el fondo. El código
        veta lo que no cabe (imagen grande con mucho texto) y usa la distribución por defecto si la IA falla."""
        info, palabras = [], {}
        for i, sl in enumerate(prs.slides):
            if i == 0 or sl.shapes.title.text.startswith(TITULO_FUENTES):
                continue
            pics = _imagenes(sl)
            n = sum(len(par.text.split()) for par in sl.placeholders[1].text_frame.paragraphs)
            palabras[i] = n
            info.append({"diapositiva": i, "titulo": sl.shapes.title.text, "palabras": n,
                         "imagen": ({"relacion_ancho_alto": round(pics[0].image.size[0] / pics[0].image.size[1], 2)}
                                    if pics else None)})
        prompt = (
            "Decide la distribución de cada diapositiva de contenido. Para las que tienen imagen elige "
            f"'distribucion' entre {list(DISTRIBUCIONES)}: alterna los lados para dar ritmo; usa las 'grande' cuando "
            "el texto es corto (hasta unas 45 palabras) o la imagen es muy visual, porque dejan menos espacio al "
            "texto. Para todas elige 'fondo' ('claro' o 'tinte') alternando para dar ritmo visual.\n"
            'Responde {"razonamiento": "<criterio>", "diapositivas": [{"diapositiva": N, "distribucion": "...", '
            '"fondo": "claro|tinte"}]}\n\nDIAPOSITIVAS:\n' + json.dumps(info, ensure_ascii=False)
        )
        layouts, vetos, razon, origen = {}, [], "", "IA"
        try:
            d = parsear_json(llamar_llm(SYS_DECORADOR_LAYOUT, con_feedback(prompt, feedback)))
            razon = str(d.get("razonamiento", ""))
            por_num = {x["diapositiva"]: x for x in d.get("diapositivas", [])
                       if isinstance(x, dict) and isinstance(x.get("diapositiva"), int)}
        except Exception as e:
            por_num, origen = {}, f"respaldo ({type(e).__name__})"
        for i, n in palabras.items():
            x = por_num.get(i) or {}
            dist = x.get("distribucion") if x.get("distribucion") in DISTRIBUCIONES else "imagen_derecha"
            if "grande" in dist and n > MAX_PALABRAS_IMAGEN_GRANDE:
                vetos.append(f"diap. {i}: imagen grande con {n} palabras → normal")
                dist = dist.replace("_grande", "")
            layouts[i] = {"distribucion": dist, "fondo": x.get("fondo") if x.get("fondo") in FONDOS else None}
        return layouts, origen, razon, vetos

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
        layouts, origen_layout, razon_layout, vetos_layout = self._proponer_layouts(prs, feedback)
        tamanos = []
        for i, slide in enumerate(prs.slides):
            slide.background.fill.solid()
            if i == 0:
                self._portada(slide, diseno, W, H)
            else:
                lay = layouts.get(i, {})
                tam = self._contenido(slide, diseno, W, H, i, es_fuentes=slide.shapes.title.text.startswith(TITULO_FUENTES),
                                      layout=lay)
                extra = f" (con imagen: {lay.get('distribucion', 'imagen_derecha')})" if _imagenes(slide) else ""
                tamanos.append(f"Diap. {i + 1}: texto a {tam} pt{extra}, fondo {lay.get('fondo') or 'automático'}")
        prs.save(ruta)

        pal = diseno["paleta"]
        detalle = (
            f"Origen del diseño: {origen}\n"
            "Paleta: " + ", ".join(f"{k} {_a_hex(v)}" for k, v in pal.items()) + "\n"
            f"Formas: {diseno['estilo_formas']} · Barra: {diseno['barra']} · "
            f"Fondos alternados: {'sí' if diseno['alternar_fondos'] else 'no'}\n"
            f"Tipografía: títulos {diseno['fuente_titulo']}, cuerpo {diseno['fuente_cuerpo']}\n"
            f"Distribución por diapositiva decidida por: {origen_layout}. {razon_layout}\n"
            + ("Vetos del código: " + "; ".join(vetos_layout) + "\n" if vetos_layout else "")
            + "Las imágenes llevan marco de color y su crédito debajo; el tamaño de letra se ajusta para que el "
            "texto quepa; las diapositivas de fuentes usan letra pequeña.\n"
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

    def _contenido(self, slide, d, W, H, numero, es_fuentes=False, layout=None):
        p, estilo = d["paleta"], d["estilo_formas"]
        pref = (layout or {}).get("fondo")
        if es_fuentes:
            alterna = False
        elif pref in FONDOS:
            alterna = pref == "tinte"
        else:
            alterna = d["alternar_fondos"] and numero % 2 == 0
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
            dist = (layout or {}).get("distribucion", "imagen_derecha")
            img_w, img_h = (Inches(6.0), Inches(4.2)) if "grande" in dist else (Inches(4.9), Inches(3.5))
            if dist.endswith("izquierda"):
                img_x = Inches(0.9)
            else:
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
            if dist.endswith("izquierda"):
                cuerpo_x = img_x + img_w + Inches(0.4)
                ancho_cuerpo = W - Inches(0.9) - cuerpo_x
            else:
                cuerpo_x = Inches(0.9)
                ancho_cuerpo = img_x - Inches(0.9) - Inches(0.4)
        else:
            cuerpo_x, ancho_cuerpo = Inches(0.9), W - Inches(1.8)
        _colocar(cuerpo, cuerpo_x, top_cuerpo, ancho_cuerpo, alto_cuerpo)

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
