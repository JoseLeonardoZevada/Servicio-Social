"""Funciones auxiliares para dibujar y escribir en diapositivas con python-pptx."""
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from pptx.util import Pt


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
