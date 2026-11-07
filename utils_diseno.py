"""Utilidades de diseño: ¿cabe el texto en su caja? y contraste de colores."""
import math
import re


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
