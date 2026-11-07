"""Acceso a Wikimedia Commons (imágenes con licencia libre y autor)."""
import re

import requests

from config import USER_AGENT
from utils import _limpiar_html


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
