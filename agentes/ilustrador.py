"""Agente Ilustrador: pone imagen en las diapositivas que el Seccionador marcó con 'lleva_imagen'.
Orden de preferencia para cada una: 1) imágenes del propio PDF (las entrega el Lector),
2) búsqueda en Wikimedia Commons (licencia libre, con autor), 3) ilustración generada con IA."""
import io
import json
import os
import re

from PIL import Image

from utils import (Agente, buscar_commons, con_feedback, descargar_imagen, generar_imagen_ia,
                   llamar_llm, parsear_json)


SYS_ILUSTRADOR_PLAN = """
Eres un editor gráfico que prepara las imágenes de una presentación. Para cada diapositiva que se te
indica, escribes una búsqueda precisa para un banco de fotos y una descripción para generar una
ilustración. Las imágenes son solo ilustrativas: sencillas, sin texto, sin personas reales
reconocibles y sin contenido violento o gráfico.
Responde ÚNICAMENTE con JSON válido.
"""

SYS_ILUSTRADOR = """
Eres un editor gráfico muy estricto. Eliges imágenes para diapositivas, pero solo
si muestran de forma clara y directa el tema. Si hay duda, no eliges ninguna.
Responde ÚNICAMENTE con JSON válido.
"""


def _texto(sl):
    return " ".join(x for x in (sl.get("title", ""), sl.get("parrafo") or "",
                                " ".join(sl.get("bullets") or [])) if x)


def _palabras(texto):
    """Palabras de 5+ letras recortadas a 6 (así 'imagen' e 'imágenes' coinciden más o menos)."""
    return {w[:6] for w in re.findall(r"[a-záéíóúñü]{5,}", texto.lower())}


class Ilustrador(Agente):
    nombre = "Ilustrador"
    rol = ("Pone imagen en las diapositivas que el Seccionador marcó. La IA decide para cada una la fuente "
           "(imagen del PDF, Wikimedia Commons o ilustración generada), qué buscar y verifica viendo cada imagen")

    def __init__(self, carpeta):
        self.carpeta = carpeta
        self._feedback = None

    # ---- 1) la IA decide, para cada diapositiva marcada, de dónde sale su imagen y qué buscar ----
    def _planear(self, slides, marcadas, imagenes_pdf, feedback):
        resumen = [{"diapositiva": i, "titulo": slides[i - 1].get("title", ""),
                    "texto": _texto(slides[i - 1])[:240]} for i in marcadas]
        catalogo = [{"id": im["id"], "pagina": im["pagina"], "tipo": im.get("tipo", ""),
                     "descripcion": im.get("descripcion", "")} for im in imagenes_pdf]
        prompt = (
            f"Estas son las diapositivas de una presentación que llevan imagen ({len(marcadas)} de {len(slides)}). "
            "Para cada una decide la FUENTE de su imagen:\n"
            "- 'pdf': una de las imágenes del PDF de abajo (pon su id en 'imagen_pdf'). Úsala solo si de verdad "
            "ilustra esa diapositiva, y no repitas una imagen en dos diapositivas.\n"
            "- 'web': una foto de un banco de imágenes libres.\n"
            "- 'ia': una ilustración generada con IA.\n"
            "El orden de preferencia es pdf, luego web, luego ia; elige otra solo si lo justificas en 'por_que'.\n"
            "Además, para cada una escribe:\n"
            "- 'consulta': búsqueda corta EN INGLÉS (2 a 5 palabras) para Wikimedia Commons, sobre algo "
            "concreto y fotografiable que ilustre su contenido.\n"
            "- 'prompt_ia': descripción EN INGLÉS (10 a 25 palabras) de una ilustración sencilla y "
            "conceptual del mismo tema, para generarla con IA (sin texto ni letras en la imagen).\n"
            'Responde {"razonamiento": "<cómo decidiste>", "imagenes": [{"diapositiva": N, "fuente": "pdf|web|ia", '
            '"imagen_pdf": "<id o null>", "consulta": "...", "prompt_ia": "...", "por_que": "<una frase>"}]}\n\n'
            "IMÁGENES DEL PDF:\n" + (json.dumps(catalogo, ensure_ascii=False) if catalogo else "(ninguna)")
            + "\n\nDIAPOSITIVAS:\n" + json.dumps(resumen, ensure_ascii=False)
        )
        d = parsear_json(llamar_llm(SYS_ILUSTRADOR_PLAN, con_feedback(prompt, feedback)))
        por_num = {x["diapositiva"]: x for x in d.get("imagenes", [])
                   if isinstance(x, dict) and isinstance(x.get("diapositiva"), int)}
        ids_pdf = {im["id"] for im in imagenes_pdf}
        plan, completadas, ya_asignadas = [], [], set()
        for i in marcadas:
            x = por_num.get(i) or {}
            consulta = " ".join(str(x.get("consulta", "")).split()[:6])
            prompt_ia = str(x.get("prompt_ia", "")).strip()
            if not consulta or not prompt_ia:           # la IA se saltó esta diapositiva: se completa con su título
                titulo = slides[i - 1].get("title", "")
                consulta = consulta or " ".join(titulo.split()[:5]) or "illustration"
                prompt_ia = prompt_ia or f"simple conceptual illustration about {titulo}"
                completadas.append(i)
            fuente = x.get("fuente") if x.get("fuente") in ("pdf", "web", "ia") else ("pdf" if ids_pdf else "web")
            img_pdf = x.get("imagen_pdf") if x.get("imagen_pdf") in ids_pdf and x.get("imagen_pdf") not in ya_asignadas else None
            if img_pdf:
                ya_asignadas.add(img_pdf)
            plan.append({"diapositiva": i, "fuente": fuente, "imagen_pdf": img_pdf, "consulta": consulta,
                         "prompt_ia": prompt_ia, "por_que": str(x.get("por_que", ""))})
        return str(d.get("razonamiento", "")), plan, completadas

    # ---- la IA mira las candidatas y elige una (o ninguna) ----
    def _elegir(self, slide, candidatas, origen):
        partes = [
            f"Diapositiva: «{slide.get('title', '')}». Contenido: {_texto(slide)}\n"
            f"Estas imágenes vienen de {origen}. Elige la que muestre de forma clara y "
            "directa el tema de la diapositiva. Descarta las que solo se relacionen vagamente, "
            "logotipos, capturas con texto ilegible o imágenes que no se entiendan. Si NINGUNA "
            'cumple, devuelve null. Responde {"elegida": <número de imagen o null>, '
            '"razon": "<una frase: qué se ve y por qué se relaciona con el tema>"}'
        ]
        for k, c in enumerate(candidatas):
            partes.append(f"Imagen {k} ({c['titulo']}):")
            partes.append({"mime_type": c["mime"], "data": c["bytes"]})
        partes[0] = con_feedback(partes[0], self._feedback)
        d = parsear_json(llamar_llm(SYS_ILUSTRADOR, partes))
        k = d.get("elegida")
        if isinstance(k, int) and not isinstance(k, bool) and 0 <= k < len(candidatas):
            return candidatas[k], d.get("razon", "")
        return None, d.get("razon", "ninguna imagen ilustraba el tema con claridad")

    # ---- fuente 1: imágenes del PDF ----
    def _desde_pdf(self, sl, imagenes_pdf, usadas, preferida=None):
        palabras = _palabras(_texto(sl))

        def afinidad(im):                                # cuánto se parece su página y su descripción a la diapositiva
            return len(palabras & _palabras(im.get("contexto", "") + " " + im.get("descripcion", ""))) / max(1, len(palabras))

        libres = [im for im in imagenes_pdf if im["id"] not in usadas and os.path.exists(im["ruta"])]
        libres.sort(key=lambda im: (im["id"] != preferida, -afinidad(im)))      # primero la que eligió la IA
        cands = []
        for im in libres[:3]:
            with open(im["ruta"], "rb") as f:
                cands.append({"titulo": f"figura del PDF, página {im['pagina']}; texto de esa página: "
                                        f"{(im.get('contexto') or '')[:200]!r}",
                              "mime": "image/jpeg", "bytes": f.read(), "im": im})
        if not cands:
            return None, "el PDF no tiene imágenes disponibles"
        elegida, razon = self._elegir(sl, cands, "el propio PDF del documento")
        if not elegida:
            return None, razon
        im = elegida["im"]
        return {"fuente": "pdf", "id": im["id"], "ruta": im["ruta"], "titulo": f"Figura del PDF (pág. {im['pagina']})",
                "autor": "", "licencia": "", "pagina": "", "pagina_pdf": im["pagina"], "razon": razon,
                "ancho": im["ancho"], "alto": im["alto"]}, razon

    # ---- fuente 2: Wikimedia Commons ----
    def _desde_web(self, sl, i, consultas, titulos_usados):
        ultima = "sin resultados con licencia libre y autor"
        for q in consultas:
            cands = [c for c in buscar_commons(q) if c["titulo"] not in titulos_usados]
            bajadas = []
            for c in cands:
                if len(bajadas) >= 3:
                    break
                try:
                    datos, mime = descargar_imagen(c["thumb"])
                    bajadas.append({**c, "bytes": datos, "mime": mime})
                except Exception:
                    continue
            if not bajadas:
                continue
            elegida, razon = self._elegir(sl, [{**b, "titulo": b["titulo"]} for b in bajadas], "Wikimedia Commons")
            if not elegida:
                ultima = razon
                continue
            ext = ".png" if elegida["mime"] == "image/png" else ".jpg"
            ruta = os.path.join(self.carpeta, f"imagen_{i + 1}{ext}")
            with open(ruta, "wb") as f:
                f.write(elegida["bytes"])
            ancho, alto = Image.open(io.BytesIO(elegida["bytes"])).size
            return {"fuente": "web", "id": "web:" + elegida["titulo"], "ruta": ruta, "titulo": elegida["titulo"],
                    "autor": elegida["autor"], "licencia": elegida["licencia"], "pagina": elegida["pagina"],
                    "razon": razon, "ancho": ancho, "alto": alto}, razon
        return None, ultima

    # ---- fuente 3: ilustración generada con IA ----
    def _generar(self, i, prompt_ia):
        datos, _ = generar_imagen_ia(prompt_ia)
        ruta = os.path.join(self.carpeta, f"imagen_{i + 1}_ia.jpg")
        with open(ruta, "wb") as f:
            f.write(datos)
        ancho, alto = Image.open(io.BytesIO(datos)).size
        return {"fuente": "ia", "id": f"ia:{i + 1}", "ruta": ruta, "titulo": "Ilustración generada con IA",
                "autor": "", "licencia": "", "pagina": "", "prompt": prompt_ia,
                "razon": "Ilustración generada con IA a partir del contenido de la diapositiva (solo ilustrativa)",
                "ancho": ancho, "alto": alto}

    def ejecutar(self, entrada, feedback=None):
        slides = [dict(sl) for sl in entrada["slides"]]
        for sl in slides:
            sl.pop("imagen", None)
        imagenes_pdf = entrada.get("imagenes_pdf", []) or []
        marcadas = [i for i, sl in enumerate(slides, 1) if sl.get("lleva_imagen")]

        if not marcadas:
            return {"salida": {**entrada, "slides": slides, "plan_imagenes": [], "razonamiento_imagenes": "",
                               "fuentes_imagenes": {"pdf": 0, "web": 0, "ia": 0}},
                    "recibio": f"Plan de {len(slides)} diapositivas; el Seccionador no marcó ninguna para imagen",
                    "entrego": "0 imágenes (nada que ilustrar)", "razonamiento": "No había diapositivas marcadas.",
                    "detalle": "", "notas": ""}

        self._feedback = feedback
        razon_plan, plan, completadas = self._planear(slides, marcadas, imagenes_pdf, feedback)
        lineas = [f"El Seccionador marcó {len(marcadas)} de {len(slides)} diapositivas · el Lector entregó "
                  f"{len(imagenes_pdf)} imágenes del PDF",
                  f"Cómo decidió qué mostrar: {razon_plan}"]
        if completadas:
            lineas.append(f"(La IA no dio búsqueda para las diapositivas {completadas}; usé su título.)")
        lineas.append("")

        usadas_pdf, titulos_web = set(), set()
        conteo = {"pdf": 0, "web": 0, "ia": 0}
        for p in plan:
            i = p["diapositiva"] - 1
            sl = slides[i]
            etiqueta = f"Diap. {i + 1} «{sl.get('title', '')}»"
            fallos, img = [], None
            orden = [p["fuente"]] + [f for f in ("pdf", "web", "ia") if f != p["fuente"]]   # elegida por la IA y luego el resto
            for fuente in orden:
                if img:
                    break
                try:
                    if fuente == "pdf":
                        img, razon = self._desde_pdf(sl, imagenes_pdf, usadas_pdf, p.get("imagen_pdf"))
                        if img:
                            usadas_pdf.add(img["id"])
                        else:
                            fallos.append(f"PDF: {razon}")
                    elif fuente == "web":
                        consultas = [p["consulta"]]
                        titulo_corto = " ".join(str(sl.get("title", "")).split()[:5])
                        if titulo_corto and titulo_corto.lower() != p["consulta"].lower():
                            consultas.append(titulo_corto)
                        img, razon = self._desde_web(sl, i, consultas, titulos_web)
                        if img:
                            titulos_web.add(img["titulo"])
                        else:
                            fallos.append(f"web («{p['consulta']}»): {razon}")
                    else:
                        img = self._generar(i, p["prompt_ia"])
                except Exception as e:
                    fallos.append(f"{fuente}: error ({type(e).__name__}: {e})")

            previo = ("  [" + " | ".join(fallos) + "]") if fallos else ""
            if img:
                sl["imagen"] = img
                conteo[img["fuente"]] += 1
                origen = {"pdf": "del PDF", "web": f"de la web ({img['autor']}, {img['licencia']})",
                          "ia": "generada con IA"}[img["fuente"]]
                lineas.append(f"{etiqueta}: ✔ {origen} (la IA eligió «{p['fuente']}»: {p['por_que']}) → "
                              f"{img['razon']}{previo}")
            else:
                lineas.append(f"{etiqueta}: ✘ sin imagen{previo}")

        total = sum(conteo.values())
        return {
            "salida": {**entrada, "slides": slides, "plan_imagenes": plan, "razonamiento_imagenes": razon_plan,
                       "fuentes_imagenes": conteo},
            "recibio": f"Plan de {len(slides)} diapositivas; {len(marcadas)} marcadas para imagen y "
                       f"{len(imagenes_pdf)} imágenes del PDF",
            "entrego": f"{total} de {len(marcadas)} imágenes colocadas (PDF: {conteo['pdf']} · "
                       f"web: {conteo['web']} · IA: {conteo['ia']})",
            "razonamiento": (
                "Para cada diapositiva marcada por el Seccionador, la IA decidió de dónde sale su imagen "
                "(imagen del PDF, Wikimedia Commons o ilustración generada) y qué buscar; la preferencia por "
                "defecto es PDF, luego web, luego IA. Cada imagen se verifica viéndola, y si la fuente elegida "
                f"falla se prueban las otras. {razon_plan}"
            ),
            "detalle": "\n".join(lineas),
            "notas": "",
        }
