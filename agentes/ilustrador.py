"""Agente Ilustrador: pone imagen en las diapositivas marcadas trabajando en SIMULTANEO (hilos paralelos)
para recortar tiempos de espera."""

from concurrent.futures import ThreadPoolExecutor, as_completed
import io
import json
import os
import re
import threading

from PIL import Image

from utils import (Agente, buscar_commons, con_feedback, descargar_imagen,
                   generar_imagen_ia, llamar_llm, parsear_json)

SYS_ILUSTRADOR_PLAN = """
Eres un editor grafico que prepara las imagenes de una presentacion. Para cada diapositiva que se te
indica, escribes una busqueda precisa para un banco de fotos y una descripcion para generar una
ilustracion. Las imagenes son solo ilustrativas: sencillas, sin texto, sin personas reales
reconocibles y sin contenido violento o grafico.
Responde UNICAMENTE con JSON valido.
"""

SYS_ILUSTRADOR = """
Eres un editor grafico muy estricto. Eliges imagenes para diapositivas, pero solo
si muestran de forma clara y directa el tema. Si hay duda, no eliges ninguna.
Responde UNICAMENTE con JSON valido.
"""


def _texto(sl):
    return " ".join(x for x in (sl.get("title", ""), sl.get("parrafo") or "",
                                " ".join(sl.get("bullets") or [])) if x)


def _palabras(texto):
    """Palabras de 5+ letras recortadas a 6 (asi 'imagen' e 'imagenes' coinciden mas o menos)."""
    return {w[:6] for w in re.findall(r"[a-záéíóúñü]{5,}", texto.lower())}


class Ilustrador(Agente):
    nombre = "Ilustrador"
    rol = ("Pone imagen en las diapositivas que el Seccionador marco ejecutando busquedas y descargas en "
           "SIMULTANEO (paralelo). La IA decide la fuente (PDF, Wikimedia Commons o IA) y verifica cada imagen")

    def __init__(self, carpeta):
        self.carpeta = carpeta
        self._feedback = None
        self._lock = threading.Lock()

    def _planear(self, slides, marcadas, imagenes_pdf, feedback):
        resumen = [{"diapositiva": i, "titulo": slides[i - 1].get("title", ""),
                    "texto": _texto(slides[i - 1])[:240]} for i in marcadas]
        catalogo = [{"id": im["id"], "pagina": im["pagina"], "tipo": im.get("tipo", ""),
                     "descripcion": im.get("descripcion", "")} for im in imagenes_pdf]
        prompt = (
            f"Estas son las diapositivas de una presentacion que llevan imagen ({len(marcadas)} de {len(slides)}). "
            "Para cada una decide la FUENTE de su imagen:\n"
            "- 'pdf': una de las imagenes del PDF de abajo (pon su id en 'imagen_pdf'). Usala solo si de verdad "
            "ilustra esa diapositiva, y no repitas una imagen en dos diapositivas.\n"
            "- 'web': una foto de un banco de imagenes libres (Wikimedia Commons).\n"
            "- 'ia': una ilustracion generada con IA.\n"
            "El orden de preferencia es pdf, luego web, luego ia; elige otra solo si lo justificas en 'por_que'.\n"
            "Ademas, para cada una escribe:\n"
            "- 'consulta': busqueda corta EN INGLES (2 a 5 palabras) para Wikimedia Commons, sobre algo "
            "concreto y fotografiable que ilustre su contenido.\n"
            "- 'prompt_ia': descripcion EN INGLES (10 a 25 palabras) de una ilustracion sencilla y "
            "conceptual del mismo tema, para generarla con IA (sin texto ni letras en la imagen).\n"
            'Responde {"razonamiento": "<como decidiste>", "imagenes": [{"diapositiva": N, "fuente": "pdf|web|ia", '
            '"imagen_pdf": "<id o null>", "consulta": "...", "prompt_ia": "...", "por_que": "<una frase>"}]}\n\n'
            "IMAGENES DEL PDF:\n" + (json.dumps(catalogo, ensure_ascii=False) if catalogo else "(ninguna)")
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
            if not consulta or not prompt_ia:
                base = slides[i - 1].get("title", "presentation topic")
                consulta = consulta or " ".join(base.split()[:4])
                prompt_ia = prompt_ia or f"Conceptual illustration about {base}"
                completadas.append(i)
            fuente = x.get("fuente") if x.get("fuente") in ("pdf", "web", "ia") else ("pdf" if imagenes_pdf else "web")
            img_pdf = x.get("imagen_pdf") if x.get("imagen_pdf") in ids_pdf and x.get("imagen_pdf") not in ya_asignadas else None
            if img_pdf:
                ya_asignadas.add(img_pdf)
            plan.append({"diapositiva": i, "fuente": fuente, "imagen_pdf": img_pdf, "consulta": consulta,
                         "prompt_ia": prompt_ia, "por_que": str(x.get("por_que", ""))})
        return str(d.get("razonamiento", "")), plan, completadas

    def _elegir(self, sl, candidatas, origen):
        """La IA mira las candidatas y elige una (o ninguna)."""
        partes = [
            f"Diapositiva «{sl.get('title', '')}»\nContenido: {_texto(sl)[:350]}\n\n"
            f"Te muestro {len(candidatas)} imagen(es) candidata(s) de {origen} (numeradas del 1 al "
            f"{len(candidatas)}). Mira cada imagen y elige la que MEJOR ilustre de forma clara y directa "
            "el contenido de la diapositiva. Si ninguna se relaciona claramente, responde elegida = 0.\n"
            + (f"Nota del Jefe sobre un intento anterior: {self._feedback}\n" if self._feedback else "")
            + 'Responde {"elegida": <0 o 1..N>, "razon": "<una frase en espanol>"}'
        ]
        for k, c in enumerate(candidatas, 1):
            partes.append(f"Candidata {k} ({c['titulo']}):")
            partes.append({"mime_type": c["mime"], "data": c["bytes"]})
        try:
            d = parsear_json(llamar_llm(SYS_ILUSTRADOR, partes))
            n = d.get("elegida", 0)
            razon = str(d.get("razon", "")).strip()
            if isinstance(n, int) and 1 <= n <= len(candidatas):
                return candidatas[n - 1], razon or "ilustra el tema de la diapositiva"
            return None, razon or "ninguna candidata encaja"
        except Exception as e:
            return None, f"{type(e).__name__}: {e}"

    def _reservar_id(self, usadas, clave):
        with self._lock:
            if clave in usadas:
                return False
            usadas.add(clave)
            return True

    def _probar_pdf(self, i, sl, imagenes_pdf, usadas, preferida=None):
        if not imagenes_pdf:
            return None, "el PDF no tiene imagenes disponibles"
        pal_sl = _palabras(_texto(sl))

        def afinidad(im):
            return len(pal_sl & _palabras(im.get("contexto", "") + " " + im.get("descripcion", "")))

        with self._lock:
            libres = [im for im in imagenes_pdf if im["id"] not in usadas and os.path.exists(im["ruta"])]
        if not libres:
            return None, "todas las imagenes del PDF ya se usaron"
        libres.sort(key=lambda im: (im["id"] != preferida, -afinidad(im)))
        cands = []
        for im in libres[:3]:
            try:
                with open(im["ruta"], "rb") as f:
                    cands.append({"titulo": f"figura del PDF, pagina {im['pagina']}; {im.get('descripcion', '')[:140]}",
                                  "mime": "image/jpeg", "bytes": f.read(), "im": im})
            except Exception:
                continue
        if not cands:
            return None, "no se pudieron leer las imagenes del PDF"
        elegida, razon = self._elegir(sl, cands, "el propio PDF")
        if not elegida:
            return None, razon
        im = elegida["im"]
        if not self._reservar_id(usadas, im["id"]):
            return None, "la imagen del PDF ya fue tomada por otro hilo"
        return {
            "id": im["id"], "fuente": "pdf", "ruta": im["ruta"],
            "titulo": f"Figura de la pag. {im['pagina']} del PDF",
            "pagina_pdf": im["pagina"], "ancho": im["ancho"], "alto": im["alto"], "razon": razon,
        }, razon

    def _probar_web(self, i, sl, consulta, usadas):
        try:
            halladas = buscar_commons(consulta, limite=6)
        except Exception as e:
            return None, f"error en Wikimedia ({type(e).__name__}: {e})"
        if not halladas:
            return None, f"sin resultados en Wikimedia para '{consulta}'"
        bajadas = []
        for h in halladas:
            with self._lock:
                ya_usada = h["titulo"] in usadas
            if ya_usada:
                continue
            try:
                datos, mime = descargar_imagen(h["thumb"])
                with Image.open(io.BytesIO(datos)) as im:
                    w, al = im.size
                if w < 400:
                    continue
                bajadas.append({**h, "bytes": datos, "mime": mime, "ancho": w, "alto": al})
                if len(bajadas) >= 3:
                    break
            except Exception:
                continue
        if not bajadas:
            return None, "no se pudieron descargar candidatas validas de Wikimedia"
        elegida, razon = self._elegir(sl, [{**b, "titulo": b["titulo"]} for b in bajadas], "Wikimedia Commons")
        if not elegida:
            return None, razon
        if not self._reservar_id(usadas, elegida["titulo"]):
            return None, "candidata web tomada por otro hilo"
        ext = ".png" if elegida["mime"] == "image/png" else ".jpg"
        ruta = os.path.join(self.carpeta, f"imagen_{i}{ext}")
        with open(ruta, "wb") as f:
            f.write(elegida["bytes"])
        return {
            "id": elegida["titulo"], "fuente": "web", "ruta": ruta, "titulo": elegida["titulo"],
            "autor": elegida["autor"], "licencia": elegida["licencia"], "pagina": elegida["pagina"],
            "ancho": elegida["ancho"], "alto": elegida["alto"], "razon": razon,
        }, razon

    def _probar_ia(self, i, sl, prompt_ia):
        try:
            datos, mime = generar_imagen_ia(prompt_ia)
            with Image.open(io.BytesIO(datos)) as im:
                w, al = im.size
        except Exception as e:
            return None, f"error ({type(e).__name__}: {e})"
        ruta = os.path.join(self.carpeta, f"imagen_ia_{i}.jpg")
        with open(ruta, "wb") as f:
            f.write(datos)
        return {
            "id": f"ia-{i}", "fuente": "ia", "ruta": ruta,
            "titulo": f"Ilustracion IA ({prompt_ia[:45]}...)",
            "ancho": w, "alto": al,
            "razon": f"Ilustracion conceptual generada con IA para '{sl.get('title', '')}'",
        }, "generada con IA"

    def _procesar_diapositiva(self, p, slides, imagenes_pdf, usadas):
        """Se ejecuta en paralelo dentro del ThreadPoolExecutor para cada diapositiva marcada."""
        i = p["diapositiva"]
        sl = slides[i - 1]
        preferida = p["fuente"]
        orden = [preferida] + [f for f in ("pdf", "web", "ia") if f != preferida]
        intentos_txt = []
        colocada = None
        for fte in orden:
            if fte == "pdf":
                img, msg = self._probar_pdf(i, sl, imagenes_pdf, usadas, preferida=p.get("imagen_pdf"))
            elif fte == "web":
                img, msg = self._probar_web(i, sl, p["consulta"], usadas)
            else:
                img, msg = self._probar_ia(i, sl, p["prompt_ia"])
            if img:
                colocada = img
                intentos_txt.append(f"{fte}: OK ({msg})")
                break
            intentos_txt.append(f"{fte}: {msg}")

        if colocada:
            linea = (f"Diap. {i} «{sl.get('title', '')}»: ✔ [{colocada['fuente'].upper()}] "
                     f"{colocada['titulo']} — {colocada['razon']}")
        else:
            linea = f"Diap. {i} «{sl.get('title', '')}»: ✘ sin imagen  [{' | '.join(intentos_txt)}]"
        return i, colocada, linea

    def ejecutar(self, entrada, feedback=None):
        self._feedback = feedback
        slides = [dict(s) for s in entrada["slides"]]
        for s in slides:
            s.pop("imagen", None)
        imagenes_pdf = entrada.get("imagenes_pdf", []) or []
        marcadas = [i for i, s in enumerate(slides, 1) if s.get("lleva_imagen")]
        if not marcadas:
            return {
                "salida": {**entrada, "slides": slides, "plan_imagenes": [], "razonamiento_imagenes": "Sin diapositivas marcadas",
                           "fuentes_imagenes": {"pdf": 0, "web": 0, "ia": 0}},
                "recibio": f"Plan de {len(slides)} diapositivas (0 marcadas para imagen)",
                "entrego": "0 imagenes colocadas",
                "razonamiento": "El Seccionador no marco ninguna diapositiva para llevar imagen.",
                "detalle": "Sin imagenes solicitadas.",
                "notas": "",
            }

        razon_plan, plan, completadas = self._planear(slides, marcadas, imagenes_pdf, feedback)
        usadas = set()
        conteo = {"pdf": 0, "web": 0, "ia": 0}
        resultados_ordenados = {}

        # Ejecucion en SIMULTANEO con hasta 3 hilos paralelos para recortar el tiempo de espera
        max_hilos = min(3, len(plan))
        print(f"[ILUSTRADOR] Procesando {len(plan)} imagenes en simultaneo ({max_hilos} hilos paralelos)...", flush=True)
        with ThreadPoolExecutor(max_workers=max_hilos) as pool:
            futuros = {pool.submit(self._procesar_diapositiva, p, slides, imagenes_pdf, usadas): p["diapositiva"]
                       for p in plan}
            for fut in as_completed(futuros):
                i, colocada, linea = fut.result()
                resultados_ordenados[i] = (colocada, linea)

        lineas = [
            f"El Seccionador marco {len(marcadas)} de {len(slides)} diapositivas · el Lector entrego {len(imagenes_pdf)} imagenes del PDF",
            f"Modo de ejecucion: Paralelo ({max_hilos} hilos simultaneos con ThreadPoolExecutor)",
            f"Como decidio que mostrar: {razon_plan}",
            "",
        ]
        for i in sorted(resultados_ordenados.keys()):
            colocada, linea = resultados_ordenados[i]
            lineas.append(linea)
            if colocada:
                slides[i - 1]["imagen"] = colocada
                conteo[colocada["fuente"]] += 1

        total = sum(conteo.values())
        return {
            "salida": {**entrada, "slides": slides, "plan_imagenes": plan, "razonamiento_imagenes": razon_plan,
                       "fuentes_imagenes": conteo},
            "recibio": f"Plan de {len(slides)} diapositivas; {len(marcadas)} marcadas para imagen y "
                       f"{len(imagenes_pdf)} imagenes del PDF",
            "entrego": f"{total} de {len(marcadas)} imagenes colocadas en paralelo (PDF: {conteo['pdf']} · "
                       f"web: {conteo['web']} · IA: {conteo['ia']})",
            "razonamiento": (
                f"Procese las {len(marcadas)} diapositivas en simultaneo ({max_hilos} hilos paralelos). "
                "Para cada una la IA decidio su fuente (PDF, Wikimedia Commons o IA) y verifico la imagen. "
                f"{razon_plan}"
            ),
            "detalle": "\n".join(lineas),
            "notas": f"Se completaron busquedas por defecto en {completadas}" if completadas else "",
        }
        