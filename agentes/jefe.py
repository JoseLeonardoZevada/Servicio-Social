"""Agente Jefe: coordina a los demás, valida su trabajo, manda a corregir y, si hace falta, corrige él."""
import json
import math
import os
import time
from html import escape

from PIL import Image
from pptx import Presentation
from pptx.enum.dml import MSO_FILL
from pptx.util import Inches

from utils import (asegurar_imagenes, _cabe, _chk, ErrorPipeline, _imagenes, llamar_llm,
                   min_imagenes, num_palabras, parsear_json, presupuesto_diapositivas,
                   _problemas_contraste, rango_de_temas, rango_diapositivas, _recortar_texto,
                   _referencia_en_texto, total_objetivo)
from agentes.decorador import Decorador
from agentes.generador import GeneradorPPTX
from agentes.ilustrador import Ilustrador
from agentes.lector import Lector
from agentes.recopilador import Recopilador
from agentes.seccionador import Seccionador
from config import (COBERTURA_MINIMA, LIMITE_TEXTO_REVISION, MAX_DIAPOSITIVAS, MAX_INTENTOS,
                    PROPORCION_IMAGENES, TITULO_FUENTES)


SYS_JEFE = """
Eres el Jefe supervisor de un equipo de agentes que convierte documentos en presentaciones.
Revisas el trabajo con criterio y sin ser quisquilloso: señalas errores reales, no gustos.
Responde ÚNICAMENTE con JSON válido.
"""


class Jefe:
    nombre = "Jefe"
    rol = "Supervisa y corrige a los agentes"

    def __init__(self):
        self.log = []
        self.t_inicio = time.time()
        self.texto_pdf = ""
        self.texto_original = ""
        self.rango_temas = (3, 8)   # cuántos temas debe sacar el Recopilador (depende de las páginas)
        self.n_temas = 0            # temas que entregó el Recopilador
        self.secciones = []
        self.presupuesto = []       # diapositivas que le tocan a cada tema (según sus caracteres)
        self.objetivo = 0           # diapositivas de contenido que corresponden en total
        self.n_min, self.n_max = 3, MAX_DIAPOSITIVAS
        self.ultimo_intento = False  # en el último intento, el veto de la IA solo avisa (para no trabar todo)

    # ---- bitácora ----
    def _anotar(self, agente, rol, intento, estado, segundos, res=None, revisiones=None, notas=""):
        res = res or {}
        self.log.append(
            {
                "paso": len(self.log) + 1,
                "agente": agente,
                "rol": rol,
                "intento": intento,
                "estado": estado,
                "recibio": res.get("recibio", "-"),
                "entrego": res.get("entrego", "-"),
                "razonamiento": res.get("razonamiento", ""),
                "detalle": res.get("detalle", ""),
                "notas": notas or res.get("notas", ""),
                "revisiones": revisiones or [],
                "segundos": round(segundos, 1),
            }
        )

    # ---- reglas de supervisión: devuelven lista de revisiones ----
    def revisar(self, nombre, s):
        c = []
        if nombre == "Lector":
            _chk(c, "Texto extraído suficiente (al menos 200 caracteres)",
                 len(s["texto"].strip()) >= 200, "Texto casi vacío (¿PDF escaneado como imagen?)")
            if s.get("solapamiento") is not None:
                _chk(c, "La lectura de Gemini coincide con la de pdfplumber (al menos 60% de palabras)",
                     s["solapamiento"] >= 0.6,
                     f"La transcripción de Gemini solo coincide en {round(s['solapamiento'] * 100)}% con "
                     "pdfplumber; puede estar incompleta o alterada")
        elif nombre == "Recopilador":
            secciones = s.get("secciones", [])
            _chk(c, "Trae título y subtítulo del documento",
                 s.get("titulo_documento") and s.get("subtitulo"), "Falta 'titulo_documento' o 'subtitulo'")
            _chk(c, "Título de máximo 12 palabras",
                 num_palabras(s.get("titulo_documento", "")) <= 12, "El título del documento es muy largo")
            _chk(c, "Trae tema general", s.get("tema_general"), "Falta 'tema_general'")
            _chk(c, "Explica su razonamiento", s.get("razonamiento"), "Falta el campo 'razonamiento'")
            mn, mx = self.rango_temas
            _chk(c, "Al menos 3 temas", len(secciones) >= 3, f"Hay {len(secciones)} temas; extrae al menos 3")
            _chk(c, f"Cantidad de temas razonable para el tamaño del PDF (referencia: de {mn} a {mx})",
                 mn <= len(secciones) <= mx,
                 f"Hay {len(secciones)} temas; la referencia para este PDF era de {mn} a {mx}", bloquea=False)
            sin_puntos = [x.get("tema", "?") for x in secciones if not x.get("puntos")]
            _chk(c, "Todos los temas tienen puntos", not sin_puntos, f"Temas sin puntos: {sin_puntos}")
            refs = s.get("referencias", [])
            _chk(c, "'referencias' es una lista de máximo 8", isinstance(refs, list) and len(refs) <= 8,
                 "'referencias' debe ser una lista de máximo 8 elementos")
            if isinstance(refs, list):
                inventadas = [r for r in refs if not _referencia_en_texto(str(r), self.texto_pdf)]
                _chk(c, "Todas las referencias aparecen realmente en el PDF (anti-invención)",
                     not inventadas,
                     f"Estas referencias no aparecen en el documento, quítalas: {inventadas}")
            if all(r["ok"] or not r["bloquea"] for r in c):      # la IA del Jefe revisa solo si pasó las reglas fijas
                try:
                    r = self._revision_ia_temas(s)
                    _chk(c, "Revisión con IA: no faltan temas importantes del documento ni hay datos inventados",
                         not (r["temas_faltantes"] or r["datos_inventados"]),
                         "La IA encontró: " + " | ".join(
                             (["Temas que faltan: " + ", ".join(r["temas_faltantes"])] if r["temas_faltantes"] else [])
                             + (["Datos que no aparecen en el documento: " + "; ".join(r["datos_inventados"])]
                                if r["datos_inventados"] else [])),
                         bloquea=not self.ultimo_intento)
                    if r["temas_mezclados"]:
                        _chk(c, "Sugerencias de la IA sobre temas mezclados (no bloquean)", False,
                             " | ".join(r["temas_mezclados"]), bloquea=False)
                except Exception as e:
                    _chk(c, "Revisión con IA de los temas", False,
                         f"No se pudo ejecutar ({type(e).__name__}: {e})", bloquea=False)
        elif nombre == "Seccionador":
            c += self._checks_plan(s["slides"], s.get("razonamiento"))
            if all(r["ok"] or not r["bloquea"] for r in c):  # la IA revisa solo si pasó las reglas fijas
                try:
                    r = self._revision_ia_plan(s)
                    _chk(c, "Revisión con IA: fiel al documento, sin ideas repetidas y coherente",
                         not r["problemas_graves"], "La revisión con IA encontró: " + " | ".join(r["problemas_graves"]),
                         bloquea=not self.ultimo_intento)
                    sc = [k for k in r["temas_sin_cubrir"] if 1 <= k <= self.n_temas]
                    cob = (self.n_temas - len(sc)) / self.n_temas if self.n_temas else 1
                    _chk(c, f"Revisión con IA: cobertura real de al menos el {round(COBERTURA_MINIMA * 100)}% de los temas",
                         cob >= COBERTURA_MINIMA,
                         "La IA considera que estos temas casi no se tratan: " + ", ".join(
                             f"{k} ({self.secciones[k - 1].get('tema', '?')})" for k in sc),
                         bloquea=not self.ultimo_intento)
                    if r["sugerencias"]:
                        _chk(c, "Sugerencias de la IA (no bloquean)", False, " | ".join(r["sugerencias"]), bloquea=False)
                    if r["veredicto"]:
                        _chk(c, "Veredicto de la IA: " + r["veredicto"], True)
                except Exception as e:
                    _chk(c, "Revisión con IA", False, f"No se pudo ejecutar ({type(e).__name__}: {e})", bloquea=False)
        elif nombre == "Ilustrador":
            slides = s["slides"]
            marcadas = [i for i, x in enumerate(slides, 1) if x.get("lleva_imagen")]
            plan = s.get("plan_imagenes", [])
            imgs = [(i, x["imagen"]) for i, x in enumerate(slides, 1) if x.get("imagen")]
            en_plan = [p.get("diapositiva") for p in plan]
            sin_plan = [i for i in marcadas if i not in en_plan]
            _chk(c, "El plan de imágenes cubre todas las diapositivas que marcó el Seccionador", not sin_plan,
                 f"Sin plan de imagen para las diapositivas {sin_plan}")
            _chk(c, "No planea dos imágenes para la misma diapositiva", len(en_plan) == len(set(en_plan)),
                 "Hay diapositivas repetidas en el plan de imágenes")
            malas_q = [p.get("diapositiva") for p in plan
                       if not p.get("consulta") or num_palabras(p["consulta"]) > 6 or not p.get("prompt_ia")]
            _chk(c, "Cada búsqueda es corta (máximo 6 palabras) y trae su descripción para la ilustración",
                 not malas_q, f"Búsquedas o descripciones vacías o muy largas en las diapositivas {malas_q}")
            if marcadas:
                _chk(c, "Explica cómo decidió qué mostrar", s.get("razonamiento_imagenes"),
                     "Falta el razonamiento de la selección de imágenes")
            _chk(c, "Colocó al menos una imagen", not marcadas or imgs,
                 "No se logró colocar ninguna imagen (fallaron las tres fuentes: PDF, web e IA)")
            n_pdf = len(s.get("imagenes_pdf", []) or [])
            _chk(c, "Si el PDF trae imágenes útiles, se aprovechó al menos una",
                 not n_pdf or any(im.get("fuente") == "pdf" for _, im in imgs),
                 f"El PDF trae {n_pdf} imágenes útiles y no se usó ninguna", bloquea=False)
            faltan = [i for i in marcadas if i not in [j for j, _ in imgs]]
            _chk(c, "Todas las diapositivas marcadas recibieron imagen", not faltan,
                 f"Quedaron sin imagen las diapositivas {faltan}", bloquea=False)
            malas, sin_cred, sin_razon = [], [], []
            for i, img in imgs:
                minimo = 250 if img.get("fuente") == "pdf" else 400
                try:
                    with Image.open(img["ruta"]) as im:
                        im.verify()
                    if img["ancho"] < minimo:
                        malas.append(i)
                except Exception:
                    malas.append(i)
                if img.get("fuente", "web") == "web" and not (img.get("autor") and img.get("licencia") and img.get("pagina")):
                    sin_cred.append(i)
                if img.get("fuente") == "pdf" and not img.get("pagina_pdf"):
                    sin_cred.append(i)
                if not img.get("razon"):
                    sin_razon.append(i)
            _chk(c, "Cada imagen es un archivo legible y con buen tamaño (400 px de ancho; 250 si viene del PDF)",
                 not malas, f"Imágenes ilegibles o muy pequeñas en las diapositivas {malas}")
            _chk(c, "Cada imagen trae su crédito (autor, licencia y enlace si es de la web; página si es del PDF)",
                 not sin_cred, f"Faltan datos de crédito en las diapositivas {sin_cred}")
            _chk(c, "Cada imagen trae la justificación de por qué ilustra el tema", not sin_razon,
                 f"Falta la justificación en las diapositivas {sin_razon}")
            ids = [img.get("id") or img["titulo"] for _, img in imgs]
            _chk(c, "No hay imágenes repetidas", len(ids) == len(set(ids)), "Imagen repetida")
            if imgs and all(r["ok"] or not r["bloquea"] for r in c):
                try:
                    rechazadas = self._revision_ia_imagenes(slides, imgs)
                    _chk(c, "Revisión con IA: cada imagen encaja con el tema de su diapositiva", not rechazadas,
                         "La IA vio que no encajan: " + " | ".join(f"diap. {n}: {m}" for n, m in rechazadas),
                         bloquea=not self.ultimo_intento)
                except Exception as e:
                    _chk(c, "Revisión con IA de las imágenes", False,
                         f"No se pudo ejecutar ({type(e).__name__}: {e})", bloquea=False)
        elif nombre == "Generador PPTX":
            existe = os.path.exists(s["ruta"])
            _chk(c, "El archivo .pptx existe", existe, "No se creó el archivo")
            if existe:
                prs = Presentation(s["ruta"])
                _chk(c, "El número de diapositivas coincide con el plan (portada + contenido + fuentes)",
                     len(prs.slides) == s["n_slides"], "El número de diapositivas no coincide con el plan")
                portada = prs.slides[0].shapes.title.text if len(prs.slides) else ""
                _chk(c, "La primera diapositiva es la portada con el título",
                     portada == s["titulo"], "La portada no tiene el título esperado")
                ultima = prs.slides[-1].shapes.title.text if len(prs.slides) else ""
                _chk(c, f"La última diapositiva es de '{TITULO_FUENTES}'", ultima.startswith(TITULO_FUENTES),
                     "Falta la diapositiva de fuentes al final")
                n_pics = sum(len(_imagenes(sl)) for sl in prs.slides)
                _chk(c, "Cada imagen del plan quedó insertada en su diapositiva",
                     n_pics == s["n_imagenes"], f"Hay {n_pics} imágenes y el plan tenía {s['n_imagenes']}")
                n_notas, n_cont = s.get("n_notas", 0), s.get("n_contenido", 0)
                _chk(c, "Cada diapositiva de contenido trae notas del orador (las escribe la IA)",
                     n_cont == 0 or n_notas >= 0.9 * n_cont,
                     f"Solo {n_notas} de {n_cont} diapositivas tienen notas del orador", bloquea=False)
        elif nombre == "Decorador":
            prs = Presentation(s["ruta"])
            W, H = prs.slide_width, prs.slide_height
            prob_c = _problemas_contraste(s["diseno"]["paleta"])
            _chk(c, "Los colores tienen contraste suficiente para leerse", not prob_c,
                 "Contraste bajo: " + "; ".join(prob_c))
            _chk(c, "El diseño lo propuso la IA (si no, se usó el diseño de respaldo)",
                 s.get("origen_diseno") == "IA",
                 "Se usó el diseño de respaldo: " + "; ".join(s.get("problemas_ia", [])), bloquea=False)
            _chk(c, "No cambió el número de diapositivas", len(prs.slides) == s["n_slides"],
                 "Cambió el número de diapositivas")
            sin_fondo = [i for i, sl in enumerate(prs.slides, 1) if sl.background.fill.type != MSO_FILL.SOLID]
            _chk(c, "Todas las diapositivas tienen fondo aplicado", not sin_fondo,
                 f"Sin fondo en las diapositivas: {sin_fondo}")
            sin_formas = [i for i, sl in enumerate(prs.slides, 1) if len(sl.shapes) <= len(sl.placeholders)]
            _chk(c, "Todas las diapositivas tienen elementos decorativos", not sin_formas,
                 f"Sin decoración en las diapositivas: {sin_formas}")
            def se_cruzan(a, b):
                return (a.left < b.left + b.width and b.left < a.left + a.width
                        and a.top < b.top + b.height and b.top < a.top + a.height)
            fuera, solapa = [], []
            for i, sl in enumerate(prs.slides, 1):
                for pic in _imagenes(sl):
                    if pic.left < 0 or pic.left + pic.width > W or pic.top + pic.height + Inches(0.8) > H:
                        fuera.append(i)
                    if se_cruzan(pic, sl.placeholders[1]) or se_cruzan(pic, sl.shapes.title):
                        solapa.append(i)
            _chk(c, "Las imágenes y sus créditos caben dentro de la diapositiva", not fuera,
                 f"Imagen fuera de los márgenes en las diapositivas {fuera}")
            _chk(c, "Las imágenes no se encimen con el texto", not solapa,
                 f"La imagen se encima con el texto en las diapositivas {solapa}")
            desborda = []
            for i, sl in enumerate(prs.slides, 1):
                if i == 1:
                    continue
                if not _cabe(sl.placeholders[1]) or not _cabe(sl.shapes.title, sangria_in=0):
                    desborda.append(i)
            _chk(c, "El texto cabe en su caja (estimación, conviene revisarlo en PowerPoint)",
                 not desborda, f"Posible texto desbordado en las diapositivas {desborda}", bloquea=False)
        return c

    def _checks_plan(self, slides, razonamiento, estricto=True):
        c = []
        n = len(slides)
        _chk(c, f"Al menos {self.n_min} diapositivas (los temas y su extensión piden unas {self.objetivo})",
             n >= self.n_min,
             f"Hay {n} diapositivas y los temas pedían al menos {self.n_min} (objetivo {self.objetivo}): "
             "desarrolla cada tema con las diapositivas que le tocan", bloquea=estricto)
        _chk(c, f"Máximo {MAX_DIAPOSITIVAS} diapositivas", n <= MAX_DIAPOSITIVAS,
             f"Hay {n} diapositivas; el máximo es {MAX_DIAPOSITIVAS}")
        _chk(c, f"No se pasa del máximo recomendado ({self.n_max})", n <= self.n_max,
             f"Hay {n} diapositivas y el objetivo era {self.objetivo}", bloquea=False)
        malos = [i for i, sl in enumerate(slides, 1) if not sl.get("title") or num_palabras(sl["title"]) > 8]
        _chk(c, "Títulos de máximo 8 palabras", not malos,
             f"Títulos vacíos o de más de 8 palabras en las diapositivas {malos}")

        tipos_mal = [i for i, sl in enumerate(slides, 1) if sl.get("tipo") not in ("parrafo", "mixto", "bullets")]
        _chk(c, "Cada diapositiva tiene un 'tipo' válido (parrafo, mixto o bullets)", not tipos_mal,
             f"'tipo' inválido o ausente en las diapositivas {tipos_mal}")

        incoherentes = []
        for i, sl in enumerate(slides, 1):
            t, par, bl = sl.get("tipo"), sl.get("parrafo", "") or "", sl.get("bullets", []) or []
            if t == "parrafo" and not 25 <= num_palabras(par) <= 55:
                incoherentes.append(f"{i} (párrafo de {num_palabras(par)} palabras; debe tener 25 a 55)")
            elif t == "mixto" and not (10 <= num_palabras(par) <= 30 and 2 <= len(bl) <= 5):
                incoherentes.append(f"{i} (mixto: párrafo de 10 a 30 palabras y 2 a 5 bullets)")
            elif t == "bullets" and not 3 <= len(bl) <= 8:
                incoherentes.append(f"{i} (bullets: debe tener 3 a 8)")
        _chk(c, "El contenido de cada diapositiva coincide con su tipo y sus límites", not incoherentes,
             "Revisa las diapositivas: " + "; ".join(incoherentes))

        solo_bullets = sum(1 for sl in slides if sl.get("tipo") == "bullets")
        permitido = math.ceil(0.6 * n)
        _chk(c, "No más del 60% de las diapositivas son solo bullets", solo_bullets <= permitido,
             f"{solo_bullets} de {n} diapositivas son solo bullets (máximo {permitido}); usa más texto corrido")

        largos = [i for i, sl in enumerate(slides, 1)
                  if any(num_palabras(b) > 15 for b in (sl.get("bullets") or []))]
        _chk(c, "Bullets de máximo 15 palabras", not largos,
             f"Bullets de más de 15 palabras en las diapositivas {largos}")
        vistos, repetidos = set(), []
        for i, sl in enumerate(slides, 1):
            for b in (sl.get("bullets") or []):
                clave = str(b).strip().lower()
                if clave in vistos:
                    repetidos.append(i)
                vistos.add(clave)
        _chk(c, "Sin bullets repetidos", not repetidos,
             f"Bullets repetidos en las diapositivas {sorted(set(repetidos))}")

        def temas_validos(sl):
            t = sl.get("temas")
            return (isinstance(t, list) and t
                    and all(isinstance(x, int) and not isinstance(x, bool) and 1 <= x <= self.n_temas for x in t))
        temas_mal = [i for i, sl in enumerate(slides, 1) if not temas_validos(sl)]
        _chk(c, "Cada diapositiva indica qué temas del Recopilador cubre (números válidos)", not temas_mal,
             f"Falta 'temas' o tiene números inválidos (1 a {self.n_temas}) en las diapositivas {temas_mal}")
        # La introducción y la conclusión resumen varios temas: la cobertura se mide con las de en medio
        medio = slides[1:-1] if n >= 3 else slides
        cubiertos = {x for sl in medio if temas_validos(sl) for x in sl["temas"]}
        minimo_cub = math.ceil(round(COBERTURA_MINIMA * self.n_temas, 6))
        faltan = sorted(set(range(1, self.n_temas + 1)) - cubiertos)
        _chk(c, f"Cobertura: al menos el {round(COBERTURA_MINIMA * 100)}% de los temas del Recopilador tiene "
                f"su propia diapositiva ({minimo_cub} de {self.n_temas})",
             len(cubiertos) >= minimo_cub,
             "Ninguna diapositiva desarrolla los temas " + ", ".join(
                 f"{k} ({self.secciones[k - 1].get('tema', '?')})" for k in faltan))

        # Imágenes: el Seccionador es quien dice en qué diapositivas va una
        marcadas = [bool(sl.get("lleva_imagen")) for sl in slides]
        minimo_img = min_imagenes(n)
        _chk(c, f"Al menos el {round(PROPORCION_IMAGENES * 100)}% de las diapositivas llevan imagen "
                f"({minimo_img} de {n})", sum(marcadas) >= minimo_img,
             f"Solo {sum(marcadas)} de {n} diapositivas llevan imagen; deben ser al menos {minimo_img}")
        seguidas = [i + 1 for i in range(1, n) if not marcadas[i] and not marcadas[i - 1]]
        _chk(c, "Las imágenes están repartidas (nunca dos diapositivas seguidas sin imagen)", not seguidas,
             f"Hay dos diapositivas seguidas sin imagen que terminan en la {seguidas}")
        _chk(c, "Explica su razonamiento", razonamiento, "Falta el campo 'razonamiento'")
        return c

    # ---- revisiones con IA del Jefe ----
    def _revision_ia_plan(self, salida):
        plan = [{"numero": i, **{k: sl.get(k) for k in ("title", "tipo", "parrafo", "bullets")}}
                for i, sl in enumerate(salida["slides"], 1)]
        temas = [{"numero": i, "tema": sec.get("tema"), "puntos": sec.get("puntos")}
                 for i, sec in enumerate(self.secciones, 1)]
        prompt = (
            "Compara el PLAN de diapositivas con el DOCUMENTO original y revisa:\n"
            "1. Fidelidad: ¿alguna afirmación del plan NO aparece en el documento o lo contradice? "
            "(Si el documento muestra '[... texto omitido ...]', no marques como inventado lo que podría "
            "estar en esa parte.)\n"
            "2. Repetición: ¿hay ideas repetidas entre diapositivas?\n"
            "3. Coherencia: ¿el orden es lógico (introducción, desarrollo, conclusión) y todo es del "
            "tema del documento?\n"
            "4. Cobertura real: el RESUMEN DE TEMAS lista los temas del documento. En 'temas_sin_cubrir' pon "
            "los NÚMEROS de los temas cuyas ideas principales no aparezcan realmente en ninguna diapositiva "
            "(aunque el plan diga cubrirlos). Si todos están tratados, deja la lista vacía.\n"
            "Pon en 'problemas_graves' SOLO lo que sea claramente un error (afirmación inventada, "
            "contradicción, idea repetida, diapositiva fuera de tema), indicando el número de diapositiva. "
            "Los asuntos de estilo o mejoras opcionales van en 'sugerencias'. Si no hay problemas graves, "
            "deja la lista vacía.\n"
            'Responde {"problemas_graves": ["diapositiva N: ..."], "sugerencias": ["..."], '
            '"temas_sin_cubrir": [N], "veredicto": "<una frase>"}\n\n'
            "DOCUMENTO:\n" + _recortar_texto(self.texto_original) + "\n\nRESUMEN DE TEMAS:\n"
            + json.dumps(temas, ensure_ascii=False) + "\n\nPLAN:\n" + json.dumps(plan, ensure_ascii=False)
        )
        d = parsear_json(llamar_llm(SYS_JEFE, prompt))
        sin_cubrir = [x for x in d.get("temas_sin_cubrir", []) if isinstance(x, int) and not isinstance(x, bool)]
        return {"problemas_graves": [str(x) for x in d.get("problemas_graves", [])],
                "sugerencias": [str(x) for x in d.get("sugerencias", [])],
                "temas_sin_cubrir": sin_cubrir,
                "veredicto": str(d.get("veredicto", ""))}

    def _revision_ia_temas(self, salida):
        temas = [{"numero": i, "tema": sec.get("tema"), "puntos": sec.get("puntos")}
                 for i, sec in enumerate(salida.get("secciones", []), 1)]
        prompt = (
            "Compara los TEMAS que extrajo el Recopilador con el DOCUMENTO original y revisa:\n"
            "1. 'temas_faltantes': secciones o ideas importantes del documento que no aparecen en ningún tema "
            "(nombres cortos). Solo lo claramente importante.\n"
            "2. 'temas_mezclados': temas de la lista que mezclan asuntos que convendría separar.\n"
            "3. 'datos_inventados': puntos que NO aparecen en el documento o lo contradicen.\n"
            "(Si el documento muestra '[... texto omitido ...]', no marques como inventado lo que podría estar en "
            "esa parte.)\n"
            'Responde {"temas_faltantes": ["..."], "temas_mezclados": ["..."], "datos_inventados": ["..."], '
            '"veredicto": "<una frase>"}\n\nDOCUMENTO:\n' + _recortar_texto(self.texto_original)
            + "\n\nTEMAS:\n" + json.dumps(temas, ensure_ascii=False)
        )
        d = parsear_json(llamar_llm(SYS_JEFE, prompt))
        lista = lambda k: [str(x) for x in d.get(k, []) if str(x).strip()]
        return {"temas_faltantes": lista("temas_faltantes"), "temas_mezclados": lista("temas_mezclados"),
                "datos_inventados": lista("datos_inventados")}

    def _revision_ia_imagenes(self, slides, imgs):
        """La IA del Jefe VE cada imagen junto al texto de su diapositiva y dice si encaja. Devuelve [(diapositiva, motivo)]."""
        rechazadas = []
        for ini in range(0, len(imgs), 4):
            lote = imgs[ini:ini + 4]
            partes = ['Para cada imagen di si ilustra con claridad el tema de su diapositiva. "encaja" es false '
                      "si no se relaciona, es confusa, ilegible o muestra algo distinto. Responde "
                      '{"resultados": [{"diapositiva": N, "encaja": true, "motivo": "<una frase>"}]}']
            for n, img in lote:
                sl = slides[n - 1]
                partes.append(f"Imagen de la diapositiva {n} «{sl.get('title', '')}» (contenido: "
                              f"{(sl.get('parrafo') or ' '.join(sl.get('bullets') or []))[:200]}):")
                with open(img["ruta"], "rb") as f:
                    partes.append({"mime_type": "image/png" if img["ruta"].lower().endswith(".png") else "image/jpeg",
                                   "data": f.read()})
            d = parsear_json(llamar_llm(SYS_JEFE, partes))
            for x in d.get("resultados", []):
                if isinstance(x, dict) and x.get("encaja") is False and isinstance(x.get("diapositiva"), int):
                    rechazadas.append((x["diapositiva"], str(x.get("motivo", "no encaja"))))
        return rechazadas

    def _corregir_plan_ia(self, salida, problemas):
        """Cuando el Seccionador agota sus intentos, la IA del Jefe rehace el plan con las objeciones."""
        plan = [{k: sl.get(k) for k in ("title", "temas", "tipo", "parrafo", "bullets", "lleva_imagen")}
                for sl in salida["slides"]]
        temas = [{"numero": i, "tema": sec.get("tema"), "puntos": sec.get("puntos")}
                 for i, sec in enumerate(self.secciones, 1)]
        prompt = (
            f"El Seccionador no logró cumplir las reglas tras {MAX_INTENTOS} intentos. Corrige TÚ el plan.\n"
            "Problemas que señaló el Jefe:\n" + str(problemas) + "\n\n"
            f"Reglas: entre {self.n_min} y {self.n_max} diapositivas; cada tema del RESUMEN tiene su propia "
            "diapositiva (su número en 'temas'), salvo la introducción (primera) y la conclusión (última); "
            "'tipo' es parrafo (25 a 55 palabras, sin bullets), mixto (párrafo de 10 a 30 palabras y 2 a 5 bullets) "
            "o bullets (3 a 8 bullets de máximo 15 palabras); títulos de máximo 8 palabras; no más del 60% solo "
            "bullets; al menos el 60% con 'lleva_imagen' true y nunca dos seguidas sin imagen; usa solo información "
            "del RESUMEN.\n"
            'Responde {"razonamiento": "...", "slides": [<TODAS las diapositivas corregidas>]}\n\n'
            "RESUMEN DE TEMAS:\n" + json.dumps(temas, ensure_ascii=False) + "\n\nPLAN ACTUAL:\n"
            + json.dumps(plan, ensure_ascii=False)
        )
        d = parsear_json(llamar_llm(SYS_JEFE, prompt))
        slides = [dict(x) for x in d.get("slides", []) if isinstance(x, dict)]
        if not slides:
            raise ValueError("la IA no devolvió diapositivas")
        return {**salida, "slides": slides, "razonamiento": str(d.get("razonamiento") or salida.get("razonamiento") or "")}

    def _revision_final(self, ruta):
        t0 = time.time()
        try:
            prs = Presentation(ruta)
            textos = []
            for i, sl in enumerate(prs.slides, 1):
                partes = [sh.text_frame.text for sh in sl.shapes
                          if sh.has_text_frame and len(sh.text_frame.text.strip()) > 2]
                textos.append(f"Diapositiva {i}: " + " | ".join(partes))
            prompt = (
                "Esta es la presentación FINAL generada a partir del DOCUMENTO. Califica del 1 al 10 su "
                "calidad (claridad, coherencia, fidelidad al documento, equilibrio entre texto e "
                "imágenes). Indica fortalezas y problemas. Si hay problemas, di qué agente debería "
                "revisarse (Lector, Recopilador, Seccionador, Ilustrador, Generador PPTX o Decorador).\n"
                'Responde {"calificacion": <1-10>, "fortalezas": ["..."], "problemas": ["..."], '
                '"agente_a_revisar": "<nombre o null>", "veredicto": "<una frase>"}\n\n'
                "DOCUMENTO:\n" + _recortar_texto(self.texto_original, LIMITE_TEXTO_REVISION)
                + "\n\nPRESENTACIÓN:\n" + "\n".join(textos)
            )
            d = parsear_json(llamar_llm(SYS_JEFE, prompt))
            detalle = ["Fortalezas:"] + [f"   + {x}" for x in d.get("fortalezas", [])]
            detalle += ["Problemas:"] + ([f"   - {x}" for x in d.get("problemas", [])] or ["   (ninguno)"])
            detalle.append(f"Agente a revisar: {d.get('agente_a_revisar') or 'ninguno'}")
            self._anotar(self.nombre, "Revisión final de la presentación con IA", "-", "🧐 Revisión final",
                         time.time() - t0,
                         {"recibio": f"Presentación final de {len(prs.slides)} diapositivas",
                          "entrego": f"Calificación {d.get('calificacion', '?')}/10",
                          "razonamiento": str(d.get("veredicto", "")), "detalle": "\n".join(detalle)})
        except Exception as e:
            self._anotar(self.nombre, "Revisión final de la presentación con IA", "-",
                         "⚠️ Revisión final no disponible", time.time() - t0,
                         {"recibio": "-", "entrego": "-",
                          "razonamiento": f"No se pudo ejecutar la revisión con IA: {type(e).__name__}: {e}"})

    # ---- corrección directa del Jefe (cuando el agente ya no puede) ----
    def _corregir_plan(self, salida):
        cambios, vistos, nuevos = [], set(), []
        slides = salida["slides"]
        if len(slides) > MAX_DIAPOSITIVAS:
            cambios.append(f"recortó el plan a {MAX_DIAPOSITIVAS} diapositivas")
        for i, sl in enumerate(slides[:MAX_DIAPOSITIVAS], 1):
            t = " ".join(str(sl.get("title", "Sin título")).split()[:8])
            par = (sl.get("parrafo") or "").strip()
            if num_palabras(par) > 55:
                par = " ".join(par.split()[:55]) + "…"
                cambios.append(f"recortó el párrafo de la diapositiva {i}")
            bullets = []
            for b in (sl.get("bullets") or []):
                palabras = str(b).split()
                if len(palabras) > 15:
                    b = " ".join(palabras[:15]) + "…"
                    cambios.append(f"recortó un bullet de la diapositiva {i}")
                clave = b.strip().lower()
                if clave in vistos:
                    cambios.append(f"quitó un bullet repetido en la diapositiva {i}")
                    continue
                vistos.add(clave)
                bullets.append(b)
            if len(bullets) > 8:
                bullets = bullets[:8]
                cambios.append(f"dejó solo 8 bullets en la diapositiva {i}")
            tipo = "mixto" if par and bullets else ("parrafo" if par else "bullets")
            if tipo != sl.get("tipo"):
                cambios.append(f"ajustó el tipo de la diapositiva {i} a '{tipo}'")
            temas = [x for x in (sl.get("temas") or [])
                     if isinstance(x, int) and not isinstance(x, bool) and 1 <= x <= self.n_temas]
            nuevos.append({"title": t, "tipo": tipo, "parrafo": par, "bullets": bullets, "temas": temas,
                           "lleva_imagen": bool(sl.get("lleva_imagen"))})

        # Cobertura: por cada tema que nadie desarrolla, el Jefe agrega una diapositiva con los puntos del Recopilador
        def cubiertos_de(lista):
            medio = lista[1:-1] if len(lista) >= 3 else lista
            return {x for sl in medio for x in sl["temas"]}
        cubiertos = cubiertos_de(nuevos)
        for k in range(1, self.n_temas + 1):
            if k in cubiertos or len(nuevos) >= MAX_DIAPOSITIVAS:
                continue
            sec = self.secciones[k - 1]
            puntos = []
            for pt in (sec.get("puntos") or [])[:8]:
                palabras = str(pt).split()
                puntos.append(" ".join(palabras[:15]) + ("…" if len(palabras) > 15 else ""))
            nueva = {"title": " ".join(str(sec.get("tema", f"Tema {k}")).split()[:8]), "tipo": "bullets",
                     "parrafo": "", "bullets": puntos, "temas": [k], "lleva_imagen": False}
            nuevos.insert(max(0, len(nuevos) - 1), nueva)   # antes de la conclusión
            cubiertos.add(k)
            cambios.append(f"agregó una diapositiva para el tema {k} («{sec.get('tema', '?')}») que nadie desarrollaba")

        marcadas = asegurar_imagenes(nuevos)
        if marcadas:
            cambios.append(f"marcó con imagen las diapositivas {marcadas} para llegar al "
                           f"{round(PROPORCION_IMAGENES * 100)}% y no dejar dos seguidas sin imagen")
        razon = salida.get("razonamiento")
        if not razon:
            cambios.append("completó el razonamiento faltante")
        return {**salida, "slides": nuevos, "razonamiento": razon or "(el agente no explicó su decisión)"}, cambios

    # ---- ciclo de supervisión de un agente ----
    def _supervisar(self, agente, entrada, opcional=False, respaldo=None):
        feedback, res = None, None
        for intento in range(1, MAX_INTENTOS + 1):
            t0 = time.time()
            res, revisiones = None, []
            self.ultimo_intento = intento == MAX_INTENTOS
            try:
                res = agente.ejecutar(entrada, feedback)
                revisiones = self.revisar(agente.nombre, res["salida"])
            except Exception as e:
                _chk(revisiones, "La ejecución termina sin errores", False, f"{type(e).__name__}: {e}")
            dt = time.time() - t0
            problemas = [r["detalle"] for r in revisiones if not r["ok"] and r["bloquea"]]

            if not problemas:
                self._anotar(agente.nombre, agente.rol, intento, "✅ Aprobado", dt, res, revisiones)
                return res["salida"]

            self._anotar(agente.nombre, agente.rol, intento, "❌ Rechazado por el Jefe", dt, res, revisiones,
                         notas="Motivos: " + "; ".join(problemas))
            feedback = "- " + "\n- ".join(problemas)

        # Se acabaron los intentos: el Jefe interviene si puede
        if agente.nombre == "Seccionador" and res:
            salida_ia, por_ia = res["salida"], False
            try:
                salida_ia, por_ia = self._corregir_plan_ia(res["salida"], feedback), True
            except Exception:
                pass                                      # si la IA no puede, el código corrige solo
            nuevo, cambios = self._corregir_plan(salida_ia)
            if por_ia:
                cambios.insert(0, "la IA rehízo el plan con las objeciones del Jefe")
            revisiones = self._checks_plan(nuevo["slides"], nuevo["razonamiento"], estricto=False)
            if all(r["ok"] or not r["bloquea"] for r in revisiones):
                self._anotar(
                    self.nombre, self.rol, "-", "🛠️ Corrigió él mismo", 0,
                    {
                        "recibio": f"Plan rechazado {MAX_INTENTOS} veces",
                        "entrego": f"Plan de {len(nuevo['slides'])} diapositivas corregido",
                        "razonamiento": "El Seccionador no logró cumplir las reglas tras varios intentos. Le pedí a la "
                                        "IA que rehiciera el plan con mis objeciones y luego apliqué yo mismo "
                                        "los ajustes mínimos que exigen las reglas.",
                        "detalle": "\n".join(sorted(set(cambios))) or
                                   ("Sin cambios de formato. Se aceptó con las observaciones pendientes:\n"
                                    + (feedback or "")),
                    },
                    revisiones,
                )
                return nuevo

        if opcional and respaldo is not None:
            self._anotar(self.nombre, self.rol, "-", "⚠️ Siguió sin este paso", 0,
                         {"recibio": f"{agente.nombre} falló {MAX_INTENTOS} veces",
                          "entrego": "Presentación sin imágenes",
                          "razonamiento": "Las imágenes son un extra, no un requisito para tener una "
                                          "presentación útil, así que continué sin ellas en vez de detener todo."},
                         notas=feedback or "")
            return respaldo

        self._anotar(self.nombre, self.rol, "-", "🛑 Detuvo el proceso", 0,
                     {"recibio": f"{agente.nombre} falló {MAX_INTENTOS} veces",
                      "razonamiento": "No pude corregirlo yo mismo, así que detengo el proceso para no "
                                      "entregar una presentación incorrecta."},
                     notas=feedback or "")
        raise ErrorPipeline(f"{agente.nombre} no logró entregar un resultado válido")

    # ---- cuántas diapositivas le tocan a cada tema ----
    def _calcular_presupuesto(self):
        self.presupuesto = presupuesto_diapositivas(self.secciones)
        self.objetivo = total_objetivo(self.presupuesto)
        self.n_min, self.n_max = rango_diapositivas(self.presupuesto)
        detalle = [f"Tema {f['numero']} «{f['tema']}»: {f['caracteres']} caracteres → {f['diapositivas']} diapositiva(s)"
                   for f in self.presupuesto]
        detalle.append(f"Más introducción y conclusión: {self.objetivo} diapositivas de contenido (referencia)")
        self._anotar(self.nombre, "Calcula una sugerencia de diapositivas por tema y fija los límites de su veto", "-",
                     "📐 Sugerencia de diapositivas", 0,
                     {"recibio": f"{self.n_temas} temas del Recopilador",
                      "entrego": f"Sugerencia de {self.objetivo} diapositivas; la IA decide el reparto y yo vetaría "
                                 f"menos de {self.n_min} (o más de {self.n_max})",
                      "razonamiento": "La fórmula de caracteres de cada tema es solo una referencia para el Seccionador. "
                                      "Él decide cuántas diapositivas merece cada tema y yo reviso que se mantenga dentro "
                                      "de los límites.",
                      "detalle": "\n".join(detalle)})

    # ---- el flujo completo ----
    def coordinar(self, pdf_path, ruta_pptx, carpeta):
        lectura = self._supervisar(Lector(carpeta), pdf_path)
        self.texto_original = lectura["texto"]
        self.rango_temas = rango_de_temas(max(1, int(lectura.get("paginas", 1))))
        self.texto_pdf = (lectura["texto"] + "\n" + lectura.get("texto_plumber", "")).lower()
        resumen = self._supervisar(Recopilador(self.rango_temas), lectura)
        self.secciones = resumen.get("secciones", [])
        self.n_temas = len(self.secciones)
        self._calcular_presupuesto()
        plan = self._supervisar(Seccionador(self.presupuesto), resumen)
        plan = {**plan, "documento": os.path.basename(pdf_path), "referencias": resumen.get("referencias", []),
                "imagenes_pdf": lectura.get("imagenes_pdf", []), "secciones": self.secciones}
        respaldo = {**plan, "slides": [{k: v for k, v in s.items() if k != "imagen"} for s in plan["slides"]]}
        plan = self._supervisar(Ilustrador(carpeta), plan, opcional=True, respaldo=respaldo)
        base = self._supervisar(GeneradorPPTX(ruta_pptx), plan)
        final = self._supervisar(Decorador(), base)
        self._revision_final(final["ruta"])
        return final["ruta"]

    # ---- reporte para ti (el "jefe de jefes") ----
    def reporte_html(self, exito):
        e = escape
        rechazos = sum(1 for x in self.log if "Rechazado" in x["estado"])
        correcciones = sum(1 for x in self.log if "Corrigió" in x["estado"])
        omitidos = sum(1 for x in self.log if "Siguió" in x["estado"])
        total = round(time.time() - self.t_inicio, 1)
        caja = "border:1px solid rgba(128,128,128,.4);border-radius:8px;margin:8px 0;padding:8px 12px;"
        celda = "border:1px solid rgba(128,128,128,.35);padding:6px 8px;text-align:left;"

        h = ['<div style="max-height:75vh;overflow-y:auto;padding-right:10px;">']
        h.append("<h3>Informe final del Jefe</h3><ul>")
        h.append(f"<li>Resultado: {'✅ presentación generada' if exito else '🛑 el proceso se detuvo'}</li>")
        h.append(f"<li>Rechazos a agentes: {rechazos}</li>")
        h.append(f"<li>Correcciones hechas por el Jefe: {correcciones}</li>")
        h.append(f"<li>Pasos opcionales omitidos: {omitidos}</li>")
        h.append(f"<li>Tiempo total: {total} s</li></ul>")

        h.append("<h3>Resumen de la corrida</h3>")
        h.append('<table style="border-collapse:collapse;width:100%;"><tr>'
                 + "".join(f'<th style="{celda}">{t}</th>' for t in ["#", "Agente", "Intento", "Estado", "Seg"])
                 + "</tr>")
        for x in self.log:
            h.append("<tr>" + "".join(
                f'<td style="{celda}">{e(str(v))}</td>'
                for v in [x["paso"], x["agente"], x["intento"], x["estado"], x["segundos"]]
            ) + "</tr>")
        h.append("</table>")

        h.append("<h3>Detalle por agente (clic para abrir)</h3>")
        for x in self.log:
            abierto = " open" if any(k in x["estado"] for k in ("Rechazado", "Corrigió", "Detuvo", "Siguió", "Revisión final")) else ""
            h.append(f'<details{abierto} style="{caja}"><summary style="cursor:pointer;font-weight:600;">'
                     f'#{x["paso"]} {e(x["agente"])} — {e(x["estado"])} '
                     f'(intento {e(str(x["intento"]))}, {x["segundos"]} s)</summary>')
            h.append(f'<p><b>Qué hace:</b> {e(x["rol"])}</p>')
            h.append(f'<p><b>Recibió:</b> {e(x["recibio"])}</p>')
            h.append(f'<p><b>Entregó:</b> {e(x["entrego"])}</p>')
            if x["razonamiento"]:
                h.append(f'<p><b>Razonamiento:</b> {e(x["razonamiento"])}</p>')
            if x["notas"]:
                h.append(f'<p><b>Notas:</b> {e(x["notas"])}</p>')
            if x["detalle"]:
                h.append('<p><b>Detalle del trabajo:</b></p>'
                         f'<pre style="white-space:pre-wrap;margin:0;padding:8px;'
                         f'border:1px solid rgba(128,128,128,.35);border-radius:6px;">{e(x["detalle"])}</pre>')
            if x["revisiones"]:
                h.append("<p><b>Revisiones del Jefe:</b></p><ul>")
                for r in x["revisiones"]:
                    marca = "✅" if r["ok"] else ("❌" if r["bloquea"] else "⚠️")
                    extra = f" — {e(r['detalle'])}" if not r["ok"] else ""
                    h.append(f"<li>{marca} {e(r['regla'])}{extra}</li>")
                h.append("</ul>")
            h.append("</details>")
        h.append("</div>")
        return "".join(h)
