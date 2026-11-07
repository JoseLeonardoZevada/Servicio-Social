"""Agente Ilustrador: elige qué diapositivas llevan imagen y las trae de Wikimedia Commons."""
import io
import json
import os

from PIL import Image

from agentes.base import Agente
from commons import buscar_commons, descargar_imagen
from utils import _max_imagenes, con_feedback, llamar_llm, parsear_json


SYS_ILUSTRADOR_PLAN = """
Eres un editor gráfico que planea las imágenes de una presentación. Eliges solo las
diapositivas donde una imagen real aporta valor, y escribes búsquedas precisas.
Responde ÚNICAMENTE con JSON válido.
"""

SYS_ILUSTRADOR = """
Eres un editor gráfico muy estricto. Eliges imágenes para diapositivas, pero solo
si muestran de forma clara y directa el tema. Si hay duda, no eliges ninguna.
Responde ÚNICAMENTE con JSON válido.
"""


class Ilustrador(Agente):
    nombre = "Ilustrador"
    rol = ("Decide qué diapositivas llevan imagen y qué buscar; la trae de Wikimedia Commons (licencia libre), "
           "verifica con IA (viendo la imagen) que ilustre el tema y guarda autor y licencia")

    def __init__(self, carpeta):
        self.carpeta = carpeta

    def _planear(self, slides, cap, feedback):
        resumen = [{"diapositiva": i, "titulo": sl.get("title", ""),
                    "texto": (sl.get("parrafo") or " ".join(sl.get("bullets", [])))[:160]}
                   for i, sl in enumerate(slides, 1)]
        prompt = (
            f"Estas son las diapositivas de una presentación ({len(slides)} en total). Elige hasta {cap} "
            "diapositivas que se beneficien de una imagen concreta y real (un objeto, lugar, escena, "
            "alimento, diagrama...), no de conceptos abstractos. No pongas imagen a todas: elige solo "
            "donde la imagen ilustre claramente el contenido. Para cada una escribe una búsqueda corta "
            "EN INGLÉS (2 a 5 palabras) para Wikimedia Commons.\n"
            'Responde {"razonamiento": "<por qué elegiste esas diapositivas>", '
            '"imagenes": [{"diapositiva": N, "consulta": "..."}]}\n\n'
            "DIAPOSITIVAS:\n" + json.dumps(resumen, ensure_ascii=False)
        )
        d = parsear_json(llamar_llm(SYS_ILUSTRADOR_PLAN, con_feedback(prompt, feedback)))
        plan = [{"diapositiva": x.get("diapositiva"), "consulta": str(x.get("consulta", "")).strip()}
                for x in d.get("imagenes", []) if isinstance(x, dict)]
        return str(d.get("razonamiento", "")), plan

    def _elegir(self, slide, candidatas):
        texto = slide.get("parrafo") or " ".join(slide.get("bullets", []))
        partes = [
            f"Diapositiva: «{slide.get('title', '')}». Contenido: {texto}\n"
            "Estas imágenes vienen de Wikimedia Commons. Elige la que muestre de forma clara y "
            "directa el tema de la diapositiva. Descarta las que solo se relacionen vagamente, "
            "logotipos, capturas con texto ilegible o imágenes que no se entiendan. Si NINGUNA "
            'cumple, devuelve null. Responde {"elegida": <número de imagen o null>, '
            '"razon": "<una frase: qué se ve y por qué se relaciona con el tema>"}'
        ]
        for k, c in enumerate(candidatas):
            partes.append(f"Imagen {k} (archivo: {c['titulo']}):")
            partes.append({"mime_type": c["mime"], "data": c["bytes"]})
        d = parsear_json(llamar_llm(SYS_ILUSTRADOR, partes))
        k = d.get("elegida")
        if isinstance(k, int) and 0 <= k < len(candidatas):
            return candidatas[k], d.get("razon", "")
        return None, d.get("razon", "ninguna imagen ilustraba el tema con claridad")

    def ejecutar(self, entrada, feedback=None):
        slides = [dict(sl) for sl in entrada["slides"]]
        for sl in slides:
            sl.pop("imagen", None)
        cap = _max_imagenes(len(slides))
        razon_plan, plan = self._planear(slides, cap, feedback)

        lineas = [f"Cupo de imágenes: hasta {cap} (para {len(slides)} diapositivas)",
                  f"Por qué eligió esas diapositivas: {razon_plan}",
                  "Plan: " + (" · ".join(f"diap. {p['diapositiva']} → «{p['consulta']}»" for p in plan) or "(vacío)"),
                  ""]
        validos = [p for p in plan if isinstance(p["diapositiva"], int)
                   and 1 <= p["diapositiva"] <= len(slides) and p["consulta"]][:cap]
        errores, elegidas = [], 0

        for p in validos:
            i = p["diapositiva"] - 1
            sl = slides[i]
            etiqueta = f"Diap. {i + 1} «{sl.get('title', '')}» (búsqueda: {p['consulta']})"
            try:
                cands = buscar_commons(p["consulta"])
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
                    lineas.append(f"{etiqueta}: sin candidatas utilizables ({len(cands)} con licencia y autor)")
                    continue
                elegida, razon = self._elegir(sl, bajadas)
                if not elegida:
                    lineas.append(f"{etiqueta}: ninguna sirvió → {razon}")
                    continue
                ext = ".png" if elegida["mime"] == "image/png" else ".jpg"
                ruta = os.path.join(self.carpeta, f"imagen_{i + 1}{ext}")
                with open(ruta, "wb") as f:
                    f.write(elegida["bytes"])
                ancho, alto = Image.open(io.BytesIO(elegida["bytes"])).size
                sl["imagen"] = {
                    "ruta": ruta, "titulo": elegida["titulo"], "autor": elegida["autor"],
                    "licencia": elegida["licencia"], "pagina": elegida["pagina"],
                    "razon": razon, "ancho": ancho, "alto": alto,
                }
                elegidas += 1
                lineas.append(f"{etiqueta}: ✔ {elegida['titulo']} — {elegida['autor']}, "
                              f"{elegida['licencia']} → {razon}")
            except Exception as e:
                errores.append(str(e))
                lineas.append(f"{etiqueta}: error ({type(e).__name__}: {e})")

        if validos and len(errores) == len(validos):
            raise RuntimeError("No se pudo consultar Wikimedia Commons: " + errores[0])

        return {
            "salida": {**entrada, "slides": slides, "plan_imagenes": plan, "max_imagenes": cap,
                       "razonamiento_imagenes": razon_plan},
            "recibio": f"Plan de {len(slides)} diapositivas (sin imágenes)",
            "entrego": f"{elegidas} imágenes aceptadas de {len(validos)} búsquedas",
            "razonamiento": (
                "Primero leí el plan y elegí las diapositivas donde una imagen real aporta valor, con su "
                "búsqueda. Luego consulté Wikimedia Commons (solo imágenes con autor y licencia libre), "
                "descargué hasta 3 candidatas por búsqueda y le pedí a la IA que las viera y eligiera solo una "
                "que muestre claramente el tema; si ninguna lo hacía, esa diapositiva queda sin imagen. "
                f"Selección: {razon_plan}"
            ),
            "detalle": "\n".join(lineas),
            "notas": "",
        }
