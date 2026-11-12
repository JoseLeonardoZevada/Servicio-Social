"""Agente Seccionador: reparte el contenido en diapositivas (según los temas y su extensión)
y marca cuáles llevan imagen."""
import json

from utils import Agente, asegurar_imagenes, con_feedback, llamar_llm, parsear_json
from config import MAX_DIAPOSITIVAS

MAX_POR_LOTE = 10   # diapositivas que se piden en una sola llamada a la IA


SYS_SECCIONADOR = """
Eres un asistente experto en diseño de presentaciones. Decides qué información va en cada
diapositiva y en qué formato, pensando en explicarlas en voz alta frente a una audiencia.

Reglas que debes seguir SIEMPRE:
- Responde ÚNICAMENTE con JSON válido. Nada de texto extra ni ``` .
- Español neutro, tono claro y profesional.
- NO incluyas portada ni diapositiva de fuentes. Tampoco buscas las imágenes: otro agente se
  encarga; tú solo marcas en cuáles diapositivas va una.
- Cada "title" máximo 8 palabras.
- Cada diapositiva tiene un "tipo":
    * "parrafo": un texto corrido de 25 a 55 palabras, redactado con fluidez
      (nada de listas). Deja "bullets" vacío.
    * "mixto": un "parrafo" corto (10 a 30 palabras) más 2 a 5 bullets.
    * "bullets": 3 a 8 bullets. Úsalo cuando el contenido sea naturalmente una lista
      (ingredientes, pasos, requisitos). En el resto de los casos prefiere texto corrido.
  Como máximo el 60% de las diapositivas pueden ser solo "bullets".
- Cada bullet máximo 15 palabras, frase corta. No repitas ideas entre diapositivas.
- Respeta EXACTAMENTE el número de diapositivas que se te pide para cada tema. Si un tema pide
  varias, reparte sus puntos entre ellas sin repetir ideas y sin omitir datos importantes.
  Usa los datos concretos de los puntos (cifras, nombres, definiciones, pasos); no los generalices.
- En "temas" pon los números (campo "numero") del tema que trata cada diapositiva.
- "lleva_imagen": true si la diapositiva trata algo concreto que una imagen ilustre bien (un objeto,
  un lugar, un proceso, un ejemplo, un esquema); false si es muy abstracta. Procura alternar
  (una con imagen, otra sin, otra con imagen...) y que en total al menos el 60% lleve imagen.
- Explica tu decisión en "razonamiento".
"""

SYS_REPARTO = """
Eres un editor de presentaciones. Decides cuántas diapositivas merece cada tema de un documento,
según su densidad, su importancia y cuánta información trae. Responde ÚNICAMENTE con JSON válido.
"""

MAX_DIAP_POR_TEMA_IA = 6

EJEMPLO_SLIDES = {
    "razonamiento": "Dos diapositivas: una explicativa del tema 1 y una de cierre.",
    "slides": [
        {
            "title": "¿Qué es el reciclaje?",
            "temas": [1],
            "tipo": "parrafo",
            "parrafo": "El reciclaje convierte materiales de desecho en nuevos productos. Así se "
                       "reduce el consumo de materias primas y se disminuye la contaminación del aire "
                       "y del agua, además de ahorrar energía en la producción.",
            "bullets": [],
            "lleva_imagen": True,
        },
        {
            "title": "Conclusión",
            "temas": [1, 2],
            "tipo": "mixto",
            "parrafo": "Reciclar bien empieza en casa y beneficia a todos.",
            "bullets": ["Separar los materiales desde el hogar", "Reducir el impacto ambiental"],
            "lleva_imagen": False,
        },
    ],
}


def _agrupar_lotes(presupuesto, maximo=MAX_POR_LOTE):
    """Junta temas consecutivos en bloques de hasta 'maximo' diapositivas."""
    lotes, actual, suma = [], [], 0
    for f in presupuesto:
        if actual and suma + f["diapositivas"] > maximo:
            lotes.append(actual)
            actual, suma = [], 0
        actual.append(f)
        suma += f["diapositivas"]
    if actual:
        lotes.append(actual)
    return lotes


def _validos(temas, n_temas):
    """Deja solo números de tema válidos (1 a n_temas)."""
    if not isinstance(temas, list):
        return []
    return [t for t in temas if isinstance(t, int) and not isinstance(t, bool) and 1 <= t <= n_temas]


class Seccionador(Agente):
    nombre = "Seccionador"
    rol = ("La IA decide cuántas diapositivas merece cada tema, redacta cada una por bloques de temas "
           "y marca cuáles llevan imagen")

    def __init__(self, presupuesto):
        self.presupuesto = presupuesto

    def _repartir(self, entrada, feedback):
        """La IA decide cuántas diapositivas lleva cada tema. La fórmula de caracteres solo es una sugerencia;
        si la IA falla, se usa la fórmula."""
        secciones = entrada.get("secciones", [])
        temas = [{"numero": f["numero"], "tema": f["tema"], "caracteres": f["caracteres"],
                  "sugeridas": f["diapositivas"],
                  "puntos": (secciones[f["numero"] - 1].get("puntos") if f["numero"] - 1 < len(secciones) else [])}
                 for f in self.presupuesto]
        sugerido = sum(f["diapositivas"] for f in self.presupuesto)
        prompt = (
            "Decide cuántas diapositivas lleva cada tema de esta presentación. Cada tema lleva al menos 1 y "
            f"como máximo {MAX_DIAP_POR_TEMA_IA}. Da más a los temas densos o importantes y menos a los ligeros "
            "o repetitivos. 'sugeridas' sale de una fórmula por caracteres: úsala como referencia, no como "
            f"obligación (suman {sugerido}; sin contar introducción y conclusión). Tu total debería quedar cerca "
            "de esa suma, pero tú decides.\n"
            'Responde {"razonamiento": "<criterio general>", "reparto": [{"numero": N, "diapositivas": M, '
            '"por_que": "<una frase>"}]}\n\nTEMAS:\n' + json.dumps(temas, ensure_ascii=False)
        )
        try:
            d = parsear_json(llamar_llm(SYS_REPARTO, con_feedback(prompt, feedback)))
            por_num = {x["numero"]: x for x in d.get("reparto", [])
                       if isinstance(x, dict) and isinstance(x.get("numero"), int)}
        except Exception as e:
            return [dict(f, sugeridas=f["diapositivas"], por_que="la IA no respondió; usé la fórmula")
                    for f in self.presupuesto], f"La IA no pudo repartir ({type(e).__name__}); usé la fórmula."
        nuevo = []
        for f in self.presupuesto:
            x = por_num.get(f["numero"]) or {}
            m = x.get("diapositivas")
            usable = isinstance(m, int) and not isinstance(m, bool) and 1 <= m <= MAX_DIAP_POR_TEMA_IA
            nuevo.append({**f, "sugeridas": f["diapositivas"], "diapositivas": m if usable else f["diapositivas"],
                          "por_que": str(x.get("por_que", "")) if usable else "sin dato de la IA; usé la fórmula"})
        while sum(f["diapositivas"] for f in nuevo) > MAX_DIAPOSITIVAS - 2:      # no pasar del tope absoluto
            mayor = max(nuevo, key=lambda f: f["diapositivas"])
            if mayor["diapositivas"] <= 1:
                break
            mayor["diapositivas"] -= 1
        return nuevo, str(d.get("razonamiento", ""))

    def _pedir_lote(self, entrada, esquema, lote, primero, ultimo, esperadas, inicio, feedback):
        secciones = entrada.get("secciones", [])
        lote_json = [{"numero": f["numero"], "tema": f["tema"],
                      "puntos": (secciones[f["numero"] - 1].get("puntos") if f["numero"] - 1 < len(secciones) else []),
                      "diapositivas": f["diapositivas"]} for f in lote]
        base = (
            "Ejemplo del formato esperado (solo la estructura, no el contenido):\n"
            + json.dumps(EJEMPLO_SLIDES, ensure_ascii=False)
            + f"\n\nEstás armando una parte de una presentación sobre «{entrada.get('titulo_documento', '')}» "
              f"({entrada.get('tema_general', '')}).\n"
              "ESQUEMA COMPLETO de temas (para que lo que escribas encaje con lo anterior y lo siguiente):\n"
            + json.dumps(esquema, ensure_ascii=False)
            + f"\n\nEste bloque produce EXACTAMENTE {esperadas} diapositivas: las números {inicio} a "
              f"{inicio + esperadas - 1} del plan completo.\n"
            + ("1) Empieza con UNA diapositiva de introducción que presente el tema general.\n" if primero
               else "1) No hagas introducción: continúa la explicación con el tema que sigue.\n")
            + "2) Luego, para CADA tema de abajo y en su orden, escribe exactamente las diapositivas que indica "
              "su campo 'diapositivas'. Si son varias, reparte los puntos del tema entre ellas sin repetir ideas "
              "y sin omitir datos importantes.\n"
            + ("3) Termina con UNA diapositiva de resumen o conclusión de todo el documento.\n" if ultimo
               else "3) No hagas conclusión: aún faltan temas por tratar.\n")
            + 'Devuelve SOLO el JSON {"razonamiento": "...", "slides": [{"title": "...", "temas": [1], '
              '"tipo": "...", "parrafo": "...", "bullets": ["..."], "lleva_imagen": true}]}.\n\n'
              "TEMAS DE ESTE BLOQUE:\n" + json.dumps(lote_json, ensure_ascii=False)
        )
        aviso, previo = "", None
        for intento in range(2):
            try:
                data = parsear_json(llamar_llm(SYS_SECCIONADOR, con_feedback(base + aviso, feedback)))
                slides = data["slides"]
                if not (isinstance(slides, list) and slides and all(isinstance(x, dict) for x in slides)):
                    raise ValueError("'slides' está vacío o mal formado")
                if abs(len(slides) - esperadas) <= 1 or intento == 1:
                    return data
                previo = data
                aviso = (f"\n\nTu respuesta anterior tenía {len(slides)} diapositivas y este bloque necesita "
                         f"exactamente {esperadas}. Corrígelo.")
            except Exception as e:
                if intento == 1:
                    if previo is not None:
                        return previo
                    raise
                aviso = (f"\n\nTu respuesta anterior no se pudo usar ({type(e).__name__}: {e}). "
                         "Responde SOLO el JSON pedido.")

    def _ajustar_temas(self, nuevas, lote, primero, ultimo, esperadas, n_temas):
        """Si el bloque trae el número pedido de diapositivas, cada una queda asignada a su tema (el orden
        se pidió explícitamente); si no, se conservan los números que puso la IA, corregidos."""
        if len(nuevas) == esperadas:
            asignacion = [f["numero"] for f in lote for _ in range(f["diapositivas"])]
            desfase = 1 if primero else 0
            for k, num in enumerate(asignacion):
                nuevas[desfase + k]["temas"] = [num]
            if primero:
                nuevas[0]["temas"] = _validos(nuevas[0].get("temas"), n_temas) or [lote[0]["numero"]]
            if ultimo:
                nuevas[-1]["temas"] = _validos(nuevas[-1].get("temas"), n_temas) or [lote[-1]["numero"]]
        else:
            for sl in nuevas:
                sl["temas"] = _validos(sl.get("temas"), n_temas) or [lote[0]["numero"]]

    def ejecutar(self, entrada, feedback=None):
        presupuesto, razon_reparto = self._repartir(entrada, feedback)
        lotes = _agrupar_lotes(presupuesto)
        esquema = [{"numero": f["numero"], "tema": f["tema"]} for f in presupuesto]
        n_temas = len(presupuesto)
        slides, razones = [], []
        for j, lote in enumerate(lotes):
            primero, ultimo = j == 0, j == len(lotes) - 1
            esperadas = sum(f["diapositivas"] for f in lote) + (1 if primero else 0) + (1 if ultimo else 0)
            data = self._pedir_lote(entrada, esquema, lote, primero, ultimo, esperadas, len(slides) + 1, feedback)
            nuevas = [dict(x) for x in data["slides"]]
            self._ajustar_temas(nuevas, lote, primero, ultimo, esperadas, n_temas)
            slides += nuevas
            razones.append(str(data.get("razonamiento", "")).strip())

        marcadas_ia = sum(1 for sl in slides if str(sl.get("lleva_imagen")).strip().lower() in ("true", "1", "si", "sí"))
        agregadas = asegurar_imagenes(slides)       # cuota del 60% y reparto: se completa en código
        n_img = sum(1 for sl in slides if sl["lleva_imagen"])
        total_chars = sum(f["caracteres"] for f in presupuesto)
        objetivo = sum(f["diapositivas"] for f in presupuesto) + 2

        detalle = ["Reparto decidido por la IA (la fórmula de caracteres del Jefe era solo una sugerencia):"]
        for f in presupuesto:
            detalle.append(f"   Tema {f['numero']} «{f['tema']}»: {f['caracteres']} caracteres, fórmula "
                           f"{f['sugeridas']} → IA {f['diapositivas']} diapositiva(s). {f.get('por_que', '')}")
        detalle.append(f"   Criterio: {razon_reparto}")
        detalle.append(f"   Total: {objetivo} (con introducción y conclusión). Se trabajó en {len(lotes)} bloque(s) de temas.")
        detalle.append("")
        for i, sl in enumerate(slides, 1):
            temas = ", ".join(str(t) for t in sl["temas"])
            marca = " 🖼" if sl["lleva_imagen"] else ""
            detalle.append(f"Diapositiva {i} [{sl.get('tipo', '?')}]{marca}: {sl.get('title', '?')}   (tema: {temas})")
            if sl.get("parrafo"):
                detalle.append(f"   Texto: {sl['parrafo']}")
            for b in sl.get("bullets", []) or []:
                detalle.append(f"   - {b}")

        razon_img = f"Marqué imagen en {n_img} de {len(slides)} diapositivas"
        if agregadas:
            razon_img += (f" (la IA marcó {marcadas_ia}; completé a mano las diapositivas {agregadas} para llegar "
                          "al 60% y no dejar dos seguidas sin imagen)")
        razonamiento = (
            f"Decidí cuántas diapositivas merece cada tema ({n_temas} temas, {total_chars} caracteres de contenido): "
            f"{objetivo - 2} diapositivas de tema, más introducción y conclusión. {razon_reparto} "
            f"Trabajé por bloques de temas para no perder ninguno. {razon_img}. "
            + " ".join(r for r in razones if r)
        ).strip()

        return {
            "salida": {
                "slides": slides,
                "titulo": entrada.get("titulo_documento", "Presentación"),
                "subtitulo": entrada.get("subtitulo", ""),
                "tema_general": entrada.get("tema_general", ""),
                "razonamiento": razonamiento,
            },
            "recibio": f"{n_temas} temas del Recopilador y la fórmula de caracteres como sugerencia",
            "entrego": f"Plan de {len(slides)} diapositivas de contenido, {n_img} con imagen",
            "razonamiento": razonamiento,
            "detalle": "\n".join(detalle),
            "notas": " / ".join(sl.get("title", "?") for sl in slides),
        }
