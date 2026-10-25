import os
import json
import re
import time
import tempfile

import gradio as gr
import pdfplumber
import google.generativeai as genai
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.dml import MSO_FILL

# La API key se lee de la variable de entorno configurada en Render
genai.configure(api_key=os.environ["GEMINI_API_KEY"])
MODEL_NAME = os.environ.get("GEMINI_MODEL", "gemini-3.5-flash-lite")

MAX_INTENTOS = 3  # cuántas veces el Jefe deja reintentar a un agente


# ----------------------------------------------------------------------
# Utilidades
# ----------------------------------------------------------------------
def llamar_llm(system_instruction, prompt):
    """Una llamada a Gemini con la instrucción de sistema del agente."""
    model = genai.GenerativeModel(MODEL_NAME, system_instruction=system_instruction)
    return model.generate_content(prompt).text.strip()


def parsear_json(raw):
    """Convierte la respuesta del modelo en JSON, quitando ``` si los puso."""
    raw = re.sub(r"^```(?:json)?", "", raw.strip())
    raw = re.sub(r"```$", "", raw).strip()
    return json.loads(raw)


def num_palabras(texto):
    return len(str(texto).split())


def con_feedback(prompt, feedback):
    """Si el Jefe rechazó un intento anterior, se lo decimos al agente."""
    if not feedback:
        return prompt
    return (
        prompt
        + "\n\nIMPORTANTE: el Jefe rechazó tu intento anterior por estos motivos:\n"
        + feedback
        + "\nCorrígelo en este nuevo intento."
    )


# ----------------------------------------------------------------------
# Agentes (cada uno hace UNA cosa y devuelve qué recibió/entregó)
# ----------------------------------------------------------------------
class Agente:
    nombre = "Agente"
    rol = ""

    def ejecutar(self, entrada, feedback=None):
        """Debe devolver {"salida":..., "recibio":str, "entrego":str, "notas":str}"""
        raise NotImplementedError


class Lector(Agente):
    nombre = "Lector"
    rol = "Extrae el texto del PDF página por página (sin IA)"

    def ejecutar(self, entrada, feedback=None):
        paginas = []
        with pdfplumber.open(entrada) as pdf:
            for p in pdf.pages:
                paginas.append(p.extract_text() or "")
        texto = "\n".join(t for t in paginas if t)
        vacias = sum(1 for t in paginas if not t)
        return {
            "salida": {"texto": texto, "paginas": len(paginas)},
            "recibio": f"PDF: {os.path.basename(entrada)}",
            "entrego": f"{len(paginas)} páginas, {len(texto)} caracteres de texto",
            "notas": f"{vacias} páginas sin texto extraíble",
        }


SYS_RECOPILADOR = """
Eres un analista que lee documentos y extrae su información clave.
Responde ÚNICAMENTE con JSON válido, sin texto extra ni marcado de código.
Escribe en español neutro. No inventes información que no esté en el texto.
"""


class Recopilador(Agente):
    nombre = "Recopilador"
    rol = "Resume el texto y lo agrupa en temas con puntos clave (IA)"

    def ejecutar(self, entrada, feedback=None):
        esquema = {
            "tema_general": "...",
            "secciones": [{"tema": "...", "puntos": ["...", "..."]}],
        }
        prompt = (
            "Lee el siguiente documento y extrae su información clave, agrupada "
            "por temas (entre 4 y 8 temas, cada uno con 2 a 5 puntos).\n"
            "Estructura exacta: " + json.dumps(esquema, ensure_ascii=False) + "\n\n"
            'DOCUMENTO:\n"""\n' + entrada["texto"][:15000] + '\n"""'
        )
        data = parsear_json(llamar_llm(SYS_RECOPILADOR, con_feedback(prompt, feedback)))
        n_puntos = sum(len(s.get("puntos", [])) for s in data.get("secciones", []))
        return {
            "salida": data,
            "recibio": f"{len(entrada['texto'])} caracteres de texto",
            "entrego": f"{len(data.get('secciones', []))} temas, {n_puntos} puntos clave",
            "notas": f"Tema general: {data.get('tema_general', '?')}",
        }


SYS_SECCIONADOR = """
Eres un asistente experto en diseño de presentaciones. Decides cuántas
diapositivas se necesitan y qué información va en cada una, pensando en
explicarlas en voz alta frente a una audiencia.

Reglas que debes seguir SIEMPRE:
- Responde ÚNICAMENTE con JSON válido. Nada de texto extra ni ``` .
- Español neutro, tono claro y profesional.
- Cada "title" máximo 8 palabras.
- Cada "bullet" máximo 15 palabras, frase corta, nunca un párrafo.
- No repitas la misma idea en dos diapositivas.
- La primera diapositiva presenta el tema general (introducción).
- La última diapositiva es un resumen o conclusión.
"""

EJEMPLO_SLIDES = {
    "slides": [
        {
            "title": "¿Qué es el reciclaje?",
            "bullets": [
                "Convierte materiales de desecho en nuevos productos",
                "Reduce el consumo de materias primas frescas",
                "Disminuye la contaminación del aire y el agua",
            ],
        },
        {
            "title": "Conclusión",
            "bullets": [
                "El reciclaje reduce el impacto ambiental",
                "La separación desde el hogar es clave",
                "Todos podemos participar",
            ],
        },
    ]
}


class Seccionador(Agente):
    nombre = "Seccionador"
    rol = "Decide cuántas diapositivas hay y qué contenido lleva cada una (IA)"

    def ejecutar(self, entrada, feedback=None):
        prompt = (
            "Ejemplo del formato esperado (solo la estructura, no el contenido):\n"
            + json.dumps(EJEMPLO_SLIDES, ensure_ascii=False)
            + "\n\nAhora, con base en este resumen estructurado del documento, "
            "arma entre 5 y 10 diapositivas, cada una con entre 3 y 5 bullets. "
            'Devuelve SOLO el JSON {"slides": [{"title": "...", "bullets": ["..."]}]}.\n\n'
            "RESUMEN:\n" + json.dumps(entrada, ensure_ascii=False)
        )
        data = parsear_json(llamar_llm(SYS_SECCIONADOR, con_feedback(prompt, feedback)))
        slides = data["slides"]
        titulos = " | ".join(s.get("title", "?") for s in slides)
        return {
            "salida": {"slides": slides, "tema_general": entrada.get("tema_general", "")},
            "recibio": f"{len(entrada.get('secciones', []))} temas del Recopilador",
            "entrego": f"Plan de {len(slides)} diapositivas",
            "notas": titulos,
        }


class GeneradorPPTX(Agente):
    nombre = "Generador PPTX"
    rol = "Construye el archivo .pptx a partir del plan (sin IA)"

    def __init__(self, ruta_salida):
        self.ruta_salida = ruta_salida

    def ejecutar(self, entrada, feedback=None):
        prs = Presentation()
        layout = prs.slide_layouts[1]
        for info in entrada["slides"]:
            slide = prs.slides.add_slide(layout)
            slide.shapes.title.text = info.get("title", "Sin titulo")
            tf = slide.placeholders[1].text_frame
            tf.clear()
            for i, bullet in enumerate(info.get("bullets", [])):
                if i == 0:
                    tf.text = bullet
                else:
                    tf.add_paragraph().text = bullet
        prs.save(self.ruta_salida)
        return {
            "salida": {
                "ruta": self.ruta_salida,
                "n_slides": len(entrada["slides"]),
                "tema": entrada.get("tema_general", ""),
            },
            "recibio": f"Plan de {len(entrada['slides'])} diapositivas",
            "entrego": f"Archivo {os.path.basename(self.ruta_salida)} (plantilla básica)",
            "notas": "",
        }


PALETAS = {
    "oceano": {"fondo": (0xF4, 0xF8, 0xFB), "titulo": (0x0B, 0x3C, 0x5D), "texto": (0x33, 0x33, 0x33), "acento": (0x32, 0x82, 0xB8)},
    "bosque": {"fondo": (0xF3, 0xF8, 0xF3), "titulo": (0x1B, 0x4D, 0x2E), "texto": (0x2E, 0x2E, 0x2E), "acento": (0x4C, 0xAF, 0x50)},
    "atardecer": {"fondo": (0xFF, 0xF6, 0xEE), "titulo": (0x8A, 0x2B, 0x0F), "texto": (0x3A, 0x2E, 0x2A), "acento": (0xE6, 0x7E, 0x22)},
    "grafito": {"fondo": (0xF2, 0xF2, 0xF4), "titulo": (0x22, 0x22, 0x2B), "texto": (0x3C, 0x3C, 0x44), "acento": (0x6C, 0x5C, 0xE7)},
}

SYS_DECORADOR = """
Eres un diseñador de presentaciones. Eliges la paleta de colores que mejor
va con el tema del documento. Responde ÚNICAMENTE con JSON válido.
"""


class Decorador(Agente):
    nombre = "Decorador"
    rol = "Elige paleta (IA) y aplica colores, fuentes y barra de acento (código)"

    def ejecutar(self, entrada, feedback=None):
        ruta, tema = entrada["ruta"], entrada["tema"]
        nombre_paleta, motivo = "oceano", "paleta por defecto"
        try:
            prompt = (
                f"Tema del documento: {tema}\n"
                f"Paletas disponibles: {list(PALETAS)}\n"
                'Responde {"paleta": "<nombre>", "motivo": "<una frase corta>"}'
            )
            d = parsear_json(llamar_llm(SYS_DECORADOR, prompt))
            if d.get("paleta") in PALETAS:
                nombre_paleta, motivo = d["paleta"], d.get("motivo", "")
        except Exception as e:  # si la IA falla, no tiramos todo el proceso
            motivo = f"la IA falló ({e}); usé la paleta por defecto"

        p = PALETAS[nombre_paleta]
        prs = Presentation(ruta)
        for slide in prs.slides:
            fill = slide.background.fill
            fill.solid()
            fill.fore_color.rgb = RGBColor(*p["fondo"])

            barra = slide.shapes.add_shape(
                MSO_SHAPE.RECTANGLE, 0, 0, Inches(0.3), prs.slide_height
            )
            barra.fill.solid()
            barra.fill.fore_color.rgb = RGBColor(*p["acento"])
            barra.line.fill.background()

            for shape in slide.placeholders:
                if not shape.has_text_frame:
                    continue
                es_titulo = shape.placeholder_format.idx == 0
                for par in shape.text_frame.paragraphs:
                    for run in par.runs:
                        run.font.size = Pt(36 if es_titulo else 24)
                        run.font.bold = es_titulo
                        run.font.color.rgb = RGBColor(*(p["titulo"] if es_titulo else p["texto"]))
        prs.save(ruta)
        return {
            "salida": {"ruta": ruta, "n_slides": len(prs.slides), "paleta": nombre_paleta},
            "recibio": f"Archivo básico de {len(prs.slides)} diapositivas",
            "entrego": f"Mismo archivo con paleta '{nombre_paleta}'",
            "notas": motivo,
        }


# ----------------------------------------------------------------------
# El Jefe: coordina, valida, manda a corregir y, si hace falta, corrige él
# ----------------------------------------------------------------------
class ErrorPipeline(Exception):
    pass


class Jefe:
    nombre = "Jefe"

    def __init__(self):
        self.log = []
        self.t_inicio = time.time()

    # ---- bitácora ----
    def _anotar(self, agente, rol, intento, estado, recibio, entrego, notas, segundos):
        self.log.append(
            {
                "paso": len(self.log) + 1,
                "agente": agente,
                "rol": rol,
                "intento": intento,
                "estado": estado,
                "recibio": recibio,
                "entrego": entrego,
                "notas": notas,
                "segundos": round(segundos, 1),
            }
        )

    # ---- reglas de supervisión: devuelven lista de problemas ----
    def validar(self, nombre, salida):
        p = []
        if nombre == "Lector":
            if len(salida["texto"].strip()) < 200:
                p.append("Texto casi vacío (¿PDF escaneado como imagen?)")
        elif nombre == "Recopilador":
            secciones = salida.get("secciones", [])
            if not salida.get("tema_general"):
                p.append("Falta 'tema_general'")
            if len(secciones) < 3:
                p.append("Hay menos de 3 temas; extrae más temas del documento")
            for s in secciones:
                if not s.get("puntos"):
                    p.append(f"El tema '{s.get('tema', '?')}' no tiene puntos")
        elif nombre == "Seccionador":
            p += self._problemas_plan(salida["slides"])
        elif nombre == "Generador PPTX":
            if not os.path.exists(salida["ruta"]):
                p.append("No se creó el archivo")
            elif len(Presentation(salida["ruta"]).slides) != salida["n_slides"]:
                p.append("El número de diapositivas no coincide con el plan")
        elif nombre == "Decorador":
            prs = Presentation(salida["ruta"])
            if len(prs.slides) != salida["n_slides"]:
                p.append("Cambió el número de diapositivas")
            for i, s in enumerate(prs.slides, 1):
                if s.background.fill.type != MSO_FILL.SOLID:
                    p.append(f"La diapositiva {i} no tiene fondo aplicado")
        return p

    def _problemas_plan(self, slides):
        p = []
        if not 5 <= len(slides) <= 10:
            p.append(f"Debe haber entre 5 y 10 diapositivas (hay {len(slides)})")
        vistos = set()
        for i, s in enumerate(slides, 1):
            t, bullets = s.get("title", ""), s.get("bullets", [])
            if not t or num_palabras(t) > 8:
                p.append(f"Diapositiva {i}: título vacío o de más de 8 palabras")
            if not 2 <= len(bullets) <= 5:
                p.append(f"Diapositiva {i}: debe tener entre 2 y 5 bullets (tiene {len(bullets)})")
            for b in bullets:
                if num_palabras(b) > 15:
                    p.append(f"Diapositiva {i}: bullet de más de 15 palabras: '{b[:40]}...'")
                clave = b.strip().lower()
                if clave in vistos:
                    p.append(f"Diapositiva {i}: bullet repetido: '{b[:40]}...'")
                vistos.add(clave)
        return p

    # ---- corrección directa del Jefe (cuando el agente ya no puede) ----
    def _corregir_plan(self, slides):
        cambios, vistos, nuevos = [], set(), []
        for i, s in enumerate(slides[:10], 1):
            t = " ".join(str(s.get("title", "Sin título")).split()[:8])
            bullets = []
            for b in s.get("bullets", []):
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
            if len(bullets) > 5:
                bullets = bullets[:5]
                cambios.append(f"dejó solo 5 bullets en la diapositiva {i}")
            nuevos.append({"title": t, "bullets": bullets})
        return nuevos, cambios

    # ---- ciclo de supervisión de un agente ----
    def _supervisar(self, agente, entrada):
        feedback, res = None, None
        for intento in range(1, MAX_INTENTOS + 1):
            t0 = time.time()
            res, problemas = None, []
            try:
                res = agente.ejecutar(entrada, feedback)
                problemas = self.validar(agente.nombre, res["salida"])
            except Exception as e:
                problemas = [f"Error: {type(e).__name__}: {e}"]
            dt = time.time() - t0

            if not problemas:
                self._anotar(agente.nombre, agente.rol, intento, "✅ Aprobado",
                             res["recibio"], res["entrego"], res["notas"], dt)
                return res["salida"]

            self._anotar(agente.nombre, agente.rol, intento, "❌ Rechazado por el Jefe",
                         res["recibio"] if res else "-", res["entrego"] if res else "-",
                         "Motivos: " + "; ".join(problemas), dt)
            feedback = "- " + "\n- ".join(problemas)

        # Se acabaron los intentos: el Jefe interviene si puede
        if agente.nombre == "Seccionador" and res:
            nuevos, cambios = self._corregir_plan(res["salida"]["slides"])
            salida = {"slides": nuevos, "tema_general": res["salida"].get("tema_general", "")}
            if not self._problemas_plan(nuevos):
                self._anotar(self.nombre, "Supervisa y corrige a los agentes", "-",
                             "🛠️ Corrigió él mismo", "Plan rechazado 3 veces",
                             f"Plan de {len(nuevos)} diapositivas corregido",
                             "; ".join(sorted(set(cambios))) or "ajustes menores", 0)
                return salida

        self._anotar(self.nombre, "Supervisa y corrige a los agentes", "-",
                     "🛑 Detuvo el proceso", f"{agente.nombre} falló {MAX_INTENTOS} veces",
                     "-", feedback or "", 0)
        raise ErrorPipeline(f"{agente.nombre} no logró entregar un resultado válido")

    # ---- el flujo completo ----
    def coordinar(self, pdf_path, ruta_pptx):
        lectura = self._supervisar(Lector(), pdf_path)
        resumen = self._supervisar(Recopilador(), lectura)
        plan = self._supervisar(Seccionador(), resumen)
        base = self._supervisar(GeneradorPPTX(ruta_pptx), plan)
        final = self._supervisar(Decorador(), base)
        return final["ruta"]

    # ---- reporte para ti (el "jefe de jefes") ----
    def reporte_markdown(self, exito):
        def c(x):
            return str(x).replace("|", "/").replace("\n", " ")

        filas = [
            "| # | Agente | Qué hace | Intento | Estado | Recibió | Entregó | Notas | Seg |",
            "|---|---|---|---|---|---|---|---|---|",
        ]
        for e in self.log:
            filas.append(
                f"| {e['paso']} | **{c(e['agente'])}** | {c(e['rol'])} | {e['intento']} | "
                f"{c(e['estado'])} | {c(e['recibio'])} | {c(e['entrego'])} | {c(e['notas'])} | {e['segundos']} |"
            )
        rechazos = sum(1 for e in self.log if "Rechazado" in e["estado"])
        correcciones = sum(1 for e in self.log if "Corrigió" in e["estado"])
        total = round(time.time() - self.t_inicio, 1)
        resumen = (
            f"### Informe final del Jefe\n"
            f"- Resultado: {'✅ presentación generada' if exito else '🛑 el proceso se detuvo'}\n"
            f"- Rechazos a agentes: {rechazos}\n"
            f"- Correcciones hechas por el Jefe: {correcciones}\n"
            f"- Tiempo total: {total} s\n\n"
        )
        return resumen + "### Bitácora por agente\n" + "\n".join(filas)


# ----------------------------------------------------------------------
# Interfaz
# ----------------------------------------------------------------------
def process_pdf(pdf_file):
    if pdf_file is None:
        raise gr.Error("Por favor sube un archivo PDF.")

    carpeta = tempfile.mkdtemp()
    ruta_pptx = os.path.join(carpeta, "presentacion_generada.pptx")
    ruta_log = os.path.join(carpeta, "bitacora.json")

    jefe = Jefe()
    try:
        jefe.coordinar(pdf_file, ruta_pptx)
        exito = True
    except ErrorPipeline:
        exito = False

    with open(ruta_log, "w", encoding="utf-8") as f:
        json.dump(jefe.log, f, ensure_ascii=False, indent=2)

    return (ruta_pptx if exito else None), jefe.reporte_markdown(exito), ruta_log


demo = gr.Interface(
    fn=process_pdf,
    inputs=gr.File(label="Sube tu PDF", file_types=[".pdf"]),
    outputs=[
        gr.File(label="Descarga tu presentación (.pptx)"),
        gr.Markdown(label="Informe del Jefe y bitácora de cada agente"),
        gr.File(label="Bitácora completa (.json)"),
    ],
    title="Sistema multiagente: PDF a Diapositivas",
    description=(
        "Sube un PDF. Seis agentes trabajan en cadena (Lector, Recopilador, "
        "Seccionador, Generador, Decorador) supervisados por un Jefe que valida, "
        "corrige y te entrega un informe de quién hizo qué."
    ),
    flagging_mode="never",
)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    demo.launch(server_name="0.0.0.0", server_port=port)
