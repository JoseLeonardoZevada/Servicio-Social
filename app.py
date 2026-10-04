import os
import json
import re
 
import gradio as gr
import pdfplumber
import google.generativeai as genai
from pptx import Presentation
 
# La API key se lee de la variable de entorno configurada en Render
genai.configure(api_key=os.environ["GEMINI_API_KEY"])
 
# "System instruction": el comportamiento y las reglas fijas del agente,
# separado de la tarea puntual que le pedimos cada vez. Esto es lo más
# parecido a "entrenar" o "instruir" a un agente basado en un LLM.
SYSTEM_INSTRUCTION = """
Eres un asistente experto en diseño de presentaciones. Tu trabajo es leer
documentos y convertirlos en diapositivas claras, concisas y bien
estructuradas, pensadas para explicarse en voz alta frente a una audiencia.
 
Reglas que debes seguir SIEMPRE, sin excepción:
- Responde ÚNICAMENTE con JSON válido. Nada de texto antes o después,
  nada de explicaciones, nada de ``` ni marcado de código.
- Escribe en español neutro, con tono claro y profesional.
- Cada "title" debe tener máximo 8 palabras.
- Cada "bullet" debe tener máximo 15 palabras y ser una frase corta,
  nunca un párrafo completo ni una oración copiada tal cual del documento.
- No repitas la misma idea en dos diapositivas distintas.
- La primera diapositiva siempre debe presentar el tema general del
  documento (a modo de introducción).
- La última diapositiva siempre debe ser un resumen o conclusión con
  los puntos más importantes.
"""
 
model = genai.GenerativeModel(
    "gemini-3.8-flash",
    system_instruction=SYSTEM_INSTRUCTION,
)
 
# Ejemplo fijo que le mostramos al modelo (few-shot prompting): le
# enseña con un caso concreto el patrón exacto que esperamos de entrada
# y salida, además de las reglas generales de arriba.
EXAMPLE_INPUT = """
El reciclaje es el proceso de convertir materiales de desecho en nuevos
productos. Ayuda a reducir el consumo de materias primas frescas, reduce
el uso de energía, reduce la contaminación del aire y del agua, y reduce
las emisiones de gases de efecto invernadero. Los materiales más comunes
que se reciclan son el papel, el vidrio, el plástico y los metales. Para
que el reciclaje funcione bien, es importante separar correctamente los
materiales desde el hogar.
"""
 
EXAMPLE_OUTPUT = {
    "slides": [
        {
            "title": "¿Qué es el reciclaje?",
            "bullets": [
                "Convierte materiales de desecho en nuevos productos",
                "Reduce el consumo de materias primas frescas",
            ],
        },
        {
            "title": "Beneficios ambientales",
            "bullets": [
                "Reduce el uso de energía",
                "Disminuye la contaminación del aire y el agua",
                "Reduce las emisiones de gases de efecto invernadero",
            ],
        },
        {
            "title": "Materiales reciclables comunes",
            "bullets": ["Papel y cartón", "Vidrio", "Plástico", "Metales"],
        },
        {
            "title": "Conclusión",
            "bullets": [
                "El reciclaje reduce el impacto ambiental",
                "La separación correcta desde el hogar es clave",
            ],
        },
    ]
}
 
 
def extract_text(pdf_path):
    """Saca todo el texto del PDF, página por página."""
    text_parts = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text_parts.append(page_text)
    return "\n".join(text_parts)
 
 
def ask_gemini_for_slides(text):
    """Le pide a Gemini que convierta el texto en una estructura de diapositivas (JSON).
 
    Usa few-shot prompting: le mostramos un ejemplo completo de
    entrada -> salida antes de darle el documento real, además de las
    reglas fijas que ya tiene en su system_instruction.
    """
    prompt = f"""
A continuación tienes un ejemplo de cómo debes transformar un texto en
diapositivas, siguiendo tus reglas:
 
TEXTO DE EJEMPLO:
\"\"\"{EXAMPLE_INPUT}\"\"\"
 
SALIDA ESPERADA PARA ESE EJEMPLO:
{json.dumps(EXAMPLE_OUTPUT, ensure_ascii=False)}
 
Ahora haz exactamente lo mismo, pero con el siguiente documento real.
Genera entre 5 y 10 diapositivas, cada una con entre 3 y 5 bullets,
devolviendo SOLO el JSON con esta estructura exacta:
 
{{"slides": [{{"title": "...", "bullets": ["...", "..."]}}]}}
 
DOCUMENTO REAL:
\"\"\"
{text[:15000]}
\"\"\"
"""
    response = model.generate_content(prompt)
    raw = response.text.strip()
 
    # Por si el modelo agrega ```json ... ``` a pesar de la instrucción
    raw = re.sub(r"^```json", "", raw)
    raw = re.sub(r"^```", "", raw)
    raw = re.sub(r"```$", "", raw)
    raw = raw.strip()
 
    data = json.loads(raw)
    return data["slides"]
 
 
def build_pptx(slides_data, output_path):
    """Arma el archivo .pptx a partir de la estructura de diapositivas."""
    prs = Presentation()
    title_and_content_layout = prs.slide_layouts[1]
 
    for slide_info in slides_data:
        slide = prs.slides.add_slide(title_and_content_layout)
        slide.shapes.title.text = slide_info.get("title", "Sin titulo")
 
        body_placeholder = slide.placeholders[1]
        tf = body_placeholder.text_frame
        tf.clear()
 
        bullets = slide_info.get("bullets", [])
        for i, bullet in enumerate(bullets):
            if i == 0:
                tf.text = bullet
            else:
                p = tf.add_paragraph()
                p.text = bullet
 
    prs.save(output_path)
 
 
def process_pdf(pdf_file):
    if pdf_file is None:
        raise gr.Error("Por favor sube un archivo PDF.")
 
    text = extract_text(pdf_file)
    if not text.strip():
        raise gr.Error(
            "No se pudo extraer texto de ese PDF (¿está escaneado como imagen?)."
        )
 
    slides_data = ask_gemini_for_slides(text)
 
    output_path = "/tmp/presentacion_generada.pptx"
    build_pptx(slides_data, output_path)
 
    return output_path
 
 
demo = gr.Interface(
    fn=process_pdf,
    inputs=gr.File(label="Sube tu PDF", file_types=[".pdf"]),
    outputs=gr.File(label="Descarga tu presentación (.pptx)"),
    title="Agente de IA: PDF a Diapositivas",
    description="Sube un PDF y el agente generará automáticamente una presentación resumida.",
)
 
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    demo.launch(server_name="0.0.0.0", server_port=port)
