"""Punto de entrada: interfaz Gradio. Los agentes viven en la carpeta agentes/."""
import json
import os
import tempfile

import gradio as gr

from utils import ErrorPipeline
from agentes.jefe import Jefe


def process_pdf(pdf_file):
    if pdf_file is None:
        raise gr.Error("Por favor sube un archivo PDF.")

    carpeta = tempfile.mkdtemp()
    ruta_pptx = os.path.join(carpeta, "presentacion_generada.pptx")
    ruta_log = os.path.join(carpeta, "bitacora.json")

    jefe = Jefe()
    try:
        jefe.coordinar(pdf_file, ruta_pptx, carpeta)
        exito = True
    except ErrorPipeline:
        exito = False

    with open(ruta_log, "w", encoding="utf-8") as f:
        json.dump(jefe.log, f, ensure_ascii=False, indent=2)

    return (ruta_pptx if exito else None), jefe.reporte_html(exito), ruta_log


demo = gr.Interface(
    fn=process_pdf,
    inputs=gr.File(label="Sube tu PDF", file_types=[".pdf"]),
    outputs=[
        gr.File(label="Descarga tu presentación (.pptx)"),
        gr.HTML(label="Informe del Jefe y bitácora de cada agente"),
        gr.File(label="Bitácora completa (.json)"),
    ],
    title="Sistema multiagente: PDF a Diapositivas",
    description=(
        "Sube un PDF. Seis agentes trabajan en cadena (Lector, Recopilador, Seccionador, "
        "Ilustrador, Generador, Decorador) supervisados por un Jefe que valida con reglas y con IA, "
        "corrige y te entrega un informe de quién hizo qué y por qué. Las imágenes salen, en este orden, "
        "del propio PDF, de Wikimedia Commons (con su crédito) o de una ilustración generada con IA."
    ),
    flagging_mode="never",
)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    demo.launch(server_name="0.0.0.0", server_port=port)
