"""Clase base de los agentes y excepción del pipeline."""



class Agente:
    nombre = "Agente"
    rol = ""

    def ejecutar(self, entrada, feedback=None):
        """Debe devolver {"salida", "recibio", "entrego", "razonamiento", "detalle", "notas"}"""
        raise NotImplementedError


class ErrorPipeline(Exception):
    pass
