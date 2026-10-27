"""Genera los scripts de Windows de la raíz del proyecto a partir de generador/scripts/.

Cada plantilla (aplicar_ultimos_cambios.vbs, ConstruirFormularios.vbs) lleva {COMUN} donde va el
código de comun.vbs, que así se mantiene en un solo sitio pero cada script queda en un único
fichero. Se escriben en Windows-1252 con CRLF, que es lo que espera Windows Script Host.

Uso: python3 generador/generar_scripts.py
"""
from pathlib import Path

CARPETA = Path(__file__).resolve().parent / "scripts"
RAIZ = Path(__file__).resolve().parent.parent

comun = (CARPETA / "comun.vbs").read_text(encoding="utf-8").strip("\n")
for nombre in ("aplicar_ultimos_cambios.vbs", "ConstruirFormularios.vbs"):
    plantilla = (CARPETA / nombre).read_text(encoding="utf-8")
    texto = plantilla.replace("{COMUN}", comun)
    (RAIZ / nombre).write_bytes(texto.replace("\r\n", "\n").replace("\n", "\r\n").encode("cp1252"))
    print("Generado", nombre)
