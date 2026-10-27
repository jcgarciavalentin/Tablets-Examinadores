' Creación inicial: crea las consultas, los formularios, el menú y el código dentro del
' GestionTablets.accdb de esta carpeta. Si ya tiene datos, pregunta si conservarlos; si no, ofrece
' importar los de la base antigua. Antes hace una copia de seguridad en la carpeta "copias".
' Uso: doble clic sobre este fichero (Access debe estar instalado y la base de datos cerrada).
' Para actualizar la base de datos con la que trabajas usa aplicar_ultimos_cambios.vbs.
'
' Generado por generador/generar_scripts.py: no lo edites a mano.
Option Explicit

{COMUN}

' ---------------------------------------------------------------------------------------------
' Inicio
' ---------------------------------------------------------------------------------------------

AplicarCambios fso.BuildPath(carpeta, "GestionTablets.accdb")
