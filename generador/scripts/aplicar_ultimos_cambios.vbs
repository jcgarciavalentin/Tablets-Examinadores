' Aplica los últimos cambios (estructura, consultas, formularios y código) a la base de datos con
' la que trabajas, esté donde esté. Antes hace una copia de seguridad y pregunta si conservar los
' datos o volver a importarlos de la base antigua.
'
' Uso: arrastra tu base de datos (.accdb) sobre este fichero, o haz doble clic y te pide la ruta
' (propone la de la última vez). La base de datos debe estar cerrada.
'
' Generado por generador/generar_scripts.py: no lo edites a mano.
Option Explicit

{COMUN}

' ---------------------------------------------------------------------------------------------
' Inicio
' ---------------------------------------------------------------------------------------------

Dim rutaBD
rutaBD = PedirBaseTrabajo()
If rutaBD <> "" Then AplicarCambios rutaBD
