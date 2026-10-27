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

Const FICHERO_RUTA_BD = "ruta_base_datos.txt"
Dim rutaBD

If WScript.Arguments.Count > 0 Then
    rutaBD = WScript.Arguments(0)
Else
    rutaBD = LeerTexto(fso.BuildPath(carpeta, FICHERO_RUTA_BD))
    If rutaBD = "" Then rutaBD = fso.BuildPath(carpeta, "GestionTablets.accdb")
    rutaBD = Replace(InputBox("Base de datos a la que aplicar los cambios:", TITULO_APP, rutaBD), """", "")
    If rutaBD = "" Then WScript.Quit 0
End If
If fso.FileExists(rutaBD) Then EscribirTexto fso.BuildPath(carpeta, FICHERO_RUTA_BD), rutaBD

AplicarCambios rutaBD
