' Crea las consultas, formularios y menú dentro de GestionTablets.accdb usando Microsoft Access.
' Uso: doble clic sobre este fichero (Access debe estar instalado y la base de datos cerrada).
' Puede volver a ejecutarse: vuelve a generar consultas y formularios sin tocar los datos.
Option Explicit

Const AC_MODULE = 5

Dim fso, carpeta, rutaBD, acc
Set fso = CreateObject("Scripting.FileSystemObject")
carpeta = fso.GetParentFolderName(WScript.ScriptFullName)
rutaBD = fso.BuildPath(carpeta, "GestionTablets.accdb")

If Not fso.FileExists(rutaBD) Then
    MsgBox "No se encuentra " & rutaBD, vbCritical, "Gestión de tablets"
    WScript.Quit 1
End If

Set acc = CreateObject("Access.Application")
acc.Visible = True
acc.OpenCurrentDatabase rutaBD

CargarModulo "modGestionTablets"
CargarModulo "modConstruirInterfaz"
acc.Run "ConstruirInterfaz"
acc.DoCmd.DeleteObject AC_MODULE, "modConstruirInterfaz"

acc.CloseCurrentDatabase
acc.Quit
Set acc = Nothing

MsgBox "Formularios creados. Abre GestionTablets.accdb para empezar.", vbInformation, "Gestión de tablets"

Sub CargarModulo(nombre)
    On Error Resume Next            ' si ya existía de una ejecución anterior, se sustituye
    acc.DoCmd.DeleteObject AC_MODULE, nombre
    On Error GoTo 0
    acc.LoadFromText AC_MODULE, nombre, fso.BuildPath(carpeta, "vba\" & nombre & ".bas")
End Sub
