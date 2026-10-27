' Carga en GestionTablets.accdb los datos de la base de datos antigua de tablets.
' Uso: arrastra la base antigua (.accdb) sobre este fichero, o haz doble clic y se busca en esta
' carpeta. Las dos bases de datos deben estar cerradas.
' ATENCIÓN: borra antes los examinadores, tablets, SIM y asignaciones de GestionTablets.accdb.
Option Explicit

Const AC_MODULE = 5
Const TITULO = "Importar datos antiguos"

Dim fso, carpeta, rutaBD, rutaAntigua, acc, resultado
Set fso = CreateObject("Scripting.FileSystemObject")
carpeta = fso.GetParentFolderName(WScript.ScriptFullName)
rutaBD = fso.BuildPath(carpeta, "GestionTablets.accdb")

If Not fso.FileExists(rutaBD) Then
    MsgBox "No se encuentra " & rutaBD, vbCritical, TITULO
    WScript.Quit 1
End If

If WScript.Arguments.Count > 0 Then
    rutaAntigua = WScript.Arguments(0)
Else
    rutaAntigua = InputBox("Ruta completa de la base de datos antigua:", TITULO, BuscarBaseAntigua())
    If rutaAntigua = "" Then WScript.Quit 0
End If
rutaAntigua = Replace(rutaAntigua, """", "")
If Not fso.FileExists(rutaAntigua) Then
    MsgBox "No se encuentra " & rutaAntigua, vbCritical, TITULO
    WScript.Quit 1
End If

If MsgBox("Se van a BORRAR los examinadores, tablets, tarjetas SIM y asignaciones de" & vbCrLf & _
          rutaBD & vbCrLf & vbCrLf & "y se cargarán desde" & vbCrLf & rutaAntigua & vbCrLf & vbCrLf & _
          "La base antigua no se modifica. ¿Continuar?", vbQuestion + vbYesNo, TITULO) = vbNo Then WScript.Quit 0

Set acc = CreateObject("Access.Application")
acc.Visible = True
acc.OpenCurrentDatabase rutaBD

On Error Resume Next            ' si ya existía de una ejecución anterior, se sustituye
acc.DoCmd.DeleteObject AC_MODULE, "modImportarDatosAntiguos"
On Error GoTo 0
acc.LoadFromText AC_MODULE, "modImportarDatosAntiguos", fso.BuildPath(carpeta, "vba\modImportarDatosAntiguos.bas")
resultado = acc.Run("ImportarDatosAntiguos", rutaAntigua)
acc.DoCmd.DeleteObject AC_MODULE, "modImportarDatosAntiguos"

acc.CloseCurrentDatabase
acc.Quit
Set acc = Nothing

If Left(resultado, 6) = "ERROR:" Then
    MsgBox resultado, vbCritical, TITULO
Else
    MsgBox resultado, vbInformation, TITULO
End If

' Propone la única otra base de datos .accdb de esta carpeta, si la hay.
Function BuscarBaseAntigua()
    Dim f, encontrada, n
    n = 0
    For Each f In fso.GetFolder(carpeta).Files
        If LCase(fso.GetExtensionName(f.Name)) = "accdb" And LCase(f.Name) <> "gestiontablets.accdb" Then
            encontrada = f.Path
            n = n + 1
        End If
    Next
    If n = 1 Then BuscarBaseAntigua = encontrada Else BuscarBaseAntigua = ""
End Function
