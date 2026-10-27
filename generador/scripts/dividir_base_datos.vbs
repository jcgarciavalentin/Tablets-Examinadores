' Divide la base de datos de trabajo en dos ficheros, para usarla entre varias personas:
'   - ..._Datos.accdb: solo las tablas con los datos. Va en una carpeta compartida de la red.
'   - el fichero de trabajo (GestionTablets.accdb): formularios, consultas y código, con las tablas
'     vinculadas al de datos. Cada usuario usa una copia de este en su equipo.
' Se usa una sola vez. Antes hace una copia de seguridad en la carpeta "copias".
'
' Uso: arrastra tu base de datos (.accdb) sobre este fichero, o haz doble clic y te pide la ruta.
'
' Generado por generador/generar_scripts.py: no lo edites a mano.
Option Explicit

{COMUN}

' ---------------------------------------------------------------------------------------------
' Inicio
' ---------------------------------------------------------------------------------------------

Dim rutaBD, acc, carpetaDatos, rutaDatos, copia, resultado

rutaBD = PedirBaseTrabajo()
If rutaBD = "" Then WScript.Quit 0
If Not PuedeEmpezar(rutaBD) Then WScript.Quit 0

Set acc = CreateObject("Access.Application")
acc.OpenCurrentDatabase rutaBD
CerrarFormularios acc
If RutaDatosDe(acc) <> "" Then
    MsgBox "Esta base de datos ya está dividida. Sus datos están en:" & vbCrLf & RutaDatosDe(acc) & _
           vbCrLf & vbCrLf & "Para cambiar esa ubicación usa el botón ""Ubicación de los datos"" del menú.", _
           vbInformation, TITULO_APP
    Cerrar acc
    WScript.Quit 0
End If

' Carpeta compartida donde dejar el fichero de datos
acc.Visible = True
carpetaDatos = ElegirCarpeta(acc, "Elige la carpeta compartida donde guardar el fichero de datos")
If carpetaDatos = "" Then
    Cerrar acc
    WScript.Quit 0
End If
rutaDatos = fso.BuildPath(carpetaDatos, fso.GetBaseName(rutaBD) & "_Datos.accdb")
If fso.FileExists(rutaDatos) Then
    Cerrar acc
    MsgBox "Ya existe " & rutaDatos & vbCrLf & vbCrLf & "Elige otra carpeta, o mueve o borra ese fichero " & _
           "si no se usa.", vbExclamation, TITULO_APP
    WScript.Quit 0
End If
If MsgBox("Se va a dividir la base de datos:" & vbCrLf & rutaBD & vbCrLf & vbCrLf & _
          "Los datos pasarán a:" & vbCrLf & rutaDatos & vbCrLf & vbCrLf & _
          "y este fichero se quedará con los formularios, vinculado a esos datos. Antes se hace una " & _
          "copia de seguridad. ¿Continuar?", vbQuestion + vbYesNo, TITULO_APP) = vbNo Then
    Cerrar acc
    WScript.Quit 0
End If
Cerrar acc

' 1. Copia de seguridad y fichero de datos (copia completa de la base)
copia = HacerCopia(rutaBD)
fso.CopyFile rutaBD, rutaDatos

' 2. En el fichero de datos se quitan formularios, consultas y código: solo quedan las tablas
Set acc = CreateObject("Access.Application")
acc.OpenCurrentDatabase rutaDatos
CerrarFormularios acc
DejarSoloTablas acc
acc.CloseCurrentDatabase

' 3. En el fichero de trabajo se quitan las tablas y se vinculan las del fichero de datos
acc.OpenCurrentDatabase rutaBD
CerrarFormularios acc
acc.Visible = True
CargarModulo acc, "modGestionTablets"
CargarModulo acc, "modConstruirInterfaz"
resultado = acc.Run("DividirBaseDatos", rutaDatos)
acc.Run "ConstruirInterfaz"
BorrarModulo acc, "modConstruirInterfaz"
Cerrar acc

' 4. Compactar el fichero de trabajo: al quitarle las tablas (y las fotos) queda espacio libre
Compactar rutaBD

MsgBox "Base de datos dividida." & vbCrLf & vbCrLf & _
       "Datos (déjalo en la carpeta compartida, no lo abras directamente):" & vbCrLf & rutaDatos & vbCrLf & vbCrLf & _
       "Formularios (da una copia a cada usuario para su equipo):" & vbCrLf & rutaBD & vbCrLf & vbCrLf & _
       "Copia de seguridad de antes de dividir:" & vbCrLf & copia, vbInformation, TITULO_APP

' Quita de la base abierta todo lo que no sean tablas, y la tabla local de incidencias.
Sub DejarSoloTablas(acc)
    Dim tipo, nombres, nombre, db
    For Each tipo In Array(2, 3, 4, 5)          ' formularios, informes, macros y módulos
        nombres = NombresDe(ColeccionDe(acc, tipo))
        For Each nombre In nombres
            acc.DoCmd.DeleteObject tipo, nombre
        Next
    Next
    For Each nombre In NombresDe(acc.CurrentData.AllQueries)
        If Left(nombre, 1) <> "~" Then acc.DoCmd.DeleteObject 1, nombre
    Next
    For Each nombre In NombresDe(acc.CurrentData.AllTables)
        If nombre = "ImportacionIncidencias" Then acc.DoCmd.DeleteObject 0, nombre
    Next
    Set db = acc.CurrentDb
    On Error Resume Next            ' puede no tener formulario de inicio
    db.Properties.Delete "StartUpForm"
    On Error GoTo 0
End Sub

Function ColeccionDe(acc, tipo)
    Select Case tipo
        Case 2: Set ColeccionDe = acc.CurrentProject.AllForms
        Case 3: Set ColeccionDe = acc.CurrentProject.AllReports
        Case 4: Set ColeccionDe = acc.CurrentProject.AllMacros
        Case 5: Set ColeccionDe = acc.CurrentProject.AllModules
    End Select
End Function

' Nombres de una colección de objetos de Access (para poder borrarlos mientras se recorren).
Function NombresDe(coleccion)
    Dim o, lista
    lista = ""
    For Each o In coleccion
        lista = lista & vbTab & o.Name
    Next
    If lista = "" Then
        NombresDe = Array()
    Else
        NombresDe = Split(Mid(lista, 2), vbTab)
    End If
End Function

Function ElegirCarpeta(acc, titulo)
    Dim dialogo
    Set dialogo = acc.FileDialog(4)         ' msoFileDialogFolderPicker
    dialogo.Title = titulo
    dialogo.AllowMultiSelect = False
    If dialogo.Show = -1 Then ElegirCarpeta = dialogo.SelectedItems(1) Else ElegirCarpeta = ""
End Function

Sub Compactar(ruta)
    Dim app, temporal
    temporal = fso.BuildPath(fso.GetParentFolderName(ruta), fso.GetBaseName(ruta) & "_compactando.accdb")
    If fso.FileExists(temporal) Then fso.DeleteFile temporal
    Set app = CreateObject("Access.Application")
    If app.CompactRepair(ruta, temporal) Then
        fso.DeleteFile ruta
        fso.MoveFile temporal, ruta
    End If
    app.Quit
End Sub
