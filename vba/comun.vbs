' Código común de ConstruirFormularios.vbs y aplicar_ultimos_cambios.vbs, que lo cargan con
' ExecuteGlobal. AplicarCambios lleva una base de datos GestionTablets a la versión actual
' (estructura, consultas, formularios y código), preguntando antes si se conservan sus datos.

Const AC_MODULE = 5
Const TITULO_APP = "Gestión de tablets"
Const FICHERO_RUTA_ANTIGUA = "ruta_base_antigua.txt"

Dim fso, carpeta
Set fso = CreateObject("Scripting.FileSystemObject")
carpeta = fso.GetParentFolderName(WScript.ScriptFullName)

Sub AplicarCambios(rutaBD)
    Dim acc, hayDatos, conservar, rutaAntigua, copia, resultado, mensaje

    ' Al abrir un script desde dentro de un ZIP, Windows lo saca solo a una carpeta temporal
    If InStr(LCase(carpeta), ".zip") > 0 Then
        MsgBox "Estás ejecutando el script desde dentro del fichero ZIP." & vbCrLf & vbCrLf & _
               "Primero descomprímelo (clic derecho sobre el ZIP > Extraer todo...) y ejecuta el " & _
               "script desde la carpeta extraída.", vbExclamation, TITULO_APP
        Exit Sub
    End If
    If Not fso.FileExists(rutaBD) Then
        MsgBox "No se encuentra " & rutaBD, vbCritical, TITULO_APP
        Exit Sub
    End If
    If fso.FileExists(fso.BuildPath(fso.GetParentFolderName(rutaBD), fso.GetBaseName(rutaBD) & ".laccdb")) Then
        If MsgBox("Parece que la base de datos está abierta (en este equipo o en otro):" & vbCrLf & rutaBD & _
                  vbCrLf & vbCrLf & "Ciérrala antes de seguir. ¿Continuar de todos modos?", _
                  vbExclamation + vbYesNo, TITULO_APP) = vbNo Then Exit Sub
    End If

    copia = HacerCopia(rutaBD)

    Set acc = CreateObject("Access.Application")
    acc.OpenCurrentDatabase rutaBD
    hayDatos = TieneDatos(acc)

    ' ¿Se conservan los datos? Si no, se importan de la base antigua o se deja vacía.
    conservar = hayDatos
    If hayDatos Then
        Select Case MsgBox("La base de datos tiene datos:" & vbCrLf & rutaBD & vbCrLf & vbCrLf & _
                           "¿Quieres CONSERVAR los datos actuales?" & vbCrLf & vbCrLf & _
                           "Sí: se aplican los cambios sin tocar los datos." & vbCrLf & _
                           "No: se borran examinadores, tablets, SIM y asignaciones (a continuación " & _
                           "podrás importarlos de la base antigua)." & vbCrLf & _
                           "Cancelar: no se hace nada." & vbCrLf & vbCrLf & _
                           "Copia de seguridad hecha en:" & vbCrLf & copia, _
                           vbQuestion + vbYesNoCancel + vbDefaultButton1, TITULO_APP)
            Case vbCancel
                Cerrar acc
                Exit Sub
            Case vbNo
                conservar = False
        End Select
    End If
    If Not conservar Then
        If MsgBox("¿Importar ahora los datos de la base de datos antigua?", _
                  vbQuestion + vbYesNo, TITULO_APP) = vbYes Then
            acc.Visible = True              ' para que el selector de archivos se vea delante
            rutaAntigua = PedirBaseAntigua(acc)
            If rutaAntigua = "" Then
                Cerrar acc
                MsgBox "Cancelado: no se ha cambiado nada.", vbInformation, TITULO_APP
                Exit Sub
            End If
        End If
    End If

    ' Visible a partir de aquí, para que se vea cualquier error de VBA
    acc.Visible = True
    CargarModulo acc, "modGestionTablets"
    CargarModulo acc, "modConstruirInterfaz"
    If Not conservar Then
        ' Los datos se cargan antes de crear los formularios: las casillas de situación de
        ' examinadores se crean con las situaciones que haya en ese momento.
        acc.Run "ActualizarEstructura"
        CargarModulo acc, "modImportarDatosAntiguos"
        If rutaAntigua <> "" Then
            resultado = acc.Run("ImportarDatosAntiguos", rutaAntigua)
        Else
            resultado = acc.Run("VaciarDatos")
        End If
        acc.DoCmd.DeleteObject AC_MODULE, "modImportarDatosAntiguos"
    End If
    acc.Run "ConstruirInterfaz"
    acc.DoCmd.DeleteObject AC_MODULE, "modConstruirInterfaz"
    Cerrar acc

    mensaje = "Cambios aplicados a:" & vbCrLf & rutaBD & vbCrLf & vbCrLf
    If conservar Then
        mensaje = mensaje & "Se han conservado los datos."
    Else
        mensaje = mensaje & resultado
    End If
    mensaje = mensaje & vbCrLf & vbCrLf & "Copia de seguridad de antes de los cambios:" & vbCrLf & copia
    If Left(resultado, 6) = "ERROR:" Then
        MsgBox mensaje, vbCritical, TITULO_APP
    Else
        MsgBox mensaje, vbInformation, TITULO_APP
    End If
End Sub

' Copia la base de datos a la carpeta "copias", junto a ella, con la fecha y la hora en el nombre.
Function HacerCopia(rutaBD)
    Dim carpetaCopias, ahora, marca, destino
    carpetaCopias = fso.BuildPath(fso.GetParentFolderName(rutaBD), "copias")
    If Not fso.FolderExists(carpetaCopias) Then fso.CreateFolder carpetaCopias
    ahora = Now
    marca = Year(ahora) & Dos(Month(ahora)) & Dos(Day(ahora)) & "_" & Dos(Hour(ahora)) & Dos(Minute(ahora)) & Dos(Second(ahora))
    destino = fso.BuildPath(carpetaCopias, fso.GetBaseName(rutaBD) & "_" & marca & ".accdb")
    fso.CopyFile rutaBD, destino
    HacerCopia = destino
End Function

Function Dos(n)
    Dos = Right("0" & n, 2)
End Function

Function TieneDatos(acc)
    TieneDatos = acc.DCount("*", "Examinadores") + acc.DCount("*", "Tablets") + acc.DCount("*", "TarjetasSIM") > 0
End Function

Sub CargarModulo(acc, nombre)
    On Error Resume Next            ' si ya existía de una ejecución anterior, se sustituye
    acc.DoCmd.DeleteObject AC_MODULE, nombre
    On Error GoTo 0
    acc.LoadFromText AC_MODULE, nombre, fso.BuildPath(carpeta, "vba\" & nombre & ".bas")
End Sub

Sub Cerrar(acc)
    acc.CloseCurrentDatabase
    acc.Quit
End Sub

' Abre el selector de archivos de Access para elegir la base antigua. Empieza en la elegida la
' última vez (o en la carpeta de los scripts) y guarda la elección para la siguiente.
' Devuelve "" si se cancela.
Function PedirBaseAntigua(acc)
    Dim fichero, ruta, dialogo
    fichero = fso.BuildPath(carpeta, FICHERO_RUTA_ANTIGUA)
    ruta = LeerTexto(fichero)
    If ruta = "" Then ruta = BuscarBaseAntigua()

    Set dialogo = acc.FileDialog(3)         ' msoFileDialogFilePicker
    dialogo.Title = "Elige la base de datos antigua"
    dialogo.AllowMultiSelect = False
    dialogo.Filters.Clear
    dialogo.Filters.Add "Bases de datos de Access", "*.accdb; *.mdb"
    If ruta <> "" And fso.FileExists(ruta) Then
        dialogo.InitialFileName = ruta
    ElseIf ruta <> "" And fso.FolderExists(fso.GetParentFolderName(ruta)) Then
        dialogo.InitialFileName = fso.GetParentFolderName(ruta) & "\"
    Else
        dialogo.InitialFileName = carpeta & "\"
    End If

    If dialogo.Show = -1 Then
        ruta = dialogo.SelectedItems(1)
        EscribirTexto fichero, ruta
    Else
        ruta = ""
    End If
    PedirBaseAntigua = ruta
End Function

' Propone la única base de datos .accdb de esta carpeta que no sea GestionTablets, si la hay.
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

Function LeerTexto(fichero)
    LeerTexto = ""
    If fso.FileExists(fichero) Then
        With fso.OpenTextFile(fichero, 1)
            If Not .AtEndOfStream Then LeerTexto = Trim(.ReadLine)
            .Close
        End With
    End If
End Function

Sub EscribirTexto(fichero, texto)
    With fso.CreateTextFile(fichero, True)
        .WriteLine texto
        .Close
    End With
End Sub
