' Aplica los últimos cambios (estructura, consultas, formularios y código) a la base de datos con
' la que trabajas, esté donde esté. Antes hace una copia de seguridad y pregunta si conservar los
' datos o volver a importarlos de la base antigua.
'
' Uso: arrastra tu base de datos (.accdb) sobre este fichero, o haz doble clic y te pide la ruta
' (propone la de la última vez). La base de datos debe estar cerrada.
'
' Generado por generador/generar_scripts.py: no lo edites a mano.
Option Explicit

' Código común de ConstruirFormularios.vbs y aplicar_ultimos_cambios.vbs. generar_scripts.py lo
' copia dentro de cada uno, detrás de su cabecera: así cada script es un único fichero y los errores
' indican la línea real. AplicarCambios lleva una base de datos GestionTablets a la versión actual
' (estructura, consultas, formularios y código), preguntando antes si se conservan sus datos.

Const AC_MODULE = 5
Const TITULO_APP = "Gestión de tablets"
Const FICHERO_RUTA_ANTIGUA = "ruta_base_antigua.txt"
Const FICHERO_RUTA_BD = "ruta_base_datos.txt"

Dim fso, carpeta
Set fso = CreateObject("Scripting.FileSystemObject")
carpeta = fso.GetParentFolderName(WScript.ScriptFullName)

Sub AplicarCambios(rutaBD)
    Dim acc, hayDatos, conservar, rutaAntigua, copia, rutaDatos, copiaDatos, resultado, mensaje, textoDatos

    If Not PuedeEmpezar(rutaBD) Then Exit Sub
    copia = HacerCopia(rutaBD)

    Set acc = CreateObject("Access.Application")
    acc.OpenCurrentDatabase rutaBD
    CerrarFormularios acc           ' el menú se abre solo al arrancar y hay que poder rehacerlo

    ' Base dividida: los datos están en otro fichero, compartido por todos los usuarios
    rutaDatos = RutaDatosDe(acc)
    If rutaDatos <> "" Then
        If Not fso.FileExists(rutaDatos) Then
            Cerrar acc
            MsgBox "Las tablas de esta base de datos están en otro fichero que no se encuentra:" & vbCrLf & _
                   rutaDatos & vbCrLf & vbCrLf & "Abre la base de datos en Access: te pedirá que indiques " & _
                   "dónde está. Después vuelve a ejecutar este script.", vbExclamation, TITULO_APP
            Exit Sub
        End If
        If EstaAbierta(rutaDatos) Then
            If MsgBox("Parece que alguien está usando el fichero de datos:" & vbCrLf & rutaDatos & vbCrLf & vbCrLf & _
                      "Si hay cambios de estructura (tablas o campos nuevos) no se podrán aplicar mientras " & _
                      "esté en uso. Pide a todos que cierren la base de datos. ¿Continuar de todos modos?", _
                      vbExclamation + vbYesNo, TITULO_APP) = vbNo Then
                Cerrar acc
                Exit Sub
            End If
        End If
        copiaDatos = HacerCopia(rutaDatos)
        textoDatos = vbCrLf & "(los datos están en " & rutaDatos & " y los comparten todos los usuarios)"
    End If
    hayDatos = TieneDatos(acc)

    ' ¿Se conservan los datos? Si no, se importan de la base antigua o se deja vacía.
    conservar = hayDatos
    If hayDatos Then
        Select Case MsgBox("La base de datos tiene datos:" & vbCrLf & rutaBD & vbCrLf & vbCrLf & _
                           textoDatos & vbCrLf & vbCrLf & "¿Quieres CONSERVAR los datos actuales?" & vbCrLf & vbCrLf & _
                           "Sí: se aplican los cambios sin tocar los datos." & vbCrLf & _
                           "No: se borran examinadores, tablets, SIM y asignaciones (a continuación " & _
                           "podrás importarlos de la base antigua)." & vbCrLf & _
                           "Cancelar: no se hace nada." & vbCrLf & vbCrLf & _
                           "Copia de seguridad hecha en:" & vbCrLf & copia & vbCrLf & copiaDatos, _
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
        BorrarModulo acc, "modImportarDatosAntiguos"
    End If
    acc.Run "ConstruirInterfaz"
    BorrarModulo acc, "modConstruirInterfaz"
    Cerrar acc

    mensaje = "Cambios aplicados a:" & vbCrLf & rutaBD & vbCrLf & vbCrLf
    If conservar Then
        mensaje = mensaje & "Se han conservado los datos."
    Else
        mensaje = mensaje & resultado
    End If
    mensaje = mensaje & vbCrLf & vbCrLf & "Copia de seguridad de antes de los cambios:" & vbCrLf & copia
    If copiaDatos <> "" Then mensaje = mensaje & vbCrLf & copiaDatos
    If Left(resultado, 6) = "ERROR:" Then
        MsgBox mensaje, vbCritical, TITULO_APP
    Else
        MsgBox mensaje, vbInformation, TITULO_APP
    End If
End Sub

' Comprobaciones antes de tocar nada: que no se ejecute desde dentro del ZIP, que la base exista
' y que no esté abierta (o que se quiera seguir igualmente).
Function PuedeEmpezar(rutaBD)
    PuedeEmpezar = False
    ' Al abrir un script desde dentro de un ZIP, Windows lo saca solo a una carpeta temporal
    If InStr(LCase(carpeta), ".zip") > 0 Then
        MsgBox "Estás ejecutando el script desde dentro del fichero ZIP." & vbCrLf & vbCrLf & _
               "Primero descomprímelo (clic derecho sobre el ZIP > Extraer todo...) y ejecuta el " & _
               "script desde la carpeta extraída.", vbExclamation, TITULO_APP
        Exit Function
    End If
    If Not fso.FileExists(rutaBD) Then
        MsgBox "No se encuentra " & rutaBD, vbCritical, TITULO_APP
        Exit Function
    End If
    If EstaAbierta(rutaBD) Then
        If MsgBox("Parece que la base de datos está abierta (en este equipo o en otro):" & vbCrLf & rutaBD & _
                  vbCrLf & vbCrLf & "Ciérrala antes de seguir. ¿Continuar de todos modos?", _
                  vbExclamation + vbYesNo, TITULO_APP) = vbNo Then Exit Function
    End If
    PuedeEmpezar = True
End Function

' Access crea un fichero .laccdb junto a la base mientras alguien la tiene abierta.
Function EstaAbierta(ruta)
    EstaAbierta = fso.FileExists(fso.BuildPath(fso.GetParentFolderName(ruta), fso.GetBaseName(ruta) & ".laccdb"))
End Function

' Fichero de datos al que están vinculadas las tablas, o "" si la base no está dividida.
Function RutaDatosDe(acc)
    Dim db, conexion, p
    RutaDatosDe = ""
    Set db = acc.CurrentDb
    On Error Resume Next
    conexion = db.TableDefs("Tablets").Connect
    On Error GoTo 0
    p = InStr(1, conexion, ";DATABASE=", vbTextCompare)
    If p = 0 Then Exit Function
    conexion = Mid(conexion, p + Len(";DATABASE="))
    If InStr(conexion, ";") > 0 Then conexion = Left(conexion, InStr(conexion, ";") - 1)
    RutaDatosDe = conexion
End Function

' Pide la base de datos de trabajo: la arrastrada sobre el script o, si no, la que se escriba
' (propone la de la última vez y la recuerda para la siguiente).
Function PedirBaseTrabajo()
    Dim fichero, ruta
    fichero = fso.BuildPath(carpeta, FICHERO_RUTA_BD)
    If WScript.Arguments.Count > 0 Then
        ruta = WScript.Arguments(0)
    Else
        ruta = LeerTexto(fichero)
        If ruta = "" Then ruta = fso.BuildPath(carpeta, "GestionTablets.accdb")
        ruta = Replace(InputBox("Base de datos de trabajo (la que abres en Access):", TITULO_APP, ruta), """", "")
    End If
    If ruta <> "" Then
        If fso.FileExists(ruta) Then EscribirTexto fichero, ruta
    End If
    PedirBaseTrabajo = ruta
End Function

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
    BorrarModulo acc, nombre        ' si ya existía de una ejecución anterior, se sustituye
    acc.LoadFromText AC_MODULE, nombre, fso.BuildPath(carpeta, "vba\" & nombre & ".bas")
End Sub

' Solo si existe: con Access visible, borrar un módulo que no existe muestra un error en pantalla
' aunque el script lo ignore.
Sub BorrarModulo(acc, nombre)
    Dim m
    For Each m In acc.CurrentProject.AllModules
        If LCase(m.Name) = LCase(nombre) Then
            acc.DoCmd.DeleteObject AC_MODULE, m.Name
            Exit Sub
        End If
    Next
End Sub

Sub CerrarFormularios(acc)
    Do While acc.Forms.Count > 0
        acc.DoCmd.Close 2, acc.Forms(0).Name        ' 2 = acForm
    Loop
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

' ---------------------------------------------------------------------------------------------
' Inicio
' ---------------------------------------------------------------------------------------------

Dim rutaBD
rutaBD = PedirBaseTrabajo()
If rutaBD <> "" Then AplicarCambios rutaBD
