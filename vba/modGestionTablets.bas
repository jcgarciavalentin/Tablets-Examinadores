Option Compare Database
Option Explicit

' Lógica de los formularios de GestionTablets.accdb. Los botones y eventos de los formularios
' llaman a estas funciones mediante expresiones (p. ej. =AsignarTablet()), por eso los
' formularios no tienen módulo de código propio.
'
' Las asignaciones se guardan como periodos: FechaFin vacía = asignación vigente, y todas las
' filas juntas forman el histórico. Al asignar se cierran, con la fecha de la nueva asignación,
' las asignaciones vigentes que entren en conflicto, de modo que en cada momento una tablet tiene
' como mucho un examinador y una SIM, y un examinador o una SIM como mucho una tablet.

Private Const FALLO_SI_ERROR As Long = 128   ' dbFailOnError

Public Function AbrirFormulario(ByVal nombre As String)
    DoCmd.OpenForm nombre
End Function

Public Function AbrirHojaDatos(ByVal nombre As String)
    DoCmd.OpenForm nombre, acFormDS
End Function

' Cuadro "Buscar": deja visibles los registros en los que alguno de los campos (separados por ";";
' pueden ser expresiones) contiene el texto escrito en txtBuscar. Sin texto, se quita el filtro.
Public Function FiltrarFormulario(ByVal campos As String)
    AplicarFiltro Screen.ActiveForm, Screen.ActiveForm, campos
End Function

' Igual, pero filtra la hoja de datos del subformulario subResultados (formularios de histórico).
Public Function FiltrarSubformulario(ByVal campos As String)
    AplicarFiltro Screen.ActiveForm, Screen.ActiveForm!subResultados.Form, campos
End Function

Public Function QuitarFiltro()
    Dim f As Object, destino As Object
    Set f = Screen.ActiveForm
    f!txtBuscar = Null
    Set destino = f
    On Error Resume Next            ' sin subResultados, el filtro está en el propio formulario
    Set destino = f!subResultados.Form
    On Error GoTo 0
    destino.FilterOn = False
End Function

' Al volver a un formulario, recarga sus listas y subformularios por si se han dado de alta
' datos en otra ventana.
Public Function RefrescarListas()
    On Error Resume Next
    RefrescarControles Screen.ActiveForm
End Function

Public Function AsignarTablet()
    Dim f As Object, idExaminador As Long, idTablet As Long, fecha As Date
    Dim actual As Variant, aviso As String

    Set f = Forms("frmAsignarTablet")
    If Not DatosCompletos(f, "cboExaminador", "cboTablet") Then Exit Function
    idExaminador = f!cboExaminador
    idTablet = f!cboTablet
    fecha = f!txtFecha

    If DCount("*", "AsignacionesTablet", "FechaFin Is Null AND IdTablet = " & idTablet & _
              " AND IdExaminador = " & idExaminador) > 0 Then
        MsgBox "Esa tablet ya está asignada a ese examinador.", vbInformation
        Exit Function
    End If
    If Not TabletOperativa(idTablet) Then Exit Function

    actual = DLookup("Examinador", "qryAsignacionesTabletActuales", "IdTablet = " & idTablet)
    If Not IsNull(actual) Then aviso = aviso & "- La tablet está asignada ahora a " & actual & "." & vbCrLf
    actual = DLookup("Tablet", "qryAsignacionesTabletActuales", "IdExaminador = " & idExaminador)
    If Not IsNull(actual) Then aviso = aviso & "- El examinador tiene ahora la tablet " & actual & "." & vbCrLf
    If Not ConfirmarCierre(aviso, fecha) Then Exit Function
    If Not FechaPosterior("AsignacionesTablet", "IdTablet = " & idTablet & " OR IdExaminador = " & idExaminador, fecha) Then Exit Function

    On Error GoTo Fallo
    DBEngine.BeginTrans
    Ejecutar "PARAMETERS pFecha DateTime, pTablet Long, pExaminador Long; " & _
             "UPDATE AsignacionesTablet SET FechaFin = pFecha " & _
             "WHERE FechaFin Is Null AND (IdTablet = pTablet OR IdExaminador = pExaminador);", _
             fecha, idTablet, idExaminador
    Ejecutar "PARAMETERS pExaminador Long, pTablet Long, pFecha DateTime, pObservaciones Text(255); " & _
             "INSERT INTO AsignacionesTablet (IdExaminador, IdTablet, FechaInicio, Observaciones) " & _
             "VALUES (pExaminador, pTablet, pFecha, pObservaciones);", _
             idExaminador, idTablet, fecha, f!txtObservaciones
    DBEngine.CommitTrans
    Terminar f, "Tablet asignada."
    Exit Function

Fallo:
    DBEngine.Rollback
    MsgBox "No se pudo guardar la asignación: " & Err.Description, vbCritical
End Function

Public Function DevolverTablet()
    Dim f As Object, idTablet As Long, fecha As Date, actual As Variant

    Set f = Forms("frmAsignarTablet")
    If Not DatosCompletos(f, "cboTablet") Then Exit Function
    idTablet = f!cboTablet
    fecha = f!txtFecha

    actual = DLookup("Examinador", "qryAsignacionesTabletActuales", "IdTablet = " & idTablet)
    If IsNull(actual) Then MsgBox "Esa tablet no está asignada a ningún examinador.", vbInformation: Exit Function
    If Not FechaPosterior("AsignacionesTablet", "IdTablet = " & idTablet, fecha) Then Exit Function
    If MsgBox("¿Registrar que " & actual & " devuelve la tablet con fecha " & Format(fecha, "dd/mm/yyyy") & "?", _
              vbQuestion + vbYesNo) = vbNo Then Exit Function

    Ejecutar "PARAMETERS pFecha DateTime, pTablet Long; " & _
             "UPDATE AsignacionesTablet SET FechaFin = pFecha WHERE FechaFin Is Null AND IdTablet = pTablet;", _
             fecha, idTablet
    Terminar f, "Devolución registrada."
End Function

Public Function AsignarSIM()
    Dim f As Object, idTablet As Long, idSIM As Long, fecha As Date
    Dim actual As Variant, aviso As String

    Set f = Forms("frmAsignarSIM")
    If Not DatosCompletos(f, "cboTablet", "cboSIM") Then Exit Function
    idTablet = f!cboTablet
    idSIM = f!cboSIM
    fecha = f!txtFecha

    If DCount("*", "AsignacionesSIM", "FechaFin Is Null AND IdTablet = " & idTablet & _
              " AND IdSIM = " & idSIM) > 0 Then
        MsgBox "Esa tarjeta SIM ya está en esa tablet.", vbInformation
        Exit Function
    End If
    If Not SIMActiva(idSIM) Then Exit Function

    actual = DLookup("Tablet", "qrySIMActuales", "IdSIM = " & idSIM)
    If Not IsNull(actual) Then aviso = aviso & "- La SIM está ahora en la tablet " & actual & "." & vbCrLf
    actual = DLookup("SIM", "qrySIMActuales", "IdTablet = " & idTablet)
    If Not IsNull(actual) Then aviso = aviso & "- La tablet tiene ahora la SIM " & actual & "." & vbCrLf
    If Not ConfirmarCierre(aviso, fecha) Then Exit Function
    If Not FechaPosterior("AsignacionesSIM", "IdTablet = " & idTablet & " OR IdSIM = " & idSIM, fecha) Then Exit Function

    On Error GoTo Fallo
    DBEngine.BeginTrans
    Ejecutar "PARAMETERS pFecha DateTime, pTablet Long, pSIM Long; " & _
             "UPDATE AsignacionesSIM SET FechaFin = pFecha " & _
             "WHERE FechaFin Is Null AND (IdTablet = pTablet OR IdSIM = pSIM);", _
             fecha, idTablet, idSIM
    Ejecutar "PARAMETERS pTablet Long, pSIM Long, pFecha DateTime, pObservaciones Text(255); " & _
             "INSERT INTO AsignacionesSIM (IdTablet, IdSIM, FechaInicio, Observaciones) " & _
             "VALUES (pTablet, pSIM, pFecha, pObservaciones);", _
             idTablet, idSIM, fecha, f!txtObservaciones
    DBEngine.CommitTrans
    Terminar f, "Tarjeta SIM asociada a la tablet."
    Exit Function

Fallo:
    DBEngine.Rollback
    MsgBox "No se pudo guardar la asociación: " & Err.Description, vbCritical
End Function

Public Function RetirarSIM()
    Dim f As Object, idTablet As Long, fecha As Date, actual As Variant

    Set f = Forms("frmAsignarSIM")
    If Not DatosCompletos(f, "cboTablet") Then Exit Function
    idTablet = f!cboTablet
    fecha = f!txtFecha

    actual = DLookup("SIM", "qrySIMActuales", "IdTablet = " & idTablet)
    If IsNull(actual) Then MsgBox "Esa tablet no tiene ninguna tarjeta SIM asociada.", vbInformation: Exit Function
    If Not FechaPosterior("AsignacionesSIM", "IdTablet = " & idTablet, fecha) Then Exit Function
    If MsgBox("¿Retirar la SIM " & actual & " de la tablet con fecha " & Format(fecha, "dd/mm/yyyy") & "?", _
              vbQuestion + vbYesNo) = vbNo Then Exit Function

    Ejecutar "PARAMETERS pFecha DateTime, pTablet Long; " & _
             "UPDATE AsignacionesSIM SET FechaFin = pFecha WHERE FechaFin Is Null AND IdTablet = pTablet;", _
             fecha, idTablet
    Terminar f, "Tarjeta SIM retirada."
End Function

Private Sub AplicarFiltro(ByVal f As Object, ByVal destino As Object, ByVal campos As String)
    Dim texto As String, campo As Variant, filtro As String
    texto = Trim$(Nz(f!txtBuscar, ""))
    If Len(texto) = 0 Then
        destino.FilterOn = False
        Exit Sub
    End If
    ' Comillas duplicadas y comodines de Like (* ? # [) tratados como texto normal
    texto = Replace(texto, "'", "''")
    texto = Replace(texto, "[", "[[]")
    texto = Replace(texto, "*", "[*]")
    texto = Replace(texto, "?", "[?]")
    texto = Replace(texto, "#", "[#]")
    For Each campo In Split(campos, ";")
        filtro = filtro & " OR (" & campo & ") Like '*" & texto & "*'"
    Next
    If destino.Dirty Then destino.Dirty = False
    destino.Filter = Mid$(filtro, 5)
    destino.FilterOn = True
    If destino.RecordsetClone.RecordCount = 0 Then MsgBox "No se ha encontrado nada con ese texto.", vbInformation
End Sub

Private Function DatosCompletos(ByVal f As Object, ParamArray controles() As Variant) As Boolean
    Dim i As Long
    For i = 0 To UBound(controles)
        If IsNull(f(controles(i)).Value) Then
            MsgBox "Selecciona " & f(controles(i)).Controls(0).Caption & " y una fecha.", vbExclamation
            Exit Function
        End If
    Next
    If Not IsDate(f!txtFecha) Then
        MsgBox "Indica una fecha válida.", vbExclamation
        Exit Function
    End If
    DatosCompletos = True
End Function

' Si la tablet no está en estado "Operativa" (tabla EstadosTablet) se pide confirmación.
Private Function TabletOperativa(ByVal idTablet As Long) As Boolean
    Dim estado As Variant
    estado = DLookup("Estado", "qryEstadoTablets", "IdTablet = " & idTablet)
    If estado = "Operativa" Then
        TabletOperativa = True
    Else
        TabletOperativa = (MsgBox("La tablet está en estado '" & estado & "'. ¿Asignarla igualmente?", _
                                  vbQuestion + vbYesNo) = vbYes)
    End If
End Function

' Si la SIM no está en estado "Activa" (tabla EstadosSIM) se pide confirmación.
Private Function SIMActiva(ByVal idSIM As Long) As Boolean
    Dim estado As Variant
    estado = DLookup("EstadoSIM", "EstadosSIM", "IdEstadoSIM = " & Nz(DLookup("IdEstadoSIM", "TarjetasSIM", "IdSIM = " & idSIM), 0))
    If estado = "Activa" Then
        SIMActiva = True
    Else
        SIMActiva = (MsgBox("La tarjeta SIM está en estado '" & estado & "'. ¿Asociarla igualmente?", _
                            vbQuestion + vbYesNo) = vbYes)
    End If
End Function

Private Function ConfirmarCierre(ByVal aviso As String, ByVal fecha As Date) As Boolean
    If Len(aviso) = 0 Then
        ConfirmarCierre = True
    Else
        ConfirmarCierre = (MsgBox(aviso & vbCrLf & "Esas asignaciones se cerrarán con fecha " & _
                                  Format(fecha, "dd/mm/yyyy") & ". ¿Continuar?", vbQuestion + vbYesNo) = vbYes)
    End If
End Function

' Impide cerrar una asignación vigente con una fecha anterior a la de su inicio.
Private Function FechaPosterior(ByVal tabla As String, ByVal filtro As String, ByVal fecha As Date) As Boolean
    Dim inicio As Variant
    inicio = DMax("FechaInicio", tabla, "FechaFin Is Null AND (" & filtro & ")")
    If Not IsNull(inicio) Then
        If fecha < inicio Then
            MsgBox "La fecha no puede ser anterior al " & Format(inicio, "dd/mm/yyyy") & _
                   ", en que empezó la asignación vigente.", vbExclamation
            Exit Function
        End If
    End If
    FechaPosterior = True
End Function

' Consulta con parámetros: evita problemas con el formato de fecha y con comillas en el texto.
Private Sub Ejecutar(ByVal sql As String, ParamArray valores() As Variant)
    Dim db As Object, qd As Object, i As Long
    Set db = CurrentDb
    Set qd = db.CreateQueryDef("", sql)
    For i = 0 To UBound(valores)
        qd.Parameters(i).Value = valores(i)
    Next
    qd.Execute FALLO_SI_ERROR
    qd.Close
End Sub

Private Sub Terminar(ByVal f As Object, ByVal mensaje As String)
    RefrescarControles f
    f!txtObservaciones = Null
    MsgBox mensaje, vbInformation
End Sub

Private Sub RefrescarControles(ByVal f As Object)
    Dim c As Object
    For Each c In f.Controls
        Select Case c.ControlType
            Case acComboBox: c.Requery
            Case acSubform: c.Form.Requery
        End Select
    Next
End Sub
