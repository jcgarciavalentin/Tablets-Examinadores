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
' Cuando un examinador devuelve una tablet se pregunta a qué estado pasa (frmElegirEstado) y el
' cambio queda en HistorialEstadosTablet.

Private Const FALLO_SI_ERROR As Long = 128   ' dbFailOnError

' Comunicación con frmElegirEstado (ver PreguntarEstado)
Private mensajeEstado As String
Private estadoInicial As Variant
Private estadoElegido As Variant
Private notaElegida As Variant

Public Function AbrirFormulario(ByVal nombre As String)
    DoCmd.OpenForm nombre
End Function

Public Function AbrirHojaDatos(ByVal nombre As String)
    DoCmd.OpenForm nombre, acFormDS
End Function

' Cuadro "Buscar": deja visibles los registros que cumplen alguna de las condiciones de búsqueda
' (CamposBusqueda) con el texto escrito en txtBuscar. Sin texto, se quita el filtro.
' clave es el nombre de la búsqueda, que pone cada formulario.
Public Function FiltrarFormulario(ByVal clave As String)
    AplicarFiltro Screen.ActiveForm, Screen.ActiveForm, CamposBusqueda(clave)
End Function

' Igual, pero filtra la hoja de datos del subformulario subResultados (formularios de histórico).
Public Function FiltrarSubformulario(ByVal clave As String)
    AplicarFiltro Screen.ActiveForm, Screen.ActiveForm!subResultados.Form, CamposBusqueda(clave)
End Function

' Dónde busca cada cuadro "Buscar", separado por ";". Un campo o expresión se compara con
' Like '*texto*'; una condición que contiene {patron} se usa entera, con {patron} = '*texto*'.
Private Function CamposBusqueda(ByVal clave As String) As String
    Select Case clave
        Case "Examinadores"
            ' También por la tablet que tiene asignada ahora (IMEI 1 o nº de serie)
            CamposBusqueda = "NumeroExaminador;Nombre;Apellidos;Nombre & ' ' & Apellidos;DNI;" & _
                "IdExaminador IN (SELECT A.IdExaminador FROM AsignacionesTablet AS A INNER JOIN Tablets AS T " & _
                "ON A.IdTablet = T.IdTablet WHERE A.FechaFin Is Null AND (T.IMEI1 Like {patron} OR T.NumeroSerie Like {patron}))"
        Case "Tablets"
            CamposBusqueda = "NumeroSerie;IMEI1;IMEI2;Marca;Modelo"
        Case "TarjetasSIM"
            CamposBusqueda = "IMEI;NumeroTelefono"
        Case "HistoricoTablets"
            CamposBusqueda = "NumeroExaminador;Nombre;Apellidos;Nombre & ' ' & Apellidos;DNI;NumeroSerie;IMEI1"
        Case "HistoricoSIM"
            CamposBusqueda = "SIM;NumeroSerie;IMEI1"
    End Select
End Function

Public Function QuitarFiltro()
    Dim f As Object, destino As Object, c As Object
    Set f = Screen.ActiveForm
    f!txtBuscar = Null
    For Each c In f.Controls
        If c.Name Like "chkEstado*" Then c.Value = True
    Next
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
    Dim actual As Variant, aviso As String, idAntigua As Variant

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
    If Not IsNull(actual) Then aviso = "- La tablet está asignada ahora a " & actual & "." & vbCrLf
    If Not ConfirmarCierre(aviso, fecha) Then Exit Function
    If Not FechaPosterior("AsignacionesTablet", "IdTablet = " & idTablet & " OR IdExaminador = " & idExaminador, fecha) Then Exit Function

    ' Si el examinador ya tenía otra tablet, la devuelve: se pregunta a qué estado pasa.
    idAntigua = DLookup("IdTablet", "qryAsignacionesTabletActuales", "IdExaminador = " & idExaminador)
    If Not IsNull(idAntigua) Then
        If Not PreguntarEstado(DLookup("Examinador", "qryAsignacionesTabletActuales", "IdExaminador = " & idExaminador) & _
                               " devuelve la tablet " & DLookup("Tablet", "qryAsignacionesTabletActuales", "IdTablet = " & idAntigua) & _
                               " el " & Format(fecha, "dd/mm/yyyy") & ".", idAntigua) Then Exit Function
    End If

    On Error GoTo Fallo
    DBEngine.BeginTrans
    If Not IsNull(idAntigua) Then RegistrarDevolucion idAntigua, idExaminador, fecha
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
    Dim f As Object, idTablet As Long, fecha As Date, actual As Variant, idExaminador As Variant

    Set f = Forms("frmAsignarTablet")
    If Not DatosCompletos(f, "cboTablet") Then Exit Function
    idTablet = f!cboTablet
    fecha = f!txtFecha

    actual = DLookup("Examinador", "qryAsignacionesTabletActuales", "IdTablet = " & idTablet)
    If IsNull(actual) Then MsgBox "Esa tablet no está asignada a ningún examinador.", vbInformation: Exit Function
    If Not FechaPosterior("AsignacionesTablet", "IdTablet = " & idTablet, fecha) Then Exit Function
    idExaminador = DLookup("IdExaminador", "qryAsignacionesTabletActuales", "IdTablet = " & idTablet)
    If Not PreguntarEstado(actual & " devuelve la tablet " & _
                           DLookup("Tablet", "qryAsignacionesTabletActuales", "IdTablet = " & idTablet) & _
                           " el " & Format(fecha, "dd/mm/yyyy") & ".", idTablet) Then Exit Function

    On Error GoTo Fallo
    DBEngine.BeginTrans
    RegistrarDevolucion idTablet, idExaminador, fecha
    ' Por si hubiera más de una asignación vigente de la tablet (datos antiguos)
    Ejecutar "PARAMETERS pFecha DateTime, pTablet Long; " & _
             "UPDATE AsignacionesTablet SET FechaFin = pFecha WHERE FechaFin Is Null AND IdTablet = pTablet;", _
             fecha, idTablet
    DBEngine.CommitTrans
    Terminar f, "Devolución registrada."
    Exit Function

Fallo:
    DBEngine.Rollback
    MsgBox "No se pudo guardar la devolución: " & Err.Description, vbCritical
End Function

' Abre frmElegirEstado para preguntar a qué estado pasa la tablet devuelta. Devuelve False si se
' cancela; si se acepta, el estado y la nota quedan en estadoElegido y notaElegida.
Private Function PreguntarEstado(ByVal mensaje As String, ByVal idTablet As Long) As Boolean
    mensajeEstado = mensaje & vbCrLf & "¿A qué estado pasa la tablet?"
    estadoInicial = DLookup("IdEstado", "Tablets", "IdTablet = " & idTablet)
    estadoElegido = Null
    notaElegida = Null
    DoCmd.OpenForm "frmElegirEstado", WindowMode:=acDialog      ' espera hasta que se cierre
    PreguntarEstado = Not IsNull(estadoElegido)
End Function

Public Function PrepararElegirEstado()
    With Forms("frmElegirEstado")
        !lblMensaje.Caption = mensajeEstado
        !cboEstado = estadoInicial
    End With
End Function

Public Function AceptarEstado()
    Dim f As Object
    Set f = Forms("frmElegirEstado")
    If IsNull(f!cboEstado) Then MsgBox "Elige el estado al que pasa la tablet.", vbExclamation: Exit Function
    estadoElegido = f!cboEstado.Value
    notaElegida = f!txtNota.Value
    DoCmd.Close acForm, "frmElegirEstado"
End Function

Public Function CancelarEstado()
    DoCmd.Close acForm, "frmElegirEstado"
End Function

' Cierra la asignación vigente de la tablet devuelta anotando en ella el estado al que pasa, cambia
' el estado de la tablet y lo registra en HistorialEstadosTablet. Va dentro de una transacción y usa
' el estado elegido en PreguntarEstado.
Private Sub RegistrarDevolucion(ByVal idTablet As Long, ByVal idExaminador As Long, ByVal fecha As Date)
    Dim estadoAnterior As Variant, nota As String
    estadoAnterior = DLookup("IdEstado", "Tablets", "IdTablet = " & idTablet)
    nota = "Devuelta. Pasa a estado: " & DLookup("Estado", "EstadosTablet", "IdEstado = " & estadoElegido)
    If Len(Nz(notaElegida, "")) > 0 Then nota = nota & " (" & notaElegida & ")"

    Ejecutar "PARAMETERS pFecha DateTime, pTablet Long, pExaminador Long, pNota Text(255); " & _
             "UPDATE AsignacionesTablet SET FechaFin = pFecha, " & _
             "Observaciones = Left(IIf(Observaciones Is Null, '', Observaciones & ' | ') & pNota, 255) " & _
             "WHERE FechaFin Is Null AND IdTablet = pTablet AND IdExaminador = pExaminador;", _
             fecha, idTablet, idExaminador, Left$(nota, 255)
    Ejecutar "PARAMETERS pEstado Long, pTablet Long; UPDATE Tablets SET IdEstado = pEstado WHERE IdTablet = pTablet;", _
             estadoElegido, idTablet
    AnotarEstado idTablet, fecha, estadoAnterior, estadoElegido, idExaminador, _
                 "Devolución" & IIf(Len(Nz(notaElegida, "")) > 0, ": " & notaElegida, "")
End Sub

Private Sub AnotarEstado(ByVal idTablet As Long, ByVal fecha As Date, ByVal estadoAnterior As Variant, _
                         ByVal estadoNuevo As Variant, ByVal idExaminador As Variant, ByVal observaciones As String)
    Ejecutar "PARAMETERS pTablet Long, pFecha DateTime, pAnterior Long, pNuevo Long, pExaminador Long, pObs Text(255); " & _
             "INSERT INTO HistorialEstadosTablet (IdTablet, Fecha, IdEstadoAnterior, IdEstadoNuevo, IdExaminador, Observaciones) " & _
             "VALUES (pTablet, pFecha, pAnterior, pNuevo, pExaminador, pObs);", _
             idTablet, fecha, estadoAnterior, estadoNuevo, idExaminador, Left$(observaciones, 255)
End Sub

' Antes de guardar la ficha de una tablet: si se ha cambiado su estado a mano, se anota.
Public Function RegistrarCambioEstadoFicha()
    Dim f As Object
    Set f = Forms("frmTablets")
    If f.NewRecord Then Exit Function
    If Nz(f!IdEstado.OldValue, 0) = Nz(f!IdEstado.Value, 0) Then Exit Function
    AnotarEstado f!IdTablet, Date, f!IdEstado.OldValue, f!IdEstado.Value, Null, "Cambio en la ficha de la tablet"
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

' Combina el texto del buscador con las casillas de situación (si el formulario las tiene).
Private Sub AplicarFiltro(ByVal f As Object, ByVal destino As Object, ByVal campos As String)
    Dim texto As String, campo As Variant, filtro As String, patron As String, casillas As String
    texto = Trim$(Nz(f!txtBuscar, ""))
    If Len(texto) > 0 Then
        ' Comillas duplicadas y comodines de Like (* ? # [) tratados como texto normal
        texto = Replace(texto, "'", "''")
        texto = Replace(texto, "[", "[[]")
        texto = Replace(texto, "*", "[*]")
        texto = Replace(texto, "?", "[?]")
        texto = Replace(texto, "#", "[#]")
        patron = "'*" & texto & "*'"
        For Each campo In Split(campos, ";")
            If InStr(campo, "{patron}") > 0 Then
                filtro = filtro & " OR (" & Replace(campo, "{patron}", patron) & ")"
            Else
                filtro = filtro & " OR (" & campo & ") Like " & patron
            End If
        Next
        filtro = "(" & Mid$(filtro, 5) & ")"
    End If
    casillas = FiltroCasillas(f)
    If Len(casillas) > 0 Then
        If Len(filtro) > 0 Then filtro = filtro & " AND "
        filtro = filtro & "(" & casillas & ")"
    End If

    If destino.Dirty Then destino.Dirty = False
    If Len(filtro) = 0 Then
        destino.FilterOn = False
        Exit Sub
    End If
    destino.Filter = filtro
    destino.FilterOn = True
    If destino.RecordsetClone.RecordCount = 0 Then MsgBox "No hay ningún registro que cumpla la búsqueda.", vbInformation
End Sub

' Casillas chkEstado* del formulario; su Tag es "campo|Id". Devuelve la condición para las marcadas,
' o "" si están todas marcadas o el formulario no tiene casillas.
Private Function FiltroCasillas(ByVal f As Object) As String
    Dim c As Object, partes() As String, campo As String, ids As String, total As Long, marcadas As Long
    For Each c In f.Controls
        If c.Name Like "chkEstado*" Then
            partes = Split(c.Tag, "|")
            campo = partes(0)
            total = total + 1
            If Nz(c.Value, False) Then
                marcadas = marcadas + 1
                ids = ids & "," & partes(1)
            End If
        End If
    Next
    If total = 0 Or marcadas = total Then Exit Function
    If marcadas = 0 Then
        FiltroCasillas = "1 = 0"            ' ninguna marcada: no se muestra nada
    Else
        FiltroCasillas = campo & " IN (" & Mid$(ids, 2) & ")"
    End If
End Function

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
