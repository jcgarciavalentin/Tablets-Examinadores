Option Compare Database
Option Explicit

' Crea las consultas, los formularios y el menú de GestionTablets.accdb.
' ConstruirFormularios.vbs carga este módulo, ejecuta ConstruirInterfaz y lo elimina después.
' Se puede volver a ejecutar: primero borra los objetos que crea.

Private Const CM As Long = 567                  ' twips por centímetro
Private Const ANCHO_ETIQUETA As Single = 3.8
Private Const VISTA_FORMULARIO As Integer = 0
Private Const VISTA_HOJA_DATOS As Integer = 2
Private Const DB_TEXT As Integer = 10

Private Const FORMULARIOS As String = "frmMenu,frmExaminadores,frmTablets,frmTarjetasSIM,frmAsignarTablet," & _
    "frmEstadosTablet,frmEstadosExaminador," & _
    "frmAsignarSIM,frmEstadoTablets,frmHistoricoTablets,frmHistoricoSIM,frmSubHistoricoTablets,frmSubHistoricoSIM," & _
    "frmSubHistorialExaminador," & _
    "frmSubHistorialTablet,frmSubHistorialSIMTablet,frmSubHistorialTabletsSIM,frmSubTabletsActuales,frmSubSIMActuales"
Private Const CONSULTAS As String = "qryEstadoTablets,qryHistorialTablets,qryHistorialSIM," & _
    "qryAsignacionesTabletActuales,qrySIMActuales"

' Descripciones legibles; una parte vacía no deja su separador suelto.
Private Const DESC_EXAMINADOR As String = "E.Apellidos & ', ' & E.Nombre & IIf(E.NumeroExaminador Is Null, '', ' (' & E.NumeroExaminador & ')')"
Private Const DESC_TABLET As String = "T.NumeroSerie & IIf(T.Marca Is Null, '', ' - ' & T.Marca) & IIf(T.Modelo Is Null, '', ' ' & T.Modelo)"
Private Const DESC_SIM As String = "M.IMEI"

Public Function ConstruirInterfaz()
    BorrarObjetosPrevios
    CrearConsultas

    ' Hojas de datos de solo lectura. Los subformularios deben existir antes que sus formularios.
    CrearHojaDatos "frmSubHistorialExaminador", "qryHistorialTablets", "Histórico de tablets", _
        Array("Tablet|Tablet|6", "FechaInicio|Entrega|2.4", "FechaFin|Devolución|2.4", "Vigente|Vigente|1.6", "Observaciones|Observaciones|5")
    CrearHojaDatos "frmSubHistorialTablet", "qryHistorialTablets", "Histórico de examinadores", _
        Array("Examinador|Examinador|6", "DNI|DNI|2.2", "FechaInicio|Entrega|2.4", "FechaFin|Devolución|2.4", "Vigente|Vigente|1.6", "Observaciones|Observaciones|4")
    CrearHojaDatos "frmSubHistorialSIMTablet", "qryHistorialSIM", "Histórico de tarjetas SIM", _
        Array("SIM|Tarjeta SIM (IMEI)|6", "FechaInicio|Instalación|2.4", "FechaFin|Retirada|2.4", "Vigente|Vigente|1.6", "Observaciones|Observaciones|4")
    CrearHojaDatos "frmSubHistorialTabletsSIM", "qryHistorialSIM", "Histórico de tablets", _
        Array("Tablet|Tablet|6", "FechaInicio|Instalación|2.4", "FechaFin|Retirada|2.4", "Vigente|Vigente|1.6", "Observaciones|Observaciones|4")
    CrearHojaDatos "frmSubTabletsActuales", "qryAsignacionesTabletActuales", "Tablets asignadas", _
        Array("Examinador|Examinador|5.5", "DNI|DNI|2.2", "Tablet|Tablet|5.5", "FechaInicio|Entrega|2.4", "Observaciones|Observaciones|4")
    CrearHojaDatos "frmSubSIMActuales", "qrySIMActuales", "SIM instaladas", _
        Array("Tablet|Tablet|6", "SIM|Tarjeta SIM (IMEI)|5", "FechaInicio|Instalación|2.4", "Observaciones|Observaciones|4")
    CrearHojaDatos "frmEstadoTablets", "qryEstadoTablets", "Situación actual de las tablets", _
        Array("NumeroSerie|Nº de serie|3.5", "Marca|Marca|3", "Modelo|Modelo|3", "Estado|Estado|2.5", _
              "Examinador|Examinador|5.5", "AsignadaDesde|Asignada desde|2.6", "SIM|Tarjeta SIM (IMEI)|5")
    CrearHojaDatos "frmSubHistoricoTablets", "qryHistorialTablets", "Histórico examinadores - tablets", _
        Array("Examinador|Examinador|5.5", "NumeroExaminador|Nº examinador|2.4", "DNI|DNI|2.2", _
              "NumeroSerie|Nº serie tablet|3", "IMEI1|IMEI 1 tablet|3.5", "Tablet|Tablet|5", "FechaInicio|Entrega|2.4", _
              "FechaFin|Devolución|2.4", "Vigente|Vigente|1.6", "Observaciones|Observaciones|4")
    CrearHojaDatos "frmSubHistoricoSIM", "qryHistorialSIM", "Histórico tablets - tarjetas SIM", _
        Array("NumeroSerie|Nº serie tablet|3", "IMEI1|IMEI 1 tablet|3.5", "Tablet|Tablet|5", "SIM|Tarjeta SIM (IMEI)|5", _
              "FechaInicio|Instalación|2.4", "FechaFin|Retirada|2.4", "Vigente|Vigente|1.6", "Observaciones|Observaciones|4")
    CrearHojaDatos "frmEstadosTablet", "SELECT * FROM EstadosTablet ORDER BY Estado;", "Estados de tablet", _
        Array("Estado|Estado|6"), True
    CrearHojaDatos "frmEstadosExaminador", "SELECT * FROM EstadosExaminador ORDER BY EstadoExaminador;", _
        "Situaciones de examinador", Array("EstadoExaminador|Situación|6"), True

    CrearFormExaminadores
    CrearFormTablets
    CrearFormTarjetasSIM
    CrearFormAsignarTablet
    CrearFormAsignarSIM
    CrearFormHistorico "frmHistoricoTablets", "Histórico examinadores - tablets", "frmSubHistoricoTablets", _
        "NumeroExaminador;Nombre;Apellidos;Nombre & ' ' & Apellidos;NumeroSerie;IMEI1", _
        "Nombre, apellidos o nº de examinador; nº de serie o IMEI 1 de la tablet"
    CrearFormHistorico "frmHistoricoSIM", "Histórico tablets - tarjetas SIM", "frmSubHistoricoSIM", _
        "SIM;NumeroSerie;IMEI1", "IMEI de la SIM; nº de serie o IMEI 1 de la tablet"
    CrearMenu
    EstablecerFormularioInicio
    Application.RefreshDatabaseWindow
End Function

' ---------------------------------------------------------------------------------------------
' Consultas
' ---------------------------------------------------------------------------------------------

Private Sub CrearConsultas()
    Dim db As Object
    Set db = CurrentDb

    Const DESDE_TABLETS As String = "FROM (AsignacionesTablet AS A INNER JOIN Examinadores AS E ON A.IdExaminador = E.IdExaminador) " & _
                                    "INNER JOIN Tablets AS T ON A.IdTablet = T.IdTablet "
    Const DESDE_SIM As String = "FROM (AsignacionesSIM AS S INNER JOIN Tablets AS T ON S.IdTablet = T.IdTablet) " & _
                                "INNER JOIN TarjetasSIM AS M ON S.IdSIM = M.IdSIM "

    db.CreateQueryDef "qryAsignacionesTabletActuales", _
        "SELECT A.IdAsignacion, A.IdExaminador, A.IdTablet, " & DESC_EXAMINADOR & " AS Examinador, E.DNI, " & _
        DESC_TABLET & " AS Tablet, A.FechaInicio, A.Observaciones " & DESDE_TABLETS & _
        "WHERE A.FechaFin Is Null ORDER BY E.Apellidos, E.Nombre;"

    db.CreateQueryDef "qryHistorialTablets", _
        "SELECT A.IdAsignacion, A.IdExaminador, A.IdTablet, " & DESC_EXAMINADOR & " AS Examinador, E.DNI, " & _
        "E.NumeroExaminador, E.Nombre, E.Apellidos, T.NumeroSerie, T.IMEI1, " & _
        DESC_TABLET & " AS Tablet, A.FechaInicio, A.FechaFin, IIf(A.FechaFin Is Null, 'Sí', 'No') AS Vigente, " & _
        "A.Observaciones " & DESDE_TABLETS & "ORDER BY A.FechaInicio DESC, A.IdAsignacion DESC;"

    db.CreateQueryDef "qrySIMActuales", _
        "SELECT S.IdAsignacionSIM, S.IdTablet, S.IdSIM, " & DESC_TABLET & " AS Tablet, " & DESC_SIM & " AS SIM, " & _
        "S.FechaInicio, S.Observaciones " & DESDE_SIM & _
        "WHERE S.FechaFin Is Null ORDER BY T.NumeroSerie;"

    db.CreateQueryDef "qryHistorialSIM", _
        "SELECT S.IdAsignacionSIM, S.IdTablet, S.IdSIM, T.NumeroSerie, T.IMEI1, " & DESC_TABLET & " AS Tablet, " & _
        DESC_SIM & " AS SIM, S.FechaInicio, S.FechaFin, IIf(S.FechaFin Is Null, 'Sí', 'No') AS Vigente, S.Observaciones " & _
        DESDE_SIM & "ORDER BY S.FechaInicio DESC, S.IdAsignacionSIM DESC;"

    db.CreateQueryDef "qryEstadoTablets", _
        "SELECT T.IdTablet, T.NumeroSerie, T.Marca, T.Modelo, ET.Estado, Q.Examinador, Q.FechaInicio AS AsignadaDesde, S.SIM " & _
        "FROM ((Tablets AS T LEFT JOIN EstadosTablet AS ET ON T.IdEstado = ET.IdEstado) " & _
        "LEFT JOIN qryAsignacionesTabletActuales AS Q ON T.IdTablet = Q.IdTablet) " & _
        "LEFT JOIN qrySIMActuales AS S ON T.IdTablet = S.IdTablet ORDER BY T.NumeroSerie;"
End Sub

' ---------------------------------------------------------------------------------------------
' Formularios de datos
' ---------------------------------------------------------------------------------------------

Private Sub CrearFormExaminadores()
    Dim f As Form, y As Single
    Set f = NuevoFormulario("SELECT * FROM Examinadores ORDER BY Apellidos, Nombre;", "Examinadores")
    PrepararFicha f
    PonerTitulo f, "Examinadores", True
    Buscador f, "FiltrarFormulario", "NumeroExaminador;Nombre;Apellidos;Nombre & ' ' & Apellidos", _
        "Nº de examinador, nombre o apellidos"

    y = 2.4
    Campo f, acTextBox, "NumeroExaminador", "NumeroExaminador", "Nº de examinador", 0.5, y, 3: y = y + 0.8
    Campo f, acTextBox, "Nombre", "Nombre", "Nombre", 0.5, y, 6: y = y + 0.8
    Campo f, acTextBox, "Apellidos", "Apellidos", "Apellidos", 0.5, y, 9: y = y + 0.8
    Campo f, acTextBox, "DNI", "DNI", "DNI/NIE", 0.5, y, 3: y = y + 0.8
    Combo f, "IdEstadoExam", "IdEstadoExam", "Situación", 0.5, y, 4, _
        "SELECT IdEstadoExam, EstadoExaminador FROM EstadosExaminador ORDER BY EstadoExaminador;", "0;4": y = y + 0.8
    Memo f, "Observaciones", "Observaciones", 0.5, y, 13, 1.6: y = y + 2

    Subformulario f, "subHistorial", "frmSubHistorialExaminador", "Histórico de tablets asignadas", 0.5, y, 17.5, 5, "IdExaminador"
    GuardarFormulario f, "frmExaminadores"
End Sub

Private Sub CrearFormTablets()
    Dim f As Form, c As Control, y As Single
    Set f = NuevoFormulario("SELECT * FROM Tablets ORDER BY NumeroSerie;", "Tablets")
    PrepararFicha f
    PonerTitulo f, "Tablets", True
    Buscador f, "FiltrarFormulario", "NumeroSerie;IMEI1;IMEI2;Marca;Modelo", "Nº de serie, IMEI, marca o modelo"

    ' Dos columnas de campos
    y = 2.4
    Campo f, acTextBox, "NumeroSerie", "NumeroSerie", "Nº de serie", 0.5, y, 4.5
    Campo f, acTextBox, "IMEI1", "IMEI1", "IMEI 1", 9.5, y, 4.5: y = y + 0.8
    Campo f, acTextBox, "Marca", "Marca", "Marca", 0.5, y, 4.5
    Campo f, acTextBox, "IMEI2", "IMEI2", "IMEI 2", 9.5, y, 4.5: y = y + 0.8
    Campo f, acTextBox, "Modelo", "Modelo", "Modelo", 0.5, y, 4.5
    Combo f, "IdEstado", "IdEstado", "Estado", 9.5, y, 4.5, _
        "SELECT IdEstado, Estado FROM EstadosTablet ORDER BY Estado;", "0;4.5": y = y + 0.8
    Set c = Campo(f, acTextBox, "FechaRecepcion", "FechaRecepcion", "Fecha de recepción", 0.5, y, 2.6)
    c.Format = "Short Date"
    Set c = Campo(f, acTextBox, "FechaDesafectacion", "FechaDesafectacion", "Fecha desafectación", 9.5, y, 2.6): y = y + 0.8
    c.Format = "Short Date"
    Memo f, "Observaciones", "Observaciones", 0.5, y, 13.7, 1.4: y = y + 1.6
    Memo f, "Incidencias", "Incidencias", 0.5, y, 13.7, 1.4: y = y + 1.8

    Subformulario f, "subHistorialExaminadores", "frmSubHistorialTablet", "Histórico de examinadores", 0.5, y, 17.5, 4, "IdTablet"
    Subformulario f, "subHistorialSIM", "frmSubHistorialSIMTablet", "Histórico de tarjetas SIM", 0.5, y + 4.9, 17.5, 3.5, "IdTablet"
    GuardarFormulario f, "frmTablets"
End Sub

Private Sub CrearFormTarjetasSIM()
    Dim f As Form, y As Single
    Set f = NuevoFormulario("SELECT * FROM TarjetasSIM ORDER BY IMEI;", "Tarjetas SIM")
    PrepararFicha f
    PonerTitulo f, "Tarjetas SIM", True
    Buscador f, "FiltrarFormulario", "IMEI", "IMEI de la tarjeta"

    y = 2.4
    Campo f, acTextBox, "IMEI", "IMEI", "IMEI", 0.5, y, 5: y = y + 0.8
    Campo f, acTextBox, "PIN", "PIN", "PIN", 0.5, y, 2.5: y = y + 0.8
    Campo f, acTextBox, "PUK", "PUK", "PUK", 0.5, y, 2.5: y = y + 1.2

    Subformulario f, "subHistorial", "frmSubHistorialTabletsSIM", "Histórico de tablets", 0.5, y, 17.5, 5, "IdSIM"
    GuardarFormulario f, "frmTarjetasSIM"
End Sub

' ---------------------------------------------------------------------------------------------
' Formularios de asignación (sin origen de datos; los botones llaman a modGestionTablets)
' ---------------------------------------------------------------------------------------------

Private Sub CrearFormAsignarTablet()
    Dim f As Form, c As Control
    Set f = NuevoFormulario("", "Asignar / devolver tablet")
    PrepararDialogo f
    PonerTitulo f, "Asignación de tablets a examinadores"

    Combo f, "cboExaminador", "", "Examinador", 0.5, 1.8, 8, _
        "SELECT E.IdExaminador, " & DESC_EXAMINADOR & " AS Examinador, EE.EstadoExaminador, Q.Tablet " & _
        "FROM (Examinadores AS E LEFT JOIN EstadosExaminador AS EE ON E.IdEstadoExam = EE.IdEstadoExam) " & _
        "LEFT JOIN qryAsignacionesTabletActuales AS Q ON E.IdExaminador = Q.IdExaminador " & _
        "ORDER BY E.Apellidos, E.Nombre;", "0;6.5;2.5;6"
    Combo f, "cboTablet", "", "Tablet", 0.5, 2.6, 8, _
        "SELECT T.IdTablet, " & DESC_TABLET & " AS Tablet, ET.Estado, Q.Examinador " & _
        "FROM (Tablets AS T LEFT JOIN EstadosTablet AS ET ON T.IdEstado = ET.IdEstado) " & _
        "LEFT JOIN qryAsignacionesTabletActuales AS Q ON T.IdTablet = Q.IdTablet " & _
        "WHERE T.FechaDesafectacion Is Null ORDER BY T.NumeroSerie;", "0;6;2.5;6.5"
    CamposFechaYObservaciones f, "Fecha"

    Boton f, "cmdAsignar", "Asignar tablet", "=AsignarTablet()", 4.3, 5.2, 4.5
    Boton f, "cmdDevolver", "Registrar devolución", "=DevolverTablet()", 9.3, 5.2, 4.5
    Nota f, "Al asignar, si la tablet o el examinador tenían otra asignación vigente, se cierra con la fecha indicada. " & _
            "Para devolver basta con elegir la tablet y la fecha.", 0.5, 6.2, 17.5

    Subformulario f, "subActuales", "frmSubTabletsActuales", "Asignaciones vigentes", 0.5, 7.4, 17.5, 6
    GuardarFormulario f, "frmAsignarTablet"
End Sub

Private Sub CrearFormAsignarSIM()
    Dim f As Form
    Set f = NuevoFormulario("", "Asociar / retirar tarjeta SIM")
    PrepararDialogo f
    PonerTitulo f, "Tarjetas SIM de las tablets"

    Combo f, "cboTablet", "", "Tablet", 0.5, 1.8, 8, _
        "SELECT T.IdTablet, " & DESC_TABLET & " AS Tablet, ET.Estado, S.SIM " & _
        "FROM (Tablets AS T LEFT JOIN EstadosTablet AS ET ON T.IdEstado = ET.IdEstado) " & _
        "LEFT JOIN qrySIMActuales AS S ON T.IdTablet = S.IdTablet " & _
        "WHERE T.FechaDesafectacion Is Null ORDER BY T.NumeroSerie;", "0;6;2.5;5"
    Combo f, "cboSIM", "", "Tarjeta SIM", 0.5, 2.6, 8, _
        "SELECT M.IdSIM, " & DESC_SIM & " AS SIM, S.Tablet " & _
        "FROM TarjetasSIM AS M LEFT JOIN qrySIMActuales AS S ON M.IdSIM = S.IdSIM " & _
        "ORDER BY M.IMEI;", "0;5;6"
    CamposFechaYObservaciones f, "Fecha"

    Boton f, "cmdAsignar", "Asociar SIM", "=AsignarSIM()", 4.3, 5.2, 4.5
    Boton f, "cmdRetirar", "Retirar SIM de la tablet", "=RetirarSIM()", 9.3, 5.2, 4.5
    Nota f, "Al asociar, si la tablet o la SIM tenían otra asociación vigente, se cierra con la fecha indicada. " & _
            "Para retirar basta con elegir la tablet y la fecha.", 0.5, 6.2, 17.5

    Subformulario f, "subActuales", "frmSubSIMActuales", "SIM instaladas actualmente", 0.5, 7.4, 17.5, 6
    GuardarFormulario f, "frmAsignarSIM"
End Sub

Private Sub CrearFormHistorico(ByVal nombre As String, ByVal titulo As String, ByVal hojaDatos As String, _
                               ByVal campos As String, ByVal ayuda As String)
    Dim f As Form
    Set f = NuevoFormulario("", titulo)
    PrepararDialogo f
    PonerTitulo f, titulo, True
    Buscador f, "FiltrarSubformulario", campos, ayuda
    Subformulario f, "subResultados", hojaDatos, "Resultados (pulsa ""Ver todos"" para quitar el filtro)", 0.5, 2.3, 30, 14
    GuardarFormulario f, nombre
End Sub

Private Sub CamposFechaYObservaciones(ByVal f As Form, ByVal etiqueta As String)
    Dim c As Control
    Set c = Campo(f, acTextBox, "txtFecha", "", etiqueta, 0.5, 3.4, 2.6)
    c.Format = "Short Date"
    c.DefaultValue = "Date()"
    Campo f, acTextBox, "txtObservaciones", "", "Observaciones", 0.5, 4.2, 13.7
End Sub

Private Sub CrearMenu()
    Dim f As Form
    Set f = NuevoFormulario("", "Gestión de tablets")
    PrepararDialogo f
    f.OnActivate = ""
    PonerTitulo f, "Gestión de tablets"

    Seccion f, "Datos", 0.5
    Boton f, "cmdExaminadores", "Examinadores", "=AbrirFormulario(""frmExaminadores"")", 0.5, 2.2, 5.5
    Boton f, "cmdTablets", "Tablets", "=AbrirFormulario(""frmTablets"")", 0.5, 3.2, 5.5
    Boton f, "cmdSIM", "Tarjetas SIM", "=AbrirFormulario(""frmTarjetasSIM"")", 0.5, 4.2, 5.5
    Boton f, "cmdEstadosTablet", "Estados de tablet", "=AbrirHojaDatos(""frmEstadosTablet"")", 0.5, 5.5, 5.5
    Boton f, "cmdEstadosExam", "Situaciones de examinador", "=AbrirHojaDatos(""frmEstadosExaminador"")", 0.5, 6.5, 5.5

    Seccion f, "Asignaciones", 6.5
    Boton f, "cmdAsignarTablet", "Asignar / devolver tablet", "=AbrirFormulario(""frmAsignarTablet"")", 6.5, 2.2, 5.5
    Boton f, "cmdAsignarSIM", "Asociar / retirar SIM", "=AbrirFormulario(""frmAsignarSIM"")", 6.5, 3.2, 5.5

    Seccion f, "Consultas", 12.5
    Boton f, "cmdEstado", "Situación actual", "=AbrirHojaDatos(""frmEstadoTablets"")", 12.5, 2.2, 5.5
    Boton f, "cmdHistTablets", "Histórico examinador - tablet", "=AbrirFormulario(""frmHistoricoTablets"")", 12.5, 3.2, 5.5
    Boton f, "cmdHistSIM", "Histórico tablet - SIM", "=AbrirFormulario(""frmHistoricoSIM"")", 12.5, 4.2, 5.5

    f.Section(acDetail).Height = 7.8 * CM
    GuardarFormulario f, "frmMenu"
End Sub

Private Sub EstablecerFormularioInicio()
    Dim db As Object
    Set db = CurrentDb
    On Error Resume Next
    db.Properties("StartUpForm") = "frmMenu"
    If Err.Number <> 0 Then
        Err.Clear
        db.Properties.Append db.CreateProperty("StartUpForm", DB_TEXT, "frmMenu")
    End If
End Sub

' ---------------------------------------------------------------------------------------------
' Utilidades de diseño (posiciones y tamaños en centímetros)
' ---------------------------------------------------------------------------------------------

Private Sub BorrarObjetosPrevios()
    Dim nombre As Variant
    For Each nombre In Split(FORMULARIOS, ",")
        If Existe(CurrentProject.AllForms, nombre) Then DoCmd.DeleteObject acForm, nombre
    Next
    For Each nombre In Split(CONSULTAS, ",")
        If Existe(CurrentData.AllQueries, nombre) Then DoCmd.DeleteObject acQuery, nombre
    Next
End Sub

Private Function Existe(ByVal coleccion As Object, ByVal nombre As String) As Boolean
    Dim o As Object
    For Each o In coleccion
        If o.Name = nombre Then Existe = True: Exit Function
    Next
End Function

Private Function NuevoFormulario(ByVal origen As String, ByVal titulo As String) As Form
    Dim f As Form
    Set f = CreateForm()
    If Len(origen) > 0 Then f.RecordSource = origen
    f.Caption = titulo
    f.AutoCenter = True
    Set NuevoFormulario = f
End Function

' CreateForm crea "Formulario1"; se guarda con ese nombre y luego se renombra.
Private Sub GuardarFormulario(ByVal f As Form, ByVal nombre As String)
    Dim temporal As String
    temporal = f.Name
    DoCmd.Close acForm, temporal, acSaveYes
    DoCmd.Rename nombre, acForm, temporal
End Sub

Private Sub PrepararFicha(ByVal f As Form)
    f.DefaultView = VISTA_FORMULARIO
    f.RecordSelectors = False
    f.Width = 18.5 * CM
    f.OnActivate = "=RefrescarListas()"
End Sub

Private Sub PrepararDialogo(ByVal f As Form)
    PrepararFicha f
    f.NavigationButtons = False
    f.DividingLines = False
    f.ScrollBars = 0
End Sub

' columnas: "Campo|Título|ancho en cm". Sin editable, la hoja de datos es de solo lectura.
Private Sub CrearHojaDatos(ByVal nombre As String, ByVal origen As String, ByVal titulo As String, _
                           ByVal columnas As Variant, Optional ByVal editable As Boolean = False)
    Dim f As Form, c As Control, l As Control, partes() As String, i As Long
    Set f = NuevoFormulario(origen, titulo)
    f.DefaultView = VISTA_HOJA_DATOS
    f.AllowEdits = editable
    f.AllowAdditions = editable
    f.AllowDeletions = editable
    For i = 0 To UBound(columnas)
        partes = Split(columnas(i), "|")
        Set c = CreateControl(f.Name, acTextBox, acDetail, , , ANCHO_ETIQUETA * CM, i * 0.7 * CM, 5 * CM, 0.6 * CM)
        c.Name = partes(0)
        c.ControlSource = partes(0)
        If partes(0) Like "Fecha*" Or partes(0) = "AsignadaDesde" Then c.Format = "Short Date"
        On Error Resume Next            ' ColumnWidth solo afecta a la vista Hoja de datos
        c.ColumnWidth = Val(partes(2)) * CM
        On Error GoTo 0
        Set l = CreateControl(f.Name, acLabel, acDetail, c.Name, , 0, i * 0.7 * CM, (ANCHO_ETIQUETA - 0.2) * CM, 0.6 * CM)
        l.Caption = partes(1)       ' título de la columna en la hoja de datos
    Next
    GuardarFormulario f, nombre
End Sub

' Control con su etiqueta asociada a la izquierda. origen vacío = control independiente.
Private Function Campo(ByVal f As Form, ByVal tipo As AcControlType, ByVal nombre As String, ByVal origen As String, _
                       ByVal etiqueta As String, ByVal x As Single, ByVal y As Single, ByVal ancho As Single, _
                       Optional ByVal alto As Single = 0.6) As Control
    Dim c As Control, l As Control
    Set c = CreateControl(f.Name, tipo, acDetail, , , (x + ANCHO_ETIQUETA) * CM, y * CM, ancho * CM, alto * CM)
    c.Name = nombre
    If Len(origen) > 0 Then c.ControlSource = origen
    Set l = CreateControl(f.Name, acLabel, acDetail, c.Name, , x * CM, y * CM, (ANCHO_ETIQUETA - 0.2) * CM, 0.6 * CM)
    l.Name = "lbl" & nombre
    l.Caption = etiqueta
    Set Campo = c
End Function

Private Sub Memo(ByVal f As Form, ByVal campoTabla As String, ByVal etiqueta As String, _
                 ByVal x As Single, ByVal y As Single, ByVal ancho As Single, ByVal alto As Single)
    Dim c As Control
    Set c = Campo(f, acTextBox, campoTabla, campoTabla, etiqueta, x, y, ancho, alto)
    c.ScrollBars = 2                    ' vertical
    c.EnterKeyBehavior = True           ' Intro = nueva línea
End Sub

' anchos: anchos de columna en cm separados por ";" (la primera, el Id, oculta con 0).
Private Function Combo(ByVal f As Form, ByVal nombre As String, ByVal origen As String, ByVal etiqueta As String, _
                       ByVal x As Single, ByVal y As Single, ByVal ancho As Single, _
                       ByVal filas As String, ByVal anchos As String) As Control
    Dim c As Control, partes() As String, i As Long, twips As String, total As Long
    Set c = Campo(f, acComboBox, nombre, origen, etiqueta, x, y, ancho)
    partes = Split(anchos, ";")
    For i = 0 To UBound(partes)
        If i > 0 Then twips = twips & ";"
        twips = twips & CLng(Val(partes(i)) * CM)
        total = total + CLng(Val(partes(i)) * CM)
    Next
    c.RowSource = filas
    c.ColumnCount = UBound(partes) + 1
    c.BoundColumn = 1
    c.ColumnWidths = twips
    c.ListWidth = total
    c.ListRows = 16
    c.LimitToList = True
    Set Combo = c
End Function

' Cuadro de texto txtBuscar arriba a la derecha; funcion es FiltrarFormulario o FiltrarSubformulario
' (modGestionTablets) y campos, los campos donde buscar separados por ";".
Private Sub Buscador(ByVal f As Form, ByVal funcion As String, ByVal campos As String, ByVal ayuda As String)
    Dim c As Control, l As Control
    Set c = Campo(f, acTextBox, "txtBuscar", "", "Buscar", 8.2, 0.4, 4.2)
    c.AfterUpdate = "=" & funcion & "(""" & campos & """)"     ' se busca al pulsar Intro
    Boton f, "cmdVerTodos", "Ver todos", "=QuitarFiltro()", 16.4, 0.3, 2
    Set l = CreateControl(f.Name, acLabel, acDetail, , , 12 * CM, 1.05 * CM, 6.3 * CM, 1 * CM)
    l.Caption = ayuda & " (escribe y pulsa Intro)"
    l.FontSize = 8
    l.FontItalic = True
End Sub

Private Sub Subformulario(ByVal f As Form, ByVal nombre As String, ByVal origen As String, ByVal titulo As String, _
                          ByVal x As Single, ByVal y As Single, ByVal ancho As Single, ByVal alto As Single, _
                          Optional ByVal vinculo As String = "")
    Dim s As Control
    Seccion f, titulo, x, y, ancho
    Set s = CreateControl(f.Name, acSubform, acDetail, , , x * CM, (y + 0.7) * CM, ancho * CM, alto * CM)
    s.Name = nombre
    s.SourceObject = origen
    If Len(vinculo) > 0 Then
        s.LinkMasterFields = vinculo
        s.LinkChildFields = vinculo
    End If
End Sub

' Con buscador el título deja libre la parte derecha y puede ocupar dos líneas.
Private Sub PonerTitulo(ByVal f As Form, ByVal texto As String, Optional ByVal conBuscador As Boolean = False)
    Dim l As Control
    Set l = CreateControl(f.Name, acLabel, acDetail, , , 0.5 * CM, 0.3 * CM, _
                          IIf(conBuscador, 7.5, 17.5) * CM, IIf(conBuscador, 1.8, 0.9) * CM)
    l.Caption = texto
    l.FontSize = 16
    l.FontBold = True
End Sub

Private Sub Seccion(ByVal f As Form, ByVal texto As String, ByVal x As Single, _
                    Optional ByVal y As Single = 1.4, Optional ByVal ancho As Single = 5.5)
    Dim l As Control
    Set l = CreateControl(f.Name, acLabel, acDetail, , , x * CM, y * CM, ancho * CM, 0.6 * CM)
    l.Caption = texto
    l.FontSize = 11
    l.FontBold = True
End Sub

Private Sub Nota(ByVal f As Form, ByVal texto As String, ByVal x As Single, ByVal y As Single, ByVal ancho As Single)
    Dim l As Control
    Set l = CreateControl(f.Name, acLabel, acDetail, , , x * CM, y * CM, ancho * CM, 1 * CM)
    l.Caption = texto
    l.FontItalic = True
End Sub

Private Sub Boton(ByVal f As Form, ByVal nombre As String, ByVal texto As String, ByVal accion As String, _
                  ByVal x As Single, ByVal y As Single, ByVal ancho As Single)
    Dim b As Control
    Set b = CreateControl(f.Name, acCommandButton, acDetail, , , x * CM, y * CM, ancho * CM, 0.8 * CM)
    b.Name = nombre
    b.Caption = texto
    b.OnClick = accion
End Sub
