Option Compare Database
Option Explicit

' Copia los datos de la base de datos antigua de tablets a GestionTablets.accdb.
' ImportarDatosAntiguos.vbs carga este módulo, ejecuta ImportarDatosAntiguos y lo elimina después.
'
' Antes de importar vacía las tablas de datos de la base nueva (examinadores, tablets, SIM y
' asignaciones), así que se puede repetir tantas veces como haga falta tras corregir datos en la
' antigua. Las tablas de estados no se vacían: se les añaden los valores de la antigua que falten.
' Lo que no se puede pasar tal cual se anota en la tabla ImportacionIncidencias.

Private Const DB_OPEN_DYNASET As Long = 2
Private Const DB_OPEN_SNAPSHOT As Long = 4
Private Const DB_FAIL_ON_ERROR As Long = 128
Private Const SIN_DATO As String = "(sin dato)"

Private baseNueva As Object
Private baseAntigua As Object
Private rsIncidencias As Object
Private numIncidencias As Long

' Claves (sin distinguir mayúsculas) -> Id en la base nueva
Private idEstadoTablet As Object
Private idSituacion As Object
Private idCompaniaNA As Variant     ' compañía por defecto de las SIM (la base antigua no la tiene)
Private idEstadoSim As Object
Private idExaminadorPorDni As Object
Private idTabletPorSerie As Object
Private idSimPorNumero As Object

Public Function ImportarDatosAntiguos(ByVal rutaAntigua As String) As String
    Dim resumen As String
    On Error GoTo Fallo

    Set baseNueva = CurrentDb
    Set baseAntigua = DBEngine.OpenDatabase(rutaAntigua, False, True)     ' solo lectura
    PrepararIncidencias
    VaciarTablasNuevas

    Set idEstadoTablet = CargarDiccionario("EstadosTablet", "Estado", "IdEstado")
    Set idSituacion = CargarDiccionario("EstadosExaminador", "EstadoExaminador", "IdEstadoExam")
    Set idEstadoSim = CargarDiccionario("EstadosSIM", "EstadoSIM", "IdEstadoSIM")
    idCompaniaNA = DLookup("IdCompania", "Companias", "Compania = 'N/A'")
    Set idExaminadorPorDni = NuevoDiccionario()
    Set idTabletPorSerie = NuevoDiccionario()
    Set idSimPorNumero = NuevoDiccionario()

    ImportarEstados
    ImportarExaminadores
    ImportarTablets
    ImportarAsignacionesTablet
    ImportarSIM
    ComprobarAsignacionesVigentes

    baseAntigua.Close
    rsIncidencias.Close

    resumen = "Importación terminada." & vbCrLf & vbCrLf & _
              "Examinadores: " & idExaminadorPorDni.Count & vbCrLf & _
              "Tablets: " & idTabletPorSerie.Count & vbCrLf & _
              "Tarjetas SIM: " & idSimPorNumero.Count & vbCrLf & _
              "Asignaciones examinador - tablet: " & DCount("*", "AsignacionesTablet") & vbCrLf & _
              "Asignaciones tablet - SIM: " & DCount("*", "AsignacionesSIM") & vbCrLf & vbCrLf
    If numIncidencias = 0 Then
        resumen = resumen & "Sin incidencias."
    Else
        resumen = resumen & numIncidencias & " incidencias: revísalas en la tabla ImportacionIncidencias."
    End If
    ImportarDatosAntiguos = resumen
    Exit Function

Fallo:
    ImportarDatosAntiguos = "ERROR: la importación se ha detenido." & vbCrLf & Err.Description
End Function

' ---------------------------------------------------------------------------------------------
' Preparación
' ---------------------------------------------------------------------------------------------

Private Sub PrepararIncidencias()
    On Error Resume Next
    baseNueva.Execute "DROP TABLE ImportacionIncidencias"
    On Error GoTo 0
    baseNueva.Execute "CREATE TABLE ImportacionIncidencias (Id COUNTER CONSTRAINT PK PRIMARY KEY, " & _
                      "Tabla TEXT(60), Registro TEXT(255), Problema MEMO)", DB_FAIL_ON_ERROR
    Set rsIncidencias = baseNueva.OpenRecordset("ImportacionIncidencias", DB_OPEN_DYNASET)
    numIncidencias = 0
End Sub

Private Sub VaciarTablasNuevas()
    Dim tabla As Variant
    On Error Resume Next            ' no existe en bases creadas antes de añadir el historial de estados
    baseNueva.Execute "DELETE FROM HistorialEstadosTablet", DB_FAIL_ON_ERROR
    On Error GoTo 0
    For Each tabla In Array("AsignacionesSIM", "AsignacionesTablet", "TarjetasSIM", "Tablets", "Examinadores")
        baseNueva.Execute "DELETE FROM " & tabla, DB_FAIL_ON_ERROR
    Next
End Sub

Private Sub ImportarEstados()
    Dim r As Object
    Set r = baseAntigua.OpenRecordset("TSituacionExaminador", DB_OPEN_SNAPSHOT)
    Do Until r.EOF
        If Len(Texto(r!Situacion)) > 0 Then IdEstado idSituacion, "EstadosExaminador", "EstadoExaminador", "IdEstadoExam", Texto(r!Situacion), ""
        r.MoveNext
    Loop
    r.Close
    Set r = baseAntigua.OpenRecordset("T-EstadoTablet", DB_OPEN_SNAPSHOT)
    Do Until r.EOF
        If Len(Texto(r!Descripcion)) > 0 Then IdEstado idEstadoTablet, "EstadosTablet", "Estado", "IdEstado", Texto(r!Descripcion), ""
        r.MoveNext
    Loop
    r.Close
    Set r = baseAntigua.OpenRecordset("T-EstadoSIM", DB_OPEN_SNAPSHOT)
    Do Until r.EOF
        If Len(Texto(r!EstadoSIM)) > 0 Then IdEstado idEstadoSim, "EstadosSIM", "EstadoSIM", "IdEstadoSIM", Texto(r!EstadoSIM), ""
        r.MoveNext
    Loop
    r.Close
End Sub

' ---------------------------------------------------------------------------------------------
' Examinadores y tablets
' ---------------------------------------------------------------------------------------------

Private Sub ImportarExaminadores()
    Dim r As Object, w As Object, numerosUsados As Object
    Set numerosUsados = NuevoDiccionario()
    Set r = baseAntigua.OpenRecordset("T-Examinadores", DB_OPEN_SNAPSHOT)
    Set w = baseNueva.OpenRecordset("Examinadores", DB_OPEN_DYNASET)
    Do Until r.EOF
        ImportarExaminador r, w, numerosUsados
        r.MoveNext
    Loop
    r.Close
    w.Close
End Sub

Private Sub ImportarExaminador(ByVal r As Object, ByVal w As Object, ByVal numerosUsados As Object)
    Dim dni As String, numero As String, registro As String
    On Error GoTo Fallo

    dni = NormalizarDni(r!DNI)
    registro = "DNI " & Texto(r!DNI) & " - " & Texto(r!Nombre) & " " & Texto(r!Apellidos)
    If Len(dni) = 0 Then Anotar "Examinadores", registro, "Sin DNI: no se importa.": Exit Sub
    If Len(dni) > 9 Then Anotar "Examinadores", registro, "El DNI tiene más de 9 caracteres: no se importa. Corrígelo en la base antigua.": Exit Sub
    If idExaminadorPorDni.Exists(dni) Then Anotar "Examinadores", registro, "DNI repetido: solo se importa el primero.": Exit Sub

    numero = Texto(r!NumExaminador)
    If Len(numero) > 0 And numerosUsados.Exists(numero) Then
        Anotar "Examinadores", registro, "Nº de examinador " & numero & " repetido: se deja vacío."
        numero = ""
    End If

    w.AddNew
    w!DNI = dni
    w!NumeroExaminador = Nulo(Recortar(numero, 20, "Examinadores", registro, "Nº de examinador"))
    w!Nombre = Obligatorio(r!Nombre, 50, "Examinadores", registro, "Nombre")
    w!Apellidos = Obligatorio(r!Apellidos, 100, "Examinadores", registro, "Apellidos")
    w!IdEstadoExam = IdEstado(idSituacion, "EstadosExaminador", "EstadoExaminador", "IdEstadoExam", _
                              Texto(r!SituacionExaminador), "Activo")
    w!Observaciones = Nulo(Texto(r!Observaciones))
    w.Update
    w.Bookmark = w.LastModified
    idExaminadorPorDni.Add dni, w!IdExaminador.Value
    If Len(numero) > 0 Then numerosUsados.Add numero, True
    Exit Sub

Fallo:
    Anotar "Examinadores", registro, "No se ha podido importar: " & Err.Description
    CancelarAlta w
End Sub

Private Sub ImportarTablets()
    Dim r As Object, w As Object, imei1Usados As Object, imei2Usados As Object
    Set imei1Usados = NuevoDiccionario()
    Set imei2Usados = NuevoDiccionario()
    Set r = baseAntigua.OpenRecordset("T-Tablets", DB_OPEN_SNAPSHOT)
    Set w = baseNueva.OpenRecordset("Tablets", DB_OPEN_DYNASET)
    Do Until r.EOF
        ImportarTablet r, w, imei1Usados, imei2Usados
        r.MoveNext
    Loop
    r.Close
    w.Close
End Sub

Private Sub ImportarTablet(ByVal r As Object, ByVal w As Object, ByVal imei1Usados As Object, ByVal imei2Usados As Object)
    Dim serie As String, registro As String, imei1 As String, imei2 As String, extra As String
    Dim recepcion As Variant, desafectacion As Variant
    On Error GoTo Fallo

    serie = Texto(r!NumSerie)
    registro = "Nº serie " & serie
    If Len(serie) = 0 Then Anotar "Tablets", registro, "Sin nº de serie: no se importa.": Exit Sub
    If idTabletPorSerie.Exists(serie) Then Anotar "Tablets", registro, "Nº de serie repetido: solo se importa el primero.": Exit Sub

    imei1 = ImeiUnico(r.Fields("IMEI-1").Value, imei1Usados, "IMEI 1", registro, extra)
    imei2 = ImeiUnico(r.Fields("IMEI-2").Value, imei2Usados, "IMEI 2", registro, extra)
    recepcion = FechaImportada(r!FechaRecepcion, "Tablets", registro, "Fecha de recepción", extra)
    desafectacion = FechaImportada(r!FechaDesafectacion, "Tablets", registro, "Fecha de desafectación", extra)

    w.AddNew
    w!NumeroSerie = Recortar(serie, 50, "Tablets", registro, "Nº de serie")
    w!IMEI1 = Nulo(imei1)
    w!IMEI2 = Nulo(imei2)
    w!Marca = Nulo(Recortar(Texto(r!Marca), 50, "Tablets", registro, "Marca"))
    w!Modelo = Nulo(Recortar(Texto(r!Modelo), 50, "Tablets", registro, "Modelo"))
    w!FechaRecepcion = recepcion
    w!FechaDesafectacion = desafectacion
    w!IdEstado = IdEstado(idEstadoTablet, "EstadosTablet", "Estado", "IdEstado", Texto(r!Estado), "Operativa")
    w!Observaciones = Nulo(Combinar(Texto(r!Observaciones), extra))
    w!Incidencias = Nulo(Texto(r!INCIDENCIAS))
    w.Update
    w.Bookmark = w.LastModified
    idTabletPorSerie.Add serie, w!IdTablet.Value
    If Len(imei1) > 0 Then imei1Usados.Add imei1, True
    If Len(imei2) > 0 Then imei2Usados.Add imei2, True
    Exit Sub

Fallo:
    Anotar "Tablets", registro, "No se ha podido importar: " & Err.Description
    CancelarAlta w
End Sub

' Un IMEI repetido no se guarda en su campo (es único) sino en las observaciones de la tablet.
Private Function ImeiUnico(ByVal valor As Variant, ByVal usados As Object, ByVal campo As String, _
                           ByVal registro As String, ByRef extra As String) As String
    Dim imei As String
    imei = Texto(valor)
    If Len(imei) > 20 Then
        Anotar "Tablets", registro, campo & " de más de 20 caracteres (" & imei & "): se pasa a observaciones."
        extra = Combinar(extra, campo & " original: " & imei)
        imei = ""
    ElseIf Len(imei) > 0 And usados.Exists(imei) Then
        Anotar "Tablets", registro, campo & " " & imei & " repetido en otra tablet: se pasa a observaciones."
        extra = Combinar(extra, campo & " (repetido): " & imei)
        imei = ""
    End If
    ImeiUnico = imei
End Function

' ---------------------------------------------------------------------------------------------
' Asignaciones examinador - tablet (dos tablas en la base antigua)
' ---------------------------------------------------------------------------------------------

Private Sub ImportarAsignacionesTablet()
    Dim r As Object, w As Object, vistas As Object
    Set vistas = NuevoDiccionario()
    Set w = baseNueva.OpenRecordset("AsignacionesTablet", DB_OPEN_DYNASET)

    Set r = baseAntigua.OpenRecordset("T-HistorialAsigTabletExaminador", DB_OPEN_SNAPSHOT)
    Do Until r.EOF
        ImportarAsignacionTablet w, vistas, "T-HistorialAsigTabletExaminador", r!DNIExaminador, r!NumSerieTablet, _
            r!FechaAltaAsignacion, r!FechaBajaAsignacion, r!EstadoAsignacion, r.Fields("Observaciones Asignación").Value
        r.MoveNext
    Loop
    r.Close

    Set r = baseAntigua.OpenRecordset("T-HistoricoTablet", DB_OPEN_SNAPSHOT)
    Do Until r.EOF
        ImportarAsignacionTablet w, vistas, "T-HistoricoTablet", r!DNI, r!NumSerieTablet, _
            r!FechaAltaHistoricoTablet, r!FechaBajaHistoricoTablet, r!EstadoHistorial, r!ObservacionesHistorico
        r.MoveNext
    Loop
    r.Close
    w.Close
End Sub

Private Sub ImportarAsignacionTablet(ByVal w As Object, ByVal vistas As Object, ByVal tabla As String, _
                                     ByVal dniOrigen As Variant, ByVal serieOrigen As Variant, ByVal alta As Variant, _
                                     ByVal baja As Variant, ByVal estado As Variant, ByVal observaciones As Variant)
    Dim registro As String, dni As String, serie As String, extra As String, clave As String
    Dim inicio As Variant, fin As Variant
    On Error GoTo Fallo

    dni = NormalizarDni(dniOrigen)
    serie = Texto(serieOrigen)
    registro = "Tablet " & serie & " / DNI " & Texto(dniOrigen) & " / alta " & Texto(alta) & " / baja " & Texto(baja)
    If Not idTabletPorSerie.Exists(serie) Then Anotar tabla, registro, "La tablet no existe en la base nueva: no se importa.": Exit Sub
    If Not idExaminadorPorDni.Exists(dni) Then Anotar tabla, registro, "El examinador (DNI) no existe en la base nueva: no se importa.": Exit Sub

    inicio = FechaImportada(alta, tabla, registro, "Fecha de alta", extra)
    If IsNull(inicio) Then Anotar tabla, registro, "Sin fecha de alta válida: no se importa. Corrígela en la base antigua.": Exit Sub
    fin = FechaImportada(baja, tabla, registro, "Fecha de baja", extra)
    If Len(extra) > 0 And IsNull(fin) Then Anotar tabla, registro, "Al no tener fecha de baja válida queda como asignación vigente."
    If Not IsNull(fin) Then
        If fin < inicio Then Anotar tabla, registro, "La fecha de baja es anterior a la de alta (se importa igualmente)."
    End If

    clave = idTabletPorSerie(serie) & "|" & idExaminadorPorDni(dni) & "|" & CLng(inicio)
    If vistas.Exists(clave) Then Exit Sub        ' la misma asignación ya estaba en la otra tabla
    vistas.Add clave, True

    If Len(Texto(estado)) > 0 Then extra = Combinar("Estado anterior: " & Texto(estado), extra)
    w.AddNew
    w!IdTablet = idTabletPorSerie(serie)
    w!IdExaminador = idExaminadorPorDni(dni)
    w!FechaInicio = inicio
    w!FechaFin = fin
    w!Observaciones = Nulo(Recortar(Combinar(Texto(observaciones), extra), 255, tabla, registro, "Observaciones"))
    w.Update
    Exit Sub

Fallo:
    Anotar tabla, registro, "No se ha podido importar: " & Err.Description
    CancelarAlta w
End Sub

' ---------------------------------------------------------------------------------------------
' Tarjetas SIM y su historial con las tablets
' ---------------------------------------------------------------------------------------------

Private Sub ImportarSIM()
    Dim r As Object, sims As Object, asignaciones As Object, vistas As Object, fechaEstado As Object
    Set vistas = NuevoDiccionario()
    Set fechaEstado = NuevoDiccionario()
    Set sims = baseNueva.OpenRecordset("TarjetasSIM", DB_OPEN_DYNASET)
    Set asignaciones = baseNueva.OpenRecordset("AsignacionesSIM", DB_OPEN_DYNASET)
    Set r = baseAntigua.OpenRecordset("T-HistoriaTabletSim", DB_OPEN_SNAPSHOT)
    Do Until r.EOF
        ImportarFilaSIM r, sims, asignaciones, vistas, fechaEstado
        r.MoveNext
    Loop
    r.Close
    sims.Close
    asignaciones.Close
End Sub

' En la base antigua cada fila es una activación de una SIM en una tablet. Se crea una tarjeta por
' número de SIM: PIN, PUK y teléfono se toman de la primera fila que los tenga, y el estado de la
' activación más reciente. fechaEstado guarda, por SIM, la fecha de la fila de la que salió su estado.
Private Sub ImportarFilaSIM(ByVal r As Object, ByVal sims As Object, ByVal asignaciones As Object, _
                            ByVal vistas As Object, ByVal fechaEstado As Object)
    Dim numero As String, serie As String, registro As String, extra As String, clave As String
    Dim telefono As String, estado As String, idSIM As Long, inicio As Variant, fin As Variant, orden As Double
    On Error GoTo Fallo

    numero = Texto(r!SIM)
    serie = Texto(r!NumSerieTablet)
    telefono = Texto(r.Fields("telefono asociado").Value)
    estado = Texto(r!EstadoSIM)
    registro = "SIM " & numero & " / tablet " & serie & " / activación " & Texto(r!FechaActivacion)
    If Len(numero) = 0 Then Anotar "T-HistoriaTabletSim", registro, "Sin número de SIM: no se importa.": Exit Sub
    If Len(numero) > 22 Then Anotar "T-HistoriaTabletSim", registro, "El número de SIM tiene más de 22 caracteres: no se importa.": Exit Sub

    inicio = FechaImportada(r!FechaActivacion, "T-HistoriaTabletSim", registro, "Fecha de activación", extra)
    If IsNull(inicio) Then orden = -1 Else orden = CDbl(inicio)

    If idSimPorNumero.Exists(numero) Then
        idSIM = idSimPorNumero(numero)
        sims.FindFirst "IdSIM = " & idSIM
        sims.Edit
        If IsNull(sims!PIN) Then sims!PIN = Nulo(Recortar(Texto(r!Pin), 8, "TarjetasSIM", registro, "PIN"))
        If IsNull(sims!PUK) Then sims!PUK = Nulo(Recortar(Texto(r!PUK), 10, "TarjetasSIM", registro, "PUK"))
        If Len(telefono) > 0 Then
            If IsNull(sims!NumeroTelefono) Then
                sims!NumeroTelefono = Recortar(telefono, 20, "TarjetasSIM", registro, "Teléfono")
            ElseIf sims!NumeroTelefono <> telefono Then
                Anotar "T-HistoriaTabletSim", registro, "La SIM tenía otro teléfono (" & sims!NumeroTelefono & _
                       "); este (" & telefono & ") se guarda en las observaciones de la asociación."
                extra = Combinar(extra, "Teléfono: " & telefono)
            End If
        End If
        If Len(estado) > 0 And orden >= fechaEstado(numero) Then
            sims!IdEstadoSIM = IdEstado(idEstadoSim, "EstadosSIM", "EstadoSIM", "IdEstadoSIM", estado, "Activa")
            fechaEstado(numero) = orden
        End If
        sims.Update
    Else
        sims.AddNew
        sims!IMEI = numero
        sims!NumeroTelefono = Nulo(Recortar(telefono, 20, "TarjetasSIM", registro, "Teléfono"))
        sims!PIN = Nulo(Recortar(Texto(r!Pin), 8, "TarjetasSIM", registro, "PIN"))
        sims!PUK = Nulo(Recortar(Texto(r!PUK), 10, "TarjetasSIM", registro, "PUK"))
        sims!IdEstadoSIM = IdEstado(idEstadoSim, "EstadosSIM", "EstadoSIM", "IdEstadoSIM", estado, "Activa")
        sims!IdCompania = idCompaniaNA
        sims.Update
        sims.Bookmark = sims.LastModified
        idSIM = sims!IdSIM
        idSimPorNumero.Add numero, idSIM
        If Len(estado) > 0 Then fechaEstado.Add numero, orden Else fechaEstado.Add numero, -2
    End If

    ' Asociación con la tablet
    If Len(serie) = 0 Then Anotar "T-HistoriaTabletSim", registro, "Sin tablet: se importa la SIM pero no la asociación.": Exit Sub
    If Not idTabletPorSerie.Exists(serie) Then Anotar "T-HistoriaTabletSim", registro, "La tablet no existe en la base nueva: se importa la SIM pero no la asociación.": Exit Sub
    If IsNull(inicio) Then Anotar "T-HistoriaTabletSim", registro, "Sin fecha de activación válida: se importa la SIM pero no la asociación.": Exit Sub
    fin = FechaImportada(r!FechaFinActivacion, "T-HistoriaTabletSim", registro, "Fecha fin de activación", extra)

    clave = idTabletPorSerie(serie) & "|" & idSIM & "|" & CLng(inicio)
    If vistas.Exists(clave) Then Exit Sub
    vistas.Add clave, True

    ' El estado de cada activación se conserva en sus observaciones (la SIM solo guarda el último).
    If Len(estado) > 0 Then extra = Combinar("Estado SIM: " & estado, extra)
    asignaciones.AddNew
    asignaciones!IdTablet = idTabletPorSerie(serie)
    asignaciones!IdSIM = idSIM
    asignaciones!FechaInicio = inicio
    asignaciones!FechaFin = fin
    asignaciones!Observaciones = Nulo(Recortar(Combinar(Texto(r!Observaciones), extra), 255, "T-HistoriaTabletSim", registro, "Observaciones"))
    asignaciones.Update
    Exit Sub

Fallo:
    Anotar "T-HistoriaTabletSim", registro, "No se ha podido importar: " & Err.Description
    CancelarAlta sims
    CancelarAlta asignaciones
End Sub

' En la aplicación cada tablet tiene como mucho un examinador y una SIM vigentes (y al revés).
' Los datos antiguos pueden no cumplirlo: se avisa para cerrarlas a mano poniendo la fecha de fin.
Private Sub ComprobarAsignacionesVigentes()
    AvisarVigentes "SELECT T.NumeroSerie AS Clave, Count(*) AS N FROM AsignacionesTablet AS A INNER JOIN Tablets AS T " & _
                   "ON A.IdTablet = T.IdTablet WHERE A.FechaFin Is Null GROUP BY T.NumeroSerie HAVING Count(*) > 1", _
                   "AsignacionesTablet", "La tablet %1 tiene %2 examinadores a la vez (asignaciones sin fecha de baja)."
    AvisarVigentes "SELECT E.DNI AS Clave, Count(*) AS N FROM AsignacionesTablet AS A INNER JOIN Examinadores AS E " & _
                   "ON A.IdExaminador = E.IdExaminador WHERE A.FechaFin Is Null GROUP BY E.DNI HAVING Count(*) > 1", _
                   "AsignacionesTablet", "El examinador con DNI %1 tiene %2 tablets a la vez (asignaciones sin fecha de baja)."
    AvisarVigentes "SELECT T.NumeroSerie AS Clave, Count(*) AS N FROM AsignacionesSIM AS A INNER JOIN Tablets AS T " & _
                   "ON A.IdTablet = T.IdTablet WHERE A.FechaFin Is Null GROUP BY T.NumeroSerie HAVING Count(*) > 1", _
                   "AsignacionesSIM", "La tablet %1 tiene %2 SIM a la vez (sin fecha fin de activación)."
    AvisarVigentes "SELECT M.IMEI AS Clave, Count(*) AS N FROM AsignacionesSIM AS A INNER JOIN TarjetasSIM AS M " & _
                   "ON A.IdSIM = M.IdSIM WHERE A.FechaFin Is Null GROUP BY M.IMEI HAVING Count(*) > 1", _
                   "AsignacionesSIM", "La SIM %1 está en %2 tablets a la vez (sin fecha fin de activación)."
End Sub

Private Sub AvisarVigentes(ByVal sql As String, ByVal tabla As String, ByVal mensaje As String)
    Dim r As Object
    Set r = baseNueva.OpenRecordset(sql, DB_OPEN_SNAPSHOT)
    Do Until r.EOF
        Anotar tabla, Texto(r!Clave), Replace(Replace(mensaje, "%1", Texto(r!Clave)), "%2", r!N)
        r.MoveNext
    Loop
    r.Close
End Sub

' ---------------------------------------------------------------------------------------------
' Utilidades
' ---------------------------------------------------------------------------------------------

Private Sub Anotar(ByVal tabla As String, ByVal registro As String, ByVal problema As String)
    rsIncidencias.AddNew
    rsIncidencias!Tabla = Left$(tabla, 60)
    rsIncidencias!Registro = Left$(registro, 255)
    rsIncidencias!Problema = problema
    rsIncidencias.Update
    numIncidencias = numIncidencias + 1
End Sub

Private Sub CancelarAlta(ByVal rs As Object)
    On Error Resume Next
    If rs.EditMode <> 0 Then rs.CancelUpdate
End Sub

Private Function NuevoDiccionario() As Object
    Set NuevoDiccionario = CreateObject("Scripting.Dictionary")
    NuevoDiccionario.CompareMode = vbTextCompare
End Function

Private Function CargarDiccionario(ByVal tabla As String, ByVal campoTexto As String, ByVal campoId As String) As Object
    Dim d As Object, r As Object
    Set d = NuevoDiccionario()
    Set r = baseNueva.OpenRecordset(tabla, DB_OPEN_SNAPSHOT)
    Do Until r.EOF
        If Not d.Exists(Texto(r(campoTexto))) Then d.Add Texto(r(campoTexto)), r(campoId).Value
        r.MoveNext
    Loop
    r.Close
    Set CargarDiccionario = d
End Function

' Id del estado con ese texto; si no existe en la tabla de estados nueva, lo añade.
Private Function IdEstado(ByVal d As Object, ByVal tabla As String, ByVal campoTexto As String, ByVal campoId As String, _
                          ByVal valor As String, ByVal porDefecto As String) As Variant
    Dim w As Object
    If Len(valor) = 0 Then valor = porDefecto
    valor = Left$(valor, 50)
    If Not d.Exists(valor) Then
        Set w = baseNueva.OpenRecordset(tabla, DB_OPEN_DYNASET)
        w.AddNew
        w(campoTexto) = valor
        w.Update
        w.Bookmark = w.LastModified
        d.Add valor, w(campoId).Value
        w.Close
    End If
    IdEstado = d(valor)
End Function

Private Function Texto(ByVal v As Variant) As String
    If Not IsNull(v) Then Texto = Trim$(CStr(v))
End Function

Private Function Nulo(ByVal s As String) As Variant
    If Len(s) = 0 Then Nulo = Null Else Nulo = s
End Function

Private Function NormalizarDni(ByVal v As Variant) As String
    NormalizarDni = UCase$(Replace(Replace(Replace(Texto(v), " ", ""), "-", ""), ".", ""))
End Function

Private Function Obligatorio(ByVal v As Variant, ByVal maximo As Long, ByVal tabla As String, _
                             ByVal registro As String, ByVal campo As String) As String
    Dim s As String
    s = Texto(v)
    If Len(s) = 0 Then
        Anotar tabla, registro, campo & " vacío: se pone '" & SIN_DATO & "'."
        s = SIN_DATO
    End If
    Obligatorio = Recortar(s, maximo, tabla, registro, campo)
End Function

Private Function Recortar(ByVal s As String, ByVal maximo As Long, ByVal tabla As String, _
                          ByVal registro As String, ByVal campo As String) As String
    If Len(s) > maximo Then
        Anotar tabla, registro, campo & " recortado a " & maximo & " caracteres. Texto completo: " & s
        s = Left$(s, maximo)
    End If
    Recortar = s
End Function

' Une las partes no vacías con " | ".
Private Function Combinar(ParamArray partes() As Variant) As String
    Dim i As Long, s As String
    For i = 0 To UBound(partes)
        If Len(partes(i)) > 0 Then
            If Len(s) > 0 Then s = s & " | "
            s = s & partes(i)
        End If
    Next
    Combinar = s
End Function

' En la base antigua muchas fechas son texto. Acepta fechas, textos como 15/03/2023, 15-03-2023 o
' 15.03.2023 y números de fecha de Excel. Si no se entiende, devuelve Nulo, lo anota y guarda el
' texto original en extra para no perderlo.
Private Function FechaImportada(ByVal v As Variant, ByVal tabla As String, ByVal registro As String, _
                                ByVal campo As String, ByRef extra As String) As Variant
    Dim s As String
    FechaImportada = Null
    If IsNull(v) Then Exit Function
    If VarType(v) = vbDate Then FechaImportada = DateValue(v): Exit Function
    s = Replace(Trim$(CStr(v)), ".", "/")
    If Len(s) = 0 Then Exit Function

    If IsNumeric(s) Then
        If Val(s) > 20000 And Val(s) < 80000 Then FechaImportada = CDate(CLng(Val(s))): Exit Function
    ElseIf IsDate(s) Then
        FechaImportada = DateValue(CDate(s))
        Exit Function
    End If
    Anotar tabla, registro, campo & " no reconocida: '" & Trim$(CStr(v)) & "'. Se guarda en observaciones."
    extra = Combinar(extra, campo & " original: " & Trim$(CStr(v)))
End Function
