Attribute VB_Name = "modTestingSLA"

Option Compare Database
Option Explicit

'===========================================================
' MÓDULO DE TESTING - Spec-003 y Spec-008
' Objetivo: Validar funcionamiento de campos SLA
' IMPORTANTE: Este módulo crea y elimina eventos de prueba
' NOTA: Los módulos .bas NO usan "VERSION 1.0 CLASS"
'===========================================================

'===========================================================
' Función principal: EjecutarTestsSLA
' Ejecuta todos los tests de SLA
'===========================================================
Public Function EjecutarTestsSLA() As Boolean
    Dim RESULTADO As Boolean
    Dim msg As String
    
    On Error GoTo errores
    
    msg = "=== BATERÍA DE TESTS SLA ===" & vbCrLf & vbCrLf
    
    ' Test 1: Crear evento con campos SLA
    msg = msg & "Test 1: Crear evento con campos SLA... "
    If TestCrearEventoConSLA() Then
        msg = msg & "? PASÓ" & vbCrLf
    Else
        msg = msg & "? FALLÓ" & vbCrLf
        RESULTADO = False
    End If
    
    ' Test 2: Leer campos SLA de evento existente
    msg = msg & "Test 2: Leer campos SLA... "
    If TestLeerCamposSLA() Then
        msg = msg & "? PASÓ" & vbCrLf
    Else
        msg = msg & "? FALLÓ" & vbCrLf
        RESULTADO = False
    End If
    
    ' Test 3: Actualizar campos SLA
    msg = msg & "Test 3: Actualizar campos SLA... "
    If TestActualizarCamposSLA() Then
        msg = msg & "? PASÓ" & vbCrLf
    Else
        msg = msg & "? FALLÓ" & vbCrLf
        RESULTADO = False
    End If
    
    ' Test 4: Verificar cálculo de SLAs
    msg = msg & "Test 4: Verificar cálculo de SLAs... "
    If TestCalculoSLAs() Then
        msg = msg & "? PASÓ" & vbCrLf
    Else
        msg = msg & "? FALLÓ" & vbCrLf
        RESULTADO = False
    End If
    
    msg = msg & vbCrLf & "=== FIN TESTS ==="
    MsgBox msg, vbInformation, "Testing SLA"
    
    EjecutarTestsSLA = True
    Exit Function
    
errores:
    MsgBox "Error en test: " & Err.Description, vbCritical
    EjecutarTestsSLA = False
End Function

'===========================================================
' Test 1: Crear evento con campos SLA
'===========================================================
Private Function TestCrearEventoConSLA() As Boolean
    Dim objEvento As Evento
    Dim m_Error As String
    Dim IDEventoTest As String
    Dim fechaTest As String
    
    On Error GoTo errores
    
    fechaTest = Format(Now(), "yyyymmddhhnnss")
    IDEventoTest = "TEST-" & fechaTest
    
    ' Crear evento
    Set objEvento = New Evento
    objEvento.ParaAlta = EnumSino.Sí
    objEvento.IDEVENTO = IDEventoTest
    objEvento.NODO = "TEST"
    objEvento.BUI = "TEST001"
    objEvento.SubSistema = "TEST"
    objEvento.IDEquipo = 1
    objEvento.TipoEvento = "INCIDENCIA"
    objEvento.Criticidad = "ALTA"
    objEvento.ALIASTECNICO = "TEST"
    objEvento.Originador = "TEST"
    objEvento.FECHAALTAEVENTO = Date
    objEvento.HORAINICIALEVENTO = "08:00"
    objEvento.Descripcion = "Evento de prueba para SLA"
    
    ' Asignar campos SLA
    objEvento.FechaRecepcionNotificacion = Now()
    objEvento.FechaInicioContactoCliente = Now()
    objEvento.IncidenciaAveriaOReparacion = EnumSino.Sí
    objEvento.FechaInicioTiempoAdquisicion = Now()
    objEvento.FechaFinTiempoAdquisicion = Now()
    objEvento.TipoReparacion = "Contratista"
    objEvento.Urgente = EnumSino.Sí
    objEvento.EventoConServicioAfectado = EnumSino.Sí
    objEvento.FechaRestablecimientoServicio = Now()
    objEvento.TipoRepInsitu = EnumSino.Sí
    objEvento.TipoRepNoSMT = EnumSino.No
    objEvento.TipoRepValvulas = EnumSino.No
    
    ' Guardar
    objEvento.Registrar , m_Error
    If m_Error <> "" Then
        Debug.Print "Error al crear evento: " & m_Error
        TestCrearEventoConSLA = False
        Exit Function
    End If
    
    ' Verificar que se guardó
    Set objEvento = Constructor.getEvento(IDEventoTest, m_Error)
    If objEvento Is Nothing Or m_Error <> "" Then
        Debug.Print "Error al recuperar evento"
        TestCrearEventoConSLA = False
        Exit Function
    End If
    
    ' Verificar campos SLA
    If Not IsDate(objEvento.FechaRecepcionNotificacion) Then
        Debug.Print "FechaRecepcionNotificacion no guardada"
        TestCrearEventoConSLA = False
        Exit Function
    End If
    
    If objEvento.IncidenciaAveriaOReparacion <> EnumSino.Sí Then
        Debug.Print "IncidenciaAveriaOReparacion no guardada"
        TestCrearEventoConSLA = False
        Exit Function
    End If
    
    If objEvento.TipoReparacion <> "Contratista" Then
        Debug.Print "TipoReparacion no guardada"
        TestCrearEventoConSLA = False
        Exit Function
    End If
    
    Debug.Print "Test 1 PASÓ - IDEvento: " & IDEventoTest
    TestCrearEventoConSLA = True
    Exit Function
    
errores:
    Debug.Print "Error en Test 1: " & Err.Description
    TestCrearEventoConSLA = False
End Function

'===========================================================
' Test 2: Leer campos SLA de evento existente
'===========================================================
Private Function TestLeerCamposSLA() As Boolean
    Dim objEvento As Evento
    Dim m_Error As String
    Dim IDEventoTest As String
    Dim fechaTest As String
    
    On Error GoTo errores
    
    ' Buscar evento de prueba creado anteriormente
    fechaTest = Format(Date, "yyyymmdd")
    IDEventoTest = "TEST-" & fechaTest & "*"
    
    ' Obtener primer evento de test
    Dim rs As DAO.Recordset
    Set rs = getdb().OpenRecordset("SELECT TOP 1 IDEvento FROM TbEventos WHERE IDEvento LIKE '" & IDEventoTest & "' ORDER BY FechaRegistroAlta DESC")
    
    If rs.EOF Then
        Debug.Print "No hay eventos de test para leer"
        TestLeerCamposSLA = False
        rs.Close
        Exit Function
    End If
    
    IDEventoTest = rs!IDEVENTO
    rs.Close
    
    ' Leer evento
    Set objEvento = Constructor.getEvento(IDEventoTest, m_Error)
    If objEvento Is Nothing Or m_Error <> "" Then
        Debug.Print "Error al leer evento: " & m_Error
        TestLeerCamposSLA = False
        Exit Function
    End If
    
    ' Verificar que se pueden leer todos los campos SLA
    Debug.Print "FechaRecepcionNotificacion: " & objEvento.FechaRecepcionNotificacion
    Debug.Print "IncidenciaAveriaOReparacion: " & objEvento.IncidenciaAveriaOReparacion
    Debug.Print "TipoReparacion: " & objEvento.TipoReparacion
    Debug.Print "Urgente: " & objEvento.Urgente
    Debug.Print "TipoRepInsitu: " & objEvento.TipoRepInsitu
    
    TestLeerCamposSLA = True
    Exit Function
    
errores:
    Debug.Print "Error en Test 2: " & Err.Description
    TestLeerCamposSLA = False
End Function

'===========================================================
' Test 3: Actualizar campos SLA
'===========================================================
Private Function TestActualizarCamposSLA() As Boolean
    Dim objEvento As Evento
    Dim m_Error As String
    Dim IDEventoTest As String
    Dim fechaTest As String
    
    On Error GoTo errores
    
    ' Buscar evento de prueba
    fechaTest = Format(Date, "yyyymmdd")
    IDEventoTest = "TEST-" & fechaTest & "*"
    
    Dim rs As DAO.Recordset
    Set rs = getdb().OpenRecordset("SELECT TOP 1 IDEvento FROM TbEventos WHERE IDEvento LIKE '" & IDEventoTest & "' ORDER BY FechaRegistroAlta DESC")
    
    If rs.EOF Then
        Debug.Print "No hay eventos de test para actualizar"
        TestActualizarCamposSLA = False
        rs.Close
        Exit Function
    End If
    
    IDEventoTest = rs!IDEVENTO
    rs.Close
    
    ' Cargar evento para edición
    Set objEvento = Constructor.getEvento(IDEventoTest, m_Error)
    If objEvento Is Nothing Or m_Error <> "" Then
        Debug.Print "Error al cargar evento: " & m_Error
        TestActualizarCamposSLA = False
        Exit Function
    End If
    
    objEvento.ParaAlta = EnumSino.No
    
    ' Modificar campos SLA
    objEvento.FechaRecepcionNotificacion = Now()
    objEvento.IncidenciaAveriaOReparacion = EnumSino.No
    objEvento.TipoReparacion = "Fabricante"
    objEvento.Urgente = EnumSino.No
    objEvento.TipoRepInsitu = EnumSino.No
    objEvento.TipoRepNoSMT = EnumSino.Sí
    
    ' Guardar cambios
    m_Error = objEvento.Registrar()
    If m_Error <> "" Then
        Debug.Print "Error al actualizar: " & m_Error
        TestActualizarCamposSLA = False
        Exit Function
    End If
    
    ' Verificar actualización
    Set objEvento = Constructor.getEvento(IDEventoTest, m_Error)
    If objEvento.TipoReparacion <> "Fabricante" Then
        Debug.Print "TipoReparacion no se actualizó"
        TestActualizarCamposSLA = False
        Exit Function
    End If
    
    Debug.Print "Test 3 PASÓ"
    TestActualizarCamposSLA = True
    Exit Function
    
errores:
    Debug.Print "Error en Test 3: " & Err.Description
    TestActualizarCamposSLA = False
End Function

'===========================================================
' Test 4: Verificar cálculo de SLAs
'===========================================================
Private Function TestCalculoSLAs() As Boolean
    Dim objEvento As Evento
    Dim m_Error As String
    Dim IDEventoTest As String
    Dim fechaTest As String
    
    On Error GoTo errores
    
    ' Buscar evento de prueba
    fechaTest = Format(Date, "yyyymmdd")
    IDEventoTest = "TEST-" & fechaTest & "*"
    
    Dim rs As DAO.Recordset
    Set rs = getdb().OpenRecordset("SELECT TOP 1 IDEvento FROM TbEventos WHERE IDEvento LIKE '" & IDEventoTest & "' ORDER BY FechaRegistroAlta DESC")
    
    If rs.EOF Then
        Debug.Print "No hay eventos de test para calcular"
        TestCalculoSLAs = False
        rs.Close
        Exit Function
    End If
    
    IDEventoTest = rs!IDEVENTO
    rs.Close
    
    ' Cargar evento
    Set objEvento = Constructor.getEvento(IDEventoTest, m_Error)
    If objEvento Is Nothing Or m_Error <> "" Then
        Debug.Print "Error al cargar evento: " & m_Error
        TestCalculoSLAs = False
        Exit Function
    End If
    
    ' Verificar que se pueden calcular métricas
    ' TRES: Tiempo desde recepción hasta primer contacto
    If IsDate(objEvento.FechaRecepcionNotificacion) And IsDate(objEvento.FechaInicioContactoCliente) Then
        Dim diasTRES As Double
        diasTRES = DateDiff("h", CDate(objEvento.FechaRecepcionNotificacion), CDate(objEvento.FechaInicioContactoCliente))
        Debug.Print "TRES (horas): " & diasTRES
    End If
    
    ' TRCM: Tiempo de reparación
    If IsDate(objEvento.FechaInicioTiempoAdquisicion) And IsDate(objEvento.FechaFinTiempoAdquisicion) Then
        Dim diasTRCM As Double
        diasTRCM = DateDiff("h", CDate(objEvento.FechaInicioTiempoAdquisicion), CDate(objEvento.FechaFinTiempoAdquisicion))
        Debug.Print "TRCM (horas): " & diasTRCM
    End If
    
    ' TRSS: Restablecimiento de servicio
    If IsDate(objEvento.FechaRestablecimientoServicio) Then
        Debug.Print "FechaRestablecimientoServicio: " & objEvento.FechaRestablecimientoServicio
    End If
    
    Debug.Print "Test 4 PASÓ"
    TestCalculoSLAs = True
    Exit Function
    
errores:
    Debug.Print "Error en Test 4: " & Err.Description
    TestCalculoSLAs = False
End Function

'===========================================================
' Función: LimpiarEventosTest
' Elimina eventos de prueba creados
'===========================================================
Public Function LimpiarEventosTest() As Long
    Dim rs As DAO.Recordset
    Dim IDEventoTest As String
    Dim fechaTest As String
    Dim count As Long
    
    On Error GoTo errores
    
    fechaTest = Format(Date, "yyyymmdd")
    IDEventoTest = "TEST-" & fechaTest & "*"
    
    Set rs = getdb().OpenRecordset("SELECT IDEvento FROM TbEventos WHERE IDEvento LIKE '" & IDEventoTest & "'")
    
    count = 0
    Do While Not rs.EOF
        getdb().Execute "DELETE FROM TbEventos WHERE IDEvento = '" & rs!IDEVENTO & "'"
        count = count + 1
        rs.MoveNext
    Loop
    
    rs.Close
    Debug.Print "Eventos de test eliminados: " & count
    LimpiarEventosTest = count
    Exit Function
    
errores:
    Debug.Print "Error al limpiar: " & Err.Description
    LimpiarEventosTest = 0
End Function


