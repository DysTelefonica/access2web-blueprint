VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO DE CLASE: LogServicio.cls (REFACTORIZADO CON VALIDACIÓN)
' ==========================================================================

Public Sub RegistrarCambio(ByVal tabla As String, ByVal registroID As Long, ByVal campo As String, _
                           ByVal valorAnterior As Variant, ByVal valorNuevo As Variant, ByVal tipoOperacion As String, _
                           Optional db As DAO.Database)
    Dim objLog As New LogCambio
    On Error GoTo Errores
    
    With objLog
        .fechaHora = Now()
        .usuario = m_ObjUsuarioActivo.nombre
        If g_blnImpersonando Then .suplantadoPor = m_ObjUsuarioReal.nombre Else .suplantadoPor = ""
        .tabla = tabla
        .registro = registroID
        .campo = campo
        .valorAnterior = Nz(valorAnterior, "")
        .valorNuevo = Nz(valorNuevo, "")
        .tipoOperacion = tipoOperacion
    End With
    
    Call ValidarCambio(objLog)
    ' Pasamos la conexión
    Call LogCambioRepositorio.GuardarLog(objLog, db)
    
LimpiarYSalir:
    On Error Resume Next
    Set objLog = Nothing
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "LogServicio.RegistrarCambio"
    errObj.Raise
End Sub

Public Sub RegistrarError(ByVal numeroError As Long, ByVal descripcionError As String, ByVal modulo As String, _
                          ByVal procedimiento As String, Optional ByVal contexto As String = "")
    Dim objLogError As New LogError
    On Error GoTo Errores
    
    With objLogError
        .fechaHora = Now()
        .usuario = IIf(m_ObjUsuarioActivo Is Nothing, "Desconocido", m_ObjUsuarioActivo.nombre)
        
        If g_blnImpersonando And Not m_ObjUsuarioReal Is Nothing Then
            .suplantadoPor = m_ObjUsuarioReal.nombre
        Else
            .suplantadoPor = ""
        End If
        
        .modulo = modulo
        .procedimiento = procedimiento
        .numeroError = numeroError
        .descripcionError = descripcionError
        .contexto = contexto
    End With
    
    ' 1. Validar el objeto antes de persistir
    Call ValidarError(objLogError)
    
    ' 2. Si la validación es exitosa, guardar
    Call LogErrorRepositorio.GuardarError(objLogError)
    
LimpiarYSalir:
    On Error Resume Next
    Set objLogError = Nothing
    Exit Sub
    
Errores:
    ' Este es un caso especial: si falla el registro de un error, no queremos
    ' entrar en un bucle infinito intentando registrar el fallo del registro de errores.
    ' En este caso, lo escribimos en la ventana de inmediato para diagnóstico.
    Debug.Print "ERROR CRÍTICO EN LogServicio.RegistrarError: " & Err.Number & " - " & Err.description
    ' Opcionalmente, podríamos mostrar un MsgBox, pero podría ser intrusivo.
    ' Por ahora, el Debug.Print es la opción más segura.
End Sub

Public Function getLogs(Optional p_Usuario As String, Optional p_Tabla As String, Optional p_RegistroID As String) As Scripting.Dictionary
    Dim sql As String
    Dim params As Object
    On Error GoTo Errores
    
    sql = getLogsSQL(p_Usuario, p_Tabla, p_RegistroID)
    
    ' Creamos parámetros solo si son necesarios
    Set params = CreateObject("Scripting.Dictionary")
    If Nz(p_Usuario, "") <> "" Then params.Add "p_Usuario", p_Usuario
    If Nz(p_Tabla, "") <> "" Then params.Add "p_Tabla", p_Tabla
    If Nz(p_RegistroID, "") <> "" And IsNumeric(p_RegistroID) Then params.Add "p_RegistroID", CLng(p_RegistroID)
    
    Set getLogs = RepositorioComun.HidratarColeccionDesdeSQL(sql, "LogCambio", "idLogCambio", getdb(), IIf(params.count > 0, params, Nothing))
    
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "LogServicio.getLogs"
    errObj.Raise
End Function

Public Function getErrores(Optional p_Usuario As String, Optional p_Modulo As String, Optional p_Procedimiento As String) As Scripting.Dictionary
    Dim sql As String
    Dim params As Object
    On Error GoTo Errores
    
    sql = getErroresSQL(p_Usuario, p_Modulo, p_Procedimiento)
    
    Set params = CreateObject("Scripting.Dictionary")
    If Nz(p_Usuario, "") <> "" Then params.Add "p_Usuario", p_Usuario
    If Nz(p_Modulo, "") <> "" Then params.Add "p_Modulo", p_Modulo
    If Nz(p_Procedimiento, "") <> "" Then params.Add "p_Procedimiento", "*" & p_Procedimiento & "*"
    
    Set getErrores = RepositorioComun.HidratarColeccionDesdeSQL(sql, "LogError", "idLogError", getdb(), IIf(params.count > 0, params, Nothing))
    
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "LogServicio.getErrores"
    errObj.Raise
End Function

Public Function getLogCambioPorID(ByVal id As Long) As LogCambio
    On Error GoTo Errores
    If id <= 0 Then
        Err.Raise 513, , "El ID del log de cambio no es válido."
    End If
    
    Set getLogCambioPorID = LogCambioRepositorio.getLogCambioPorID(id)
    Exit Function
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "LogServicio.getLogCambioPorID"
    errObj.Raise
End Function

Public Function getLogErrorPorID(ByVal id As Long) As LogError
    On Error GoTo Errores
    If id <= 0 Then
        Err.Raise 513, , "El ID del log de error no es válido."
    End If
    
    Set getLogErrorPorID = LogErrorRepositorio.getLogErrorPorID(id)
    Exit Function
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "LogServicio.getLogErrorPorID"
    errObj.Raise
End Function

Private Function getLogsSQL(Optional p_Usuario As String, Optional p_Tabla As String, Optional p_RegistroID As String) As String
    Dim sql As String
    Dim wheres As String
    
    sql = "SELECT * FROM tbLogCambios "
    wheres = "WHERE 1=1 "
    
    If Nz(p_Usuario, "") <> "" Then wheres = wheres & "AND usuario = [p_Usuario] "
    If Nz(p_Tabla, "") <> "" Then wheres = wheres & "AND tabla = [p_Tabla] "
    If Nz(p_RegistroID, "") <> "" And IsNumeric(p_RegistroID) Then wheres = wheres & "AND registro = [p_RegistroID] "
    
    getLogsSQL = sql & wheres & "ORDER BY fechaHora DESC;"
End Function

Private Function getErroresSQL(Optional p_Usuario As String, Optional p_Modulo As String, Optional p_Procedimiento As String) As String
    Dim sql As String
    Dim wheres As String
    
    sql = "SELECT * FROM tbLogErrores "
    wheres = "WHERE 1=1 "
    
    If Nz(p_Usuario, "") <> "" Then wheres = wheres & "AND usuario = [p_Usuario] "
    If Nz(p_Modulo, "") <> "" Then wheres = wheres & "AND modulo = [p_Modulo] "
    If Nz(p_Procedimiento, "") <> "" Then wheres = wheres & "AND procedimiento LIKE [p_Procedimiento] "
    
    getErroresSQL = sql & wheres & "ORDER BY fechaHora DESC;"
End Function

' --- MÉTODOS DE VALIDACIÓN PRIVADOS ---

Private Sub ValidarCambio(ByRef objLog As LogCambio)
    On Error GoTo Errores
    
    If objLog Is Nothing Then Err.Raise 513, "Validación", "El objeto LogCambio a validar es nulo."
    If Trim(Nz(objLog.usuario, "")) = "" Then Err.Raise 513, "Validación", "El usuario es obligatorio para un registro de cambio."
    If Trim(Nz(objLog.tabla, "")) = "" Then Err.Raise 513, "Validación", "La tabla es obligatoria para un registro de cambio."
    If objLog.registro <= 0 Then Err.Raise 513, "Validación", "El ID del registro es obligatorio para un registro de cambio."
    If Trim(Nz(objLog.campo, "")) = "" Then Err.Raise 513, "Validación", "El campo modificado es obligatorio para un registro de cambio."
    
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "LogServicio.ValidarCambio"
    errObj.Raise
End Sub

Private Sub ValidarError(ByRef objLogError As LogError)
    On Error GoTo Errores
    
    If objLogError Is Nothing Then Err.Raise 513, "Validación", "El objeto LogError a validar es nulo."
    If Trim(Nz(objLogError.usuario, "")) = "" Then Err.Raise 513, "Validación", "El usuario es obligatorio para un registro de error."
    If Trim(Nz(objLogError.modulo, "")) = "" Then Err.Raise 513, "Validación", "El módulo es obligatorio para un registro de error."
    If Trim(Nz(objLogError.procedimiento, "")) = "" Then Err.Raise 513, "Validación", "El procedimiento es obligatorio para un registro de error."
    
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "LogServicio.ValidarError"
    errObj.Raise
End Sub


