VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RevisionServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' ==========================================================================
' MÓDULO: RevisionServicio.cls (NUEVO)
' ACCIÓN: Encapsula la lógica de negocio para el sub-flujo de revisión de Calidad.
' ==========================================================================
Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: RevisionServicio.cls
' ACCIÓN: Transacción Híbrida (Autogestionada o Heredada)
' ==========================================================================
Public Sub GuardarDecisionRevision(ByVal idSolicitud As Long, ByVal nuevoEstadoRevision As String, ByVal comentarios As String, Optional db As DAO.Database)
    Dim solServ As New SolicitudServicio
    Dim logServ As New LogServicio
    Dim sol As Solicitud
    Dim valorAnterior As String
    
    ' Variables de transacción
    Dim ws As DAO.Workspace
    Dim dbConexion As DAO.Database
    Dim gestionaTransaccionLocal As Boolean
    
    On Error GoTo Errores
    
    ' 1. CONFIGURACIÓN DE TRANSACCIÓN INTELIGENTE
    If db Is Nothing Then
        ' Modo Stand-alone (Botón "Registrar Decisión") -> Creamos transacción
        gestionaTransaccionLocal = True
        Set ws = getWorkspace()
        Set dbConexion = ws.OpenDatabase(getdb().name, False, False, "MS Access;PWD=" & GetPasswordDB())
        ws.BeginTrans
    Else
        ' Modo Heredado (Desde Workflow) -> Usamos la existente
        gestionaTransaccionLocal = False
        Set dbConexion = db
    End If
    
    ' --- INICIO DE OPERACIONES ATÓMICAS ---
    
    ' 2. Obtener solicitud (Lectura segura dentro de la conexión)
    ' Nota: SolicitudServicio.getSolicitudPorID usa getdb(), que es una conexión nueva.
    ' Para lectura no es crítico en este punto, pero idealmente pasariamos dbConexion.
    Set sol = solServ.getSolicitudPorID(idSolicitud, dbConexion)
    
    If sol Is Nothing Then Err.Raise 513, "GuardarDecisionRevision", "Solicitud no encontrada."
    If sol.idEstadoInterno <> enumEstados.estadoModificacion Then Err.Raise 513, "GuardarDecisionRevision", "Estado incorrecto."
    
    valorAnterior = sol.revisionCalidadEstado
    
    ' 3. Actualizar objeto en memoria
    sol.revisionCalidadEstado = nuevoEstadoRevision
    sol.revisionCalidadComentarios = comentarios
    sol.fechaModificacion = Now()
    sol.usuarioModificacion = m_ObjUsuarioActivo.nombre
    
    ' 4. Persistir Solicitud (UPDATE 1)
    Call SolicitudRepositorio.GuardarSolicitud(sol, dbConexion)
    
    ' 5. Registrar Log (INSERT 2)
    Call logServ.RegistrarCambio("tbSolicitudes", sol.idSolicitud, "revisionCalidadEstado", valorAnterior, nuevoEstadoRevision, "REVISION_CALIDAD", dbConexion)
    
    If Trim(comentarios) <> "" Then
        Call logServ.RegistrarCambio("tbSolicitudes", sol.idSolicitud, "revisionCalidadComentarios", "[...]", Left(comentarios, 255), "REVISION_CALIDAD", dbConexion)
    End If
    
    ' --- FIN DE OPERACIONES ATÓMICAS ---
    
    ' 6. COMMIT (Solo si somos dueños de la transacción)
    If gestionaTransaccionLocal Then
        ws.CommitTrans
    End If
    
LimpiarYSalir:
    On Error Resume Next
    If gestionaTransaccionLocal And Not dbConexion Is Nothing Then dbConexion.Close
    Set ws = Nothing
    Set solServ = Nothing: Set logServ = Nothing: Set sol = Nothing
    Exit Sub
    
Errores:
    ' 7. ROLLBACK (Solo si somos dueños)
    If gestionaTransaccionLocal And Not ws Is Nothing Then
        ws.Rollback
    End If
    
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "RevisionServicio.GuardarDecisionRevision"
    errObj.Raise
End Sub

Public Sub ActualizarMotivoRechazoActivo(ByVal idSolicitud As Long, ByVal comentarios As String, Optional db As DAO.Database)
    Dim solServ As New SolicitudServicio
    Dim sol As Solicitud
    Dim ws As DAO.Workspace
    Dim dbConexion As DAO.Database
    Dim gestionaTransaccionLocal As Boolean
    
    On Error GoTo Errores
    
    If db Is Nothing Then
        gestionaTransaccionLocal = True
        Set ws = getWorkspace()
        Set dbConexion = ws.OpenDatabase(getdb().name, False, False, "MS Access;PWD=" & GetPasswordDB())
        ws.BeginTrans
    Else
        gestionaTransaccionLocal = False
        Set dbConexion = db
    End If
    
    Set sol = solServ.getSolicitudPorID(idSolicitud, dbConexion)
    If sol Is Nothing Then Err.Raise 513, "ActualizarMotivoRechazoActivo", "Solicitud no encontrada."
    
    Call RechazoRepositorio.ActualizarComentarioRechazoActivo(idSolicitud, comentarios, dbConexion)
    sol.revisionCalidadComentarios = comentarios
    sol.fechaModificacion = Now()
    sol.usuarioModificacion = m_ObjUsuarioActivo.nombre
    Call SolicitudRepositorio.GuardarSolicitud(sol, dbConexion)
    
    If gestionaTransaccionLocal Then
        ws.CommitTrans
    End If
    
LimpiarYSalir:
    On Error Resume Next
    If gestionaTransaccionLocal And Not dbConexion Is Nothing Then dbConexion.Close
    Set ws = Nothing
    Set solServ = Nothing: Set sol = Nothing
    Exit Sub
    
Errores:
    If gestionaTransaccionLocal And Not ws Is Nothing Then
        ws.Rollback
    End If
    
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "RevisionServicio.ActualizarMotivoRechazoActivo"
    errObj.Raise
End Sub

Public Function ObtenerRechazoActivo(ByVal idSolicitud As Long, Optional ByRef db As DAO.Database) As rechazo
    On Error GoTo Errores
    
    Set ObtenerRechazoActivo = RechazoRepositorio.GetUltimoRechazoActivo(idSolicitud, db)
    Exit Function
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "RevisionServicio.ObtenerRechazoActivo"
    errObj.Raise
End Function

