VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AdjuntosServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: AdjuntosServicio.cls
' ACCIÓN: Reemplazar GuardarAdjuntoDesdeArchivo para que acepte una DB transaccional externa.
' ==========================================================================
Public Sub GuardarAdjuntoDesdeArchivo(ByVal idSolicitud As Long, _
                                     ByVal etapaNombre As String, _
                                     ByVal rutaArchivoOrigen As String, _
                                     Optional ByVal tipoAccion As String = "Adjuntado", _
                                     Optional ByRef db As DAO.Database = Nothing, _
                                     Optional ByRef idAdjuntoOut As Long = 0, _
                                     Optional ByVal descripcion As String = "", _
                                     Optional ByVal forceFailAfterFileCopy As Boolean = False)
    Dim adj As Adjunto, fso As Object, nombre As String, destinoDir As String, destinoFull As String
    Dim ws As DAO.Workspace, archivoCopiado As Boolean
    Dim dbConexion As DAO.Database ' <<< NUEVA VARIABLE
    Dim gestionaTransaccionLocal As Boolean
    
    On Error GoTo Errores
    
   ' --- LÓGICA DE GESTIÓN DE TRANSACCIÓN CORREGIDA ---
    If db Is Nothing Then
        gestionaTransaccionLocal = True
        ' Usamos el workspace aislado
        Set ws = getWorkspace()
        ' Usamos la cadena de conexión correcta y abrimos en el mismo workspace
        Set dbConexion = ws.OpenDatabase(getdb().name, False, False, "MS Access;PWD=" & GetPasswordDB())
        ws.BeginTrans
    Else
        gestionaTransaccionLocal = False
        Set dbConexion = db
    End If
    ' --------------------------------------------------
    
    archivoCopiado = False
    
    ' --- Lógica de validación y copia de archivo (sin cambios) ---
    If idSolicitud <= 0 Then Err.Raise 513, , "El ID de la Solicitud no es válido."
    If Nz(rutaArchivoOrigen, "") = "" Then Err.Raise 513, , "La ruta del archivo de origen no puede estar vacía."
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(rutaArchivoOrigen) Then Err.Raise 513, , "El archivo de origen no existe: " & rutaArchivoOrigen

    If etapaNombre = "Documento Final Firmado" Then
        Dim ext As String
        ext = LCase$(fso.GetExtensionName(rutaArchivoOrigen))
        If ext <> "pdf" Then
            Err.Raise 513, , "El archivo seleccionado debe ser un documento PDF. El archivo tiene extensión '." & ext & "'."
        End If
    End If

    nombre = fso.GetFileName(rutaArchivoOrigen)
    destinoDir = m_ObjEntorno.URLDirectorioDocumentacion
    If Right$(destinoDir, 1) <> "\" Then destinoDir = destinoDir & "\"
    destinoFull = destinoDir & nombre
    
    If fso.FileExists(destinoFull) Then
        Dim base As String
        base = fso.GetBaseName(destinoFull): ext = fso.GetExtensionName(destinoFull)
        nombre = base & "_" & Format(Now, "yyyymmdd_hhnnss") & IIf(ext <> "", "." & ext, "")
        destinoFull = destinoDir & nombre
    End If
    fso.CopyFile rutaArchivoOrigen, destinoFull, True
    archivoCopiado = True

    ' Hook de test para BR-006 (rollback). Solo activo si el caller lo pide
    ' explicitamente; en producción queda False (no-op). Ver tests
    ' Test_Adjuntos_Strict_GuardarAdjuntoDesdeArchivo_TriggersRollback en
    ' docs/epics/strict-tdd-coverage.md.
    If forceFailAfterFileCopy Then
        Err.Raise 513, "GuardarAdjuntoDesdeArchivo", "Falla forzada por test (forceFailAfterFileCopy=True)"
    End If

    ' --- Lógica de creación del objeto adjunto y persistencia ---
    Set adj = New Adjunto
    adj.idSolicitud = idSolicitud
    adj.etapaWF = etapaNombre
    adj.nombreArchivo = nombre
    adj.fechaSubida = Now()
    adj.usuarioSubida = IIf(Len(Nz(m_ObjUsuarioActivo.nombre, "")) > 0, m_ObjUsuarioActivo.nombre, Environ$("Username"))
    
    If Trim$(Nz(descripcion, "")) <> "" Then
        adj.descripcion = descripcion
    Else
        ' REGLA CANÓNICA: Si no hay descripción, se asigna una por defecto informativa.
        adj.descripcion = "Adjunto de la etapa " & etapaNombre
    End If
    
    adj.tipoAccion = tipoAccion

    Call Validar(adj)
    ' Usamos la conexión correcta (sea la local o la externa)
    Call AdjuntoRepositorio.Guardar(adj, dbConexion)
    idAdjuntoOut = adj.idAdjunto
    
    ' Solo hacemos Commit si gestionamos la transacción localmente
    If gestionaTransaccionLocal Then ws.CommitTrans

LimpiarYSalir:
    On Error Resume Next
    ' Solo cerramos la conexión si la hemos creado aquí
    If gestionaTransaccionLocal And Not dbConexion Is Nothing Then dbConexion.Close
    Set adj = Nothing: Set fso = Nothing
    If gestionaTransaccionLocal Then Set ws = Nothing
    Exit Sub
    
Errores:
     ' 1. CAPTURA
    Dim savedErr As Long, savedDesc As String
    savedErr = Err.Number: savedDesc = Err.description

    ' 2. LIMPIEZA / ROLLBACK
    If gestionaTransaccionLocal And Not ws Is Nothing Then
        On Error Resume Next ' Prevenir fallo en rollback
        ws.Rollback
    End If
    
    ' Compensación de archivo
    If archivoCopiado Then
        On Error Resume Next
        If Not fso Is Nothing Then fso.DeleteFile destinoFull, True
    End If
    
    ' 3. PROPAGACIÓN
    Dim errObj As New CondorError
    errObj.Create savedErr, savedDesc, "AdjuntosServicio.GuardarAdjuntoDesdeArchivo"
    errObj.Raise
End Sub


Public Sub ActualizarFicheroAdjunto(ByVal idAdjunto As Long, _
                                    ByVal rutaArchivoOrigen As String, _
                                    Optional ByRef dbTransaccional As DAO.Database)
    Dim adj As Adjunto
    Dim fso As Object
    Dim destinoDir As String
    Dim destinoFull As String
    Dim ws As DAO.Workspace
    Dim dbConexion As DAO.Database
    Dim gestionaTransaccionLocal As Boolean
    Dim archivoCopiado As Boolean
    
    On Error GoTo Errores
    
    If idAdjunto <= 0 Then Err.Raise 513, , "Invalid attachment ID."
    If Nz(rutaArchivoOrigen, "") = "" Then Err.Raise 513, , "Source path cannot be empty."
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(rutaArchivoOrigen) Then Err.Raise 513, , "Source file does not exist: " & rutaArchivoOrigen
    
    If dbTransaccional Is Nothing Then
        gestionaTransaccionLocal = True
        Set ws = getWorkspace()
        Set dbConexion = ws.OpenDatabase(getdb().name, False, False, "MS Access;PWD=" & GetPasswordDB())
        ws.BeginTrans
    Else
        gestionaTransaccionLocal = False
        Set dbConexion = dbTransaccional
    End If
    
    Set adj = AdjuntoRepositorio.getPorID(idAdjunto, dbConexion)
    If adj Is Nothing Then Err.Raise 513, , "Attachment not found."
    
    destinoDir = m_ObjEntorno.URLDirectorioDocumentacion
    If Right$(destinoDir, 1) <> "\" Then destinoDir = destinoDir & "\"
    destinoFull = destinoDir & adj.nombreArchivo
    
    If fso.FileExists(destinoFull) Then fso.DeleteFile destinoFull, True
    fso.CopyFile rutaArchivoOrigen, destinoFull, True
    archivoCopiado = True
    
    If gestionaTransaccionLocal Then ws.CommitTrans
    
LimpiarYSalir:
    On Error Resume Next
    If gestionaTransaccionLocal And Not dbConexion Is Nothing Then dbConexion.Close
    Set adj = Nothing
    Set fso = Nothing
    If gestionaTransaccionLocal Then Set ws = Nothing
    Exit Sub
    
Errores:
    If gestionaTransaccionLocal And Not ws Is Nothing Then
        On Error Resume Next
        ws.Rollback
    End If
    If archivoCopiado Then
        On Error Resume Next
        If Not fso Is Nothing Then fso.DeleteFile destinoFull, True
    End If
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "AdjuntosServicio.ActualizarFicheroAdjunto"
    errObj.Raise
End Sub

Public Function ExisteAdjuntoEtapa(ByVal idSolicitud As Long, ByVal etapaNombre As String, Optional db As DAO.Database) As Boolean
    On Error GoTo Errores
    ExisteAdjuntoEtapa = AdjuntoRepositorio.ExisteAdjuntoEtapa(idSolicitud, etapaNombre, db)
    Exit Function
Errores:
    Dim errObj As New CondorError: errObj.Create Err.Number, Err.description, "AdjuntosServicio.ExisteAdjuntoEtapa": errObj.Raise
End Function

Public Function GetIdAdjuntoPorEtapa(ByVal idSolicitud As Long, ByVal etapaNombre As String, Optional db As DAO.Database) As Long
    Dim adj As Adjunto
    On Error GoTo Errores
    
    Set adj = AdjuntoRepositorio.getPorEtapa(idSolicitud, etapaNombre, db)
    If adj Is Nothing Then
        GetIdAdjuntoPorEtapa = 0
    Else
        GetIdAdjuntoPorEtapa = adj.idAdjunto
    End If
    
    Exit Function
Errores:
    GetIdAdjuntoPorEtapa = 0
    Dim errObj As New CondorError: errObj.Create Err.Number, Err.description, "AdjuntosServicio.GetIdAdjuntoPorEtapa": errObj.Raise
End Function

Public Function ExisteAdjuntoGeneradoCierre(ByVal idSolicitud As Long) As Boolean
    Dim adjs As Collection
    Dim adj As Variant
    Dim fso As Object
    
    On Error GoTo Errores
    ExisteAdjuntoGeneradoCierre = False
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set adjs = getAdjuntosPorSolicitud(idSolicitud)
    
    For Each adj In adjs
        If adj.descripcion = "Borrador" And adj.tipoAccion = "Generado" Then
            If fso.FileExists(adj.nombreArchivo) Then
                ExisteAdjuntoGeneradoCierre = True
                Exit For
            End If
        End If
    Next
    
    Set adjs = Nothing
    Set fso = Nothing
    Exit Function
    
Errores:
    Set adjs = Nothing
    Set fso = Nothing
    Dim errObj As New CondorError: errObj.Create Err.Number, Err.description, "AdjuntosServicio.ExisteAdjuntoGeneradoCierre": errObj.Raise
End Function

Public Function GeneradoPDFParaFirma(ByVal idSolicitud As Long) As Boolean
    Dim etapaFormalizacion As String
    
    On Error GoTo Errores
    
    etapaFormalizacion = m_ObjEntorno.estados("7").nombreEstado
    
    GeneradoPDFParaFirma = AdjuntoRepositorio.ExisteAdjuntoEtapa(idSolicitud, etapaFormalizacion)
    
    Exit Function
    
Errores:
    Dim errObj As New CondorError: errObj.Create Err.Number, Err.description, "AdjuntosServicio.GeneradoPDFParaFirma": errObj.Raise
End Function

Public Function getAdjuntoPorID(ByVal idAdjunto As Long, Optional ByRef dbConexion As DAO.Database) As Adjunto
    On Error GoTo Errores
    Set getAdjuntoPorID = AdjuntoRepositorio.getPorID(idAdjunto, dbConexion)
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "AdjuntosServicio.getAdjuntoPorID"
    errObj.AddToCallStack "idAdjunto: " & idAdjunto
    errObj.Raise
End Function
Private Sub Validar(ByRef adj As Adjunto)

    ' Responsabilidad: Aplicar reglas de negocio a la entidad Adjunto.
    '                  Lanza un error si una regla no se cumple.
    On Error GoTo Errores
    
    If adj Is Nothing Then Err.Raise 513, "Validación", "El objeto Adjunto a validar no puede ser nulo."

    ' Regla 1: ID de solicitud debe ser válido.
    If adj.idSolicitud <= 0 Then Err.Raise 513, "Validación", "El ID de la Solicitud asociada al adjunto no es válido."
    
    ' Regla 2: La etapa del flujo de trabajo no puede estar vacía.
    If Trim(Nz(adj.etapaWF, "")) = "" Then Err.Raise 513, "Validación", "La etapa del flujo de trabajo (etapaWF) es obligatoria para el adjunto."
    
    ' Regla 3: El nombre del archivo es obligatorio y debe tener una longitud razonable (ej. 255 caracteres).
    If Trim(Nz(adj.nombreArchivo, "")) = "" Then
        Err.Raise 513, "Validación", "El nombre del archivo adjunto es obligatorio."
    ElseIf Len(adj.nombreArchivo) > 255 Then
        Err.Raise 513, "Validación", "El nombre del archivo adjunto no puede exceder los 255 caracteres."
    End If
    
    ' Regla 4: El usuario que sube el archivo es obligatorio.
    If Trim(Nz(adj.usuarioSubida, "")) = "" Then Err.Raise 513, "Validación", "No se ha podido determinar el usuario que sube el adjunto."
    
    ' Regla 5: La fecha de subida debe ser una fecha válida.
    If Not IsDate(adj.fechaSubida) Or adj.fechaSubida <= 0 Then Err.Raise 513, "Validación", "La fecha de subida del adjunto no es válida."

    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "AdjuntosServicio.Validar"
    errObj.Raise
End Sub
Public Function getAdjuntosPorSolicitud(ByVal idSolicitud As Long) As Collection
    ' RESPONSABILIDAD: Versión pública clásica sin contexto transaccional.
    On Error GoTo Errores
    
    If idSolicitud <= 0 Then
        Err.Raise 513, "getAdjuntosPorSolicitud", "El ID de la solicitud no es válido."
    End If
    
    Set getAdjuntosPorSolicitud = AdjuntoRepositorio.getPorIdSolicitud(idSolicitud)
    
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "AdjuntosServicio.getAdjuntosPorSolicitud"
    errObj.AddToCallStack "idSolicitud: " & idSolicitud
    errObj.Raise
End Function

Public Function getAdjuntosPorSolicitudTx(ByVal idSolicitud As Long, ByRef db As DAO.Database) As Collection
    ' RESPONSABILIDAD: Versión transaccional para usar dentro de una DB ya abierta.
    On Error GoTo Errores
    
    If idSolicitud <= 0 Then
        Err.Raise 513, "getAdjuntosPorSolicitudTx", "El ID de la solicitud no es válido."
    End If
    
    Set getAdjuntosPorSolicitudTx = AdjuntoRepositorio.getPorIdSolicitud(idSolicitud, db)
    
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "AdjuntosServicio.getAdjuntosPorSolicitudTx"
    errObj.AddToCallStack "idSolicitud: " & idSolicitud
    errObj.Raise
End Function
' ==========================================================================
' MÓDULO: AdjuntosServicio.cls
' ACCIÓN: Reemplazar la función getAdjuntosViewModelPorSolicitud (VERSIÓN FINAL SIMPLIFICADA)
' ==========================================================================

' ==========================================================================
' MÓDULO: AdjuntosServicio.cls
' ACCIÓN: Reemplazar getAdjuntosViewModelPorSolicitud con la lógica de traducción de ID a Nombre.
' ==========================================================================

Public Function getAdjuntosViewModelPorSolicitud(ByVal idSolicitud As Long) As Collection
    On Error GoTo Errores
    
    If idSolicitud <= 0 Then
        Err.Raise 513, "getAdjuntosViewModelPorSolicitud", "El ID de la solicitud no es válido."
    End If
    
    Dim colAdjuntos As Collection
    Set colAdjuntos = AdjuntoRepositorio.getPorIdSolicitud(idSolicitud)
    
    Dim colViewModels As New Collection
    Dim adj As Adjunto
    Dim vm As AdjuntoViewModel
    
    If Not colAdjuntos Is Nothing Then
        For Each adj In colAdjuntos
            Set vm = New AdjuntoViewModel
            vm.idAdjunto = adj.idAdjunto
            vm.idSolicitud = adj.idSolicitud
            vm.nombreArchivo = adj.nombreArchivo
            vm.fechaSubida = adj.fechaSubida
            vm.usuarioSubida = adj.usuarioSubida
            vm.descripcion = adj.descripcion
            vm.tipoAccion = adj.tipoAccion
            ' --- INICIO DE LA LÓGICA DE ENRIQUECIMIENTO ---
            Dim etapaKey As String
            etapaKey = CStr(adj.etapaWF)
            
            ' Comprobamos si la "etapa" guardada es un número (ID de estado)
            If IsNumeric(etapaKey) Then
                ' Si es un número, lo buscamos en la caché de estados.
                If m_ObjEntorno.estados.Exists(etapaKey) Then
                    vm.etapaNombre = m_ObjEntorno.estados(etapaKey).nombreEstado
                Else
                    ' Fallback: si el ID no se encuentra, mostramos el ID para depuración.
                    vm.etapaNombre = "Estado ID " & etapaKey
                End If
            Else
                ' Si ya es un texto, lo usamos directamente.
                vm.etapaNombre = adj.etapaWF
            End If
            ' --- FIN DE LA LÓGICA DE ENRIQUECIMIENTO ---
            
            colViewModels.Add vm
        Next adj
    End If
    
    Set getAdjuntosViewModelPorSolicitud = colViewModels
    
LimpiarYSalir:
    Exit Function
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "AdjuntosServicio.getAdjuntosViewModelPorSolicitud"
    errObj.AddToCallStack "idSolicitud: " & idSolicitud
    errObj.Raise
End Function
Public Sub EliminarAdjuntosPorEtapas(ByVal idSolicitud As Long, ByVal etapas As Variant, Optional ByRef db As DAO.Database)
    Dim colAdjuntos As Collection
    Dim adj As Adjunto
    Dim etapaObjetivo As String
    Dim i As Long
    
    On Error GoTo Errores
    
    If idSolicitud <= 0 Then
        Err.Raise 513, "EliminarAdjuntosPorEtapas", "El ID de la solicitud no es válido."
    End If
    
    Set colAdjuntos = AdjuntoRepositorio.getPorIdSolicitud(idSolicitud, db)
    If colAdjuntos Is Nothing Then
        Exit Sub
    End If
    
    For Each adj In colAdjuntos
        If IsArray(etapas) Then
            For i = LBound(etapas) To UBound(etapas)
                etapaObjetivo = CStr(etapas(i))
                If EtapaAdjuntoCoincide(adj, etapaObjetivo) Then
                    Call EliminarAdjunto(adj.idAdjunto, db)
                    Exit For
                End If
            Next i
        Else
            etapaObjetivo = CStr(etapas)
            If EtapaAdjuntoCoincide(adj, etapaObjetivo) Then
                Call EliminarAdjunto(adj.idAdjunto, db)
            End If
        End If
    Next adj
    
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "AdjuntosServicio.EliminarAdjuntosPorEtapas"
    errObj.AddToCallStack "idSolicitud: " & idSolicitud
    errObj.Raise
End Sub

Private Function EtapaAdjuntoCoincide(ByRef adj As Adjunto, ByVal etapaObjetivo As String) As Boolean
    Dim etapaKey As String
    Dim etapaNombre As String
    
    If Len(Trim$(etapaObjetivo)) = 0 Then
        Exit Function
    End If
    
    etapaKey = CStr(adj.etapaWF)
    
    If StrComp(etapaKey, etapaObjetivo, vbTextCompare) = 0 Then
        EtapaAdjuntoCoincide = True
        Exit Function
    End If
    
    If IsNumeric(etapaKey) Then
        If m_ObjEntorno.estados.Exists(etapaKey) Then
            etapaNombre = m_ObjEntorno.estados(etapaKey).nombreEstado
        Else
            etapaNombre = ""
        End If
        
        If StrComp(etapaNombre, etapaObjetivo, vbTextCompare) = 0 Then
            EtapaAdjuntoCoincide = True
            Exit Function
        End If
    End If
End Function

Public Sub EliminarAdjunto(ByVal idAdjunto As Long, Optional ByRef dbTransaccional As DAO.Database)
    Dim adj As Adjunto, ws As DAO.Workspace, rutaCompleta As String, fso As Object
    Dim dbConexion As DAO.Database
    Dim gestionaTransaccionLocal As Boolean
    
    On Error GoTo Errores
    
    If idAdjunto <= 0 Then
        Err.Raise 513, "EliminarAdjunto", "El ID del adjunto no es válido."
    End If
    
    If dbTransaccional Is Nothing Then
        gestionaTransaccionLocal = True
        Set ws = getWorkspace()
        Set dbConexion = ws.OpenDatabase(getdb().name, False, False, "MS Access;PWD=" & GetPasswordDB())
        ws.BeginTrans
    Else
        gestionaTransaccionLocal = False
        Set dbConexion = dbTransaccional
    End If
    
    Set adj = AdjuntoRepositorio.getPorID(idAdjunto, dbConexion)
    If adj Is Nothing Then
        GoTo LimpiarYSalir
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    rutaCompleta = m_ObjEntorno.URLDirectorioDocumentacion & adj.nombreArchivo
    
    Call AdjuntoRepositorio.Eliminar(idAdjunto, dbConexion)
    If fso.FileExists(rutaCompleta) Then
        fso.DeleteFile rutaCompleta, True
    End If
    
    If gestionaTransaccionLocal Then
        ws.CommitTrans
    End If
    
LimpiarYSalir:
    On Error Resume Next
    If gestionaTransaccionLocal And Not dbConexion Is Nothing Then
        dbConexion.Close
    End If
    Set adj = Nothing: Set ws = Nothing: Set fso = Nothing
    Exit Sub
Errores:
    If gestionaTransaccionLocal And Not ws Is Nothing Then
        ws.Rollback
    End If
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "AdjuntosServicio.EliminarAdjunto"
    errObj.AddToCallStack "idAdjunto: " & idAdjunto
    errObj.Raise
End Sub

' ==========================================================================
' SubirYCerrar: adjunta el PDF de cierre y, si la etapa es "Documento Final
' Firmado", ejecuta la callback atómica de aprobación dentro de una sola
' transacción DAO: fija decisionFinal="APROBADO" en el tipo correspondiente
' (PC, CD_CA, CD_CA_SUB, PC_SUB) y transiciona a estadoAprobada. NO hace
' MsgBox ni NotificarCambioExterno — la UI feedback queda en el caller
' (Form_frmGestionAdjuntos.EjecutarFlujoSubida).
'
' Slice A5 (2026-06-15): seam extraído de EjecutarFlujoSubida. Permite tests
' fixture-first sin depender del formulario.
' ==========================================================================
Public Sub SubirYCerrar(ByVal idSolicitud As Long, _
                        ByVal etapaNombre As String, _
                        ByVal rutaArchivoOrigen As String, _
                        Optional ByVal tipoAccion As String = "Adjuntado", _
                        Optional ByVal descripcion As String = "", _
                        Optional ByRef dbTransaccional As DAO.Database, _
                        Optional ByRef idAdjuntoOut As Long = 0)

    Dim idAdjunto As Long
    Me.GuardarAdjuntoDesdeArchivo idSolicitud, etapaNombre, rutaArchivoOrigen, tipoAccion, dbTransaccional, idAdjunto, descripcion
    idAdjuntoOut = idAdjunto

    If etapaNombre = "Documento Final Firmado" Then
        Call EjecutarCallbackAprobacion(idSolicitud, dbTransaccional)
    End If
End Sub

' ==========================================================================
' EjecutarCallbackAprobacion: abre transacción (si db es Nothing), fija
' decisionFinal="APROBADO" en el gemelo correspondiente y transiciona a
' estadoAprobada. Esta callback es la que estaba acoplada en
' Form_frmGestionAdjuntos.EjecutarFlujoSubida líneas 175-225 (Slice A5).
' ==========================================================================
Public Sub EjecutarCallbackAprobacion(ByVal idSolicitud As Long, Optional ByRef dbTransaccional As DAO.Database)
    Dim ws As DAO.Workspace
    Dim db As DAO.Database
    Dim gestionaTransaccionLocal As Boolean
    Dim sol As Solicitud
    Dim wfServ As New WorkflowServicio
    Dim solServ As New SolicitudServicio
    Dim pcServ As New DatosPCServicio
    Dim cdcaServ As New DatosCDCAServicio
    Dim cdcasubServ As New DatosCDCASUBServicio
    Dim pcsubServ As New DatosPCSUBServicio
    Dim pc As DatosPC
    Dim cd As DatosCDCA
    Dim cdc As DatosCDCASUB
    Dim pcs As DatosPCSUB
    Dim errObj As CondorError

    On Error GoTo Errores

    If dbTransaccional Is Nothing Then
        gestionaTransaccionLocal = True
        Set ws = getWorkspace()
        Set db = ws.OpenDatabase(getdb().name, False, False, "MS Access;PWD=" & GetPasswordDB())
        ws.BeginTrans
    Else
        gestionaTransaccionLocal = False
        Set db = dbTransaccional
    End If

    Set sol = solServ.getSolicitudPorID(idSolicitud)
    If sol Is Nothing Then
        Set errObj = New CondorError
        errObj.Create 513, "EjecutarCallbackAprobacion", _
            "No se encontró la solicitud " & idSolicitud & " para ejecutar la callback de aprobación."
        errObj.Raise
    End If

    Select Case sol.tipoSolicitud
        Case "PC"
            Set pc = New DatosPC
            pc.idSolicitud = sol.idSolicitud
            pc.decisionFinal = "APROBADO"
            pcServ.GuardarDecisionFinal pc, db
        Case "CD_CA"
            Set cd = New DatosCDCA
            cd.idSolicitud = sol.idSolicitud
            cd.decisionFinal = "APROBADO"
            cdcaServ.GuardarDecisionFinal cd, db
        Case "CD_CA_SUB"
            Set cdc = New DatosCDCASUB
            cdc.idSolicitud = sol.idSolicitud
            cdc.decisionFinal = "APROBADO"
            cdcasubServ.GuardarDecisionFinal cdc, db
        Case "PC_SUB"
            Set pcs = New DatosPCSUB
            pcs.idSolicitud = sol.idSolicitud
            pcs.decisionFinal = "APROBADO"
            pcsubServ.GuardarDecisionFinal pcs, db
        Case Else
            Set errObj = New CondorError
            errObj.Create 513, "EjecutarCallbackAprobacion", _
                "tipoSolicitud no soportado por la callback de aprobación: '" & sol.tipoSolicitud & "'"
            errObj.Raise
    End Select

    wfServ.EjecutarTransicion sol, enumEstados.estadoAprobada, m_ObjUsuarioActivo, db

    If gestionaTransaccionLocal Then ws.CommitTrans

LimpiarYSalir:
    On Error Resume Next
    If gestionaTransaccionLocal And Not db Is Nothing Then db.Close
    Set ws = Nothing: Set db = Nothing
    Set sol = Nothing
    Set pc = Nothing: Set cd = Nothing: Set cdc = Nothing: Set pcs = Nothing
    Set wfServ = Nothing: Set solServ = Nothing
    Set pcServ = Nothing: Set cdcaServ = Nothing: Set cdcasubServ = Nothing: Set pcsubServ = Nothing
    Exit Sub

Errores:
    If gestionaTransaccionLocal And Not ws Is Nothing Then ws.Rollback
    Set errObj = New CondorError
    errObj.Create Err.Number, Err.description, "AdjuntosServicio.EjecutarCallbackAprobacion"
    errObj.Raise
End Sub




