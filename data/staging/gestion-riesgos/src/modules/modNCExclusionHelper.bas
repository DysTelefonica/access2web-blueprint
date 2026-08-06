Attribute VB_Name = "modNCExclusionHelper"
' =============================================================================
' modNCExclusionHelper.bas
'
' Helper module para exclusion y gestion de decisiones de NC en materializaciones.
' Project: gestion_riesgos
' Branch: feat/metodologia-e2e-riesgos-2026-06-19
' Refactor: 2026-06-19
'
' Helpers publicos (3):
'   #1 RegistrarMaterializacionNoNC — marca materializacion como NoNC
'   #2 RevocarDecisionMaterializacion — quita la decision NoNC
'   #3 EstablecerBotoneraNC — calcula estado para visibilidad de botones
'
' CORRECCION 2026-06-19 (helper #3):
'   La firma original del audit (p_EstadoActual As EnumSiNo) no matcheaba
'   el codigo real. Verificado contra Form_FormCalidadRiesgoMaterializaciones.cls
'   (lineas 298-369) y RiesgoMaterializacion.cls (ParaNCCalculado, NC):
'   Helper #3 NO recibe estado externo; resuelve internamente los 4 campos
'   (EsMaterializacionCalcuado, FechaDecison, ParaNCCalculado, .NC).
'   Retorna uno de: "SinCalcular" / "PorTomar" / "NoSeGeneraNC" / "VinculadaNC"
' =============================================================================
Option Compare Database
Option Explicit

' --- Helper #1: RegistrarMaterializacionNoNC ---
Public Function RegistrarMaterializacionNoNC( _
    ByVal p_IDMaterializacion As String, _
    Optional ByRef db As DAO.Database = Nothing, _
    Optional ByRef p_PromptResult As Long) As String

    ' Resolucion de dependencias
    Dim m_RiesgoMat As RiesgoMaterializacion
    Dim m_Error As String
    Dim m_IDEdicion As String
    Dim logs(0 To 9) As String
    Dim logIdx As Long

    On Error GoTo errores
    m_Error = ""
    logIdx = 0
    logs(logIdx) = "RegistrarMaterializacionNoNC: id=" & p_IDMaterializacion
    logIdx = logIdx + 1

    ' --- Bypass de MsgBox (p_PromptResult injected por tests TDD) ---
    If p_PromptResult <> vbYes Then
        logs(logIdx) = "2. Usuario no confirmo MsgBox (p_PromptResult <> vbYes)"
        logIdx = logIdx + 1
        RegistrarMaterializacionNoNC = Test_Helper.BuildJsonOk("cancelled", logs)
        Exit Function
    End If

    ' --- Resolver materializacion ---
    Set m_RiesgoMat = Constructor.getRiesgoMaterializado(p_IDMaterializacion, m_Error)
    If m_Error <> "" Then
        logs(logIdx) = "3. Constructor.getRiesgoMaterializado error: " & m_Error
        logIdx = logIdx + 1
        RegistrarMaterializacionNoNC = Test_Helper.BuildJsonFail(m_Error, logs)
        Exit Function
    End If
    If m_RiesgoMat Is Nothing Then
        m_Error = "No se encontro la materializacion con ID: " & p_IDMaterializacion
        logs(logIdx) = "3. Materializacion Nothing"
        logIdx = logIdx + 1
        RegistrarMaterializacionNoNC = Test_Helper.BuildJsonFail(m_Error, logs)
        Exit Function
    End If

    ' --- Validar estado: si ya tiene ParaNC populado, es idempotente ---
    ' (El boton NoParaNC no se muestra si ParaNC ya esta populado,
    ' pero por seguridad el helper verifica y treata como exito si ya esta)
    If Nz(m_RiesgoMat.ParaNC, "") <> "" Then
        logs(logIdx) = "4. ParaNC ya populado (idempotente): ParaNC=" & m_RiesgoMat.ParaNC
        logIdx = logIdx + 1
        RegistrarMaterializacionNoNC = Test_Helper.BuildJsonOk("already-no-nc", logs)
        Exit Function
    End If

    ' --- Registrar como NoNC (llama al metodo del dominio) ---
    ' RegistrarParaNONC sin parametros: fecha de decision se setea a Date internamente
    m_Error = ""
    m_RiesgoMat.RegistrarParaNONC , m_Error
    If m_Error <> "" Then
        logs(logIdx) = "5. RegistrarParaNONC error: " & m_Error
        logIdx = logIdx + 1
        RegistrarMaterializacionNoNC = Test_Helper.BuildJsonFail(m_Error, logs)
        Exit Function
    End If

    ' --- Refrescar cache de publicabilidad ---
    m_IDEdicion = m_RiesgoMat.IDEdicion
    If m_IDEdicion <> "" Then
        m_Error = ""
        Call InvalidarPublicabilidadPorCambioEvidenciaEdicion(m_IDEdicion, m_Error)
        If m_Error <> "" Then
            logs(logIdx) = "6. InvalidarPublicabilidad warning (no fatal): " & m_Error
            logIdx = logIdx + 1
            ' No bloqueamos por error en invalidacion de cache
        Else
            logs(logIdx) = "6. Cache invalidado para IDEdicion=" & m_IDEdicion
            logIdx = logIdx + 1
        End If
    End If

    logs(logIdx) = "7. Registro exitoso: ParaNC=No"
    logIdx = logIdx + 1
    RegistrarMaterializacionNoNC = Test_Helper.BuildJsonOk("registered-no-nc", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        m_Error = "RegistrarMaterializacionNoNC: " & Err.Number & " - " & Err.description
    End If
    logs(logIdx) = "Error: " & m_Error
    logIdx = logIdx + 1
    RegistrarMaterializacionNoNC = Test_Helper.BuildJsonFail(m_Error, logs)
End Function

' --- Helper #2: RevocarDecisionMaterializacion ---
Public Function RevocarDecisionMaterializacion( _
    ByVal p_IDMaterializacion As String, _
    Optional ByRef db As DAO.Database = Nothing) As String

    ' Resolucion de dependencias
    Dim m_RiesgoMat As RiesgoMaterializacion
    Dim m_Error As String
    Dim m_IDEdicion As String
    Dim logs(0 To 9) As String
    Dim logIdx As Long

    On Error GoTo errores
    m_Error = ""
    logIdx = 0
    logs(logIdx) = "RevocarDecisionMaterializacion: id=" & p_IDMaterializacion
    logIdx = logIdx + 1

    ' --- Resolver materializacion ---
    Set m_RiesgoMat = Constructor.getRiesgoMaterializado(p_IDMaterializacion, m_Error)
    If m_Error <> "" Then
        logs(logIdx) = "2. Constructor.getRiesgoMaterializado error: " & m_Error
        logIdx = logIdx + 1
        RevocarDecisionMaterializacion = Test_Helper.BuildJsonFail(m_Error, logs)
        Exit Function
    End If
    If m_RiesgoMat Is Nothing Then
        m_Error = "No se encontro la materializacion con ID: " & p_IDMaterializacion
        logs(logIdx) = "2. Materializacion Nothing"
        logIdx = logIdx + 1
        RevocarDecisionMaterializacion = Test_Helper.BuildJsonFail(m_Error, logs)
        Exit Function
    End If

    ' --- Validar estado: ParaNC debe estar populado para poder revocar ---
    If Nz(m_RiesgoMat.ParaNC, "") = "" Then
        logs(logIdx) = "3. Estado idempotente: ParaNC vacio (ya por decidir)"
        logIdx = logIdx + 1
        RevocarDecisionMaterializacion = Test_Helper.BuildJsonOk("already-undetermined", logs)
        Exit Function
    End If

    ' --- Revocar decision (llama al metodo del dominio RegistrarPorDecidir) ---
    ' Este metodo setea ParaNC="", FechaDecison="", IDNC=Null
    m_Error = ""
    m_RiesgoMat.RegistrarPorDecidir m_Error
    If m_Error <> "" Then
        logs(logIdx) = "4. RegistrarPorDecidir error: " & m_Error
        logIdx = logIdx + 1
        RevocarDecisionMaterializacion = Test_Helper.BuildJsonFail(m_Error, logs)
        Exit Function
    End If

    ' --- Refrescar cache de publicabilidad ---
    m_IDEdicion = m_RiesgoMat.IDEdicion
    If m_IDEdicion <> "" Then
        m_Error = ""
        Call InvalidarPublicabilidadPorCambioEvidenciaEdicion(m_IDEdicion, m_Error)
        If m_Error <> "" Then
            logs(logIdx) = "5. InvalidarPublicabilidad warning (no fatal): " & m_Error
            logIdx = logIdx + 1
        Else
            logs(logIdx) = "5. Cache invalidado para IDEdicion=" & m_IDEdicion
            logIdx = logIdx + 1
        End If
    End If

    logs(logIdx) = "6. Revocacion exitosa: ParaNC=Null"
    logIdx = logIdx + 1
    RevocarDecisionMaterializacion = Test_Helper.BuildJsonOk("revoked", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        m_Error = "RevocarDecisionMaterializacion: " & Err.Number & " - " & Err.description
    End If
    logs(logIdx) = "Error: " & m_Error
    logIdx = logIdx + 1
    RevocarDecisionMaterializacion = Test_Helper.BuildJsonFail(m_Error, logs)
End Function

' --- Helper #3: EstablecerBotoneraNC (firma corregida 2026-06-19) ---
' NO recibe p_EstadoActual externo; resuelve internamente desde el RiesgoMaterializacion
' Retorna: "SinCalcular" | "PorTomar" | "NoSeGeneraNC" | "VinculadaNC"
Public Function EstablecerBotoneraNC( _
    ByVal p_IDMaterializacion As String, _
    Optional ByRef db As DAO.Database = Nothing, _
    Optional ByRef p_Error As String) As String

    ' Resolucion de dependencias
    Dim m_RiesgoMat As RiesgoMaterializacion
    Dim m_NC As nc
    Dim m_IDEdicion As String

    On Error GoTo errores
    p_Error = ""

    ' --- ID vacio o materializacion no existe > SinCalcular ---
    If p_IDMaterializacion = "" Then
        EstablecerBotoneraNC = "SinCalcular"
        Exit Function
    End If

    ' --- Resolver materializacion ---
    Set m_RiesgoMat = Constructor.getRiesgoMaterializado(p_IDMaterializacion, p_Error)
    If p_Error <> "" Then
        EstablecerBotoneraNC = ""
        Exit Function
    End If
    If m_RiesgoMat Is Nothing Then
        EstablecerBotoneraNC = "SinCalcular"
        Exit Function
    End If

    ' --- Paso 1: EsMaterializacionCalcuado <> EnumSiNo.Sí > SinCalcular ---
    If m_RiesgoMat.EsMaterializacionCalcuado <> EnumSiNo.Sí Then
        EstablecerBotoneraNC = "SinCalcular"
        Exit Function
    End If

    ' --- Paso 2: FechaDecison no es fecha valida > PorTomar ---
    If Not IsDate(m_RiesgoMat.FechaDecison) Then
        EstablecerBotoneraNC = "PorTomar"
        Exit Function
    End If

    ' --- Paso 3: ParaNCCalculado = EnumSiNo.No > NoSeGeneraNC ---
    If m_RiesgoMat.ParaNCCalculado = EnumSiNo.No Then
        EstablecerBotoneraNC = "NoSeGeneraNC"
        Exit Function
    End If

    ' --- Paso 4: NC existe (no es Nothing) > VinculadaNC ---
    Set m_NC = m_RiesgoMat.nc
    If Not m_NC Is Nothing Then
        EstablecerBotoneraNC = "VinculadaNC"
        Exit Function
    End If

    ' Fallback inesperado: ninguno de los casos anteriores aplica
    ' Esto puede ocurrir si ParaNCCalculado = EnumSiNo.Sí pero NC es Nothing
    ' y FechaDecison esta seteada (estado inconsistente pero posible)
    p_Error = "Estado inconsistente: ParaNCCalculado=Si pero NC es Nothing para id=" & p_IDMaterializacion
    EstablecerBotoneraNC = ""
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "EstablecerBotoneraNC: " & Err.Number & " - " & Err.description
    End If
    EstablecerBotoneraNC = ""
End Function

