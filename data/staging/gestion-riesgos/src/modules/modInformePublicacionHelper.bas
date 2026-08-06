Attribute VB_Name = "modInformePublicacionHelper"
Option Compare Database
Option Explicit

' =============================================================================
' modInformePublicacionHelper.bas
'
' Helper module para validación y generación de informe de publicación.
' Project: gestion_riesgos
' Branch: feat/metodologia-e2e-riesgos-2026-06-19
' Refactor: 2026-06-19
'
' Helper público (1):
'   ValidarYGenerarInformeEdicion - valida priorización + genera HTML/PDF
'
' CORRECCIÓN 2026-06-19: la firma del audit original no tenía
' Optional ByRef p_Error As String. Se agrega como último parámetro.
' EnumControlCambiosAlcance: Resumen3=1, Completo=2 (verificado).
' =============================================================================

' =============================================================================
' Helper: ValidarYGenerarInformeEdicion
'
' Recibe un IDEdicion, resuelve la edición, valida priorización de riesgos,
' valida el alcance del control de cambios, genera el informe HTML y retorna
' la URL del informe generado.
'
' Parámetros:
'   p_IDEdicion          - ID de la edición a procesar
'   p_Alcance            - Alcance del control de cambios (Resumen3 o Completo)
'   p_URLInforme (out)   - URL del informe HTML generado
'   db                   - DAO.Database opcional (para inyección en tests)
'   p_Error (out)        - Descripción de error si falla la validación o generación
'
' Retorna:
'   JSON con {ok, value: URL, error: null} si OK
'   JSON con {ok: false, error: msg} si falla
'
' Validaciones:
'   1. La edición debe existir y tener todos sus riesgos priorizados
'   2. El alcance debe ser Resumen3 (1) o Completo (2)
'
' Escalado:
'   Constructor.getEdicion(p_IDEdicion, p_Error) > para obtener la edición
'   Edicion.TodosLosRiesgosPriorizados > para validar priorización
'   InformeRiesgoHTML.GenerarInformeEdicionHTML > para generar el informe
' =============================================================================
Public Function ValidarYGenerarInformeEdicion( _
    ByVal p_IDEdicion As String, _
    ByVal p_Alcance As EnumControlCambiosAlcance, _
    Optional ByRef p_URLInforme As String, _
    Optional ByRef db As DAO.Database = Nothing, _
    Optional ByRef p_Error As String) As String

    Dim m_Edicion As Edicion
    Dim m_URL As String
    Dim m_Logs(0) As String

    On Error GoTo errores
    p_Error = ""
    p_URLInforme = ""
    m_Logs(0) = "1. Resolver edición IDEdicion=" & CStr(p_IDEdicion)

    ' 1. Resolver la edición
    Set m_Edicion = Constructor.getEdicion(p_IDEdicion:=CLng(p_IDEdicion), p_Error:=p_Error)
    If p_Error <> "" Then
        m_Logs(0) = m_Logs(0) & " > falla"
        ValidarYGenerarInformeEdicion = Test_Helper.BuildJsonFail(p_Error, m_Logs)
        Exit Function
    End If
    If m_Edicion Is Nothing Then
        p_Error = "No se encontró la edición con IDEdicion=" & CStr(p_IDEdicion)
        m_Logs(0) = m_Logs(0) & " > no encontrada"
        ValidarYGenerarInformeEdicion = Test_Helper.BuildJsonFail(p_Error, m_Logs)
        Exit Function
    End If
    m_Logs(0) = m_Logs(0) & " > ok"

    ' 2. Validar que todos los riesgos estón priorizados
    If m_Edicion.TodosLosRiesgosPriorizados <> EnumSiNo.Sí Then
        p_Error = "Los riesgos deben estar priorizados"
        m_Logs(0) = "2. Validación priorización > " & p_Error
        ValidarYGenerarInformeEdicion = Test_Helper.BuildJsonFail(p_Error, m_Logs)
        Exit Function
    End If

    ' 3. Validar que el alcance sea válido (solo Resumen3=1 o Completo=2)
    If p_Alcance <> EnumControlCambiosAlcanceResumen3 And _
       p_Alcance <> EnumControlCambiosAlcanceCompleto Then
        p_Error = "Alcance inválido"
        m_Logs(0) = "3. Validación alcance > " & p_Error & " (valor=" & CStr(p_Alcance) & ")"
        ValidarYGenerarInformeEdicion = Test_Helper.BuildJsonFail(p_Error, m_Logs)
        Exit Function
    End If

    ' 4. Generar el informe HTML (siempre PDF: la publicación es HTML+PDF)
    m_Logs(0) = "4. Generar informe HTML con alcance=" & CStr(p_Alcance)
    m_URL = InformeRiesgoHTML.GenerarInformeEdicionHTML( _
                p_Edicion:=m_Edicion, _
                p_hWnd:=0, _
                p_FechaCierre:="", _
                p_FechaPublicacion:="", _
                p_Error:=p_Error, _
                p_GenerarPDF:=True, _
                p_ControlCambiosAlcance:=p_Alcance)

    If p_Error <> "" Then
        m_Logs(0) = m_Logs(0) & " > falla: " & p_Error
        ValidarYGenerarInformeEdicion = Test_Helper.BuildJsonFail(p_Error, m_Logs)
        Exit Function
    End If

    ' 5. Asignar la URL al parámetro de salida
    p_URLInforme = m_URL
    m_Logs(0) = "5. Informe generado > " & m_URL

    ' 6. Retornar JSON de éxito
    ValidarYGenerarInformeEdicion = Test_Helper.BuildJsonOk(m_URL, m_Logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "ValidarYGenerarInformeEdicion: error " & Err.Number & " - " & Err.description
    End If
    m_Logs(0) = "Error: " & p_Error
    ValidarYGenerarInformeEdicion = Test_Helper.BuildJsonFail(p_Error, m_Logs)
End Function

' =============================================================================
' Helper: ObtenerMotivoNoPublicable
'
' Devuelve el motivo persistido por el cual un riesgo no es publicable,
' leído desde tbCambiosParaPublicacion (filtro NombreCampo="MotivoNoPublicable"
' y Riesgo=p_IDRiesgo, ordenado por FechaRegistro DESC LIMIT 1).
' Aplica HTMLSafe al valor para que el reporte lo pueda pintar sin
' inyección de tags. Si no hay motivo persistido, devuelve
' "sin motivo registrado" (el form lo muestra tal cual en la tarjeta).
'
' Decisión de fuente (2026-06-22, sdd-apply Bloque 3 / PR-3):
'   * TbRiesgos no tiene campo MotivoNoPublicable (verificado vía
'     dysflow_get_schema table=TbRiesgos).
'   * TbRiesgosNC, tbCambios y tbCambiosParaPublicacion son las candidatas
'     que cargan descripciones de cambios a nivel de riesgo.
'   * tbCambiosParaPublicacion es la tabla designada por la arquitectura
'     para cambios "para publicación" — tiene Descripcion (Memo) y
'     Riesgo (Text) como columnas naturales para guardar el motivo.
'   * Si en el futuro Calidad decide moverlo a TbRiesgos.MotivoNoPublicable,
'     este helper se ajusta sin tocar la firma.
'
' Parámetros:
'   p_IDRiesgo  - ID del riesgo (String por convención del proyecto).
'   db          - DAO.Database opcional (inyectado por el átomo TDD).
'   p_Error (out) - String, ByRef. Descripción del error si falla la consulta
'                    o el ID es inválido; cadena vacía si OK.
'
' Retorna:
'   Motivo literal (HTMLSafe'd) si existe en tbCambiosParaPublicacion.
'   "sin motivo registrado" si no hay motivo persistido o si el motivo
'   quedó en blanco.
'   "" si p_Error está poblado (validación fallida antes de consultar).
' =============================================================================
Public Function ObtenerMotivoNoPublicable( _
    ByRef p_IDRiesgo As String, _
    Optional ByRef db As DAO.Database = Nothing, _
    Optional ByRef p_Error As String) As String

    Dim m_Db As DAO.Database
    Dim m_Rs As DAO.Recordset
    Dim m_Motivo As String

    Const PLACEHOLDER_SIN_MOTIVO As String = "sin motivo registrado"

    On Error GoTo errores
    p_Error = ""
    ObtenerMotivoNoPublicable = ""

    ' --- 1. Validación: p_IDRiesgo no vacío ---
    If Len(Trim$(Nz(p_IDRiesgo, ""))) = 0 Then
        p_Error = "ObtenerMotivoNoPublicable: p_IDRiesgo está vacío"
        Exit Function
    End If

    ' --- 2. Resolver db (inyectado por el átomo, o CurrentDb en producción) ---
    If db Is Nothing Then
        Set m_Db = CurrentDb
    Else
        Set m_Db = db
    End If

    ' --- 3. Consultar tbCambiosParaPublicacion ---
    '     Top-1 más reciente por FechaRegistro DESC, IdCambio DESC (desempate
    '     determinista por si varios registros en el mismo timestamp).
    '     Tabla: tbCambiosParaPublicacion (per publica risk-level changes).
    '     Columna NombreCampo actúa como discriminador semántico ("MotivoNoPublicable"
    '     vs otros motivos que Calidad pueda persistir en el futuro).
    Dim m_SQL As String
    m_SQL = "SELECT TOP 1 [Descripcion] " & _
            "FROM tbCambiosParaPublicacion " & _
            "WHERE [Riesgo]='" & Test_Helper.SqlStr(p_IDRiesgo) & "' " & _
            "  AND [NombreCampo]='MotivoNoPublicable' " & _
            "ORDER BY [FechaRegistro] DESC, [IDCambio] DESC"

    Set m_Rs = m_Db.OpenRecordset(m_SQL, dbOpenSnapshot)
    If m_Rs.EOF Then
        m_Rs.Close
        Set m_Rs = Nothing
        Set m_Db = Nothing
        ObtenerMotivoNoPublicable = PLACEHOLDER_SIN_MOTIVO
        Exit Function
    End If

    m_Motivo = Nz(m_Rs.fields("Descripcion").value, "")
    m_Rs.Close
    Set m_Rs = Nothing

    ' --- 4. Si está vacío, devolver placeholder ---
    If Len(Trim$(m_Motivo)) = 0 Then
        Set m_Db = Nothing
        ObtenerMotivoNoPublicable = PLACEHOLDER_SIN_MOTIVO
        Exit Function
    End If

    ' --- 5. Aplicar HTMLSafe para render seguro en el reporte ---
    Set m_Db = Nothing
    ObtenerMotivoNoPublicable = HTMLSafe(m_Motivo)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "ObtenerMotivoNoPublicable: " & Err.Number & " - " & Err.description
    Else
        p_Error = "ObtenerMotivoNoPublicable: " & p_Error
    End If
    ObtenerMotivoNoPublicable = ""
End Function

' =============================================================================
' Helper: ObtenerDetallePorRiesgo
'
' Devuelve un bloque JSON con el detalle de publicabilidad de un riesgo individual:
'   esPublicable              - boolean: true si el riesgo puede incluirse en el
'                               informe de publicación sin más bloqueo.
'   motivo                    - string: razón consolidada del estado. Si el
'                               riesgo es no publicable, contiene la causa
'                               concreta (motivo persistido + decisiones de NC
'                               pendientes). Si es publicable, contiene un
'                               resumen afirmativo.
'   tieneMaterializaciones    - boolean: true si TbRiesgosMaterializaciones
'                               tiene al menos una fila para este riesgo.
'   cantidadMaterializaciones - number: total de materializaciones del riesgo
'                               (incluye las que están pendientes y las que ya
'                               tienen decisión de NC).
'
' Decisión de fuente (2026-06-22, sdd-apply Bloque 3 / PR-4):
'   * TbRiesgosMaterializaciones NO tiene columna IDRiesgo (verificado via
'     Constructor.bas SELECTs — usa IDProyecto + CodigoRiesgo como natural key).
'     Por eso el helper hace JOIN con TbRiesgos para resolver IDProyecto +
'     CodigoRiesgo antes de consultar las materializaciones.
'   * tbCambiosParaPublicacion se consulta igual que en ObtenerMotivoNoPublicable
'     (misma query: filtro NombreCampo='MotivoNoPublicable', Riesgo=p_IDRiesgo,
'     ORDER BY FechaRegistro DESC, IDCambio DESC LIMIT 1).
'   * Veredicto de publicabilidad:
'       - Si hay motivo persistido en tbCambiosParaPublicacion ? no publicable.
'       - Si hay materializaciones y al menos una con ParaNC vacío/Null ? no
'         publicable (decisión de NC pendiente).
'       - En cualquier otro caso (sin motivo, sin materializaciones, o todas las
'         materializaciones con ParaNC decidido) ? publicable.
'
' Autorización (simplificación documentada, fuera de scope SDD):
'   El helper exige que m_ObjUsuarioConectado.UsuarioRed esté en
'   m_ObjEntorno.ColUsuariosCalidad.Exists(...). Si alguna de las dos globales
'   es Nothing, el deny es conservador. La capa completa de permisos (roles
'   finos por proyecto, RAC, administradores con override) queda fuera de scope
'   — el caller del form es responsable de aplicar la lógica de rol definitiva.
'
' Parámetros:
'   p_IDRiesgo       - ID del riesgo (String por convención del proyecto). Si
'                      es numérico se valida contra TbRiesgos.IDRiesgo.
'   db               - DAO.Database opcional (inyectado por el átomo TDD).
'   p_PromptResult   - Long, ByRef. Reservado para futura migración del MsgBox
'                      del form que muestra el detalle por riesgo. No se usa
'                      en este helper (la consulta es no interactiva).
'   p_Error (out)    - String, ByRef. Cadena de error si falla la validación,
'                      la autorización o el DAO. Cadena vacía si OK.
'
' Retorna:
'   JSON con {esPublicable, motivo, tieneMaterializaciones,
'             cantidadMaterializaciones} si OK.
'   "" si p_Error está poblado (validación / autorización / DAO fallido).
' =============================================================================
Public Function ObtenerDetallePorRiesgo( _
    ByRef p_IDRiesgo As String, _
    Optional ByRef db As DAO.Database = Nothing, _
    Optional ByRef p_PromptResult As Long, _
    Optional ByRef p_Error As String) As String

    Dim m_Db As DAO.Database
    Dim m_Rs As DAO.Recordset
    Dim m_SQL As String
    Dim m_LngRiesgo As Long

    Dim m_IDProyecto As Long
    Dim m_IDEdicion As Long
    Dim m_CodigoRiesgo As String

    Dim m_CantidadMat As Long
    Dim m_ConNC As Long
    Dim m_Pendientes As Long
    Dim m_ParaNC As String
    Dim m_EstadoMat As String
    Dim m_HayEstadoProblematico As Boolean

    Dim m_MotivoPersistido As String

    Dim m_EsPublicable As Boolean
    Dim m_MotivoOut As String
    Dim m_JSON As String

    On Error GoTo errores
    p_Error = ""
    ObtenerDetallePorRiesgo = ""

    ' --- 1. Validación: p_IDRiesgo no vacío ---
    If Len(Trim$(Nz(p_IDRiesgo, ""))) = 0 Then
        p_Error = "ObtenerDetallePorRiesgo: p_IDRiesgo está vacío"
        Exit Function
    End If
    If Not IsNumeric(p_IDRiesgo) Then
        p_Error = "ObtenerDetallePorRiesgo: p_IDRiesgo no es numérico"
        Exit Function
    End If
    m_LngRiesgo = CLng(p_IDRiesgo)

    ' --- 2. Resolver db (inyectado por el átomo, o CurrentDb en producción) ---
    If db Is Nothing Then
        Set m_Db = CurrentDb
    Else
        Set m_Db = db
    End If

    ' --- 3. Authorization guard (simplificación — ver header del helper) ---
    '     Deny conservador si cualquier global requerida falta o si el usuario
    '     no está en la colección de Calidad.
    If (m_ObjEntorno Is Nothing) Then
        Set m_Db = Nothing
        p_Error = "ObtenerDetallePorRiesgo: el usuario no tiene permisos para ver el detalle del riesgo"
        Exit Function
    End If
    If (m_ObjEntorno.ColUsuariosCalidad Is Nothing) Then
        Set m_Db = Nothing
        p_Error = "ObtenerDetallePorRiesgo: el usuario no tiene permisos para ver el detalle del riesgo"
        Exit Function
    End If
    If (m_ObjUsuarioConectado Is Nothing) Then
        Set m_Db = Nothing
        p_Error = "ObtenerDetallePorRiesgo: el usuario no tiene permisos para ver el detalle del riesgo"
        Exit Function
    End If
    If Not m_ObjEntorno.ColUsuariosCalidad.Exists(m_ObjUsuarioConectado.UsuarioRed) Then
        Set m_Db = Nothing
        p_Error = "ObtenerDetallePorRiesgo: el usuario no tiene permisos para ver el detalle del riesgo"
        Exit Function
    End If

    ' --- 4. Risk exists check: query TbRiesgos ---
    m_SQL = "SELECT IDRiesgo, IDEdicion, IDProyecto, CodigoRiesgo " & _
            "FROM TbRiesgos " & _
            "WHERE IDRiesgo=" & m_LngRiesgo
    Set m_Rs = m_Db.OpenRecordset(m_SQL, dbOpenSnapshot)
    If m_Rs.EOF Then
        m_Rs.Close
        Set m_Rs = Nothing
        Set m_Db = Nothing
        p_Error = "ObtenerDetallePorRiesgo: no se encontró el riesgo"
        Exit Function
    End If
    m_IDProyecto = CLng(Nz(m_Rs.fields("IDProyecto").value, 0))
    m_IDEdicion = CLng(Nz(m_Rs.fields("IDEdicion").value, 0))
    m_CodigoRiesgo = CStr(Nz(m_Rs.fields("CodigoRiesgo").value, ""))
    m_Rs.Close
    Set m_Rs = Nothing

    ' --- 5. Materializaciones: query TbRiesgosMaterializaciones ---
    '     SQL documentado: filtro por (IDProyecto, CodigoRiesgo) porque la tabla
    '     no tiene columna IDRiesgo (ver Constructor.bas SELECTs).
    '     Ordenamos por Fecha DESC para que el helper pueda reportar el último
    '     estado primero si el caller lo necesita en el futuro.
    m_SQL = "SELECT Fecha, Estado, IDPlanContingencia, IDNC, ParaNC " & _
            "FROM TbRiesgosMaterializaciones " & _
            "WHERE IDProyecto=" & m_IDProyecto & " " & _
            "  AND CodigoRiesgo='" & Test_Helper.SqlStr(m_CodigoRiesgo) & "' " & _
            "ORDER BY Fecha DESC"
    Set m_Rs = m_Db.OpenRecordset(m_SQL, dbOpenSnapshot)

    m_CantidadMat = 0
    m_ConNC = 0
    m_Pendientes = 0
    m_HayEstadoProblematico = False
    Do While Not m_Rs.EOF
        m_CantidadMat = m_CantidadMat + 1
        m_ParaNC = Trim$(CStr(Nz(m_Rs.fields("ParaNC").value, "")))
        m_EstadoMat = Trim$(CStr(Nz(m_Rs.fields("Estado").value, "")))
        If Len(m_ParaNC) = 0 Then
            ' Decisión de NC aún no tomada ? pendiente.
            m_Pendientes = m_Pendientes + 1
        Else
            m_ConNC = m_ConNC + 1
        End If
        ' Heurística de estado problemático: si Estado es algo distinto de
        ' "Resuelto"/"Cerrado" (estados canónicos de la app), marcamos como
        ' bloqueante. Documentado en el helper: la lista canónica de estados
        ' "OK" la mantiene Calidad; aquí solo rechazamos cadenas vacías.
        If Len(m_EstadoMat) = 0 Then
            m_HayEstadoProblematico = True
        End If
        m_Rs.MoveNext
    Loop
    m_Rs.Close
    Set m_Rs = Nothing

    ' --- 6. MotivoNoPublicable: misma query que ObtenerMotivoNoPublicable ---
    m_SQL = "SELECT TOP 1 [Descripcion] " & _
            "FROM tbCambiosParaPublicacion " & _
            "WHERE [Riesgo]='" & Test_Helper.SqlStr(CStr(m_LngRiesgo)) & "' " & _
            "  AND [NombreCampo]='MotivoNoPublicable' " & _
            "ORDER BY [FechaRegistro] DESC, [IDCambio] DESC"
    Set m_Rs = m_Db.OpenRecordset(m_SQL, dbOpenSnapshot)
    If m_Rs.EOF Then
        m_MotivoPersistido = ""
    Else
        m_MotivoPersistido = CStr(Nz(m_Rs.fields("Descripcion").value, ""))
    End If
    m_Rs.Close
    Set m_Rs = Nothing

    ' --- 7. Veredicto consolidado de publicabilidad ---
    '     Reglas (en orden de prioridad):
    '       a) Si hay motivo persistido ? no publicable (motivo manda).
    '       b) Si hay materializaciones pendientes de decisión NC ? no publicable.
    '       c) Si hay alguna materialización con estado vacío (Estado problemático)
    '          ? no publicable.
    '       d) En cualquier otro caso ? publicable.
    If Len(Trim$(m_MotivoPersistido)) > 0 Then
        m_EsPublicable = False
        m_MotivoOut = "No publicable — motivo: " & HTMLSafe(m_MotivoPersistido)
    ElseIf m_Pendientes > 0 Then
        m_EsPublicable = False
        m_MotivoOut = "No publicable — hay " & CStr(m_Pendientes) & _
                      " materializacion(es) pendiente(s) de decisión de NC " & _
                      "(total=" & CStr(m_CantidadMat) & _
                      ", con NC decidida=" & CStr(m_ConNC) & ")"
    ElseIf m_HayEstadoProblematico Then
        m_EsPublicable = False
        m_MotivoOut = "No publicable — alguna materialización tiene Estado vacío " & _
                      "(total=" & CStr(m_CantidadMat) & ")"
    ElseIf m_CantidadMat = 0 Then
        m_EsPublicable = True
        m_MotivoOut = "Publicable — sin riesgos materializados para este riesgo"
    Else
        m_EsPublicable = True
        m_MotivoOut = "Publicable — las " & CStr(m_CantidadMat) & _
                      " materializacion(es) tienen decisión de NC tomada"
    End If

    ' --- 8. Componer JSON ---
    '     BuildJsonOk-style: ok/value/payload/error/logs envelope NO aplica acá
    '     (el caller es un form, no el runner TDD). Devolvemos JSON crudo con
    '     campos serializados; usamos Test_Helper.EscapeJsonString para
    '     user-controlled fields (motivo).
    m_JSON = "{""esPublicable"":" & LCase$(CStr(m_EsPublicable)) & _
             ",""motivo"":""" & Test_Helper.EscapeJsonString(m_MotivoOut) & """" & _
             ",""tieneMaterializaciones"":" & LCase$(CStr(m_CantidadMat > 0)) & _
             ",""cantidadMaterializaciones"":" & CStr(m_CantidadMat) & "}"

    Set m_Db = Nothing
    ObtenerDetallePorRiesgo = m_JSON
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "ObtenerDetallePorRiesgo: " & Err.Number & " - " & Err.description
    End If
    ObtenerDetallePorRiesgo = ""
End Function

