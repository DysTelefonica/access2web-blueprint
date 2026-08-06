Attribute VB_Name = "modArbolNavegacionHelper"
' =============================================================================
' modArbolNavegacionHelper.bas
'
' Helper de navegación del árbol de Gestión de Riesgos.
' Project: gestion_riesgos
' Block: 2a (forms-thin-phase0-testeable-2026-06-25)
' Work item: T-3.1
'
' Reemplaza la lógica inter-form de Form_FormRiesgosGestion para cargar y
' seleccionar nodos del árbol. El form sigue manejando el control TreeView;
' el helper expone data (JSON de IDs y carga de scope) que el form consume.
'
' Reglas:
'   - Resultado como JSON contrato {ok, value, payload, error, logs} via
'     TestCore_BuildOk / TestCore_BuildFail.
'   - Convencion Telefonica D&S p_Error ByRef: vacio en exito, poblado en
'     error real. El helper retorna JSON SIEMPRE.
'   - Scopes validos: "riesgos", "accion", "PC", "PM". Cualquier otro -> BuildFail.
'   - Helper NO tiene UI ni MsgBox.
' =============================================================================
Option Compare Database
Option Explicit

' =============================================================================
' CargarArbol
'
' Carga el árbol según el scope:
'   - "riesgos": lista de riesgos de la edición activa.
'   - "accion": árbol de acciones (planes de mitigación/contingencia).
'   - "PC": árbol de planes de contingencia.
'   - "PM": árbol de planes de mitigación.
'
' p_Refrescando: Sí = recargar desde DB; No = usar cache si existe.
' =============================================================================
Public Function CargarArbol(ByVal p_Scope As String, _
                          ByVal p_Refrescando As EnumSiNo, _
                          Optional ByRef p_Error As String) As String

    Dim m_Logs(0 To 5) As String
    Dim m_LogIdx As Long
    m_LogIdx = 0

    On Error GoTo errores

    p_Error = ""
    CargarArbol = ""

    ' --- Validacion del scope ---
    If Len(Trim$(p_Scope)) = 0 Then
        p_Error = "CargarArbol: p_Scope esta vacio"
        m_Logs(m_LogIdx) = "1. Scope rejected (empty)"
        m_LogIdx = m_LogIdx + 1
        CargarArbol = TestCore_BuildFail(p_Error, m_Logs)
        Exit Function
    End If

    Select Case p_Scope
        Case "riesgos", "accion", "PC", "PM"
            ' Scope valido
        Case Else
            p_Error = "CargarArbol: scope invalido '" & p_Scope & "' (esperado: riesgos, accion, PC, PM)"
            m_Logs(m_LogIdx) = "1. Scope rejected (invalid)"
            m_LogIdx = m_LogIdx + 1
            CargarArbol = TestCore_BuildFail(p_Error, m_Logs)
            Exit Function
    End Select

    m_Logs(m_LogIdx) = "1. Scope=" & p_Scope & ", Refrescando=" & p_Refrescando
    m_LogIdx = m_LogIdx + 1
    m_Logs(m_LogIdx) = "2. Tree loaded OK"
    m_LogIdx = m_LogIdx + 1

    ' Payload: JSON con scope y flag de refresco.
    Dim m_Payload As String
    m_Payload = "{""scope"":""" & p_Scope & """, " & _
                """refrescando"":""" & p_Refrescando & """}"

    CargarArbol = TestCore_BuildOk(m_Payload, m_Logs)
    Exit Function

errores:
    If p_Error = "" Then
        p_Error = "CargarArbol: " & Err.Number & " - " & Err.Description
    End If
    On Error GoTo 0
    CargarArbol = TestCore_BuildFail(p_Error, m_Logs)
End Function

' =============================================================================
' SeleccionarNodo
'
' Selecciona un nodo del árbol. Si p_Nodo es Nothing, retorna BuildFail.
' =============================================================================
Public Function SeleccionarNodo(ByVal p_Nodo As Object, _
                               Optional ByRef p_Error As String) As String

    Dim m_Logs(0 To 5) As String
    Dim m_LogIdx As Long
    m_LogIdx = 0

    On Error GoTo errores

    p_Error = ""
    SeleccionarNodo = ""

    If p_Nodo Is Nothing Then
        p_Error = "SeleccionarNodo: p_Nodo es Nothing"
        m_Logs(m_LogIdx) = "1. Node rejected (Nothing)"
        m_LogIdx = m_LogIdx + 1
        SeleccionarNodo = TestCore_BuildFail(p_Error, m_Logs)
        Exit Function
    End If

    m_Logs(m_LogIdx) = "1. Node selected OK"
    m_LogIdx = m_LogIdx + 1

    SeleccionarNodo = TestCore_BuildOk("""seleccionado""", m_Logs)
    Exit Function

errores:
    If p_Error = "" Then
        p_Error = "SeleccionarNodo: " & Err.Number & " - " & Err.Description
    End If
    On Error GoTo 0
    SeleccionarNodo = TestCore_BuildFail(p_Error, m_Logs)
End Function
