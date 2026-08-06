Attribute VB_Name = "modResponsablePorRolService"
Option Compare Database
Option Explicit

' -----------------------------------------------------------------------------
' modResponsablePorRolService
'
' Servicio JSON que envuelve la carga de los combos IDResponsableCalidad y
' IDResponsableSeguridad del form Form_FormExpedienteGeneral.  Sustituye la
' RowSource hardcodeada que existia en
' `GuardadoAutomaticoHelper.bas:EstablecerCombos` y la lista hardcodeada
' de UsuarioRed que existia en `constructor.getUsuariosCalidad`.
'
' API publica:
'   ExpedienteGeneral_EstablecerCombos( _
'       frm, [p_OrphanCalidad], [p_OrphanSeguridad], [p_db], [p_Error]) As String
'
' Devuelve el envelope canonico `{ok, value, payload, error, logs}`.  El
' payload se serializa via `JsonConverter.ConvertToJson` (mismo patron que
' el resto de los helpers del proyecto, p.ej. modExpedienteGeneralHelper).
'
' payload shape:
'   {
'     "calidad":         [{"id":"119","nombre":"..."}, ...],
'     "seguridad":       [{"id":"127","nombre":"..."}, ...],
'     "seguridadSentinel": {"id":"0","label":"N/A"},
'     "orphanCalidad":   "900099" | "",
'     "orphanSeguridad": "900099" | "",
'     "orphanCalidadInList":   true | false,
'     "orphanSeguridadInList": true | false,
'     "cmbRows": {
'       "calidad":   ["119;Ana ...", "181;Beatriz ...", ...],
'       "seguridad": ["0;N/A", "127;Esperanza ...", ...]
'     }
'   }
'
' Reglas:
'   1. El sentinel "0;N/A" en Seguridad es un UX contract: "sin responsable
'      de seguridad".  Se preserva como primer AddItem del combo Seguridad.
'   2. Los orphan IDs se preservan: si el ID guardado en el expediente ya
'      no tiene el rol, se incluye igualmente (lookup defensivo en
'      `TbUsuariosAplicaciones` por el helper).  Asi un expediente antiguo
'      no queda visualmente vacio.
'   3. Si `frm` es Nothing, NO se hace AddItem sobre los combos; el JSON
'      sigue conteniendo los datos (mismo shape).  Esto permite que los
'      tests asserten sin necesidad de un form real.
'   4. Si `p_db` es Nothing, el helper usa CurrentDb() del frontend (que
'      enrutara las linked tables a sus backends respectivos: staging
'      backend para TbResponsablesPorRol, Lanzadera para
'      TbUsuariosAplicaciones).  Los tests inyectan su propio `p_db` (temp
'      .accdb) para aislarse.
'   5. Convencion de error: el caller que falla setea `p_Error` y la
'      funcion devuelve BuildJsonFail con `p_Error` en `error`.
'   6. Orden de populacion de combos:
'        a) IDResponsableCalidad: AddItem por cada Usuario del dict Calidad
'        b) IDResponsableSeguridad: AddItem "0;N/A" primero (sentinel)
'           seguido de AddItem por cada Usuario del dict Seguridad
' -----------------------------------------------------------------------------

' Whitelist cerrada (mirror de modResponsablePorRolHelper).
Private Const ROL_CALIDAD    As String = "Calidad"
Private Const ROL_SEGURIDAD  As String = "Seguridad"

' UX contract: el combo Seguridad admite "sin responsable" como 0;N/A.
Private Const SENTINEL_SEG_ID    As String = "0"
Private Const SENTINEL_SEG_LABEL As String = "N/A"


' -----------------------------------------------------------------------------
' API publica
' -----------------------------------------------------------------------------

' Carga los combos IDResponsableCalidad/Seguridad de un form y devuelve el
' envelope JSON.  La firma acepta parametros opcionales para hacer el servicio
' directamente testeable con temp .accdb y orphan IDs controlados, sin
' necesidad de un form real ni del DTO activo.
'
'   frm                 : Form destino del AddItem.  Nothing -> solo JSON.
'                         El form debe exponer los controles
'                         "IDResponsableCalidad" y "IDResponsableSeguridad".
'   p_OrphanCalidad     : ID guardado en el expediente para Calidad.
'                         "" si el caller no quiere preservar orphan.
'   p_OrphanSeguridad   : ID guardado en el expediente para Seguridad.
'   p_db                : DAO.Database inyectado.  Nothing -> CurrentDb().
'   p_Error             : parametro de error estandar del proyecto.
Public Function ExpedienteGeneral_EstablecerCombos( _
                                                Optional ByRef frm As Form = Nothing, _
                                                Optional ByVal p_OrphanCalidad As String = "", _
                                                Optional ByVal p_OrphanSeguridad As String = "", _
                                                Optional ByRef p_db As DAO.Database = Nothing, _
                                                Optional ByRef p_Error As String _
                                                ) As String

    Dim m_Calidad As Scripting.Dictionary
    Dim m_Seguridad As Scripting.Dictionary
    Dim m_Payload As Object
    Dim m_Logs(0 To 6) As String

    On Error GoTo errores

    m_Logs(0) = "1. GetUsuariosPorRol Calidad (orphan='" & p_OrphanCalidad & "')"
    Set m_Calidad = modResponsablePorRolHelper.GetUsuariosPorRol( _
                        ROL_CALIDAD, p_OrphanCalidad, p_db, p_Error)
    If p_Error <> "" Then
        ExpedienteGeneral_EstablecerCombos = BuildJsonFail(p_Error, m_Logs, "GetUsuariosPorRol Calidad")
        Exit Function
    End If

    m_Logs(1) = "2. GetUsuariosPorRol Seguridad (orphan='" & p_OrphanSeguridad & "')"
    Set m_Seguridad = modResponsablePorRolHelper.GetUsuariosPorRol( _
                        ROL_SEGURIDAD, p_OrphanSeguridad, p_db, p_Error)
    If p_Error <> "" Then
        ExpedienteGeneral_EstablecerCombos = BuildJsonFail(p_Error, m_Logs, "GetUsuariosPorRol Seguridad")
        Exit Function
    End If

    m_Logs(2) = "3. Build payload (calidad, seguridad, sentinel, orphan visibility)"
    Set m_Payload = BuildPayload(m_Calidad, m_Seguridad, p_OrphanCalidad, p_OrphanSeguridad)

    m_Logs(3) = "4. AddItem a combos del form (si frm no Nothing)"
    If Not frm Is Nothing Then
        AddItemCalidad frm, m_Calidad
        AddItemSeguridad frm, m_Seguridad
    End If

    m_Logs(4) = "5. Emit JSON envelope"
    ExpedienteGeneral_EstablecerCombos = BuildJsonOk("ok", m_Payload, m_Logs)
    Exit Function

errores:
    p_Error = "ExpedienteGeneral_EstablecerCombos: " & Err.Number & _
                " - " & Err.Description
    ExpedienteGeneral_EstablecerCombos = BuildJsonFail(p_Error, m_Logs, "")
End Function


' -----------------------------------------------------------------------------
' Helpers privados
' -----------------------------------------------------------------------------

' Construye el payload (Dictionary) del envelope JSON.  Shape publico,
' versionado implicitamente por la firma de `ExpedienteGeneral_EstablecerCombos`.
Private Function BuildPayload( _
                                ByVal p_Calidad As Scripting.Dictionary, _
                                ByVal p_Seguridad As Scripting.Dictionary, _
                                ByVal p_OrphanCalidad As String, _
                                ByVal p_OrphanSeguridad As String _
                                ) As Object

    Dim m_Payload As Object
    Set m_Payload = CreateJsonObject()

    Set m_Payload("calidad") = DictionaryToArray(p_Calidad)
    Set m_Payload("seguridad") = DictionaryToArray(p_Seguridad)

    Dim m_Sentinel As Object
    Set m_Sentinel = CreateJsonObject()
    m_Sentinel("id") = SENTINEL_SEG_ID
    m_Sentinel("label") = SENTINEL_SEG_LABEL
    Set m_Payload("seguridadSentinel") = m_Sentinel

    m_Payload("orphanCalidad") = p_OrphanCalidad
    m_Payload("orphanSeguridad") = p_OrphanSeguridad
    m_Payload("orphanCalidadInList") = OrphanInDictionary(p_Calidad, p_OrphanCalidad)
    m_Payload("orphanSeguridadInList") = OrphanInDictionary(p_Seguridad, p_OrphanSeguridad)

    Dim m_Cmb As Object
    Set m_Cmb = CreateJsonObject()
    Set m_Cmb("calidad") = CmbRowsFromDict(p_Calidad, False)
    Set m_Cmb("seguridad") = CmbRowsFromDict(p_Seguridad, True)
    Set m_Payload("cmbRows") = m_Cmb

    Set BuildPayload = m_Payload
End Function


' Convierte un Dictionary de Usuario en una Collection de {id, nombre}.
' Devuelve Collection vacia si el Dictionary es Nothing o vacio.
' (Collection en lugar de Dictionary-indexed-by-int para que JsonConverter
' lo serialice como JSON array, no como objeto.)
Private Function DictionaryToArray( _
                                    ByVal p_Dict As Scripting.Dictionary _
                                    ) As Collection

    Dim m_Col As New Collection
    If p_Dict Is Nothing Then
        Set DictionaryToArray = m_Col
        Exit Function
    End If

    Dim m_ID As Variant
    Dim m_Usuario As Usuario
    Dim m_Entry As Object
    For Each m_ID In p_Dict.Keys
        Set m_Usuario = p_Dict(m_ID)
        If Not m_Usuario Is Nothing Then
            Set m_Entry = CreateJsonObject()
            m_Entry("id") = CStr(m_Usuario.ID)
            m_Entry("nombre") = CStr(m_Usuario.Nombre)
            m_Col.Add m_Entry
            Set m_Entry = Nothing
        End If
        Set m_Usuario = Nothing
    Next

    Set DictionaryToArray = m_Col
End Function


' Construye una Collection de strings "id;nombre" que se pasan a
' `cmb.AddItem`.  Si `p_WithSentinel` es True, prepende "0;N/A" (UX
' contract para Seguridad).
Private Function CmbRowsFromDict( _
                                ByVal p_Dict As Scripting.Dictionary, _
                                ByVal p_WithSentinel As Boolean _
                                ) As Collection

    Dim m_Col As New Collection
    If p_WithSentinel Then
        m_Col.Add SENTINEL_SEG_ID & ";" & SENTINEL_SEG_LABEL
    End If

    If p_Dict Is Nothing Then
        Set CmbRowsFromDict = m_Col
        Exit Function
    End If

    Dim m_ID As Variant
    Dim m_Usuario As Usuario
    For Each m_ID In p_Dict.Keys
        Set m_Usuario = p_Dict(m_ID)
        If Not m_Usuario Is Nothing Then
            m_Col.Add CStr(m_Usuario.ID) & ";" & CStr(m_Usuario.Nombre)
        End If
        Set m_Usuario = Nothing
    Next

    Set CmbRowsFromDict = m_Col
End Function


' Indica si el orphan ID esta presente en el Dictionary resultante.
' True si (a) el orphan es no vacio Y (b) existe como clave en el dict.
Private Function OrphanInDictionary( _
                                    ByVal p_Dict As Scripting.Dictionary, _
                                    ByVal p_OrphanID As String _
                                    ) As Boolean

    If Len(Trim$(p_OrphanID)) = 0 Then
        OrphanInDictionary = False
        Exit Function
    End If
    If p_Dict Is Nothing Then
        OrphanInDictionary = False
        Exit Function
    End If
    OrphanInDictionary = p_Dict.Exists(Trim$(p_OrphanID))
End Function


' AddItem al combo Calidad.  NO incluye sentinel (Calidad no admite
' "ninguno").
Private Sub AddItemCalidad( _
                            ByRef frm As Form, _
                            ByVal p_Dict As Scripting.Dictionary _
                            )
    Dim cmb As ComboBox
    Set cmb = frm.Controls("IDResponsableCalidad")
    cmb.RowSource = ""
    If p_Dict Is Nothing Then Exit Sub
    Dim m_ID As Variant
    Dim m_Usuario As Usuario
    For Each m_ID In p_Dict.Keys
        Set m_Usuario = p_Dict(m_ID)
        If Not m_Usuario Is Nothing Then
            cmb.AddItem CStr(m_Usuario.ID) & ";" & CStr(m_Usuario.Nombre)
        End If
        Set m_Usuario = Nothing
    Next
    Set cmb = Nothing
End Sub


' AddItem al combo Seguridad.  Inicia con el sentinel "0;N/A" (UX
' contract).
Private Sub AddItemSeguridad( _
                                ByRef frm As Form, _
                                ByVal p_Dict As Scripting.Dictionary _
                                )
    Dim cmb As ComboBox
    Set cmb = frm.Controls("IDResponsableSeguridad")
    cmb.RowSource = ""
    cmb.AddItem SENTINEL_SEG_ID & ";" & SENTINEL_SEG_LABEL
    cmb.DefaultValue = SENTINEL_SEG_ID
    If p_Dict Is Nothing Then Exit Sub
    Dim m_ID As Variant
    Dim m_Usuario As Usuario
    For Each m_ID In p_Dict.Keys
        Set m_Usuario = p_Dict(m_ID)
        If Not m_Usuario Is Nothing Then
            cmb.AddItem CStr(m_Usuario.ID) & ";" & CStr(m_Usuario.Nombre)
        End If
        Set m_Usuario = Nothing
    Next
    Set cmb = Nothing
End Sub


' -----------------------------------------------------------------------------
' Wrappers JSON (envelope canonico)
'
' Construye `{ok, value, payload, error, logs}` con JsonConverter para el
' payload (mismo patron que el resto de los helpers del proyecto).
' -----------------------------------------------------------------------------

Private Function BuildJsonOk( _
                                ByVal p_Value As String, _
                                ByVal p_Payload As Object, _
                                ByRef p_Logs() As String _
                                ) As String
    Dim m_PayloadJson As String
    If p_Payload Is Nothing Then
        m_PayloadJson = "null"
    Else
        m_PayloadJson = JsonConverter.ConvertToJson(p_Payload)
    End If
    BuildJsonOk = "{""ok"":true,""value"":""" & EscapeJson(p_Value) & """," & _
                    """payload"":" & m_PayloadJson & ",""error"":null,""logs"":[" & _
                    LogsToJson(p_Logs) & "]}"
End Function


Private Function BuildJsonFail( _
                                ByVal p_Error As String, _
                                ByRef p_Logs() As String, _
                                ByVal p_Det As String _
                                ) As String
    Dim m_ErrJson As String
    Dim m_MsgCompleto As String
    m_MsgCompleto = p_Error
    If Len(p_Det) > 0 Then m_MsgCompleto = m_MsgCompleto & " | " & p_Det
    m_ErrJson = """" & EscapeJson(m_MsgCompleto) & """"
    BuildJsonFail = "{""ok"":false,""value"":null,""payload"":null,""error"":" & _
                    m_ErrJson & ",""logs"":[" & LogsToJson(p_Logs) & "]}"
End Function


Private Function LogsToJson(ByRef p_Logs() As String) As String
    Dim i As Long, s As String
    For i = LBound(p_Logs) To UBound(p_Logs)
        If Len(p_Logs(i)) > 0 Then
            If Len(s) > 0 Then s = s & ","
            s = s & """" & EscapeJson(p_Logs(i)) & """"
        End If
    Next
    LogsToJson = s
End Function


Private Function EscapeJson(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbCr, "\n")
    EscapeJson = s
End Function


' Helper: crea un Dictionary (object) compatible con JsonConverter.
Private Function CreateJsonObject() As Object
    Dim d As Object
    Set d = New Scripting.Dictionary
    d.CompareMode = TextCompare
    Set CreateJsonObject = d
End Function