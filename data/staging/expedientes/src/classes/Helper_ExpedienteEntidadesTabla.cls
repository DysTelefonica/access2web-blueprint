VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Helper_ExpedienteEntidadesTabla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private m_IDExpediente As Long
Private m_Entidades As Object

Private Sub EnsureInitialized()
    If m_Entidades Is Nothing Then
        Set m_Entidades = CreateObject("Scripting.Dictionary")
        m_Entidades.CompareMode = TextCompare
    End If
End Sub

Private Function BuildPayloadEnvelope( _
    ByVal p_Ok As Boolean, _
    ByVal p_Payload As Object, _
    ByVal p_Error As String _
) As String
    Dim payloadJson As String
    If p_Payload Is Nothing Then
        payloadJson = "null"
    Else
        payloadJson = JsonConverter.ConvertToJson(p_Payload)
    End If

    Dim errorJson As String
    If p_Ok Then
        errorJson = "null"
    Else
        errorJson = """" & TestHelper.EscapeJsonString(p_Error) & """"
    End If

    Dim okJson As String
    If p_Ok Then
        okJson = "true"
    Else
        okJson = "false"
    End If

    BuildPayloadEnvelope = "{""ok"":" & okJson & _
                           ",""value"":null" & _
                           ",""payload"":" & payloadJson & _
                           ",""error"":" & errorJson & _
                           ",""logs"":[]}"
End Function

Private Function BuildSerializedPayload() As Object
    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim entities As New Collection
    Dim key As Variant
    For Each key In m_Entidades.Keys
        entities.Add m_Entidades(key)
    Next key

    payload("IDExpediente") = m_IDExpediente
    payload("count") = m_Entidades.Count
    Set payload("entities") = entities
    Set BuildSerializedPayload = payload
End Function

Private Function EntityID( _
    ByVal p_Entidad As Object, _
    Optional ByRef p_Error As String _
) As String
    On Error GoTo errores

    If p_Entidad Is Nothing Then
        p_Error = "Entity is Nothing"
        Err.Raise 1000
    End If
    If Not p_Entidad.Exists("IDEntidad") Then
        p_Error = "Entity requires IDEntidad"
        Err.Raise 1000
    End If
    If Len(Trim$("" & p_Entidad("IDEntidad"))) = 0 Then
        p_Error = "Entity requires IDEntidad"
        Err.Raise 1000
    End If

    EntityID = Trim$("" & p_Entidad("IDEntidad"))
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "EntityID: " & Err.Description
    End If
    Err.Raise 1000
End Function

Public Sub Init( _
    ByVal p_IDExpediente As Long, _
    Optional ByRef p_Error As String _
)
    On Error GoTo errores
    p_Error = ""
    m_IDExpediente = p_IDExpediente
    Set m_Entidades = CreateObject("Scripting.Dictionary")
    m_Entidades.CompareMode = TextCompare
    Exit Sub

errores:
    If Err.Number <> 1000 Then
        p_Error = "Init: " & Err.Description
    End If
    Err.Raise 1000
End Sub

Public Function Cargar( _
    Optional ByVal p_Entidades As Object = Nothing, _
    Optional ByRef p_Error As String _
) As Boolean
    Dim stage As String

    On Error GoTo errores
    p_Error = ""
    stage = "ensure-initialized"
    EnsureInitialized

    stage = "setup-collection"
    If Not p_Entidades Is Nothing Then
        Set m_Entidades = CreateObject("Scripting.Dictionary")
        m_Entidades.CompareMode = TextCompare

        stage = "iterate-collection"

        Dim key As Variant
        Dim entidad As Object
        For Each key In p_Entidades.Keys
            stage = "read-entity:" & CStr(key)
            Err.Clear
            Set entidad = p_Entidades(key)
            If Err.Number <> 0 Then
                p_Error = "Cargar: " & stage & ": " & Err.Description
                Err.Clear
                Err.Raise 1000
            End If

            stage = "agregar-entidad:" & CStr(key)
            Set entidad = p_Entidades(key)
            AgregarEntidad entidad, p_Error
            If p_Error <> "" Then
                p_Error = "Cargar: " & stage & ": " & p_Error
                Err.Raise 1000
            End If
        Next key
    End If

    Cargar = True
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "Cargar: " & stage & ": " & Err.Description
    End If
    Err.Raise 1000
End Function

Public Function AgregarEntidad( _
    ByVal p_Entidad As Object, _
    Optional ByRef p_Error As String _
) As Boolean
    Dim stage As String

    On Error GoTo errores
    p_Error = ""
    EnsureInitialized

    stage = "resolve-id"
    Dim idEntidad As String
    idEntidad = EntityID(p_Entidad, p_Error)
    If p_Error <> "" Then
        p_Error = "AgregarEntidad: " & stage & ": " & p_Error
        Err.Raise 1000
    End If

    stage = "check-existing"
    If m_Entidades.Exists(idEntidad) Then
        AgregarEntidad = False
        Exit Function
    End If

    stage = "add-entity"
    m_Entidades.Add idEntidad, p_Entidad
    AgregarEntidad = True
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "AgregarEntidad: " & stage & ": " & Err.Description
    End If
    Err.Raise 1000
End Function

Public Function EliminarEntidad( _
    ByVal p_IDEntidad As String, _
    Optional ByRef p_Error As String _
) As Boolean
    On Error GoTo errores
    p_Error = ""
    EnsureInitialized

    If Len(Trim$(p_IDEntidad)) = 0 Then
        p_Error = "EliminarEntidad requires IDEntidad"
        Err.Raise 1000
    End If
    If Not m_Entidades.Exists(Trim$(p_IDEntidad)) Then
        EliminarEntidad = False
        Exit Function
    End If

    m_Entidades.Remove Trim$(p_IDEntidad)
    EliminarEntidad = True
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "EliminarEntidad: " & Err.Description
    End If
    Err.Raise 1000
End Function

Public Function Serializar( _
    Optional ByRef p_Error As String _
) As String
    On Error GoTo errores
    p_Error = ""
    EnsureInitialized
    Serializar = BuildPayloadEnvelope(True, BuildSerializedPayload(), "")
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "Serializar: " & Err.Description
    End If
    Serializar = BuildPayloadEnvelope(False, Nothing, p_Error)
    Err.Raise 1000
End Function
