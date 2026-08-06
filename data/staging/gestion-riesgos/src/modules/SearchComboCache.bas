Attribute VB_Name = "SearchComboCache"
Option Compare Database
Option Explicit

Private Const SEARCH_COMBO_TABLE_NC As String = "TBNOConformidades"
Private Const SEARCH_COMBO_FIELD_EXPEDIENTE As String = "EXPEDIENTE"
Private Const SEARCH_COMBO_FIELD_PROYECTO As String = "PROYECTO"
Private Const SEARCH_COMBO_FIELD_VEHICULO As String = "VEHICULO"

Private m_Cache As Scripting.Dictionary
Private m_MissCount As Long
Private m_HitCount As Long
Private m_InvalidationCount As Long

Public Function SearchComboCache_GetDistinctValues( _
    ByVal p_NombreTabla As String, _
    ByVal p_NombreCampo As String, _
    Optional ByVal p_db As DAO.Database, _
    Optional ByRef p_Error As String) As Scripting.Dictionary

    Dim cacheKey As String
    Dim loadedValues As Scripting.Dictionary

    On Error GoTo EH

    p_Error = ""
    cacheKey = SearchComboCache_Key(p_NombreTabla, p_NombreCampo, p_Error)
    If p_Error <> "" Then Err.Raise 1000

    SearchComboCache_EnsureCache
    If m_Cache.Exists(cacheKey) Then
        m_HitCount = m_HitCount + 1
        Set SearchComboCache_GetDistinctValues = m_Cache(cacheKey)
        Exit Function
    End If

    Set loadedValues = SearchComboCache_LoadDistinctValues(cacheKey, p_db, p_Error)
    If p_Error <> "" Then Err.Raise 1000

    m_Cache.Add cacheKey, loadedValues
    m_MissCount = m_MissCount + 1
    Set SearchComboCache_GetDistinctValues = loadedValues
    Exit Function

EH:
    If Err.Number <> 1000 Then
        p_Error = "SearchComboCache_GetDistinctValues: " & Err.Number & " - " & Err.Description
    End If
    Set SearchComboCache_GetDistinctValues = Nothing
End Function

Public Sub SearchComboCache_Invalidar( _
    Optional ByVal p_NombreTabla As String = "", _
    Optional ByVal p_NombreCampo As String = "")

    Dim cacheKey As String
    Dim errMsg As String

    On Error GoTo EH

    SearchComboCache_EnsureCache
    If p_NombreTabla = "" Or p_NombreCampo = "" Then
        If m_Cache.Count > 0 Then m_InvalidationCount = m_InvalidationCount + 1
        Set m_Cache = New Scripting.Dictionary
        m_Cache.CompareMode = TextCompare
        Exit Sub
    End If

    cacheKey = SearchComboCache_Key(p_NombreTabla, p_NombreCampo, errMsg)
    If errMsg <> "" Then Exit Sub
    If m_Cache.Exists(cacheKey) Then
        m_Cache.Remove cacheKey
        m_InvalidationCount = m_InvalidationCount + 1
    End If
    Exit Sub

EH:
    ' Invalidation is intentionally best-effort for UI callers.
End Sub

Public Sub SearchComboCache_Reset(Optional ByRef p_Error As String)
    p_Error = ""
    Set m_Cache = New Scripting.Dictionary
    m_Cache.CompareMode = TextCompare
    m_MissCount = 0
    m_HitCount = 0
    m_InvalidationCount = 0
End Sub

Public Function SearchComboCache_MissCount() As Long
    SearchComboCache_MissCount = m_MissCount
End Function

Public Function SearchComboCache_HitCount() As Long
    SearchComboCache_HitCount = m_HitCount
End Function

Public Function SearchComboCache_InvalidationCount() As Long
    SearchComboCache_InvalidationCount = m_InvalidationCount
End Function

Private Sub SearchComboCache_EnsureCache()
    If m_Cache Is Nothing Then
        Set m_Cache = New Scripting.Dictionary
        m_Cache.CompareMode = TextCompare
    End If
End Sub

Private Function SearchComboCache_Key( _
    ByVal p_NombreTabla As String, _
    ByVal p_NombreCampo As String, _
    ByRef p_Error As String) As String

    Dim tableName As String
    Dim fieldName As String

    p_Error = ""
    tableName = UCase$(Trim$(p_NombreTabla))
    fieldName = UCase$(Trim$(p_NombreCampo))

    If tableName <> UCase$(SEARCH_COMBO_TABLE_NC) Then
        p_Error = "Unsupported search combo cache table: " & p_NombreTabla
        Exit Function
    End If

    Select Case fieldName
        Case SEARCH_COMBO_FIELD_EXPEDIENTE, SEARCH_COMBO_FIELD_PROYECTO, SEARCH_COMBO_FIELD_VEHICULO
            SearchComboCache_Key = tableName & "." & fieldName
        Case Else
            p_Error = "Unsupported search combo cache field: " & p_NombreCampo
    End Select
End Function

Private Function SearchComboCache_LoadDistinctValues( _
    ByVal p_CacheKey As String, _
    Optional ByVal p_db As DAO.Database, _
    Optional ByRef p_Error As String) As Scripting.Dictionary

    Dim fieldName As String
    Dim values As Scripting.Dictionary

    On Error GoTo EH

    p_Error = ""
    fieldName = Mid$(p_CacheKey, InStr(1, p_CacheKey, ".", vbTextCompare) + 1)
    Set values = Constructor.getValoresDistintos("TbNoConformidades", fieldName, p_db, p_Error)
    If p_Error <> "" Then Err.Raise 1000
    If values Is Nothing Then
        Set values = New Scripting.Dictionary
        values.CompareMode = TextCompare
    End If

    Set SearchComboCache_LoadDistinctValues = values
    Exit Function

EH:
    If Err.Number <> 1000 Then
        p_Error = "SearchComboCache_LoadDistinctValues: " & Err.Number & " - " & Err.Description
    End If
    Set SearchComboCache_LoadDistinctValues = Nothing
End Function
