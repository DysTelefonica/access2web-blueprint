VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RiskTreeSelectionResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Function ResolveFallbackKey( _
    ByVal p_RequestedKey As String, _
    ByVal p_CandidateKeysSerialized As String, _
    ByVal p_ExistingKeysSerialized As String, _
    Optional ByRef p_Error As String _
) As String
    Dim candidateKeys() As String
    Dim item As Variant
    Dim key As String
    Dim requestedKey As String

    On Error GoTo EH

    p_Error = vbNullString
    requestedKey = Trim$(Nz(p_RequestedKey, vbNullString))
    candidateKeys = Split(Nz(p_CandidateKeysSerialized, vbNullString), vbLf)

    If requestedKey <> vbNullString Then
        If KeyExistsInSerialized(requestedKey, p_ExistingKeysSerialized) Then
            ResolveFallbackKey = requestedKey
            Exit Function
        End If
    End If

    For Each item In candidateKeys
        key = Trim$(CStr(item))
        If key <> vbNullString Then
            If KeyExistsInSerialized(key, p_ExistingKeysSerialized) Then
                ResolveFallbackKey = key
                Exit Function
            End If
        End If
    Next item

    ResolveFallbackKey = vbNullString
    Exit Function

EH:
    If p_Error = vbNullString Then
        p_Error = "RiskTreeSelectionResolver.ResolveFallbackKey: " & Err.Description
    End If
    Err.Raise IIf(Err.Number = 0, 1000, Err.Number), "RiskTreeSelectionResolver.ResolveFallbackKey", p_Error
End Function

Public Function ResolveCandidateFallbackKey( _
    ByVal p_RequestedKey As String, _
    ByVal p_CandidateKeysSerialized As String, _
    Optional ByRef p_Error As String _
) As String
    Dim candidateKeys() As String
    Dim item As Variant
    Dim key As String
    Dim requestedKey As String

    On Error GoTo EH

    p_Error = vbNullString
    requestedKey = Trim$(Nz(p_RequestedKey, vbNullString))
    candidateKeys = Split(Nz(p_CandidateKeysSerialized, vbNullString), vbLf)

    If requestedKey <> vbNullString Then
        For Each item In candidateKeys
            key = Trim$(CStr(item))
            If key <> vbNullString Then
                If StrComp(key, requestedKey, vbTextCompare) = 0 Then
                    ResolveCandidateFallbackKey = key
                    Exit Function
                End If
            End If
        Next item
    End If

    For Each item In candidateKeys
        key = Trim$(CStr(item))
        If key <> vbNullString Then
            ResolveCandidateFallbackKey = key
            Exit Function
        End If
    Next item

    ResolveCandidateFallbackKey = vbNullString
    Exit Function

EH:
    If p_Error = vbNullString Then
        p_Error = "RiskTreeSelectionResolver.ResolveCandidateFallbackKey: " & Err.Description
    End If
    Err.Raise IIf(Err.Number = 0, 1000, Err.Number), "RiskTreeSelectionResolver.ResolveCandidateFallbackKey", p_Error
End Function

Private Function KeyExistsInSerialized( _
    ByVal p_Key As String, _
    ByVal p_KeysSerialized As String _
) As Boolean
    Dim keys() As String
    Dim item As Variant
    Dim key As String

    key = Trim$(Nz(p_Key, vbNullString))
    If key = vbNullString Then Exit Function

    keys = Split(Nz(p_KeysSerialized, vbNullString), vbLf)
    For Each item In keys
        If StrComp(Trim$(CStr(item)), key, vbTextCompare) = 0 Then
            KeyExistsInSerialized = True
            Exit Function
        End If
    Next item
End Function
