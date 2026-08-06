VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Helper_ExpedienteSuministradoresArbol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' Helper_ExpedienteSuministradoresArbol — pure in-memory tree model for the
' supplier chain of an expediente.
'
' Contract:
'   Init             Optional organ name; primes empty state.
'   Reset            Clears all state (nodes, selection, organ name).
'   Cargar           Loads pre-materialized rows; atomic validation; returns count.
'   GetNodo          Returns the stored Dictionary for a node key, or Nothing.
'   SeleccionarNodo  True if the key exists and updates selection;
'                    False otherwise, leaving the previous selection intact.
'   Serializar       Returns the canonical JSON envelope describing the model.
'
' Cargar is atomic: any validation failure (missing ID, duplicate key,
' orphan parent, cycle) leaves m_Nodos unchanged. Serialized payload is
' independent of input iteration order (sorted by depth, then key).

Private Const K_ROOT_ORGANO As String = "ROOT_ORGANO"
Private Const K_PREFIX_NODO As String = "N-"

Private m_OrganoNombre As String
Private m_Nodos As Object
Private m_NodoSeleccionado As String

' --- Public API -------------------------------------------------------------

Public Sub Init(Optional ByVal p_OrganoNombre As String = "", Optional ByRef p_Error As String)
    On Error GoTo errores
    p_Error = ""
    m_OrganoNombre = p_OrganoNombre
    Set m_Nodos = NewDict()
    m_NodoSeleccionado = ""
    Exit Sub
errores:
    If Err.Number <> 1000 Then p_Error = "Init: " & Err.Description
    Err.Raise 1000
End Sub

Public Sub Reset(Optional ByRef p_Error As String)
    On Error GoTo errores
    p_Error = ""
    Set m_Nodos = NewDict()
    m_NodoSeleccionado = ""
    m_OrganoNombre = ""
    Exit Sub
errores:
    If Err.Number <> 1000 Then p_Error = "Reset: " & Err.Description
    Err.Raise 1000
End Sub

Public Function Cargar(ByVal p_Items As Object, Optional ByRef p_Error As String) As Long
    Dim stage As String
    Dim tmp As Object, keyMap As Object
    Dim k As Variant, raw As Object
    Dim nodeKey As String, parentKey As String

    On Error GoTo errores
    p_Error = ""
    Cargar = 0

    If m_Nodos Is Nothing Then Set m_Nodos = NewDict()
    If p_Items Is Nothing Then Exit Function

    Set tmp = NewDict()
    Set keyMap = NewDict()

    stage = "build-tmp"
    For Each k In p_Items.Keys
        Set raw = p_Items(k)
        If raw Is Nothing Then
            p_Error = "Cargar: " & stage & ": item at key '" & CStr(k) & "' is Nothing"
            Err.Raise 1000
        End If
        If Not raw.Exists("IDExpedienteSuministrador") Then
            p_Error = "Cargar: " & stage & ": item at key '" & CStr(k) & "' missing IDExpedienteSuministrador"
            Err.Raise 1000
        End If
        If IsNull(raw("IDExpedienteSuministrador")) Then
            p_Error = "Cargar: " & stage & ": item at key '" & CStr(k) & "' has Null IDExpedienteSuministrador"
            Err.Raise 1000
        End If
        nodeKey = K_PREFIX_NODO & CStr(CLng(raw("IDExpedienteSuministrador")))
        If CLng(raw("IDExpedienteSuministrador")) <= 0 Then
            p_Error = "Cargar: " & stage & ": item at key '" & CStr(k) & "' has non-positive IDExpedienteSuministrador"
            Err.Raise 1000
        End If
        If tmp.Exists(nodeKey) Then
            p_Error = "Cargar: " & stage & ": duplicate IDExpedienteSuministrador " & nodeKey
            Err.Raise 1000
        End If
        tmp.Add nodeKey, raw
        If raw.Exists("IdPadre") And Not IsNull(raw("IdPadre")) And CLng(raw("IdPadre")) > 0 Then
            keyMap.Add nodeKey, K_PREFIX_NODO & CStr(CLng(raw("IdPadre")))
        Else
            keyMap.Add nodeKey, K_ROOT_ORGANO
        End If
    Next k

    stage = "orphan-check"
    For Each k In keyMap.Keys
        parentKey = CStr(keyMap(k))
        If parentKey <> K_ROOT_ORGANO Then
            If Not tmp.Exists(parentKey) Then
                p_Error = "Cargar: " & stage & ": orphan parent " & parentKey & " for node " & CStr(k)
                Err.Raise 1000
            End If
        End If
    Next k

    stage = "cycle-check"
    For Each k In tmp.Keys
        If HasCycleFrom(CStr(k), tmp) Then
            p_Error = "Cargar: " & stage & ": cycle detected starting from " & CStr(k)
            Err.Raise 1000
        End If
    Next k

    Set m_Nodos = tmp
    Cargar = m_Nodos.Count
    Exit Function
errores:
    If Err.Number <> 1000 Then p_Error = "Cargar: " & stage & ": " & Err.Description
    Err.Raise 1000
End Function

Public Function GetNodo(ByVal p_NodoKey As String, Optional ByRef p_Error As String) As Object
    On Error GoTo errores
    p_Error = ""
    If m_Nodos Is Nothing Then Set GetNodo = Nothing: Exit Function
    If Not m_Nodos.Exists(p_NodoKey) Then Set GetNodo = Nothing: Exit Function
    Set GetNodo = m_Nodos(p_NodoKey)
    Exit Function
errores:
    If Err.Number <> 1000 Then p_Error = "GetNodo: " & Err.Description
    Err.Raise 1000
End Function

Public Function SeleccionarNodo(ByVal p_NodoKey As String, Optional ByRef p_Error As String) As Boolean
    On Error GoTo errores
    p_Error = ""
    If m_Nodos Is Nothing Then SeleccionarNodo = False: Exit Function
    If Not m_Nodos.Exists(p_NodoKey) Then SeleccionarNodo = False: Exit Function
    m_NodoSeleccionado = p_NodoKey
    SeleccionarNodo = True
    Exit Function
errores:
    If Err.Number <> 1000 Then p_Error = "SeleccionarNodo: " & Err.Description
    Err.Raise 1000
End Function

Public Function Serializar(Optional ByRef p_Error As String) As String
    Dim payload As Object, depthMap As Object, children As Object, sortedKeys As Collection
    Dim nodes As New Collection, view As Object
    Dim k As Variant, pk As String, viewKey As Variant

    On Error GoTo errores
    p_Error = ""
    If m_Nodos Is Nothing Then Set m_Nodos = NewDict()

    Set depthMap = NewDict()
    For Each k In m_Nodos.Keys
        If ComputeDepth(CStr(k), m_Nodos, depthMap) < 0 Then
            p_Error = "Serializar: cycle encountered while computing depths"
            Err.Raise 1000
        End If
    Next k

    Set children = NewDict()
    For Each k In m_Nodos.Keys
        pk = ParentKeyOf(m_Nodos(CStr(k)))
        If Not children.Exists(pk) Then children.Add pk, NewDict()
        children(pk).Add CStr(k), True
    Next k

    Set sortedKeys = SortByDepthThenKey(m_Nodos, depthMap)

    Set payload = NewDict()
    payload("IDExpediente") = 0
    payload("OrganoNombre") = m_OrganoNombre
    payload("NodoSeleccionado") = m_NodoSeleccionado
    payload("Count") = m_Nodos.Count

    nodes.Add BuildView(K_ROOT_ORGANO, "", False, "Órgano de Contratación (" & m_OrganoNombre & ")", "", 0, ChildrenOf(children, K_ROOT_ORGANO))

    For Each viewKey In sortedKeys
        nodes.Add BuildView( _
            CStr(viewKey), _
            ParentKeyOf(m_Nodos(CStr(viewKey))), _
            RawBool(m_Nodos(CStr(viewKey)), "EsEmpresaPropia"), _
            RawStr(m_Nodos(CStr(viewKey)), "Text"), _
            RawStr(m_Nodos(CStr(viewKey)), "Tag"), _
            CLng(depthMap(CStr(viewKey))), _
            ChildrenOf(children, CStr(viewKey)))
    Next viewKey

    Set payload("nodes") = nodes
    Serializar = BuildPayloadEnvelope(True, payload, "")
    Exit Function
errores:
    If Err.Number <> 1000 Then p_Error = "Serializar: " & Err.Description
    Serializar = BuildPayloadEnvelope(False, Nothing, p_Error)
    Err.Raise 1000
End Function

' --- Internal helpers -------------------------------------------------------

Private Function NewDict() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = TextCompare
    Set NewDict = d
End Function

Private Function ParentKeyOf(ByVal p_Item As Object) As String
    If p_Item.Exists("IdPadre") And Not IsNull(p_Item("IdPadre")) And CLng(p_Item("IdPadre")) > 0 Then
        ParentKeyOf = K_PREFIX_NODO & CStr(CLng(p_Item("IdPadre")))
    Else
        ParentKeyOf = K_ROOT_ORGANO
    End If
End Function

Private Function RawStr(ByVal p_Item As Object, ByVal p_Key As String) As String
    If p_Item.Exists(p_Key) And Not IsNull(p_Item(p_Key)) Then RawStr = CStr(p_Item(p_Key)) Else RawStr = ""
End Function

Private Function RawBool(ByVal p_Item As Object, ByVal p_Key As String) As Boolean
    If p_Item.Exists(p_Key) And Not IsNull(p_Item(p_Key)) Then RawBool = CBool(p_Item(p_Key)) Else RawBool = False
End Function

Private Function BuildView( _
    ByVal p_Key As String, _
    ByVal p_ParentKey As String, _
    ByVal p_EsEmpresaPropia As Boolean, _
    ByVal p_Text As String, _
    ByVal p_Tag As String, _
    ByVal p_Depth As Long, _
    ByVal p_ChildKeys As Object _
) As Object
    Dim view As Object
    Set view = NewDict()
    view("key") = p_Key
    If p_ParentKey = "" Then view("parentKey") = Null Else view("parentKey") = p_ParentKey
    view("text") = p_Text
    view("tag") = p_Tag
    view("esEmpresaPropia") = p_EsEmpresaPropia
    view("depth") = p_Depth
    Set view("childKeys") = SortedCollectionFromDict(p_ChildKeys)
    Set BuildView = view
End Function

Private Function ChildrenOf(ByVal p_Children As Object, ByVal p_Key As String) As Object
    If p_Children.Exists(p_Key) Then Set ChildrenOf = p_Children(p_Key) Else Set ChildrenOf = NewDict()
End Function

Private Function SortedCollectionFromDict(ByVal p_Dict As Object) As Collection
    Dim col As New Collection
    If p_Dict Is Nothing Then Set SortedCollectionFromDict = col: Exit Function
    Dim n As Long: n = p_Dict.Count
    If n = 0 Then Set SortedCollectionFromDict = col: Exit Function

    Dim arr() As String, idx As Long, k As Variant
    ReDim arr(0 To n - 1)
    idx = 0
    For Each k In p_Dict.Keys
        arr(idx) = CStr(k)
        idx = idx + 1
    Next k

    Dim i As Long, j As Long, tmp As String
    For i = 0 To n - 2
        For j = 0 To n - 2 - i
            If StrComp(arr(j), arr(j + 1), vbBinaryCompare) > 0 Then
                tmp = arr(j): arr(j) = arr(j + 1): arr(j + 1) = tmp
            End If
        Next j
    Next i

    For i = 0 To n - 1
        col.Add arr(i)
    Next i
    Set SortedCollectionFromDict = col
End Function

Private Function SortByDepthThenKey(ByVal p_Nodos As Object, ByVal p_DepthMap As Object) As Collection
    Dim n As Long: n = p_Nodos.Count
    If n = 0 Then
        Dim vacio As New Collection: Set SortByDepthThenKey = vacio: Exit Function
    End If
    Dim arr() As String, k As Variant, idx As Long
    ReDim arr(0 To n - 1)
    idx = 0
    For Each k In p_Nodos.Keys
        arr(idx) = CStr(k)
        idx = idx + 1
    Next k
    Dim i As Long, j As Long, tmp As String
    Dim da As Long, db As Long, a As String, b As String
    For i = 0 To n - 2
        For j = 0 To n - 2 - i
            a = arr(j): b = arr(j + 1)
            da = CLng(p_DepthMap(a)): db = CLng(p_DepthMap(b))
            If da > db Or (da = db And StrComp(a, b, vbBinaryCompare) > 0) Then
                tmp = arr(j): arr(j) = arr(j + 1): arr(j + 1) = tmp
            End If
        Next j
    Next i
    Dim sorted As New Collection
    For i = 0 To n - 1
        sorted.Add arr(i)
    Next i
    Set SortByDepthThenKey = sorted
End Function

Private Function ComputeDepth(ByVal p_Key As String, ByVal p_Nodos As Object, ByRef p_DepthMap As Object) As Long
    Dim visited As Object: Set visited = NewDict()
    Dim current As String: current = p_Key
    Do While True
        If p_DepthMap.Exists(current) Then ComputeDepth = CLng(p_DepthMap(current)): Exit Function
        If visited.Exists(current) Then ComputeDepth = -1: Exit Function
        visited.Add current, True
        Dim parentKey As String
        parentKey = ParentKeyOf(p_Nodos(current))
        If parentKey = K_ROOT_ORGANO Then
            p_DepthMap.Add current, CLng(1)
            ComputeDepth = 1
            Exit Function
        End If
        If Not p_Nodos.Exists(parentKey) Then ComputeDepth = -1: Exit Function
        current = parentKey
    Loop
End Function

Private Function HasCycleFrom(ByVal p_StartKey As String, ByVal p_Nodos As Object) As Boolean
    Dim visited As Object: Set visited = NewDict()
    Dim current As String: current = p_StartKey
    Do While True
        If visited.Exists(current) Then HasCycleFrom = True: Exit Function
        visited.Add current, True
        Dim parentKey As String
        parentKey = ParentKeyOf(p_Nodos(current))
        If parentKey = K_ROOT_ORGANO Then HasCycleFrom = False: Exit Function
        If Not p_Nodos.Exists(parentKey) Then HasCycleFrom = False: Exit Function
        current = parentKey
    Loop
End Function

Private Function BuildPayloadEnvelope(ByVal p_Ok As Boolean, ByVal p_Payload As Object, ByVal p_Error As String) As String
    Dim payloadJson As String
    If p_Payload Is Nothing Then
        payloadJson = "null"
    Else
        payloadJson = JsonConverter.ConvertToJson(p_Payload)
    End If
    Dim errorJson As String
    If p_Ok Then errorJson = "null" Else errorJson = """" & TestHelper.EscapeJsonString(p_Error) & """"
    Dim okJson As String
    If p_Ok Then okJson = "true" Else okJson = "false"
    BuildPayloadEnvelope = "{""ok"":" & okJson & ",""value"":null,""payload"":" & payloadJson & ",""error"":" & errorJson & ",""logs"":[]}"
End Function
