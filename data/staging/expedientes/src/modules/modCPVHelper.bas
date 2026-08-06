Attribute VB_Name = "modCPVHelper"
Option Compare Database
Option Explicit

' modCPVHelper — REWORK (2026-06-26)
' Pure-data helpers for Form_FormCPVsGestion.
'
' Implements access-vba-e2e-methodology rule #1 (forms are thin UI wiring),
' rule #2 (honest helper signatures), rule #5 (per-module prefix: CPV_*),
' and rule #9 (no `ByRef p_Form`).
'
' New design (5 helpers, mirror of modComercialHelper):
'   1. CPV_Abrir_Inicializar(p_EsAdministrador, p_HasOpenArgs, p_Error)
'   2. CPV_Buscar_Listar(p_CPVs, p_Filter, p_Error)
'      Filter field: CPV (text name only — legacy listbox is 2-col)
'   3. CPV_Seleccionar_Cargar(p_IDSeleccionado, p_CPVs, p_EsAdministrador, p_Error)
'   4. CPV_Eliminar_Borrar(p_CPV, p_PromptResult, p_Error)
'   5. CPV_DobleClick_AbrirEdicion(p_HasElegir, p_EditarEnabled, p_Error)
'
' UI orchestration that stays in the form:
'   - Alta button: DoCmd.OpenForm "FormCPV"
'   - Edición button: DoCmd.OpenForm "FormCPV" with m_ObjCPVActivo set
'   - Limpiar button: Me.CPV = Null
'
' Telefonica D&S convention (vba-access §1.4.1): every Public Function ends with
' `Optional ByRef p_Error As String` as the LAST parameter.

' === Module-level constants (all at top per vba-access §10.1) ============================

' Legacy listbox shows ID + CPV only (DESCRIPCION NOT in listbox for CPV — 2-col shape).
Private Const CPV_HEADERS As String = "IDCPV;CPV"

Private Const CPV_FIELDS_SEP As String = ";"

Private Const CPV_MAX_ROWS As Long = 5000

Private Const CPV_TEST_ID_BASE As Long = 900510


' === Local helpers (all at top per vba-access §10.1) ====================================

Private Function BuildJsonPayload( _
    ByVal p_Ok As Boolean, _
    ByVal p_Payload As Object, _
    ByVal p_ErrorMsg As String, _
    ByRef p_Logs() As String _
) As String
    Dim payloadJson As String
    If p_Payload Is Nothing Then
        payloadJson = "null"
    Else
        payloadJson = JsonConverter.ConvertToJson(p_Payload)
    End If

    Dim logsJson As String
    logsJson = TestHelper.JsonStringArray(p_Logs)

    Dim errorJson As String
    If p_Ok Then
        errorJson = "null"
    Else
        errorJson = """" & TestHelper.EscapeJsonString(p_ErrorMsg) & """"
    End If

    BuildJsonPayload = "{""ok"":" & LCase$(CStr(p_Ok)) & _
                       ",""value"":null" & _
                       ",""payload"":" & payloadJson & _
                       ",""error"":" & errorJson & _
                       ",""logs"":" & logsJson & "}"
End Function

' --- CPV_SerializarEntidad -----------------------------------------------------
Private Function CPV_SerializarEntidad(ByVal p_Entidad As Object) As Object
    If p_Entidad Is Nothing Then
        Set CPV_SerializarEntidad = Nothing
        Exit Function
    End If

    Dim out As Object
    Set out = CreateObject("Scripting.Dictionary")

    If p_Entidad.Exists("IDCPV") Then
        out("IDCPV") = CStr(p_Entidad("IDCPV"))
    Else
        out("IDCPV") = ""
    End If

    If p_Entidad.Exists("CPV") Then
        out("CPV") = CStr(p_Entidad("CPV"))
    Else
        out("CPV") = ""
    End If

    If p_Entidad.Exists("DESCRIPCION") Then
        out("DESCRIPCION") = CStr(p_Entidad("DESCRIPCION"))
    Else
        out("DESCRIPCION") = ""
    End If

    Set CPV_SerializarEntidad = out
End Function

' --- CPV_BuildRowLine ---------------------------------------------------------
Private Function CPV_BuildRowLine( _
    ByVal p_ID As String, _
    ByVal p_Nombre As String _
) As String
    Dim safeNombre As String
    safeNombre = Replace(p_Nombre, CPV_FIELDS_SEP, ":")

    CPV_BuildRowLine = p_ID & CPV_FIELDS_SEP & safeNombre
End Function


' === Public API =========================================================================

' --- CPV_Abrir_Inicializar ---------------------------------------------------
Public Function CPV_Abrir_Inicializar( _
    ByVal p_EsAdministrador As Boolean, _
    ByVal p_HasOpenArgs As Boolean, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(4)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim showAlta As Boolean
    Dim focusAlta As Boolean
    Dim showElegir As Boolean

    If p_EsAdministrador Then
        showAlta = True
        focusAlta = True
    Else
        showAlta = False
        focusAlta = False
    End If

    If p_HasOpenArgs Then
        showElegir = True
    Else
        showElegir = False
    End If

    payload("showAlta") = showAlta
    payload("focusAlta") = focusAlta
    payload("showElegir") = showElegir

    logs(0) = "Abrir_Inicializar: admin=" & CStr(p_EsAdministrador) & _
              ", hasOpenArgs=" & CStr(p_HasOpenArgs)
    logs(1) = "Abrir_Inicializar: showAlta=" & CStr(showAlta) & _
              ", focusAlta=" & CStr(focusAlta) & _
              ", showElegir=" & CStr(showElegir)

    CPV_Abrir_Inicializar = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "CPV_Abrir_Inicializar: " & Err.Description
    End If
    CPV_Abrir_Inicializar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- CPV_Buscar_Listar --------------------------------------------------------
Public Function CPV_Buscar_Listar( _
    ByVal p_CPVs As Object, _
    ByVal p_Filter As String, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(6)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim rowSource As String
    rowSource = ""

    Dim rowCount As Long
    rowCount = 0

    If p_CPVs Is Nothing Then
        logs(0) = "Buscar_Listar: p_CPVs is Nothing -> 0 rows"
        payload("rowSource") = ""
        payload("count") = 0
        CPV_Buscar_Listar = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    Dim filterText As String
    filterText = "" & p_Filter

    Dim k As Variant
    Dim entity As Object
    Dim currentId As String
    Dim currentNombre As String

    For Each k In p_CPVs.Keys
        If rowCount >= CPV_MAX_ROWS Then
            logs(3) = "Buscar_Listar: hit MAX_ROWS=" & CPV_MAX_ROWS & ", truncating"
            Exit For
        End If

        Set entity = p_CPVs(k)
        If entity Is Nothing Then
            ' Skip malformed entries.
        Else
            If entity.Exists("IDCPV") Then
                currentId = CStr(entity("IDCPV"))
            Else
                currentId = ""
            End If

            If entity.Exists("CPV") Then
                currentNombre = CStr(entity("CPV"))
            Else
                currentNombre = ""
            End If

            If Len(filterText) > 0 Then
                If InStr(1, currentNombre, filterText, vbTextCompare) = 0 Then
                    ' No match — skip without counting.
                Else
                    If Len(rowSource) > 0 Then
                        rowSource = rowSource & vbCrLf
                    End If
                    rowSource = rowSource & CPV_BuildRowLine(currentId, currentNombre)
                    rowCount = rowCount + 1
                End If
            Else
                ' No filter — include everything.
                If Len(rowSource) > 0 Then
                    rowSource = rowSource & vbCrLf
                End If
                rowSource = rowSource & CPV_BuildRowLine(currentId, currentNombre)
                rowCount = rowCount + 1
            End If
        End If
        Set entity = Nothing
    Next k

    payload("rowSource") = rowSource
    payload("count") = rowCount

    logs(0) = "Buscar_Listar: filterText=" & filterText & ", rows=" & rowCount
    logs(1) = "Buscar_Listar: rowSource length=" & Len(rowSource)

    CPV_Buscar_Listar = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "CPV_Buscar_Listar: " & Err.Description
    End If
    CPV_Buscar_Listar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- CPV_Seleccionar_Cargar ---------------------------------------------------
Public Function CPV_Seleccionar_Cargar( _
    ByVal p_IDSeleccionado As String, _
    ByVal p_CPVs As Object, _
    ByVal p_EsAdministrador As Boolean, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(5)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim selectedId As String
    selectedId = Trim$("" & p_IDSeleccionado)

    If Len(selectedId) = 0 Then
        Set payload("entity") = Nothing
        payload("enableEditar") = False
        payload("enableEliminar") = False
        logs(0) = "Seleccionar_Cargar: empty selection"
        CPV_Seleccionar_Cargar = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    Dim entity As Object
    If p_CPVs Is Nothing Then
        Set entity = Nothing
    Else
        If p_CPVs.Exists(selectedId) Then
            Set entity = p_CPVs(selectedId)
        Else
            Set entity = Nothing
        End If
    End If

    If entity Is Nothing Then
        logs(0) = "Seleccionar_Cargar: id=" & selectedId & " not found in collection"
        Set payload("entity") = Nothing
        payload("enableEditar") = False
        payload("enableEliminar") = False
        CPV_Seleccionar_Cargar = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    Dim serializedEntity As Object
    Set serializedEntity = CPV_SerializarEntidad(entity)

    Set payload("entity") = serializedEntity
    payload("enableEditar") = True
    If p_EsAdministrador Then
        payload("enableEliminar") = True
    Else
        payload("enableEliminar") = False
    End If

    logs(0) = "Seleccionar_Cargar: id=" & selectedId
    logs(1) = "Seleccionar_Cargar: isAdmin=" & CStr(p_EsAdministrador)

    CPV_Seleccionar_Cargar = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "CPV_Seleccionar_Cargar: " & Err.Description
    End If
    CPV_Seleccionar_Cargar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- CPV_Eliminar_Borrar -----------------------------------------------------
Public Function CPV_Eliminar_Borrar( _
    ByVal p_CPV As Object, _
    Optional ByRef p_PromptResult As Long = 0, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(6)
    p_Error = ""

    On Error GoTo errores

    If p_CPV Is Nothing Then
        p_Error = "CPV_Eliminar_Borrar: p_CPV is Nothing"
        CPV_Eliminar_Borrar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim idCPV As String
    Dim cpvNombre As String

    If p_CPV.Exists("IDCPV") Then
        idCPV = CStr(p_CPV("IDCPV"))
    Else
        idCPV = ""
    End If

    If p_CPV.Exists("CPV") Then
        cpvNombre = CStr(p_CPV("CPV"))
    Else
        cpvNombre = ""
    End If

    If Len(idCPV) = 0 Then
        p_Error = "CPV_Eliminar_Borrar: missing IDCPV"
        CPV_Eliminar_Borrar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    logs(0) = "Eliminar_Borrar: id=" & idCPV & ", name=" & cpvNombre

    Dim promptResult As Long
    If p_PromptResult <> 0 Then
        promptResult = p_PromptResult
    Else
        promptResult = MsgBox("¿Desea realmente borrar al CPV seleccionado?", _
                              vbExclamation + vbYesNo + vbDefaultButton2, "Eliminar")
    End If

    logs(2) = "Eliminar_Borrar: promptResult=" & promptResult

    If promptResult <> vbYes Then
        logs(3) = "Eliminar_Borrar: cancelled by user"
        CPV_Eliminar_Borrar = BuildJsonPayload(True, Nothing, "", logs)
        Exit Function
    End If

    Dim m_CPVOP As New CPVOperaciones

    Dim wrapped As New CPV
    If p_CPV.Exists("IDCPV") Then
        wrapped.IDCPV = CStr(p_CPV("IDCPV"))
    End If
    If p_CPV.Exists("CPV") Then
        wrapped.CPV = CStr(p_CPV("CPV"))
    End If
    If p_CPV.Exists("DESCRIPCION") Then
        wrapped.DESCRIPCION = CStr(p_CPV("DESCRIPCION"))
    End If

    Set m_CPVOP.CPV = wrapped

    Dim daoErr As String
    Dim daoResult As String
    daoResult = Helper_EntidadCRUD.EliminarEntidadGenerico( _
        m_CPVOP, "CPV", wrapped, "", daoErr)
    If daoErr <> "" Then
        p_Error = daoErr
        logs(3) = "Eliminar_Borrar: DAO error=" & daoErr
        CPV_Eliminar_Borrar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    If Len(daoResult) > 0 And daoResult <> "OK" Then
        logs(3) = "Eliminar_Borrar: DAO returned '" & daoResult & "' — treating as cancelled"
        CPV_Eliminar_Borrar = BuildJsonPayload(True, Nothing, "", logs)
        Exit Function
    End If

    logs(3) = "Eliminar_Borrar: deleted"
    CPV_Eliminar_Borrar = BuildJsonPayload(True, Nothing, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "CPV_Eliminar_Borrar: " & Err.Description
    End If
    CPV_Eliminar_Borrar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- CPV_DobleClick_AbrirEdicion ----------------------------------------------
Public Function CPV_DobleClick_AbrirEdicion( _
    ByVal p_HasElegir As Boolean, _
    ByVal p_EditarEnabled As Boolean, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(4)
    p_Error = ""

    On Error GoTo errores

    Dim payload As Object
    Set payload = CreateObject("Scripting.Dictionary")

    Dim action As String

    If p_HasElegir Then
        action = "choose"
        logs(0) = "DobleClick_AbrirEdicion: choose dispatched (cmdElegir visible)"
    Else
        If p_EditarEnabled Then
            action = "edit"
            logs(0) = "DobleClick_AbrirEdicion: edit dispatched (ComandoEditar enabled)"
        Else
            action = "none"
            logs(0) = "DobleClick_AbrirEdicion: no action (cmdElegir hidden, ComandoEditar disabled)"
        End If
    End If

    payload("action") = action

    CPV_DobleClick_AbrirEdicion = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "CPV_DobleClick_AbrirEdicion: " & Err.Description
    End If
    CPV_DobleClick_AbrirEdicion = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function
