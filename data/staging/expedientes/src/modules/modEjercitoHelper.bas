Attribute VB_Name = "modEjercitoHelper"
Option Compare Database
Option Explicit

' modEjercitoHelper — REWORK (2026-06-26)
' Pure-data helpers for Form_FormEjercitosGestion.
'
' Implements access-vba-e2e-methodology rule #1 (forms are thin UI wiring),
' rule #2 (honest helper signatures), rule #5 (per-module prefix: Ejercito_*),
' and rule #9 (no `ByRef p_Form`).
'
' New design (5 helpers, mirror of modComercialHelper):
'   1. Ejercito_Abrir_Inicializar(p_EsAdministrador, p_HasOpenArgs, p_Error)
'   2. Ejercito_Buscar_Listar(p_Ejercitos, p_Filter, p_Error)
'      Filter field: Ejercito (text name only — legacy listbox is 3-col,
'      includes DESCRIPCION)
'   3. Ejercito_Seleccionar_Cargar(p_IDSeleccionado, p_Ejercitos, p_EsAdministrador, p_Error)
'   4. Ejercito_Eliminar_Borrar(p_Ejercito, p_PromptResult, p_Error)
'   5. Ejercito_DobleClick_AbrirEdicion(p_HasElegir, p_EditarEnabled, p_Error)
'
' UI orchestration that stays in the form:
'   - Alta button: DoCmd.OpenForm "FormEjercito"
'   - Edición button: DoCmd.OpenForm "FormEjercito" with m_ObjEjercitoActivo set
'   - Limpiar button: Me.Ejercito = Null
'
' Telefonica D&S convention (vba-access §1.4.1): every Public Function ends with
' `Optional ByRef p_Error As String` as the LAST parameter.

' === Module-level constants (all at top per vba-access §10.1) ============================

' Legacy listbox shows ID + Ejercito + DESCRIPCION (3-col shape — unique among the 3).
Private Const EJERCITO_HEADERS As String = "IDEjercito;Ejercito;DESCRIPCIÓN"

Private Const EJERCITO_FIELDS_SEP As String = ";"

Private Const EJERCITO_MAX_ROWS As Long = 5000

Private Const EJERCITO_TEST_ID_BASE As Long = 900520


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

' --- Ejercito_SerializarEntidad -----------------------------------------------------
Private Function Ejercito_SerializarEntidad(ByVal p_Entidad As Object) As Object
    If p_Entidad Is Nothing Then
        Set Ejercito_SerializarEntidad = Nothing
        Exit Function
    End If

    Dim out As Object
    Set out = CreateObject("Scripting.Dictionary")

    If p_Entidad.Exists("IDEjercito") Then
        out("IDEjercito") = CStr(p_Entidad("IDEjercito"))
    Else
        out("IDEjercito") = ""
    End If

    If p_Entidad.Exists("Ejercito") Then
        out("Ejercito") = CStr(p_Entidad("Ejercito"))
    Else
        out("Ejercito") = ""
    End If

    If p_Entidad.Exists("DESCRIPCION") Then
        out("DESCRIPCION") = CStr(p_Entidad("DESCRIPCION"))
    Else
        out("DESCRIPCION") = ""
    End If

    Set Ejercito_SerializarEntidad = out
End Function

' --- Ejercito_BuildRowLine ---------------------------------------------------------
' 3-col rowSource: ID + Ejercito + DESCRIPCION.
Private Function Ejercito_BuildRowLine( _
    ByVal p_ID As String, _
    ByVal p_Nombre As String, _
    ByVal p_Desc As String _
) As String
    Dim safeNombre As String
    safeNombre = Replace(p_Nombre, EJERCITO_FIELDS_SEP, ":")

    Dim safeDesc As String
    safeDesc = Replace(p_Desc, EJERCITO_FIELDS_SEP, ":")

    Ejercito_BuildRowLine = p_ID & EJERCITO_FIELDS_SEP & safeNombre & _
                            EJERCITO_FIELDS_SEP & safeDesc
End Function


' === Public API =========================================================================

' --- Ejercito_Abrir_Inicializar ---------------------------------------------------
Public Function Ejercito_Abrir_Inicializar( _
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

    Ejercito_Abrir_Inicializar = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "Ejercito_Abrir_Inicializar: " & Err.Description
    End If
    Ejercito_Abrir_Inicializar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- Ejercito_Buscar_Listar --------------------------------------------------------
Public Function Ejercito_Buscar_Listar( _
    ByVal p_Ejercitos As Object, _
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

    If p_Ejercitos Is Nothing Then
        logs(0) = "Buscar_Listar: p_Ejercitos is Nothing -> 0 rows"
        payload("rowSource") = ""
        payload("count") = 0
        Ejercito_Buscar_Listar = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    Dim filterText As String
    filterText = "" & p_Filter

    Dim k As Variant
    Dim entity As Object
    Dim currentId As String
    Dim currentNombre As String
    Dim currentDesc As String

    For Each k In p_Ejercitos.Keys
        If rowCount >= EJERCITO_MAX_ROWS Then
            logs(3) = "Buscar_Listar: hit MAX_ROWS=" & EJERCITO_MAX_ROWS & ", truncating"
            Exit For
        End If

        Set entity = p_Ejercitos(k)
        If entity Is Nothing Then
            ' Skip malformed entries.
        Else
            If entity.Exists("IDEjercito") Then
                currentId = CStr(entity("IDEjercito"))
            Else
                currentId = ""
            End If

            If entity.Exists("Ejercito") Then
                currentNombre = CStr(entity("Ejercito"))
            Else
                currentNombre = ""
            End If

            If entity.Exists("DESCRIPCION") Then
                currentDesc = CStr(entity("DESCRIPCION"))
            Else
                currentDesc = ""
            End If

            If Len(filterText) > 0 Then
                If InStr(1, currentNombre, filterText, vbTextCompare) = 0 Then
                    ' No match — skip without counting.
                Else
                    If Len(rowSource) > 0 Then
                        rowSource = rowSource & vbCrLf
                    End If
                    rowSource = rowSource & Ejercito_BuildRowLine(currentId, currentNombre, currentDesc)
                    rowCount = rowCount + 1
                End If
            Else
                ' No filter — include everything.
                If Len(rowSource) > 0 Then
                    rowSource = rowSource & vbCrLf
                End If
                rowSource = rowSource & Ejercito_BuildRowLine(currentId, currentNombre, currentDesc)
                rowCount = rowCount + 1
            End If
        End If
        Set entity = Nothing
    Next k

    payload("rowSource") = rowSource
    payload("count") = rowCount

    logs(0) = "Buscar_Listar: filterText=" & filterText & ", rows=" & rowCount
    logs(1) = "Buscar_Listar: rowSource length=" & Len(rowSource)

    Ejercito_Buscar_Listar = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "Ejercito_Buscar_Listar: " & Err.Description
    End If
    Ejercito_Buscar_Listar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- Ejercito_Seleccionar_Cargar ---------------------------------------------------
Public Function Ejercito_Seleccionar_Cargar( _
    ByVal p_IDSeleccionado As String, _
    ByVal p_Ejercitos As Object, _
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
        Ejercito_Seleccionar_Cargar = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    Dim entity As Object
    If p_Ejercitos Is Nothing Then
        Set entity = Nothing
    Else
        If p_Ejercitos.Exists(selectedId) Then
            Set entity = p_Ejercitos(selectedId)
        Else
            Set entity = Nothing
        End If
    End If

    If entity Is Nothing Then
        logs(0) = "Seleccionar_Cargar: id=" & selectedId & " not found in collection"
        Set payload("entity") = Nothing
        payload("enableEditar") = False
        payload("enableEliminar") = False
        Ejercito_Seleccionar_Cargar = BuildJsonPayload(True, payload, "", logs)
        Exit Function
    End If

    Dim serializedEntity As Object
    Set serializedEntity = Ejercito_SerializarEntidad(entity)

    Set payload("entity") = serializedEntity
    payload("enableEditar") = True
    If p_EsAdministrador Then
        payload("enableEliminar") = True
    Else
        payload("enableEliminar") = False
    End If

    logs(0) = "Seleccionar_Cargar: id=" & selectedId
    logs(1) = "Seleccionar_Cargar: isAdmin=" & CStr(p_EsAdministrador)

    Ejercito_Seleccionar_Cargar = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "Ejercito_Seleccionar_Cargar: " & Err.Description
    End If
    Ejercito_Seleccionar_Cargar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- Ejercito_Eliminar_Borrar -----------------------------------------------------
Public Function Ejercito_Eliminar_Borrar( _
    ByVal p_Ejercito As Object, _
    Optional ByRef p_PromptResult As Long = 0, _
    Optional ByRef p_Error As String _
) As String
    Dim logs() As String
    logs = TestingCore_InitLogs(6)
    p_Error = ""

    On Error GoTo errores

    If p_Ejercito Is Nothing Then
        p_Error = "Ejercito_Eliminar_Borrar: p_Ejercito is Nothing"
        Ejercito_Eliminar_Borrar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    Dim idEjercito As String
    Dim ejercitoNombre As String

    If p_Ejercito.Exists("IDEjercito") Then
        idEjercito = CStr(p_Ejercito("IDEjercito"))
    Else
        idEjercito = ""
    End If

    If p_Ejercito.Exists("Ejercito") Then
        ejercitoNombre = CStr(p_Ejercito("Ejercito"))
    Else
        ejercitoNombre = ""
    End If

    If Len(idEjercito) = 0 Then
        p_Error = "Ejercito_Eliminar_Borrar: missing IDEjercito"
        Ejercito_Eliminar_Borrar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    logs(0) = "Eliminar_Borrar: id=" & idEjercito & ", name=" & ejercitoNombre

    Dim promptResult As Long
    If p_PromptResult <> 0 Then
        promptResult = p_PromptResult
    Else
        promptResult = MsgBox("¿Desea realmente borrar al Ejército seleccionado?", _
                              vbExclamation + vbYesNo + vbDefaultButton2, "Eliminar")
    End If

    logs(2) = "Eliminar_Borrar: promptResult=" & promptResult

    If promptResult <> vbYes Then
        logs(3) = "Eliminar_Borrar: cancelled by user"
        Ejercito_Eliminar_Borrar = BuildJsonPayload(True, Nothing, "", logs)
        Exit Function
    End If

    Dim m_EjercitoOP As New EjercitoOperaciones

    Dim wrapped As New Ejercito
    If p_Ejercito.Exists("IDEjercito") Then
        wrapped.IDEjercito = CStr(p_Ejercito("IDEjercito"))
    End If
    If p_Ejercito.Exists("Ejercito") Then
        wrapped.Ejercito = CStr(p_Ejercito("Ejercito"))
    End If
    If p_Ejercito.Exists("DESCRIPCION") Then
        wrapped.DESCRIPCION = CStr(p_Ejercito("DESCRIPCION"))
    End If

    Set m_EjercitoOP.Ejercito = wrapped

    Dim daoErr As String
    Dim daoResult As String
    daoResult = Helper_EntidadCRUD.EliminarEntidadGenerico( _
        m_EjercitoOP, "Ejercito", wrapped, "", daoErr)
    If daoErr <> "" Then
        p_Error = daoErr
        logs(3) = "Eliminar_Borrar: DAO error=" & daoErr
        Ejercito_Eliminar_Borrar = BuildJsonPayload(False, Nothing, p_Error, logs)
        Exit Function
    End If

    If Len(daoResult) > 0 And daoResult <> "OK" Then
        logs(3) = "Eliminar_Borrar: DAO returned '" & daoResult & "' — treating as cancelled"
        Ejercito_Eliminar_Borrar = BuildJsonPayload(True, Nothing, "", logs)
        Exit Function
    End If

    logs(3) = "Eliminar_Borrar: deleted"
    Ejercito_Eliminar_Borrar = BuildJsonPayload(True, Nothing, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "Ejercito_Eliminar_Borrar: " & Err.Description
    End If
    Ejercito_Eliminar_Borrar = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function

' --- Ejercito_DobleClick_AbrirEdicion ----------------------------------------------
Public Function Ejercito_DobleClick_AbrirEdicion( _
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

    Ejercito_DobleClick_AbrirEdicion = BuildJsonPayload(True, payload, "", logs)
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "Ejercito_DobleClick_AbrirEdicion: " & Err.Description
    End If
    Ejercito_DobleClick_AbrirEdicion = BuildJsonPayload(False, Nothing, p_Error, logs)
End Function
