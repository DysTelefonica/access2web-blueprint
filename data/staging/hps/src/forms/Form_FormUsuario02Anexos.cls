Attribute VB_Name = "Form_FormUsuario02Anexos"
Option Compare Database
Option Explicit
Private m_Error As String
Private m_AnexoSeleccionado As AnexoUsuarioHPS


Private Function EstablecerListas(Optional ByRef p_Error As String) As String

    Dim lst As ListBox
    Dim m_IDAnexo As Variant
    Dim m_ObjAnexoUsuarioHPS As AnexoUsuarioHPS
    Dim m_Col As Scripting.Dictionary

    On Error GoTo errores

    Set lst = Me.ListaAnexosActuales
    lst.RowSource = ""
    Set m_ObjUsuarioActivo.colAnexosActuales = Nothing
    Set m_Col = m_ObjUsuarioActivo.colAnexosActuales
    p_Error = m_ObjUsuarioActivo.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_Col Is Nothing Then
        For Each m_IDAnexo In m_Col
            Set m_ObjAnexoUsuarioHPS = m_Col(m_IDAnexo)
            lst.AddItem m_IDAnexo & ";" & m_ObjAnexoUsuarioHPS.NombreAnexo
            Set m_ObjAnexoUsuarioHPS = Nothing
        Next
        If lst.ListCount = 1 Then
            lst.Selected(1) = True
            ListaAnexosActuales_Click
        End If
    End If
    Set lst = Me.ListaAnexosHistoricos
    lst.RowSource = ""
    Set m_ObjUsuarioActivo.colAnexosHistoricos = Nothing
    Set m_Col = m_ObjUsuarioActivo.colAnexosHistoricos
    p_Error = m_ObjUsuarioActivo.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_Col Is Nothing Then

        For Each m_IDAnexo In m_Col
            Set m_ObjAnexoUsuarioHPS = m_Col(m_IDAnexo)
            lst.AddItem m_IDAnexo & ";" & m_ObjAnexoUsuarioHPS.NombreAnexo
            Set m_ObjAnexoUsuarioHPS = Nothing
        Next
        If lst.ListCount = 1 Then
            lst.Selected(1) = True
            ListaAnexosHistoricos_Click
        End If
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerListas ha dado el error: " & vbNewLine & Err.Description
    End If
End Function

' [anexos-batch-hotfix] Snapshot the currently-selected rows of p_ListBox into
' a Collection of Long IDAnexo. The snapshot is an immutable copy: we MUST NOT
' keep reading ItemsSelected during the move, because successful moves
' invalidate the lists (cache + RowSource). Each item is validated as a
' numeric IDAnexo — anything else (header row, blank line) is silently
' dropped so the coordinator never sees a non-numeric id.
Private Function CapturarIDsSeleccionados(ByVal p_ListBox As ListBox) As Collection
    Dim ids As New Collection
    Dim vItem As Variant
    Dim raw As String
    Dim parsed As Long

    If p_ListBox Is Nothing Then
        Set CapturarIDsSeleccionados = ids
        Exit Function
    End If
    For Each vItem In p_ListBox.ItemsSelected
        raw = CStr(Nz(p_ListBox.Column(0, vItem), ""))
        If Len(raw) > 0 Then
            If IsNumeric(raw) Then
                parsed = CLng(raw)
                ids.Add parsed
            End If
        End If
    Next vItem
    Set CapturarIDsSeleccionados = ids
End Function

' [anexos-batch-hotfix] Reset the form selection + per-row command buttons to
' their "nothing selected" state. Called ONLY after a confirmed commit, when
' the list contents have been rebuilt from the DB and the previously selected
' rows no longer exist on either side.
Private Sub LimpiarSeleccionTrasCommit()
    Set m_AnexoSeleccionado = Nothing
    Me.ComandoPasarAHistorico.Enabled = False
    Me.ComandoPasarAActual.Enabled = False
    Me.ComandoEliminarAnexoActual.Enabled = False
    Me.ComandoEliminarAnexoHistorico.Enabled = False
End Sub


Private Sub ComandoAnexar_Click()
    
    Dim m_URL As String
    Dim VarItem As Variant
    Dim dato
    Dim lstActual As ListBox
    Dim m_AnexoUsuarioHPS As AnexoUsuarioHPS
    On Error GoTo errores
    
    m_Error = ""
    Seleccionar True, "Seleccione el archivo a anexar", m_URL, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_URL = "" Then
        Exit Sub
    End If
    Set lstActual = Me.ListaAnexosActuales
    dato = Split(m_URL, ";")
    For Each VarItem In dato
        Set m_AnexoUsuarioHPS = New AnexoUsuarioHPS
        With m_AnexoUsuarioHPS
            
            .URLLocal = VarItem
            .idUsuario = m_ObjUsuarioActivo.ID
            .EsHistorico = "No"
            .Registrar m_Error
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            
            lstActual.AddItem .IDAnexo & ";" & .NombreAnexo
        End With
        Set m_AnexoUsuarioHPS = Nothing
    Next

    EstablecerListas m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Sub
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoAnexar_Click se ha dado el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoEliminarAnexoActual_Click()
    
    On Error GoTo errores
    
    If m_AnexoSeleccionado Is Nothing Then
        m_Error = "No se ha seleccionado ningún anexo"
        Err.Raise 1000
    End If
    If m_AnexoSeleccionado.EsHistorico <> "No" Then
        m_Error = "El anexo seleccionado no es Actual"
        Err.Raise 1000
    End If
    pregunta = MsgBox("¿Desea eliminar definitivamente del Sistema el anexo seleccionado?", _
                    vbExclamation + vbYesNo + vbDefaultButton2, "Eliminar Anexo")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    m_AnexoSeleccionado.Borrar m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    EstablecerListas m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Sub
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método EliminarAnexoActual ha dado el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoEliminarAnexoHistorico_Click()
    On Error GoTo errores
    
    If m_AnexoSeleccionado Is Nothing Then
        m_Error = "Seleccione una nexo de la tabla"
        Err.Raise 1000
    End If
    If m_AnexoSeleccionado.EsHistorico <> "Sí" Then
        m_Error = "El anexo seleccionado no es histórico"
        Err.Raise 1000
    End If
    pregunta = MsgBox("¿Desea eliminar definitivamente del Sistema el anexo seleccionado?", _
                    vbExclamation + vbYesNo + vbDefaultButton2, "Eliminar Anexo")
    If pregunta <> vbYes Then
        Exit Sub
    End If
    m_AnexoSeleccionado.Borrar m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    EstablecerListas m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Sub
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método EliminarAnexoHistorico ha dado el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoPasarAActual_Click()

    ' [anexos-batch-hotfix] Atomic batch move (HISTORICO -> Actual). We capture
    ' the selected IDs ONCE at the start (the previous implementation read
    ' ItemsSelected twice — once for the precheck, once for the move — and
    ' could miss selections on a busy UI). One coordinator call replaces the
    ' two-phase per-row loop. On success we invalidate both caches, rebuild
    ' both lists from the DB, clear m_AnexoSeleccionado, and reset the
    ' per-row commands. On failure we keep the visual state intact; if the
    ' error message indicates compensation failed, we additionally lock the
    ' move/delete commands so no further mutation runs on an inconsistent
    ' DB/filesystem view.
    Dim lst As ListBox
    Dim ids As Collection
    Dim m_Coordinador As AnexoSelectionTransactionCoordinator

    On Error GoTo errores

    Set lst = Me.ListaAnexosHistoricos
    Set ids = CapturarIDsSeleccionados(lst)
    If ids.Count = 0 Then
        m_Error = "No hay ningún anexo histórico seleccionado para pasar a Actual."
        Err.Raise 1000
    End If

    Set m_Coordinador = New AnexoSelectionTransactionCoordinator
    If Not m_Coordinador.MoverLote(ids, False, Nothing, "", m_Error) Then
        If InStr(1, m_Error, "Rollback", vbTextCompare) > 0 _
            Or InStr(1, m_Error, "compensación", vbTextCompare) > 0 _
            Or InStr(1, m_Error, "compensacion", vbTextCompare) > 0 Then
            ' Rollback failed or compensation failed — lock mutations.
            Me.ComandoPasarAHistorico.Enabled = False
            Me.ComandoPasarAActual.Enabled = False
            Me.ComandoEliminarAnexoActual.Enabled = False
            Me.ComandoEliminarAnexoHistorico.Enabled = False
        End If
        Err.Raise 1000
    End If

    Set m_ObjUsuarioActivo.colAnexosActuales = Nothing
    Set m_ObjUsuarioActivo.colAnexosHistoricos = Nothing
    EstablecerListas m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    LimpiarSeleccionTrasCommit
    Indicadores_NotificarMovimiento "h2c", CStr(m_ObjUsuarioActivo.ID)
    Exit Sub

errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoPasarAActual_Click se ha dado el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComandoPasarAHistorico_Click()

    ' [anexos-batch-hotfix] Atomic batch move (Actual -> HISTORICO). See the
    ' symmetry notes in ComandoPasarAActual_Click.
    Dim lst As ListBox
    Dim ids As Collection
    Dim m_Coordinador As AnexoSelectionTransactionCoordinator

    On Error GoTo errores

    Set lst = Me.ListaAnexosActuales
    Set ids = CapturarIDsSeleccionados(lst)
    If ids.Count = 0 Then
        m_Error = "No hay ningún anexo actual seleccionado para pasar a Histórico."
        Err.Raise 1000
    End If

    Set m_Coordinador = New AnexoSelectionTransactionCoordinator
    If Not m_Coordinador.MoverLote(ids, True, Nothing, "", m_Error) Then
        If InStr(1, m_Error, "Rollback", vbTextCompare) > 0 _
            Or InStr(1, m_Error, "compensación", vbTextCompare) > 0 _
            Or InStr(1, m_Error, "compensacion", vbTextCompare) > 0 Then
            Me.ComandoPasarAHistorico.Enabled = False
            Me.ComandoPasarAActual.Enabled = False
            Me.ComandoEliminarAnexoActual.Enabled = False
            Me.ComandoEliminarAnexoHistorico.Enabled = False
        End If
        Err.Raise 1000
    End If

    Set m_ObjUsuarioActivo.colAnexosActuales = Nothing
    Set m_ObjUsuarioActivo.colAnexosHistoricos = Nothing
    EstablecerListas m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    LimpiarSeleccionTrasCommit
    Indicadores_NotificarMovimiento "c2h", CStr(m_ObjUsuarioActivo.ID)
    Exit Sub

errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoPasarAHistorico_Click se ha dado el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub Form_Open(Cancel As Integer)
    
    On Error GoTo errores
    
    Me.AllowEdits = False
    AjustarTamaño Me
    EstablecerControlesDatos Me, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If Application.TempVars("SoloLectura") = "Sí" Then
        Me.ComandoPasarAHistorico.Enabled = False
        Me.ComandoPasarAActual.Enabled = False
        Me.ComandoAnexar.Enabled = False
        Me.ComandoEliminarAnexoActual.Enabled = False
        Me.ComandoEliminarAnexoHistorico.Enabled = False
    Else
        Me.ComandoAnexar.Enabled = True
    End If
    If m_ObjUsuarioActivo Is Nothing Then
        m_Error = "No se ha seleccionado ningún usuario"
        Err.Raise 1000
    End If
    EstablecerListas m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.AllowEdits = True
    Me.ListaAnexosActuales.Enabled = True
    Me.ListaAnexosHistoricos.Enabled = True
    CopiarEstiloBoton Me.ComandoAnexar
    CopiarEstiloBoton Me.ComandoPasarAHistorico
    CopiarEstiloBoton Me.ComandoPasarAActual
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método abrir formulario ha dado el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.AllowEdits = True
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo

End Sub

Private Sub ListaAnexosActuales_Click()

    ' [anexos-batch-hotfix] Selection reconciliation. The previous implementation
    ' checked `m_IDAnexoSeleccionado <> m_AnexoSeleccionado.IDAnexo` and then
    ' re-checked the SAME condition inside the If — meaning the EsHistorico
    ' guard was never evaluated when the IDs matched. After a successful
    ' move, m_AnexoSeleccionado still pointed at the now-historical row while
    ' the user clicked on the ACTUALES list, and the click was accepted
    ' against the stale instance. Fix: re-load whenever the cached EsHistorico
    ' does not match the list that fired the event, even if the IDAnexo is
    ' still equal.
    Dim m_IDAnexoSeleccionado As String

    On Error GoTo errores
    m_Error = ""
    m_IDAnexoSeleccionado = Nz(Me.ListaAnexosActuales.Column(0), "")
    If m_IDAnexoSeleccionado = "" Then
        Set m_AnexoSeleccionado = Nothing
        Me.ComandoPasarAHistorico.Enabled = False
        Exit Sub
    End If
    If m_AnexoSeleccionado Is Nothing Then
        Set m_AnexoSeleccionado = getAnexoUsuarioHPS(m_IDAnexoSeleccionado, m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_AnexoSeleccionado Is Nothing Then
            m_Error = "El Anexo seleccionado no parece estar registrado"
            Err.Raise 1000
        End If
    Else
        If m_IDAnexoSeleccionado <> m_AnexoSeleccionado.IDAnexo _
            Or m_AnexoSeleccionado.EsHistorico <> "No" Then
            Set m_AnexoSeleccionado = getAnexoUsuarioHPS(m_IDAnexoSeleccionado, m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            If m_AnexoSeleccionado Is Nothing Then
                m_Error = "El Anexo seleccionado no parece estar registrado"
                Err.Raise 1000
            End If
        End If
    End If
    If m_EstadoConexion <> "Solo lectura" Then
        Me.ComandoEliminarAnexoActual.Enabled = True
        Me.ComandoPasarAHistorico.Enabled = True
    End If
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ListaAnexosActuales_Click se ha dado el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ListaAnexosActuales_DblClick(Cancel As Integer)

    ' [anexos-batch-hotfix] Autonomous handler — never trust m_AnexoSeleccionado.
    ' We always read the IDAnexo from the listbox that fired the event, reload
    ' the row, and verify EsHistorico='No' before opening the file. The previous
    ' implementation could open a HISTORICO file when the user double-clicked
    ' the ACTUALES list right after a move left m_AnexoSeleccionado stale.
    Dim m_IDAnexo As String
    Dim m_AnexoLocal As AnexoUsuarioHPS

    On Error GoTo errores
    m_Error = ""
    Application.Echo False

    m_IDAnexo = Nz(Me.ListaAnexosActuales.Column(0), "")
    If m_IDAnexo = "" Then
        Application.Echo True
        Exit Sub
    End If

    Set m_AnexoLocal = getAnexoUsuarioHPS(m_IDAnexo, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_AnexoLocal Is Nothing Then
        m_Error = "El Anexo del Usuario seleccionado no parece estar registrado (IDAnexo=" & m_IDAnexo & ")."
        Err.Raise 1000
    End If

    If m_AnexoLocal.EsHistorico <> "No" Then
        m_Error = "El anexo seleccionado no es Actual (EsHistorico='" & m_AnexoLocal.EsHistorico & _
            "'). Refresque la lista antes de abrir el archivo."
        Err.Raise 1000
    End If

    If Not fso.FileExists(m_AnexoLocal.URLAnexo) Then
        m_Error = "No es alcanzable en estos momentos el anexo seleccionado" & vbNewLine & m_AnexoLocal.URLAnexo
        Err.Raise 1000
    End If
    Ejecutar Me.hWnd, "open", m_AnexoLocal.URLAnexo, "", "", 1
    Set m_AnexoLocal = Nothing

    Application.Echo True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ListaAnexosActuales_DblClick se ha dado el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Application.Echo True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ListaAnexosHistoricos_Click()

    ' [anexos-batch-hotfix] Selection reconciliation (H2C mirror of
    ' ListaAnexosActuales_Click). Reload whenever the cached EsHistorico is
    ' not 'Sí' — the previous implementation re-used a stale
    ' m_AnexoSeleccionado from the actual list after a c2h move.
    Dim m_IDAnexoSeleccionado As String

    On Error GoTo errores
    m_Error = ""
    m_IDAnexoSeleccionado = Nz(Me.ListaAnexosHistoricos.Column(0), "")
    If m_IDAnexoSeleccionado = "" Then
        Set m_AnexoSeleccionado = Nothing
        Me.ComandoPasarAHistorico.Enabled = False
        Exit Sub
    End If
    If m_AnexoSeleccionado Is Nothing Then
        Set m_AnexoSeleccionado = getAnexoUsuarioHPS(m_IDAnexoSeleccionado, m_Error)
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        If m_AnexoSeleccionado Is Nothing Then
            m_Error = "El Anexo seleccionado no parece estar registrado"
            Err.Raise 1000
        End If
    Else
        If m_IDAnexoSeleccionado <> m_AnexoSeleccionado.IDAnexo _
            Or m_AnexoSeleccionado.EsHistorico <> "Sí" Then
            Set m_AnexoSeleccionado = getAnexoUsuarioHPS(m_IDAnexoSeleccionado, m_Error)
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            If m_AnexoSeleccionado Is Nothing Then
                m_Error = "El Anexo seleccionado no parece estar registrado"
                Err.Raise 1000
            End If
        End If
    End If
    If m_EstadoConexion <> "Solo lectura" Then
        Me.ComandoEliminarAnexoHistorico.Enabled = True
        Me.ComandoPasarAActual.Enabled = True
    End If
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ListaAnexosHistoricos_Click se ha dado el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ListaAnexoshistoricos_DblClick(Cancel As Integer)

    ' [anexos-batch-hotfix] Autonomous HISTORICO-side handler. Always reload
    ' from the listbox that fired the event, verify EsHistorico='Sí' before
    ' opening.
    Dim m_IDAnexo As String
    Dim m_AnexoLocal As AnexoUsuarioHPS

    On Error GoTo errores
    Application.Echo False

    m_IDAnexo = Nz(Me.ListaAnexosHistoricos.Column(0), "")
    If m_IDAnexo = "" Then
        Application.Echo True
        Exit Sub
    End If

    Set m_AnexoLocal = getAnexoUsuarioHPS(m_IDAnexo, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_AnexoLocal Is Nothing Then
        m_Error = "El Anexo Histórico seleccionado no parece estar registrado (IDAnexo=" & m_IDAnexo & ")."
        Err.Raise 1000
    End If

    If m_AnexoLocal.EsHistorico <> "Sí" Then
        m_Error = "El anexo seleccionado no es histórico (EsHistorico='" & m_AnexoLocal.EsHistorico & _
            "'). Refresque la lista antes de abrir el archivo."
        Err.Raise 1000
    End If

    If Not fso.FileExists(m_AnexoLocal.URLAnexo) Then
        m_Error = "No es alcanzable en estos momentos el anexo seleccionado" & vbNewLine & m_AnexoLocal.URLAnexo
        Err.Raise 1000
    End If
    Ejecutar Me.hWnd, "open", m_AnexoLocal.URLAnexo, "", "", 1
    Set m_AnexoLocal = Nothing

    Application.Echo True
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "El intentar ver el anexo ha dado el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Application.Echo True
    DoCmd.Hourglass False
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub
