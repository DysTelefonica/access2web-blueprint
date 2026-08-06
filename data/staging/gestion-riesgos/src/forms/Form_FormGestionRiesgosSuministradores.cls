Option Compare Database
Option Explicit
Private m_Error As String

Private Sub ComandoActualizar_Click()
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    ' Trasladar nuevos subcontratistas de la jerarquia a TbProyectosEdicionesSuministradores
    If Not m_ObjProyectoActivo Is Nothing Then
        If Not m_ObjProyectoActivo.EdicionActiva Is Nothing Then
            Dim m_IDEdicion As String
            Dim m_idExpediente As String
            m_IDEdicion = m_ObjProyectoActivo.EdicionActiva.IDEdicion
            m_idExpediente = m_ObjProyectoActivo.idExpediente
            SincronizarSuministradoresEnEdicion , m_ObjProyectoActivo.EdicionActiva, m_Error
            
            ' Invalidate the cache so next read rebuilds from hierarchy
            m_ObjProyectoActivo.EdicionActiva.ColSuministradoresVaciar
        End If
    End If
    EstablecerLista m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al ComandoActualizar_Click se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
End Sub


Private Sub Form_Load()
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    EstablecerDatos m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Load se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Function EstablecerDatos( _
                                    Optional ByRef p_Error As String _
                                    ) As String
    On Error GoTo errores
    If m_ObjProyectoActivo Is Nothing Or m_ObjProyectoAlInicio Is Nothing Then
        m_Error = "Es una edición y no se conoce el proyecto activo"
        Err.Raise 1000
    End If
   
    EstablecerLista p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Establecerdatos ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function

Private Function EstablecerLista( _
                                    Optional ByRef p_Error As String _
                                    ) As String
    Dim m_ColSuministradores As Scripting.Dictionary
    Dim m_IDSuministrador As Variant
    Dim m_Suministrador As Suministrador
    
    On Error GoTo errores
     Me.LIstaSuministradores.RowSource = "Suministradores"
    Me.LIstaSuministradores.ColumnCount = 1
    Me.LIstaSuministradores.ColumnWidths = "13cm"
    Me.LIstaSuministradores.RowSourceType = "Lista de valores"
    ' Get IDEdicion from the active edition
    If m_ObjProyectoActivo Is Nothing Then
        p_Error = "No hay proyecto activo"
        Err.Raise 1000
    End If
    m_ObjProyectoActivo.ColSuministradoresVaciar
    m_ObjProyectoActivo.Expediente.SuministradoresVaciar
    Set m_ColSuministradores = m_ObjProyectoActivo.ColSuministradores
    p_Error = m_ObjProyectoActivo.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ColSuministradores Is Nothing Then
       Exit Function
    End If
    ' Clear and set up the list
   For Each m_IDSuministrador In m_ColSuministradores
        Set m_Suministrador = m_ColSuministradores(m_IDSuministrador)
        Me.LIstaSuministradores.AddItem m_Suministrador.Nombre
    Next
    Exit Function
errores:
    
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerLista ha devuelto el error: " & vbNewLine & Err.description
    End If
End Function
