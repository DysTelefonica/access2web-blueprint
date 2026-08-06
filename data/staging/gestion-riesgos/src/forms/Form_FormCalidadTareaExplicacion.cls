Attribute VB_Name = "Form_FormCalidadTareaExplicacion"
Option Compare Database
Option Explicit
Private m_NodoSeleccionado As MSComctlLib.Node
Private m_Error As String

Private Sub Form_Open(Cancel As Integer)

    Dim m_Titulo As String
    Dim m_Explicacion As String
    Dim m_valor As String
    Dim dato
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    m_Error = ""
    Me.lblTitulo.Caption = "TÍTULO DEL APARTADO"
    Me.lblExplicacion.Caption = "EXPLICACIÓN DEL APARTADO"

    Set m_NodoSeleccionado = Form_FormCalidadTareas.m_NodoSeleccionado
    If m_NodoSeleccionado Is Nothing Then
        m_Error = "No se sabe el nodo seleccionado"
        Err.Raise 1000
    End If
    If m_ObjEntorno.ColExplicacionesArbolTareas.Exists(m_NodoSeleccionado.key) Then
        m_valor = m_ObjEntorno.ColExplicacionesArbolTareas(m_NodoSeleccionado.key)
        dato = Split(m_valor, "|")
        m_Titulo = dato(0)
        m_Explicacion = dato(1)
        If Len(m_Titulo) < 50 Then
            Me.lblTitulo.FontSize = 20
        Else
            If Len(m_Titulo) > 50 And Len(m_Titulo) < 60 Then
                Me.lblTitulo.FontSize = 15
            Else
                Me.lblTitulo.FontSize = 12
            End If
        End If

        Me.lblTitulo.Caption = m_Titulo
        Me.lblExplicacion.Caption = m_Explicacion
        'Debug.Print m_Titulo, Len(m_Titulo), "FontSize" & Me.lblTitulo.FontSize
    End If
    Me.Caption = m_ObjEntorno.TituloUsuarioConectado
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Open se ha producido el error n: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        CorreoAlAdministrador m_Error
        pregunta = MsgBox(m_Error, vbCritical, "Error")
    Else
        pregunta = MsgBox(m_Error, vbExclamation, "Advertencia")
    End If
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub