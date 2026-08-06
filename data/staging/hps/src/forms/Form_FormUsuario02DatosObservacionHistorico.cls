Attribute VB_Name = "Form_FormUsuario02DatosObservacionHistorico"
Option Compare Database
Option Explicit

Private m_Error As String




Private Sub CmdSalir_Click()
    
   
    DoCmd.Close acForm, Me.Name, acSaveNo
    
    
End Sub

Private Sub Form_Load()
    
    Dim m_Titulo As String
    Dim m_Argumento As String
    Dim m_ObsHistorica As ObservacionHistorica
    On Error GoTo errores
    
    m_Error = ""
    AjustarTamaño Me
    Me.AllowEdits = False
    m_Argumento = Nz(Me.OpenArgs(), "")
    Set m_ObsHistorica = Constructor.getObservacionHistorica(m_Argumento, m_Error)
    If m_Error <> "" Then Err.Raise 1000
    
    If m_ObsHistorica Is Nothing Then
        m_Error = "Se ha abierto con parámetros insuficientes"
        Err.Raise 1000
    End If
    If IsDate(m_ObsHistorica.Fecha) Then Me.Fecha = m_ObsHistorica.Fecha
    If m_ObsHistorica.TIPO <> "" Then Me.TIPO = m_ObsHistorica.TIPO
    If m_ObsHistorica.Observacion <> "" Then Me.Observacion = m_ObsHistorica.Observacion
    
    CopiarEstiloBoton Me.cmdSalir, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.EncabezadoDelFormulario.BackColor = m_ObjEntorno.CabeceraBackColor
    Me.PieDelFormulario.BackColor = m_ObjEntorno.CabeceraBackColor
    EstablecerControlesDatos Me, m_Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Me.Caption = "Observación Histórica"
    Me.lblTitulo.Caption = Me.Caption
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al abrir el formulario se ha producido el error: " & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub
