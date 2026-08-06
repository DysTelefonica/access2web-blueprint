VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TipoDia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Fecha As String
Public Alias As String

Private m_ObjTecnico As Tecnico
Private m_sDiaDeLaSemana As String
Private m_sTipoDia As String
Public ObservacionesLibranza As String
Public ObservacionesFestivo As String

Public Error As String
Private m_Error As String

Public Property Get Tecnico() As Tecnico
    
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If Not m_ObjTecnico Is Nothing Then
        Set Tecnico = m_ObjTecnico
        Exit Property
    End If
    If Me.Alias = "" Then
        Exit Property
    End If
    Set m_ObjTecnico = Constructor.getTecnico(Me.Alias, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Tecnico = m_ObjTecnico
    
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método TipoDia.Tecnico ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Get DiaDeLaSemana() As String
    
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_sDiaDeLaSemana <> "" Then
        DiaDeLaSemana = m_sDiaDeLaSemana
        Exit Property
    End If
    If Not IsDate(Me.Fecha) Then
        Exit Property
    End If
    m_sDiaDeLaSemana = WeekdayName(Weekday(CDate(Me.Fecha) - 1))
    
    DiaDeLaSemana = m_sDiaDeLaSemana
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método TipoDia.DiaDeLaSemana ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Get TipoDia() As String
    
    Dim m_EsDiaFestivo As String
    Dim m_EsDiaLibranza As String
    
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_sTipoDia <> "" Then
        TipoDia = m_sTipoDia
        Exit Property
    End If
    If Me.DiaDeLaSemana = "" Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    If Me.DiaDeLaSemana = "Sábado" Or Me.DiaDeLaSemana = "Domingo" Then
        m_sTipoDia = "No Laborable"
        TipoDia = m_sTipoDia
        Exit Property
    End If
    m_EsDiaFestivo = EsDiaFestivo(Me.Fecha, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_EsDiaFestivo <> "" Then
        Me.ObservacionesFestivo = m_EsDiaFestivo
        m_sTipoDia = "No Laborable"
        TipoDia = m_sTipoDia
        Exit Property
    End If
    m_sTipoDia = "Laborable"
    m_EsDiaLibranza = EsDiaLibranza(Me.Alias, Me.Fecha, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_EsDiaLibranza <> "" Then
        Me.ObservacionesLibranza = m_EsDiaLibranza
    End If
    
    TipoDia = m_sTipoDia
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método TipoDia.TipoDia ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error

End Property


    
Private Function EsDiaFestivo(p_Fecha As String, Optional ByRef p_Error As String) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    
    On Error GoTo errores
    m_SQL = "SELECT * FROM TbTecnicosFiestas " & _
            "Where FechaFiesta=#" & Format(p_Fecha, "mm/dd/yyyy") & "#;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    If rcdDatos.EOF Then
        EsDiaFestivo = ""
    Else
        EsDiaFestivo = Nz(rcdDatos.Fields("Observaciones"), "")
    End If
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EsDiaFestivo ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function

Private Function EsDiaLibranza(p_Alias As String, p_Fecha As String, Optional ByRef p_Error As String) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    
    On Error GoTo errores
    m_SQL = "SELECT * FROM TbTecnicosAusencias " & _
            "Where Alias='" & p_Alias & "' AND FechaLibranza=#" & Format(p_Fecha, "mm/dd/yyyy") & "#;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    If rcdDatos.EOF Then
        EsDiaLibranza = ""
    Else
        EsDiaLibranza = Nz(rcdDatos.Fields("Motivo"), "")
    End If
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EsDiaLibranza ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function
