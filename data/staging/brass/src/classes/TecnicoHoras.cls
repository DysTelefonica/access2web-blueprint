VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TecnicoHoras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public Alias As String
Public Fecha As String
Private m_sHlab As String
Private m_sHExt As String

Private m_sDiaDeLaSemana As String
Private m_sTipoDia As String
Private m_eLibranza As EnumSino
Private m_eFestivo As EnumSino

Public ObservacionesFestivo As String

Private m_sActividadesTextoLab As String
Private m_sActividadesTextoExt As String
Private m_ObjColActividades As Scripting.Dictionary


Private m_objColCampos As Collection
Public Error As String
Private m_Error As String


Public Property Get DiaDeLaSemana() As String
    
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If Me.Fecha = "" Then
        Exit Property
    End If
    m_sDiaDeLaSemana = WeekdayName(Weekday(CDate(Me.Fecha) - 1))
    
    
    DiaDeLaSemana = m_sDiaDeLaSemana
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método TecnicoHoras.DiaDeLaSemana ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Get TipoDia() As String
    
    Dim m_DiaDeLaSemana As String
    Dim m_Libranza As EnumSino
    Dim m_Fiesta As EnumSino
    
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If Me.Fecha = "" Then
        Exit Property
    End If
    If Me.Alias = "" Then
        Exit Property
    End If
    m_DiaDeLaSemana = Me.DiaDeLaSemana
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_DiaDeLaSemana = "Sábado" Or m_DiaDeLaSemana = "Domingo" Then
        m_sTipoDia = "No Laborable"
        TipoDia = m_sTipoDia
        Exit Property
    End If
    m_Fiesta = Me.Festivo
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Fiesta = EnumSino.Sí Then
        Me.ObservacionesFestivo = Festivo
        m_sTipoDia = "No Laborable"
        TipoDia = m_sTipoDia
        Exit Property
    End If
    m_sTipoDia = "Laborable"
    
    m_Libranza = Me.LIbranza
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Libranza = EnumSino.Sí Then
        Me.ObservacionesFestivo = "Libranza"
        
    End If
    TipoDia = m_sTipoDia
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método TecnicoHoras.TipoDia ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Get LIbranza() As EnumSino
    
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If m_eLibranza <> 0 Then
        LIbranza = m_eLibranza
        Exit Property
    End If
    If Me.Fecha = "" Then
        Exit Property
    End If
    If Me.Alias = "" Then
        Exit Property
    End If
    m_eLibranza = EsLibranza(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    LIbranza = m_eLibranza
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método TecnicoHoras.Libranza ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Get Festivo() As EnumSino
    
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If m_eFestivo <> 0 Then
        Festivo = m_eFestivo
        Exit Property
    End If
    If Me.Fecha = "" Then
        Exit Property
    End If
    
    m_eFestivo = EsFestivo(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Festivo = m_eFestivo
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método TecnicoHoras.Festivo ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property



Public Property Get ColActividades() As Scripting.Dictionary
    
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If Not m_ObjColActividades Is Nothing Then
        Set ColActividades = m_ObjColActividades
        Exit Property
    End If
    Set m_ObjColActividades = Constructor.getActividadesTecnicoDia(Me.Alias, Me.Fecha, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set ColActividades = m_ObjColActividades
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método TecnicoHoras.ColActividades ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Set ColActividades(ByVal objNewValue As Variant)

    Set m_ObjColActividades = objNewValue

End Property

Public Property Get Hlab() As String
    
    Dim m_Col As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_ObjActividad As Actividad
    Dim m_HLAb As Double
    
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    Set m_Col = Me.ColActividades
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        m_sHlab = "0"
        Hlab = m_sHlab
        Exit Property
    End If
    For Each m_ID In m_Col
        Set m_ObjActividad = m_Col(m_ID)
        With m_ObjActividad
            If IsNumeric(.HorasLaborables) Then
                m_HLAb = m_HLAb + CDbl(.HorasLaborables)
            End If
        End With
        Set m_ObjActividad = Nothing
    Next
    m_sHlab = CStr(m_HLAb)
    Hlab = m_sHlab
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método TecnicoHoras.Hlab ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Get HExt() As String

    Dim m_Col As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_ObjActividad As Actividad
    Dim m_HExt As Double
    
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    Set m_Col = Me.ColActividades
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        m_sHExt = "0"
        HExt = m_sHExt
        Exit Property
    End If
    For Each m_ID In m_Col
        Set m_ObjActividad = m_Col(m_ID)
        With m_ObjActividad
            If IsNumeric(.HorasExtras) Then
                m_HExt = m_HExt + CDbl(.HorasExtras)
            End If
        End With
        Set m_ObjActividad = Nothing
    Next
    m_sHExt = CStr(m_HExt)
    HExt = m_sHExt
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método TecnicoHoras.HExt ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Get ActividadesTextoLab() As String
    
    Dim m_Col As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_ObjActividad As Actividad
    
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_sActividadesTextoLab <> "" Then
        ActividadesTextoLab = m_sActividadesTextoLab
        Exit Property
    End If
    Set m_Col = Me.ColActividades
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Exit Property
    End If
    For Each m_ID In m_Col
        Set m_ObjActividad = m_Col(m_ID)
        If IsNumeric(m_ObjActividad.HorasLaborables) Then
            If CDbl(m_ObjActividad.HorasLaborables) > 0 Then
                If m_sActividadesTextoLab = "" Then
                    m_sActividadesTextoLab = m_ObjActividad.IDActividad
                Else
                    m_sActividadesTextoLab = m_sActividadesTextoLab & ";" & m_ObjActividad.IDActividad
                End If
            End If
        End If
        
        Set m_ObjActividad = Nothing
    Next
    ActividadesTextoLab = m_sActividadesTextoLab
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método TecnicoHoras.ActividadesTextoLab ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Get ActividadesTextoExt() As String
    
    Dim m_Col As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_ObjActividad As Actividad
    
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_sActividadesTextoExt <> "" Then
        ActividadesTextoExt = m_sActividadesTextoExt
        Exit Property
    End If
    Set m_Col = Me.ColActividades
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Exit Property
    End If
    For Each m_ID In m_Col
        Set m_ObjActividad = m_Col(m_ID)
        If IsNumeric(m_ObjActividad.HorasExtras) Then
            If CDbl(m_ObjActividad.HorasExtras) > 0 Then
                If m_sActividadesTextoExt = "" Then
                    m_sActividadesTextoExt = m_ObjActividad.IDActividad
                Else
                    m_sActividadesTextoExt = m_sActividadesTextoExt & ";" & m_ObjActividad.IDActividad
                End If
            End If
        End If
        
        Set m_ObjActividad = Nothing
    Next
    ActividadesTextoExt = m_sActividadesTextoExt
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método TecnicoHoras.ActividadesTextoExt ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Private Function EsLibranza(Optional ByRef p_Error As String) As EnumSino
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If Me.Fecha = "" Then
        Exit Function
    End If
    If Me.Alias = "" Then
        Exit Function
    End If
    m_SQL = "SELECT TbTecnicosAusencias.Alias, TbTecnicosAusencias.FechaLibranza " & _
            "FROM TbTecnicosAusencias " & _
            "WHERE Alias='" & Me.Alias & "' AND FechaLibranza=#" & Format(Me.Fecha, "mm/dd/yyyy") & "#;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            EsLibranza = EnumSino.No
        Else
            EsLibranza = EnumSino.Sí
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método TecnicoHoras.EsLibranza ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function

Private Function EsFestivo(Optional ByRef p_Error As String) As EnumSino
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If Me.Fecha = "" Then
        Exit Function
    End If
    
    m_SQL = "SELECT TbTecnicosFiestas.FechaFiesta " & _
            "FROM TbTecnicosFiestas " & _
            "WHERE FechaFiesta=#" & Format(Me.Fecha, "mm/dd/yyyy") & "#;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            EsFestivo = EnumSino.No
        Else
            EsFestivo = EnumSino.Sí
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método TecnicoHoras.EsFestivo ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function
Public Property Get ColCampos() As Collection

    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If Not m_objColCampos Is Nothing Then
        Set ColCampos = m_objColCampos
        Exit Property
    End If
    Set m_objColCampos = New Collection
    With m_objColCampos
        .Add "Alias"
        .Add "Fecha"
        .Add "Hlab"
        .Add "HExt"
        .Add "DiaDeLaSemana"
        .Add "TipoDia"
        .Add "Libranza"
        .Add "Festivo"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método TecnicoHoras.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "Alias" Then
        getPropiedad = Me.Alias
        Exit Function
    End If
    If m_NombrePropiedad = "Fecha" Then
        getPropiedad = Me.Fecha
        Exit Function
    End If
    
    p_Error = "Propiedad no reconocida en TecnicoHoras.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método TecnicoHoras.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "Alias" Then
        Me.Alias = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Fecha" Then
        Me.Fecha = m_ValorPropiedad
        Exit Function
    End If
    
    p_Error = "Propiedad no reconocida en TecnicoHoras.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método TecnicoHoras.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function

