VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Configuracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public id As String
Public DiasParaRecordatorioExcel1 As String
Public DiasParaRecordatorioExcel2 As String
Public DiasCancelacionPreMARGA As String 'SON LOS DÍAS PARA CANCELAR EL PROCESO DESDSE FechaEnvioExcel
Public DiasParaRecordatorioRellenoMarga1 As String
Public DiasParaRecordatorioRellenoMarga2 As String
Public DiasCancelacionMARGA As String 'Son los días que da Marga para cancelar el proceso una vez que se le da de alta por tramitador
Public BuzonSeguridad As String
Public EmailDirectorSeguridad As String

Public VersionPlantillasHTML As String
Public VersionPlantillasEXCEL As String
Public AutocancelacionPreMARGA As String
Public AutocancelacionMARGA As String
Public CorreosAutomaticos As String

Private m_sID_Calculado As String

Private m_ObjColRecursosHTML As Scripting.Dictionary
Private m_sColRecursosHTMLTexto As String
Private m_objColCampos As Collection
Public Error As String
Private m_Error As String


Public Property Get ID_Calculado() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    m_sID_Calculado = DameID("TbConfiguracion", "ID", , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    ID_Calculado = m_sID_Calculado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Configuracion.ID_Calculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

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
        .Add "ID"
        .Add "DiasParaRecordatorioExcel1"
        .Add "DiasParaRecordatorioExcel2"
        .Add "DiasCancelacionPreMARGA"
        .Add "DiasParaRecordatorioRellenoMarga1"
        .Add "DiasParaRecordatorioRellenoMarga2"
        .Add "DiasCancelacionMARGA"
        .Add "BuzonSeguridad"
        .Add "EmailDirectorSeguridad"
        .Add "AutocancelacionPreMARGA"
        .Add "AutocancelacionMARGA"
        .Add "CorreosAutomaticos"
        .Add "VersionPlantillasHTML"
        .Add "VersionPlantillasEXCEL"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Configuracion.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "ID" Then
        getPropiedad = Me.id
        Exit Function
    End If
    If m_NombrePropiedad = "DiasParaRecordatorioExcel1" Then
        getPropiedad = Me.DiasParaRecordatorioExcel1
        Exit Function
    End If
    If m_NombrePropiedad = "DiasParaRecordatorioExcel2" Then
        getPropiedad = Me.DiasParaRecordatorioExcel2
        Exit Function
    End If
    If m_NombrePropiedad = "DiasCancelacionPreMARGA" Then
        getPropiedad = Me.DiasCancelacionPreMARGA
        Exit Function
    End If
    If m_NombrePropiedad = "DiasParaRecordatorioRellenoMarga1" Then
        getPropiedad = Me.DiasParaRecordatorioRellenoMarga1
        Exit Function
    End If
    If m_NombrePropiedad = "DiasParaRecordatorioRellenoMarga2" Then
        getPropiedad = Me.DiasParaRecordatorioRellenoMarga2
        Exit Function
    End If
    
    If m_NombrePropiedad = "DiasCancelacionMARGA" Then
        getPropiedad = Me.DiasCancelacionMARGA
        Exit Function
    End If
    If m_NombrePropiedad = "BuzonSeguridad" Then
        getPropiedad = Me.BuzonSeguridad
        Exit Function
    End If
    If m_NombrePropiedad = "EmailDirectorSeguridad" Then
        getPropiedad = Me.EmailDirectorSeguridad
        Exit Function
    End If
    
    If m_NombrePropiedad = "AutocancelacionPreMARGA" Then
        getPropiedad = Me.AutocancelacionPreMARGA
        Exit Function
    End If
    If m_NombrePropiedad = "AutocancelacionMARGA" Then
        getPropiedad = Me.AutocancelacionMARGA
        Exit Function
    End If
    If m_NombrePropiedad = "CorreosAutomaticos" Then
        getPropiedad = Me.CorreosAutomaticos
        Exit Function
    End If
    If m_NombrePropiedad = "VersionPlantillasHTML" Then
        getPropiedad = Me.VersionPlantillasHTML
        Exit Function
    End If
    If m_NombrePropiedad = "VersionPlantillasEXCEL" Then
        getPropiedad = Me.VersionPlantillasEXCEL
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en Configuracion.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Configuracion.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


Public Function setPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "ID" Then
        Me.id = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "DiasParaRecordatorioExcel1" Then
        Me.DiasParaRecordatorioExcel1 = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "DiasParaRecordatorioExcel2" Then
        Me.DiasParaRecordatorioExcel2 = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "DiasCancelacionPreMARGA" Then
        Me.DiasCancelacionPreMARGA = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "DiasParaRecordatorioRellenoMarga1" Then
        Me.DiasParaRecordatorioRellenoMarga1 = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "DiasParaRecordatorioRellenoMarga2" Then
        Me.DiasParaRecordatorioRellenoMarga2 = m_ValorPropiedad
        Exit Function
    End If
    
    If m_NombrePropiedad = "DiasCancelacionMARGA" Then
        Me.DiasCancelacionMARGA = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "BuzonSeguridad" Then
        Me.BuzonSeguridad = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "EmailDirectorSeguridad" Then
        Me.EmailDirectorSeguridad = m_ValorPropiedad
        Exit Function
    End If
    
    If m_NombrePropiedad = "AutocancelacionPreMARGA" Then
        Me.AutocancelacionPreMARGA = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "AutocancelacionMARGA" Then
        Me.AutocancelacionMARGA = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CorreosAutomaticos" Then
        Me.CorreosAutomaticos = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "VersionPlantillasHTML" Then
        Me.VersionPlantillasHTML = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "VersionPlantillasEXCEL" Then
        Me.VersionPlantillasEXCEL = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en Configuracion.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Configuracion.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function

