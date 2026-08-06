VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SolicitudFechas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public IDSolicitud As String
Public FechaEnvioExcel As String
Public FechaRecepcionExcel As String
Public FechaTramitacionAltaMarga As String
Public FechaEnvioDPS As String
Public FechaEnvioONS As String
Public FechaRegistroEnHPS As String

Public FechaCorreoRecordatorioExcel1 As String
Public FechaCorreoRecordatorioExcel2 As String
Public FechaAutocancelacion As String
Public FechaCorreoRecordatorioRellenoMarga1 As String
Public FechaCorreoRecordatorioRellenoMarga2 As String
Public FechaDesestimado As String

Public FechaPrevistaCorreoRecordatorioExcel1 As String
Public FechaPrevistaCorreoRecordatorioExcel2 As String
Public FechaPrevistaCorreoRecordatorioRellenoMarga1 As String
Public FechaPrevistaCorreoRecordatorioRellenoMarga2 As String
Public FechaPrevistaAutocancelacionPreMarga As String
Public FechaPrevistaAutocancelacionMarga As String
Public FechaCancelado As String




Private m_sFechaPrevistaCorreoRecordatorioExcel1Calculado As String
Private m_sFechaPrevistaCorreoRecordatorioExcel2Calculado As String
Private m_sFechaPrevistaCorreoRecordatorioRellenoMarga1Calculado As String
Private m_sFechaPrevistaCorreoRecordatorioRellenoMarga2Calculado As String
Private m_sFechaPrevistaAutocancelacionPreMargaCalculado As String
Private m_sFechaPrevistaAutocancelacionMargaCalculado As String



Private m_eEstadoCalculado As EnumEstado
Private m_eEstadoEnum As EnumEstado
Private m_sEstadoCalculadoTexto As String
Private m_sDiasParaCancelacionPreMarga As String
Private m_sDiasParaCancelacionMarga As String

Private m_sDiasParaCancelacion As String

Private m_sDiasDesdeEnvioExcel As String
Private m_sDiasDesdeAltaEnMarga As String
Private m_objSolicitud As solicitud

Private m_objColCampos As Collection
Public Error As String
Private m_Error As String

Public Property Get FechaPrevistaCorreoRecordatorioExcel1Calculado() As String
    
    Dim dblDiasTotales As Double
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If m_sFechaPrevistaCorreoRecordatorioExcel1Calculado <> "" Then
        FechaPrevistaCorreoRecordatorioExcel1Calculado = m_sFechaPrevistaCorreoRecordatorioExcel1Calculado
        Exit Property
    End If
    If Not IsDate(Me.FechaEnvioExcel) Then
        Exit Property
    End If
    If Not IsNumeric(m_ObjEntorno.DiasParaRecordatorioExcel1) Then
        Exit Property
    End If
    dblDiasTotales = CDbl(m_ObjEntorno.DiasParaRecordatorioExcel1)
    m_sFechaPrevistaCorreoRecordatorioExcel1Calculado = DateAdd("d", CDate(Me.FechaEnvioExcel), dblDiasTotales)
    FechaPrevistaCorreoRecordatorioExcel1Calculado = m_sFechaPrevistaCorreoRecordatorioExcel1Calculado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método SolicitudFechas.FechaPrevistaCorreoRecordatorioExcel1Calculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get FechaPrevistaCorreoRecordatorioExcel2Calculado() As String
    
    
    Dim dblDiasTotales As Double
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If m_sFechaPrevistaCorreoRecordatorioExcel2Calculado <> "" Then
        FechaPrevistaCorreoRecordatorioExcel2Calculado = m_sFechaPrevistaCorreoRecordatorioExcel2Calculado
        Exit Property
    End If
    If Not IsDate(Me.FechaEnvioExcel) Then
        Exit Property
    End If
    If Not IsNumeric(m_ObjEntorno.DiasParaRecordatorioExcel1) Then
        Exit Property
    End If
    If Not IsNumeric(m_ObjEntorno.DiasParaRecordatorioExcel2) Then
        Exit Property
    End If
    dblDiasTotales = CDbl(m_ObjEntorno.DiasParaRecordatorioExcel1) + CDbl(m_ObjEntorno.DiasParaRecordatorioExcel2)
    m_sFechaPrevistaCorreoRecordatorioExcel2Calculado = DateAdd("d", CDate(Me.FechaEnvioExcel), dblDiasTotales)
    FechaPrevistaCorreoRecordatorioExcel2Calculado = m_sFechaPrevistaCorreoRecordatorioExcel2Calculado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método SolicitudFechas.FechaPrevistaCorreoRecordatorioExcel2Calculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get FechaPrevistaCorreoRecordatorioRellenoMarga1Calculado() As String
    
    Dim dblDiasTotales As Double
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If m_sFechaPrevistaCorreoRecordatorioRellenoMarga1Calculado <> "" Then
        FechaPrevistaCorreoRecordatorioRellenoMarga1Calculado = m_sFechaPrevistaCorreoRecordatorioRellenoMarga1Calculado
        Exit Property
    End If
    If Not IsDate(Me.FechaTramitacionAltaMarga) Then
        Exit Property
    End If
    If Not IsNumeric(m_ObjEntorno.DiasParaRecordatorioRellenoMarga1) Then
        Exit Property
    End If
    dblDiasTotales = CDbl(m_ObjEntorno.DiasParaRecordatorioRellenoMarga1)
    m_sFechaPrevistaCorreoRecordatorioRellenoMarga1Calculado = DateAdd("d", CDate(Me.FechaTramitacionAltaMarga), dblDiasTotales)
    FechaPrevistaCorreoRecordatorioRellenoMarga1Calculado = m_sFechaPrevistaCorreoRecordatorioRellenoMarga1Calculado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método SolicitudFechas.FechaPrevistaCorreoRecordatorioRellenoMarga1Calculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get FechaPrevistaCorreoRecordatorioRellenoMarga2Calculado() As String
    
    Dim dblDiasTotales As Double
   
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If m_sFechaPrevistaCorreoRecordatorioRellenoMarga2Calculado <> "" Then
        FechaPrevistaCorreoRecordatorioRellenoMarga2Calculado = m_sFechaPrevistaCorreoRecordatorioRellenoMarga2Calculado
        Exit Property
    End If
    If Not IsDate(Me.FechaTramitacionAltaMarga) Then
        Exit Property
    End If
    If Not IsNumeric(m_ObjEntorno.DiasParaRecordatorioRellenoMarga1) Then
        Exit Property
    End If
    If Not IsNumeric(m_ObjEntorno.DiasParaRecordatorioRellenoMarga2) Then
        Exit Property
    End If
    dblDiasTotales = CDbl(m_ObjEntorno.DiasParaRecordatorioRellenoMarga1) + CDbl(m_ObjEntorno.DiasParaRecordatorioRellenoMarga2)
    m_sFechaPrevistaCorreoRecordatorioRellenoMarga2Calculado = DateAdd("d", CDate(Me.FechaTramitacionAltaMarga), dblDiasTotales)
    FechaPrevistaCorreoRecordatorioRellenoMarga2Calculado = m_sFechaPrevistaCorreoRecordatorioRellenoMarga2Calculado
    
    Exit Property
errores:
    If Err.Number <> 2000 Then
        m_Error = "El método SolicitudFechas.FechaPrevistaCorreoRecordatorioRellenoMarga2Calculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get FechaPrevistaAutocancelacionPreMargaCalculado() As String
    
    Dim dblDiasTotales As Double
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If m_sFechaPrevistaAutocancelacionPreMargaCalculado <> "" Then
        FechaPrevistaAutocancelacionPreMargaCalculado = m_sFechaPrevistaAutocancelacionPreMargaCalculado
        Exit Property
    End If
    If Not IsDate(Me.FechaEnvioExcel) Then
        Exit Property
    End If
    If Not IsNumeric(m_ObjEntorno.DiasCancelacionPreMARGA) Then
        Exit Property
    End If
    dblDiasTotales = CDbl(m_ObjEntorno.DiasCancelacionPreMARGA)
    m_sFechaPrevistaAutocancelacionPreMargaCalculado = DateAdd("d", CDate(Me.FechaEnvioExcel), dblDiasTotales)
    FechaPrevistaAutocancelacionPreMargaCalculado = m_sFechaPrevistaAutocancelacionPreMargaCalculado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método SolicitudFechas.FechaPrevistaAutocancelacionPreMargaCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get FechaPrevistaAutocancelacionMargaCalculado() As String
    
    Dim dblDiasTotales As Double
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If m_sFechaPrevistaAutocancelacionMargaCalculado <> "" Then
        FechaPrevistaAutocancelacionMargaCalculado = m_sFechaPrevistaAutocancelacionMargaCalculado
        Exit Property
    End If
    If Not IsDate(Me.FechaTramitacionAltaMarga) Then
        Exit Property
    End If
    dblDiasTotales = CDbl(m_ObjEntorno.DiasCancelacionMARGA)
    m_sFechaPrevistaAutocancelacionMargaCalculado = DateAdd("d", CDate(Me.FechaTramitacionAltaMarga), dblDiasTotales)
    FechaPrevistaAutocancelacionMargaCalculado = m_sFechaPrevistaAutocancelacionMargaCalculado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método SolicitudFechas.FechaPrevistaAutocancelacionMargaCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get DiasParaCancelacion() As String
    
   
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Not IsDate(Me.FechaEnvioExcel) And Not IsDate(Me.FechaTramitacionAltaMarga) Then
        Exit Property
    End If
    If IsDate(Me.FechaDesestimado) Then
        Exit Property
    End If
    If IsDate(Me.FechaAutocancelacion) Then
        Exit Property
    End If
    If IsDate(FechaEnvioONS) Then
        Exit Property
    End If
    If IsDate(Me.FechaTramitacionAltaMarga) Then
        m_sDiasParaCancelacion = Me.DiasParaCancelacionMarga
    Else
        m_sDiasParaCancelacion = Me.DiasParaCancelacionPreMarga
    End If
    DiasParaCancelacion = m_sDiasParaCancelacion
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método SolicitudFechas.DiasParaCancelacion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get DiasDesdeEnvioExcel() As String
    
   
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Not IsDate(Me.FechaEnvioExcel) Then
        Exit Property
    End If
    m_sDiasDesdeEnvioExcel = DateDiff("d", CDate(Date), CDate(Me.FechaEnvioExcel))
    DiasDesdeEnvioExcel = m_sDiasDesdeEnvioExcel
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método SolicitudFechas.DiasDesdeEnvioExcel ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get DiasDesdeAltaEnMarga() As String
    
   
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Not IsDate(Me.FechaTramitacionAltaMarga) Then
        Exit Property
    End If
    m_sDiasDesdeAltaEnMarga = DateDiff("d", CDate(Date), CDate(Me.FechaTramitacionAltaMarga))
    DiasDesdeAltaEnMarga = m_sDiasDesdeAltaEnMarga
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método SolicitudFechas.DiasDesdeAltaEnMarga ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property
Public Property Get DiasParaCancelacionPreMarga() As String
    
    Dim m_DiasDesdeElInicio As Double
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Not IsDate(Me.FechaTramitacionAltaMarga) Then
        Exit Property
    End If
    If Not IsNumeric(m_ObjEntorno.DiasCancelacionPreMARGA) Then
        Exit Property
    End If
    'if m_ObjEntorno.di
    m_DiasDesdeElInicio = DateDiff("d", Date, CDate(Me.FechaTramitacionAltaMarga))
    If m_DiasDesdeElInicio < 0 Then m_DiasDesdeElInicio = -1 * m_DiasDesdeElInicio
    
    m_sDiasParaCancelacionPreMarga = CStr(CDbl(m_ObjEntorno.DiasCancelacionPreMARGA) - m_DiasDesdeElInicio)
    
    
    DiasParaCancelacionPreMarga = m_sDiasParaCancelacionPreMarga
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método SolicitudFechas.DiasParaCancelacionPreMarga ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get DiasParaCancelacionMarga() As String
    
    Dim m_DiasDesdeElInicio As Double
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If Not IsDate(Me.FechaEnvioExcel) Then
        Exit Property
    End If
    If Not IsNumeric(m_ObjEntorno.DiasCancelacionPreMARGA) Then
        Exit Property
    End If
    'if m_ObjEntorno.di
    m_DiasDesdeElInicio = DateDiff("d", Date, CDate(Me.FechaEnvioExcel))
    If m_DiasDesdeElInicio < 0 Then m_DiasDesdeElInicio = -1 * m_DiasDesdeElInicio
    
    m_sDiasParaCancelacionMarga = CStr(CDbl(m_ObjEntorno.DiasCancelacionPreMARGA) - m_DiasDesdeElInicio)
    
    
    DiasParaCancelacionMarga = m_sDiasParaCancelacionMarga
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método SolicitudFechas.DiasParaCancelacionMarga ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get EstadoEnum() As EnumEstado
    
    Dim m_Solicitud As solicitud
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    Set m_Solicitud = Me.solicitud
    If m_Solicitud Is Nothing Then
        Exit Property
    End If
    If m_Solicitud.Estado = "" Then
        Exit Property
    End If
    m_eEstadoEnum = m_ObjEntorno.ColTextoEstados(m_Solicitud.Estado)
    EstadoEnum = m_eEstadoEnum
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método SolicitudFechas.EstadoEnum ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get EstadoCalculadoTexto() As String
    Dim m_EstadoCalculado As EnumEstado
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    m_EstadoCalculado = Me.EstadoCalculado
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    m_sEstadoCalculadoTexto = m_ObjEntorno.ColEstadosTexto(CStr(m_EstadoCalculado))
    EstadoCalculadoTexto = m_sEstadoCalculadoTexto
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método SolicitudFechas.EstadoCalculadoTexto ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get EstadoCalculado() As EnumEstado
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    With Me
        If IsDate(.FechaDesestimado) Then
            m_eEstadoCalculado = EnumEstado.Desestimado
            EstadoCalculado = m_eEstadoCalculado
            Exit Property
        End If
        If IsDate(.FechaCancelado) Then
            m_eEstadoCalculado = EnumEstado.Cancelado
            EstadoCalculado = m_eEstadoCalculado
            Exit Property
        End If
        If Not IsDate(.FechaEnvioExcel) Then
            m_eEstadoCalculado = EnumEstado.PendienteEnvioExcel
            EstadoCalculado = m_eEstadoCalculado
            Exit Property
        End If
        If Not IsDate(.FechaRecepcionExcel) And IsDate(.FechaEnvioExcel) Then
            m_eEstadoCalculado = EnumEstado.PendienteExcelRelleno
            EstadoCalculado = m_eEstadoCalculado
            Exit Property
        End If
        If .solicitud.TIPO = "Alta" Or .solicitud.TIPO = "Renovación" Or .solicitud.TIPO = "Elevación Grado" Then
            If Not IsDate(.FechaTramitacionAltaMarga) And IsDate(.FechaRecepcionExcel) Then
                m_eEstadoCalculado = EnumEstado.PendienteRegistroMarga
                EstadoCalculado = m_eEstadoCalculado
                Exit Property
            End If
            If Not IsDate(.FechaEnvioDPS) And IsDate(.FechaTramitacionAltaMarga) Then
                m_eEstadoCalculado = EnumEstado.PendienteEnvioDPS
                EstadoCalculado = m_eEstadoCalculado
                Exit Property
            End If
            If Not IsDate(.FechaEnvioONS) And IsDate(.FechaTramitacionAltaMarga) Then
                m_eEstadoCalculado = EnumEstado.PendienteRegistroONS
                EstadoCalculado = m_eEstadoCalculado
                Exit Property
            End If
            If IsDate(.FechaEnvioONS) And Not IsDate(.FechaRegistroEnHPS) Then
                m_eEstadoCalculado = EnumEstado.PendienteRegistroEnHPS
                EstadoCalculado = m_eEstadoCalculado
                Exit Property
            Else
                m_eEstadoCalculado = EnumEstado.Tramitado
                EstadoCalculado = m_eEstadoCalculado
                Exit Property
            End If
        Else
            If Not IsDate(.FechaEnvioONS) And IsDate(.FechaRecepcionExcel) Then
                m_eEstadoCalculado = EnumEstado.PendienteEnvioTraspasoONS
                EstadoCalculado = m_eEstadoCalculado
                Exit Property
            End If
            If IsDate(.FechaEnvioONS) And Not IsDate(.FechaRegistroEnHPS) Then
                m_eEstadoCalculado = EnumEstado.PendienteRegistroEnHPS
                EstadoCalculado = m_eEstadoCalculado
                Exit Property
            Else
                m_eEstadoCalculado = EnumEstado.Tramitado
                EstadoCalculado = m_eEstadoCalculado
                Exit Property
            End If
            
        End If
        
        
    End With
    
    
    
    EstadoCalculado = m_eEstadoCalculado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método SolicitudFechas.EstadoCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get solicitud() As solicitud
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objSolicitud Is Nothing Then
        Set solicitud = m_objSolicitud
        Exit Property
    End If
    Set m_objSolicitud = constructor.getSolicitud(p_IDSolicitud:=Me.IDSolicitud, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set solicitud = m_objSolicitud
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método SolicitudFechas.Solicitud ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set solicitud(p_Valor As Variant)
    
    Set m_objSolicitud = p_Valor
    
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
        .Add "IDSolicitud"
        .Add "FechaEnvioExcel"
        .Add "FechaRecepcionExcel"
        .Add "FechaTramitacionAltaMarga"
        .Add "FechaEnvioDPS"
        .Add "FechaEnvioONS"
        .Add "FechaRegistroEnHPS"
        .Add "FechaCorreoRecordatorioExcel1"
        .Add "FechaCorreoRecordatorioExcel2"
        .Add "FechaAutocancelacion"
        .Add "FechaCorreoRecordatorioRellenoMarga1"
        .Add "FechaCorreoRecordatorioRellenoMarga2"
        .Add "FechaDesestimado"
        
        .Add "FechaPrevistaCorreoRecordatorioExcel1"
        .Add "FechaPrevistaCorreoRecordatorioExcel2"
        .Add "FechaPrevistaCorreoRecordatorioRellenoMarga1"
        .Add "FechaPrevistaCorreoRecordatorioRellenoMarga2"
        .Add "FechaPrevistaAutocancelacionPreMarga"
        .Add "FechaPrevistaAutocancelacionMarga"
        .Add "FechaCancelado"
        
        
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método SolicitudFechas.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "IDSolicitud" Then
        getPropiedad = Me.IDSolicitud
        Exit Function
    End If
    If m_NombrePropiedad = "FechaEnvioExcel" Then
        getPropiedad = Me.FechaEnvioExcel
        Exit Function
    End If
    If m_NombrePropiedad = "FechaRecepcionExcel" Then
        getPropiedad = Me.FechaRecepcionExcel
        Exit Function
    End If
    If m_NombrePropiedad = "FechaTramitacionAltaMarga" Then
        getPropiedad = Me.FechaTramitacionAltaMarga
        Exit Function
    End If
    If m_NombrePropiedad = "FechaEnvioDPS" Then
        getPropiedad = Me.FechaEnvioDPS
        Exit Function
    End If
    If m_NombrePropiedad = "FechaEnvioONS" Then
        getPropiedad = Me.FechaEnvioONS
        Exit Function
    End If
    If m_NombrePropiedad = "FechaRegistroEnHPS" Then
        getPropiedad = Me.FechaRegistroEnHPS
        Exit Function
    End If
    If m_NombrePropiedad = "FechaCorreoRecordatorioExcel1" Then
        getPropiedad = Me.FechaCorreoRecordatorioExcel1
        Exit Function
    End If
    If m_NombrePropiedad = "FechaCorreoRecordatorioExcel2" Then
        getPropiedad = Me.FechaCorreoRecordatorioExcel2
        Exit Function
    End If
    If m_NombrePropiedad = "FechaAutocancelacion" Then
        getPropiedad = Me.FechaAutocancelacion
        Exit Function
    End If
    If m_NombrePropiedad = "FechaCorreoRecordatorioRellenoMarga1" Then
        getPropiedad = Me.FechaCorreoRecordatorioRellenoMarga1
        Exit Function
    End If
    If m_NombrePropiedad = "FechaCorreoRecordatorioRellenoMarga2" Then
        getPropiedad = Me.FechaCorreoRecordatorioRellenoMarga2
        Exit Function
    End If
    If m_NombrePropiedad = "FechaDesestimado" Then
        getPropiedad = Me.FechaDesestimado
        Exit Function
    End If
    If m_NombrePropiedad = "FechaPrevistaCorreoRecordatorioExcel1" Then
        getPropiedad = Me.FechaPrevistaCorreoRecordatorioExcel1
        Exit Function
    End If
    If m_NombrePropiedad = "FechaPrevistaCorreoRecordatorioExcel2" Then
        getPropiedad = Me.FechaPrevistaCorreoRecordatorioExcel2
        Exit Function
    End If
    If m_NombrePropiedad = "FechaPrevistaCorreoRecordatorioRellenoMarga1" Then
        getPropiedad = Me.FechaPrevistaCorreoRecordatorioRellenoMarga1
        Exit Function
    End If
    If m_NombrePropiedad = "FechaPrevistaCorreoRecordatorioRellenoMarga2" Then
        getPropiedad = Me.FechaPrevistaCorreoRecordatorioRellenoMarga2
        Exit Function
    End If
    If m_NombrePropiedad = "FechaPrevistaAutocancelacionPreMarga" Then
        getPropiedad = Me.FechaPrevistaAutocancelacionPreMarga
        Exit Function
    End If
    If m_NombrePropiedad = "FechaPrevistaAutocancelacionMarga" Then
        getPropiedad = Me.FechaPrevistaAutocancelacionMarga
        Exit Function
    End If
    If m_NombrePropiedad = "FechaCancelado" Then
        getPropiedad = Me.FechaCancelado
        Exit Function
    End If
    
    
    p_Error = "Propiedad no reconocida en SolicitudFechas.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SolicitudFechas.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


Public Function setPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "IDSolicitud" Then
        Me.IDSolicitud = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaEnvioExcel" Then
        Me.FechaEnvioExcel = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaRecepcionExcel" Then
        Me.FechaRecepcionExcel = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaTramitacionAltaMarga" Then
        Me.FechaTramitacionAltaMarga = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaEnvioDPS" Then
        Me.FechaEnvioDPS = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaEnvioONS" Then
        Me.FechaEnvioONS = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaRegistroEnHPS" Then
        Me.FechaRegistroEnHPS = m_ValorPropiedad
        Exit Function
    End If
    
    If m_NombrePropiedad = "FechaCorreoRecordatorioExcel1" Then
        Me.FechaCorreoRecordatorioExcel1 = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaCorreoRecordatorioExcel2" Then
        Me.FechaCorreoRecordatorioExcel2 = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaAutocancelacion" Then
        Me.FechaAutocancelacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaCorreoRecordatorioRellenoMarga1" Then
        Me.FechaCorreoRecordatorioRellenoMarga1 = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaCorreoRecordatorioRellenoMarga2" Then
        Me.FechaCorreoRecordatorioRellenoMarga2 = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaDesestimado" Then
        Me.FechaDesestimado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaPrevistaCorreoRecordatorioExcel1" Then
        Me.FechaPrevistaCorreoRecordatorioExcel1 = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaPrevistaCorreoRecordatorioExcel2" Then
        Me.FechaPrevistaCorreoRecordatorioExcel2 = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaPrevistaCorreoRecordatorioRellenoMarga1" Then
        Me.FechaPrevistaCorreoRecordatorioRellenoMarga1 = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaPrevistaCorreoRecordatorioRellenoMarga2" Then
        Me.FechaPrevistaCorreoRecordatorioRellenoMarga2 = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaPrevistaAutocancelacionPreMarga" Then
        Me.FechaPrevistaAutocancelacionPreMarga = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaPrevistaAutocancelacionMarga" Then
        Me.FechaPrevistaAutocancelacionMarga = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaCancelado" Then
        Me.FechaCancelado = m_ValorPropiedad
        Exit Function
    End If
    
    
    p_Error = "Propiedad no reconocida en SolicitudFechas.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método SolicitudFechas.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function

