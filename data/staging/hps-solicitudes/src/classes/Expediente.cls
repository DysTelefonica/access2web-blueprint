VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Expediente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public IDExpediente As String
Public IDExpedientePadre As String
Public Nemotecnico As String
Public Titulo As String
Public ImporteLicitacion As String
Public ImporteContratacion As String
Public CodProyecto As String
Public CodExp As String
Public CodExpLargo As String
Public CodS4H As String
Public FechaInicioContrato As String
Public FechaFinContrato As String
Public FechaFinGarantia As String
Public EsAM As String
Public EsLote As String
Public EsBasado As String
Public EsExpediente As String
Public Ordinal As String
Public IdGradoClasificacion As String
Public IDOrganoContratacion As String
Public IDOficinaPrograma As String
Public IDEjercito As String
Public AccesoSharepoint As String
Public Observaciones As String
Public FechaCreacion As String
Public IDUsuarioCreacion As String
Public FechaUltimoCambio As String
Public IDUsuarioUltimoCambio As String
Public IDEstado As String
Public Ambito As String
Public NPedido As String
Public IDResponsableCalidad As String
Public Adjudicado As String
Public EnPeriodoDeAdjudicacion As String
Public TIPO As String
Public TipoInforme As String
Public AGEDYSAplica As String
Public AGEDYSGenerico As String
Public HPSAplica As String
Public CadenaPecal As String
Public Pecal As String
Public POSTAGEDO As String

Private m_sTextoExpediente As String
Private m_sCadenaContratistas As String
Private m_objColCampos As Collection
Public Error As String
Private m_Error As String

Public Property Get CadenaContratistas() As String
    
    
    
        
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If m_sCadenaContratistas <> "" Then
        CadenaContratistas = m_sCadenaContratistas
        Exit Property
    End If
    m_sCadenaContratistas = constructor.getcadenaContratistas(p_IDExp:=Me.IDExpediente, p_Error:=m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    
    CadenaContratistas = m_sCadenaContratistas
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Expediente.CadenaContratistas ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Property Get TextoExpediente() As String
    
    
    
        
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If m_sTextoExpediente <> "" Then
        TextoExpediente = m_sTextoExpediente
        Exit Property
    End If
    If Me.Nemotecnico <> "" Then
        m_sTextoExpediente = Me.Nemotecnico
    Else
        m_sTextoExpediente = Me.CodExp
    End If
    
    
    TextoExpediente = m_sTextoExpediente
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Expediente.TextoExpediente ha devuelto el error: " & vbNewLine & Err.Description
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
        .Add "IDExpediente"
        .Add "IDExpedientePadre"
        .Add "Nemotecnico"
        .Add "Titulo"
        .Add "ImporteLicitacion"
        .Add "ImporteContratacion"
        .Add "CodProyecto"
        .Add "CodExp"
        .Add "CodExpLargo"
        .Add "CodS4H"
        .Add "FechaInicioContrato"
        .Add "FechaFinContrato"
        .Add "FechaFinGarantia"
        .Add "EsAM"
        .Add "EsLote"
        .Add "EsBasado"
        .Add "EsExpediente"
        .Add "Ordinal"
        .Add "IdGradoClasificacion"
        .Add "IDOrganoContratacion"
        .Add "IDOficinaPrograma"
        .Add "IDEjercito"
        .Add "AccesoSharepoint"
        .Add "Observaciones"
        .Add "FechaCreacion"
        .Add "IDUsuarioCreacion"
        .Add "FechaUltimoCambio"
        .Add "IDUsuarioUltimoCambio"
        .Add "IDEstado"
        .Add "Ambito"
        .Add "NPedido"
        .Add "IDResponsableCalidad"
        .Add "Adjudicado"
        .Add "EnPeriodoDeAdjudicacion"
        .Add "Tipo"
        .Add "TipoInforme"
        .Add "AGEDYSAplica"
        .Add "AGEDYSGenerico"
        .Add "HPSAplica"
        .Add "CadenaPecal"
        .Add "Pecal"
        .Add "POSTAGEDO"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Expediente.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "IDExpediente" Then
        getPropiedad = Me.IDExpediente
        Exit Function
    End If
    If m_NombrePropiedad = "IDExpedientePadre" Then
        getPropiedad = Me.IDExpedientePadre
        Exit Function
    End If
    If m_NombrePropiedad = "Nemotecnico" Then
        getPropiedad = Me.Nemotecnico
        Exit Function
    End If
    If m_NombrePropiedad = "Titulo" Then
        getPropiedad = Me.Titulo
        Exit Function
    End If
    If m_NombrePropiedad = "ImporteLicitacion" Then
        getPropiedad = Me.ImporteLicitacion
        Exit Function
    End If
    If m_NombrePropiedad = "ImporteContratacion" Then
        getPropiedad = Me.ImporteContratacion
        Exit Function
    End If
    If m_NombrePropiedad = "CodProyecto" Then
        getPropiedad = Me.CodProyecto
        Exit Function
    End If
    If m_NombrePropiedad = "CodExp" Then
        getPropiedad = Me.CodExp
        Exit Function
    End If
    If m_NombrePropiedad = "CodExpLargo" Then
        getPropiedad = Me.CodExpLargo
        Exit Function
    End If
    If m_NombrePropiedad = "CodS4H" Then
        getPropiedad = Me.CodS4H
        Exit Function
    End If
    If m_NombrePropiedad = "FechaInicioContrato" Then
        getPropiedad = Me.FechaInicioContrato
        Exit Function
    End If
    If m_NombrePropiedad = "FechaFinContrato" Then
        getPropiedad = Me.FechaFinContrato
        Exit Function
    End If
    If m_NombrePropiedad = "FechaFinGarantia" Then
        getPropiedad = Me.FechaFinGarantia
        Exit Function
    End If
    If m_NombrePropiedad = "EsAM" Then
        getPropiedad = Me.EsAM
        Exit Function
    End If
    If m_NombrePropiedad = "EsLote" Then
        getPropiedad = Me.EsLote
        Exit Function
    End If
    If m_NombrePropiedad = "EsBasado" Then
        getPropiedad = Me.EsBasado
        Exit Function
    End If
    If m_NombrePropiedad = "EsExpediente" Then
        getPropiedad = Me.EsExpediente
        Exit Function
    End If
    If m_NombrePropiedad = "Ordinal" Then
        getPropiedad = Me.Ordinal
        Exit Function
    End If
    If m_NombrePropiedad = "IdGradoClasificacion" Then
        getPropiedad = Me.IdGradoClasificacion
        Exit Function
    End If
    If m_NombrePropiedad = "IDOrganoContratacion" Then
        getPropiedad = Me.IDOrganoContratacion
        Exit Function
    End If
    If m_NombrePropiedad = "IDOficinaPrograma" Then
        getPropiedad = Me.IDOficinaPrograma
        Exit Function
    End If
    If m_NombrePropiedad = "IDEjercito" Then
        getPropiedad = Me.IDEjercito
        Exit Function
    End If
    If m_NombrePropiedad = "AccesoSharepoint" Then
        getPropiedad = Me.AccesoSharepoint
        Exit Function
    End If
    If m_NombrePropiedad = "Observaciones" Then
        getPropiedad = Me.Observaciones
        Exit Function
    End If
    If m_NombrePropiedad = "FechaCreacion" Then
        getPropiedad = Me.FechaCreacion
        Exit Function
    End If
    If m_NombrePropiedad = "IDUsuarioCreacion" Then
        getPropiedad = Me.IDUsuarioCreacion
        Exit Function
    End If
    If m_NombrePropiedad = "FechaUltimoCambio" Then
        getPropiedad = Me.FechaUltimoCambio
        Exit Function
    End If
    If m_NombrePropiedad = "IDUsuarioUltimoCambio" Then
        getPropiedad = Me.IDUsuarioUltimoCambio
        Exit Function
    End If
    If m_NombrePropiedad = "IDEstado" Then
        getPropiedad = Me.IDEstado
        Exit Function
    End If
    If m_NombrePropiedad = "Ambito" Then
        getPropiedad = Me.Ambito
        Exit Function
    End If
    If m_NombrePropiedad = "NPedido" Then
        getPropiedad = Me.NPedido
        Exit Function
    End If
    If m_NombrePropiedad = "IDResponsableCalidad" Then
        getPropiedad = Me.IDResponsableCalidad
        Exit Function
    End If
    If m_NombrePropiedad = "Adjudicado" Then
        getPropiedad = Me.Adjudicado
        Exit Function
    End If
    If m_NombrePropiedad = "EnPeriodoDeAdjudicacion" Then
        getPropiedad = Me.EnPeriodoDeAdjudicacion
        Exit Function
    End If
    If m_NombrePropiedad = "Tipo" Then
        getPropiedad = Me.TIPO
        Exit Function
    End If
    If m_NombrePropiedad = "TipoInforme" Then
        getPropiedad = Me.TipoInforme
        Exit Function
    End If
    If m_NombrePropiedad = "AGEDYSAplica" Then
        getPropiedad = Me.AGEDYSAplica
        Exit Function
    End If
    If m_NombrePropiedad = "AGEDYSGenerico" Then
        getPropiedad = Me.AGEDYSGenerico
        Exit Function
    End If
    If m_NombrePropiedad = "HPSAplica" Then
        getPropiedad = Me.HPSAplica
        Exit Function
    End If
    If m_NombrePropiedad = "CadenaPecal" Then
        getPropiedad = Me.CadenaPecal
        Exit Function
    End If
    If m_NombrePropiedad = "Pecal" Then
        getPropiedad = Me.Pecal
        Exit Function
    End If
    If m_NombrePropiedad = "POSTAGEDO" Then
        getPropiedad = Me.POSTAGEDO
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en Expediente.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Expediente.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


Public Function setPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "IDExpediente" Then
        Me.IDExpediente = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDExpedientePadre" Then
        Me.IDExpedientePadre = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Nemotecnico" Then
        Me.Nemotecnico = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Titulo" Then
        Me.Titulo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "ImporteLicitacion" Then
        Me.ImporteLicitacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "ImporteContratacion" Then
        Me.ImporteContratacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CodProyecto" Then
        Me.CodProyecto = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CodExp" Then
        Me.CodExp = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CodExpLargo" Then
        Me.CodExpLargo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CodS4H" Then
        Me.CodS4H = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaInicioContrato" Then
        Me.FechaInicioContrato = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaFinContrato" Then
        Me.FechaFinContrato = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaFinGarantia" Then
        Me.FechaFinGarantia = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "EsAM" Then
        Me.EsAM = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "EsLote" Then
        Me.EsLote = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "EsBasado" Then
        Me.EsBasado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "EsExpediente" Then
        Me.EsExpediente = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Ordinal" Then
        Me.Ordinal = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IdGradoClasificacion" Then
        Me.IdGradoClasificacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDOrganoContratacion" Then
        Me.IDOrganoContratacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDOficinaPrograma" Then
        Me.IDOficinaPrograma = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDEjercito" Then
        Me.IDEjercito = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "AccesoSharepoint" Then
        Me.AccesoSharepoint = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Observaciones" Then
        Me.Observaciones = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaCreacion" Then
        Me.FechaCreacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDUsuarioCreacion" Then
        Me.IDUsuarioCreacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaUltimoCambio" Then
        Me.FechaUltimoCambio = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDUsuarioUltimoCambio" Then
        Me.IDUsuarioUltimoCambio = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDEstado" Then
        Me.IDEstado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Ambito" Then
        Me.Ambito = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NPedido" Then
        Me.NPedido = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDResponsableCalidad" Then
        Me.IDResponsableCalidad = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Adjudicado" Then
        Me.Adjudicado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "EnPeriodoDeAdjudicacion" Then
        Me.EnPeriodoDeAdjudicacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Tipo" Then
        Me.TIPO = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "TipoInforme" Then
        Me.TipoInforme = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "AGEDYSAplica" Then
        Me.AGEDYSAplica = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "AGEDYSGenerico" Then
        Me.AGEDYSGenerico = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "HPSAplica" Then
        Me.HPSAplica = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CadenaPecal" Then
        Me.CadenaPecal = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Pecal" Then
        Me.Pecal = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "POSTAGEDO" Then
        Me.POSTAGEDO = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en Expediente.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Expediente.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function

