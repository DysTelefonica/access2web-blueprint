VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public ID_PROYECTOS As String
Public CODPROYECTOS As String
Public NAcreedorSAP As String
Public email As String
Public CODCENTRO As String
Public NOMBRECENTRO As String
Public CODARTICULO As String
Public DEPENDENCIA As String
Public Descripcion As String
Public Expediente As String
Public PETICIONARIO As String
Public FECHAPETICION As String
Public OFERTARECIBIDA As Boolean
Public OFERTARECIBIDAFECHA As String
Public FECHARECEPCIONECONOMICA As String
Public OBSERVACIONESECONOMICAS As String
Public OFERTAACEPTADA As Boolean
Public OFERTAACEPTADAFECHA As String
Public OBRAREPAROS As Boolean
Public REPAROSSUBSANADOS As Boolean
Public FECHACERTIFICADOCONFORMIDAD As String
Public TEXTOCERTIFICADOCONFORMIDAD As String
Public OBSERVACIONESCALIDAD As String
Public FECHAREPAROSSUBSANADOS As String
Public FECHAOBRAENTREGASINREPAROS As String
Public CODOFERTASUMINISTRADOR As String
Public CRITERIOCOMPRAS As String
Public IMPORTESINIVA As Double
Public FACTURAPAGADA As Boolean
Public Observaciones As String
Public FREGISTRO As String
Public CENTRO As String
Public ELIMINADO As Boolean
Public FELIMINADO As String
Public FINICIONECESIDAD As String
Public FFINNECESIDAD As String
Public CONTINUAENVERSIONDOS As Boolean
Public AntesDeVersion2 As Boolean
Public AÑOPRESUPUESTO As String
Public FIRMADELTECNICO As String
Public CODCONTRATOGTV As String
Public POSICIONCONTRATOGTV As String
Public TIPOPEDIDO As String
Public AnexoOfertaSuministrador As String
Public FechaAnexoOfertaSuministrador As String
Public IDExpediente As String
Public RistraResponsablesExpTemp As String
Public AplicaCalidadEnOfertas As Boolean
Public AplicaCartaAdjudicacion As Boolean
Public AplicaCondicionamiento As Boolean
Public NombreClienteParaDPD As String
Public FechaFinAgendaTecnica As String
Public EstadoCorto As String
Public EstadoLargo As String
Public EstadoPedidoAdjudicacion As String
Public EstadoVisado As String
Public EstadoCondicionamiento As String

Private m_sSuministrador As String

Private m_objColCampos As Collection
Public Error As String
Private m_Error As String


Public Property Get Suministrador() As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sSuministrador <> "" Then
        Suministrador = m_sSuministrador
        Exit Property
    End If
    If Me.NAcreedorSAP = "" Then
        Exit Property
    End If
    m_SQL = "SELECT Suministrador " & _
            " from TbSuministradorSAP " & _
            "WHERE NAcreedorSAP=" & Me.NAcreedorSAP & ";"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            m_sSuministrador = Nz(.Fields("Suministrador"), "")
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Suministrador = m_sSuministrador
    
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Pedido.Suministrador ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
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
        .Add "ID_PROYECTOS"
        .Add "CODPROYECTOS"
        .Add "NAcreedorSAP"
        .Add "email"
        .Add "CODCENTRO"
        .Add "NOMBRECENTRO"
        .Add "CODARTICULO"
        .Add "DEPENDENCIA"
        .Add "DESCRIPCION"
        .Add "EXPEDIENTE"
        .Add "PETICIONARIO"
        .Add "FECHAPETICION"
        .Add "OFERTARECIBIDA"
        .Add "OFERTARECIBIDAFECHA"
        .Add "FECHARECEPCIONECONOMICA"
        .Add "OBSERVACIONESECONOMICAS"
        .Add "OFERTAACEPTADA"
        .Add "OFERTAACEPTADAFECHA"
        .Add "OBRAREPAROS"
        .Add "REPAROSSUBSANADOS"
        .Add "FECHACERTIFICADOCONFORMIDAD"
        .Add "TEXTOCERTIFICADOCONFORMIDAD"
        .Add "OBSERVACIONESCALIDAD"
        .Add "FECHAREPAROSSUBSANADOS"
        .Add "FECHAOBRAENTREGASINREPAROS"
        .Add "CODOFERTASUMINISTRADOR"
        .Add "CRITERIOCOMPRAS"
        .Add "IMPORTESINIVA"
        .Add "FACTURAPAGADA"
        .Add "OBSERVACIONES"
        .Add "FREGISTRO"
        .Add "CENTRO"
        .Add "ELIMINADO"
        .Add "FELIMINADO"
        .Add "FINICIONECESIDAD"
        .Add "FFINNECESIDAD"
        .Add "CONTINUAENVERSIONDOS"
        .Add "AntesDeVersion2"
        .Add "AÑOPRESUPUESTO"
        .Add "FIRMADELTECNICO"
        .Add "CODCONTRATOGTV"
        .Add "POSICIONCONTRATOGTV"
        .Add "TIPOPEDIDO"
        .Add "AnexoOfertaSuministrador"
        .Add "FechaAnexoOfertaSuministrador"
        .Add "IDExpediente"
        .Add "RistraResponsablesExpTemp"
        
        .Add "AplicaCalidadEnOfertas"
        .Add "AplicaCartaAdjudicacion"
        .Add "AplicaCondicionamiento"
        .Add "NombreClienteParaDPD"
        .Add "FechaFinAgendaTecnica"
        .Add "EstadoCorto"
        .Add "EstadoLargo"
        .Add "EstadoPedidoAdjudicacion"
        .Add "EstadoVisado"
        .Add "EstadoCondicionamiento"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Pedido.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property





Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "ID_PROYECTOS" Then
        getPropiedad = Me.ID_PROYECTOS
        Exit Function
    End If
    If m_NombrePropiedad = "CODPROYECTOS" Then
        getPropiedad = Me.CODPROYECTOS
        Exit Function
    End If
    If m_NombrePropiedad = "NAcreedorSAP" Then
        getPropiedad = Me.NAcreedorSAP
        Exit Function
    End If
    If m_NombrePropiedad = "email" Then
        getPropiedad = Me.email
        Exit Function
    End If
    If m_NombrePropiedad = "CODCENTRO" Then
        getPropiedad = Me.CODCENTRO
        Exit Function
    End If
    If m_NombrePropiedad = "NOMBRECENTRO" Then
        getPropiedad = Me.NOMBRECENTRO
        Exit Function
    End If
    If m_NombrePropiedad = "CODARTICULO" Then
        getPropiedad = Me.CODARTICULO
        Exit Function
    End If
    If m_NombrePropiedad = "DEPENDENCIA" Then
        getPropiedad = Me.DEPENDENCIA
        Exit Function
    End If
    If m_NombrePropiedad = "DESCRIPCION" Then
        getPropiedad = Me.Descripcion
        Exit Function
    End If
    If m_NombrePropiedad = "EXPEDIENTE" Then
        getPropiedad = Me.Expediente
        Exit Function
    End If
    If m_NombrePropiedad = "PETICIONARIO" Then
        getPropiedad = Me.PETICIONARIO
        Exit Function
    End If
    If m_NombrePropiedad = "FECHAPETICION" Then
        getPropiedad = Me.FECHAPETICION
        Exit Function
    End If
    If m_NombrePropiedad = "OFERTARECIBIDA" Then
        getPropiedad = Me.OFERTARECIBIDA
        Exit Function
    End If
    If m_NombrePropiedad = "OFERTARECIBIDAFECHA" Then
        getPropiedad = Me.OFERTARECIBIDAFECHA
        Exit Function
    End If
    If m_NombrePropiedad = "FECHARECEPCIONECONOMICA" Then
        getPropiedad = Me.FECHARECEPCIONECONOMICA
        Exit Function
    End If
    If m_NombrePropiedad = "OBSERVACIONESECONOMICAS" Then
        getPropiedad = Me.OBSERVACIONESECONOMICAS
        Exit Function
    End If
    If m_NombrePropiedad = "OFERTAACEPTADA" Then
        getPropiedad = Me.OFERTAACEPTADA
        Exit Function
    End If
    If m_NombrePropiedad = "OFERTAACEPTADAFECHA" Then
        getPropiedad = Me.OFERTAACEPTADAFECHA
        Exit Function
    End If
    If m_NombrePropiedad = "OBRAREPAROS" Then
        getPropiedad = Me.OBRAREPAROS
        Exit Function
    End If
    If m_NombrePropiedad = "REPAROSSUBSANADOS" Then
        getPropiedad = Me.REPAROSSUBSANADOS
        Exit Function
    End If
    If m_NombrePropiedad = "FECHACERTIFICADOCONFORMIDAD" Then
        getPropiedad = Me.FECHACERTIFICADOCONFORMIDAD
        Exit Function
    End If
    If m_NombrePropiedad = "TEXTOCERTIFICADOCONFORMIDAD" Then
        getPropiedad = Me.TEXTOCERTIFICADOCONFORMIDAD
        Exit Function
    End If
    If m_NombrePropiedad = "OBSERVACIONESCALIDAD" Then
        getPropiedad = Me.OBSERVACIONESCALIDAD
        Exit Function
    End If
    If m_NombrePropiedad = "FECHAREPAROSSUBSANADOS" Then
        getPropiedad = Me.FECHAREPAROSSUBSANADOS
        Exit Function
    End If
    If m_NombrePropiedad = "FECHAOBRAENTREGASINREPAROS" Then
        getPropiedad = Me.FECHAOBRAENTREGASINREPAROS
        Exit Function
    End If
    If m_NombrePropiedad = "CODOFERTASUMINISTRADOR" Then
        getPropiedad = Me.CODOFERTASUMINISTRADOR
        Exit Function
    End If
    If m_NombrePropiedad = "CRITERIOCOMPRAS" Then
        getPropiedad = Me.CRITERIOCOMPRAS
        Exit Function
    End If
    If m_NombrePropiedad = "IMPORTESINIVA" Then
        getPropiedad = Me.IMPORTESINIVA
        Exit Function
    End If
    If m_NombrePropiedad = "FACTURAPAGADA" Then
        getPropiedad = Me.FACTURAPAGADA
        Exit Function
    End If
    If m_NombrePropiedad = "OBSERVACIONES" Then
        getPropiedad = Me.Observaciones
        Exit Function
    End If
    If m_NombrePropiedad = "FREGISTRO" Then
        getPropiedad = Me.FREGISTRO
        Exit Function
    End If
    If m_NombrePropiedad = "CENTRO" Then
        getPropiedad = Me.CENTRO
        Exit Function
    End If
    If m_NombrePropiedad = "ELIMINADO" Then
        getPropiedad = Me.ELIMINADO
        Exit Function
    End If
    If m_NombrePropiedad = "FELIMINADO" Then
        getPropiedad = Me.FELIMINADO
        Exit Function
    End If
    If m_NombrePropiedad = "FINICIONECESIDAD" Then
        getPropiedad = Me.FINICIONECESIDAD
        Exit Function
    End If
    If m_NombrePropiedad = "FFINNECESIDAD" Then
        getPropiedad = Me.FFINNECESIDAD
        Exit Function
    End If
    If m_NombrePropiedad = "CONTINUAENVERSIONDOS" Then
        getPropiedad = Me.CONTINUAENVERSIONDOS
        Exit Function
    End If
    If m_NombrePropiedad = "AntesDeVersion2" Then
        getPropiedad = Me.AntesDeVersion2
        Exit Function
    End If
    If m_NombrePropiedad = "AÑOPRESUPUESTO" Then
        getPropiedad = Me.AÑOPRESUPUESTO
        Exit Function
    End If
    If m_NombrePropiedad = "FIRMADELTECNICO" Then
        getPropiedad = Me.FIRMADELTECNICO
        Exit Function
    End If
    If m_NombrePropiedad = "CODCONTRATOGTV" Then
        getPropiedad = Me.CODCONTRATOGTV
        Exit Function
    End If
    If m_NombrePropiedad = "POSICIONCONTRATOGTV" Then
        getPropiedad = Me.POSICIONCONTRATOGTV
        Exit Function
    End If
    If m_NombrePropiedad = "TIPOPEDIDO" Then
        getPropiedad = Me.TIPOPEDIDO
        Exit Function
    End If
    If m_NombrePropiedad = "AnexoOfertaSuministrador" Then
        getPropiedad = Me.AnexoOfertaSuministrador
        Exit Function
    End If
    If m_NombrePropiedad = "FechaAnexoOfertaSuministrador" Then
        getPropiedad = Me.FechaAnexoOfertaSuministrador
        Exit Function
    End If
    If m_NombrePropiedad = "IDExpediente" Then
        getPropiedad = Me.IDExpediente
        Exit Function
    End If
    If m_NombrePropiedad = "RistraResponsablesExpTemp" Then
        getPropiedad = Me.RistraResponsablesExpTemp
        Exit Function
    End If
    If m_NombrePropiedad = "AplicaCalidadEnOfertas" Then
        getPropiedad = Me.AplicaCalidadEnOfertas
        Exit Function
    End If
    If m_NombrePropiedad = "AplicaCartaAdjudicacion" Then
        getPropiedad = Me.AplicaCartaAdjudicacion
        Exit Function
    End If
    If m_NombrePropiedad = "AplicaCondicionamiento" Then
        getPropiedad = Me.AplicaCondicionamiento
        Exit Function
    End If
    If m_NombrePropiedad = "NombreClienteParaDPD" Then
        getPropiedad = Me.NombreClienteParaDPD
        Exit Function
    End If
    If m_NombrePropiedad = "FechaFinAgendaTecnica" Then
        getPropiedad = Me.FechaFinAgendaTecnica
        Exit Function
    End If
    If m_NombrePropiedad = "EstadoCorto" Then
        getPropiedad = Me.EstadoCorto
        Exit Function
    End If
    If m_NombrePropiedad = "EstadoLargo" Then
        getPropiedad = Me.EstadoLargo
        Exit Function
    End If
    If m_NombrePropiedad = "EstadoPedidoAdjudicacion" Then
        getPropiedad = Me.EstadoPedidoAdjudicacion
        Exit Function
    End If
    If m_NombrePropiedad = "EstadoVisado" Then
        getPropiedad = Me.EstadoVisado
        Exit Function
    End If
    If m_NombrePropiedad = "EstadoCondicionamiento" Then
        getPropiedad = Me.EstadoCondicionamiento
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en Pedido.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Pedido.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "ID_PROYECTOS" Then
        Me.ID_PROYECTOS = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CODPROYECTOS" Then
        Me.CODPROYECTOS = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NAcreedorSAP" Then
        Me.NAcreedorSAP = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "email" Then
        Me.email = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CODCENTRO" Then
        Me.CODCENTRO = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NOMBRECENTRO" Then
        Me.NOMBRECENTRO = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CODARTICULO" Then
        Me.CODARTICULO = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "DEPENDENCIA" Then
        Me.DEPENDENCIA = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "DESCRIPCION" Then
        Me.Descripcion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "EXPEDIENTE" Then
        Me.Expediente = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "PETICIONARIO" Then
        Me.PETICIONARIO = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FECHAPETICION" Then
        Me.FECHAPETICION = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "OFERTARECIBIDA" Then
        Me.OFERTARECIBIDA = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "OFERTARECIBIDAFECHA" Then
        Me.OFERTARECIBIDAFECHA = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FECHARECEPCIONECONOMICA" Then
        Me.FECHARECEPCIONECONOMICA = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "OBSERVACIONESECONOMICAS" Then
        Me.OBSERVACIONESECONOMICAS = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "OFERTAACEPTADA" Then
        Me.OFERTAACEPTADA = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "OFERTAACEPTADAFECHA" Then
        Me.OFERTAACEPTADAFECHA = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "OBRAREPAROS" Then
        Me.OBRAREPAROS = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "REPAROSSUBSANADOS" Then
        Me.REPAROSSUBSANADOS = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FECHACERTIFICADOCONFORMIDAD" Then
        Me.FECHACERTIFICADOCONFORMIDAD = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "TEXTOCERTIFICADOCONFORMIDAD" Then
        Me.TEXTOCERTIFICADOCONFORMIDAD = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "OBSERVACIONESCALIDAD" Then
        Me.OBSERVACIONESCALIDAD = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FECHAREPAROSSUBSANADOS" Then
        Me.FECHAREPAROSSUBSANADOS = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FECHAOBRAENTREGASINREPAROS" Then
        Me.FECHAOBRAENTREGASINREPAROS = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CODOFERTASUMINISTRADOR" Then
        Me.CODOFERTASUMINISTRADOR = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CRITERIOCOMPRAS" Then
        Me.CRITERIOCOMPRAS = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IMPORTESINIVA" Then
        Me.IMPORTESINIVA = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FACTURAPAGADA" Then
        Me.FACTURAPAGADA = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "OBSERVACIONES" Then
        Me.Observaciones = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FREGISTRO" Then
        Me.FREGISTRO = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CENTRO" Then
        Me.CENTRO = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "ELIMINADO" Then
        Me.ELIMINADO = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FELIMINADO" Then
        Me.FELIMINADO = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FINICIONECESIDAD" Then
        Me.FINICIONECESIDAD = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FFINNECESIDAD" Then
        Me.FFINNECESIDAD = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CONTINUAENVERSIONDOS" Then
        Me.CONTINUAENVERSIONDOS = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "AntesDeVersion2" Then
        Me.AntesDeVersion2 = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "AÑOPRESUPUESTO" Then
        Me.AÑOPRESUPUESTO = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FIRMADELTECNICO" Then
        Me.FIRMADELTECNICO = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CODCONTRATOGTV" Then
        Me.CODCONTRATOGTV = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "POSICIONCONTRATOGTV" Then
        Me.POSICIONCONTRATOGTV = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "TIPOPEDIDO" Then
        Me.TIPOPEDIDO = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "AnexoOfertaSuministrador" Then
        Me.AnexoOfertaSuministrador = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaAnexoOfertaSuministrador" Then
        Me.FechaAnexoOfertaSuministrador = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDExpediente" Then
        Me.IDExpediente = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "RistraResponsablesExpTemp" Then
        Me.RistraResponsablesExpTemp = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "AplicaCalidadEnOfertas" Then
        Me.AplicaCalidadEnOfertas = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "AplicaCartaAdjudicacion" Then
        Me.AplicaCartaAdjudicacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "AplicaCondicionamiento" Then
        Me.AplicaCondicionamiento = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NombreClienteParaDPD" Then
        Me.NombreClienteParaDPD = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaFinAgendaTecnica" Then
        Me.FechaFinAgendaTecnica = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "EstadoCorto" Then
        Me.EstadoCorto = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "EstadoLargo" Then
        Me.EstadoLargo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "EstadoPedidoAdjudicacion" Then
        Me.EstadoPedidoAdjudicacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "EstadoVisado" Then
        Me.EstadoVisado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "EstadoCondicionamiento" Then
        Me.EstadoCondicionamiento = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en Pedido.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Pedido.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function

