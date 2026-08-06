VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Expediente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ==========================================================================
' MÓDULO: Expediente.cls (Refactorizado sin campo Estado)
' ==========================================================================

Public idExpediente As Long
Public Nemotecnico As String
Public titulo As String
Public objetoContrato As String
Public CodExp As String
Public CodExpLargo As String
Public CodigoActividad As String
Public ContratistaPrincipal As String
Public FechaInicioContrato As Date
Public FechaFinContrato As Date
Public FechaFinGarantia As Date
Public ResponsableCalidad As String
Public ResponsableTecnico As String
Public EmailResponsableTecnico As String
Public EmailResponsableCalidad As String
Public RacAsignado As String

Private m_objColCampos As Collection


Public Property Get ColCampos() As Collection
    If m_objColCampos Is Nothing Then
        Set m_objColCampos = New Collection
        With m_objColCampos
            .Add "IDExpediente"
            .Add "Nemotecnico"
            .Add "Titulo"
            .Add "ObjetoContrato"
            .Add "CodExp"
            .Add "CodExpLargo"
            .Add "CodigoActividad"
            .Add "ContratistaPrincipal"
            .Add "FechaInicioContrato"
            .Add "FechaFinContrato"
            .Add "FechaFinGarantia"
            .Add "ResponsableCalidad"
            .Add "ResponsableTecnico"
            .Add "EmailResponsableTecnico"
            .Add "EmailResponsableCalidad"
            .Add "RacAsignado"
        End With
    End If
    Set ColCampos = m_objColCampos
End Property
' Sin manejo de errores. Cualquier fallo se propagará al llamador.
Public Function getPropiedad(ByVal m_NombrePropiedad As String) As Variant
    getPropiedad = CallByName(Me, m_NombrePropiedad, VbGet)
End Function

' Sin manejo de errores. Cualquier fallo se propagará al llamador.
Public Sub SetPropiedad(ByVal m_NombrePropiedad As String, ByVal m_ValorPropiedad As Variant)
    CallByName Me, m_NombrePropiedad, VbLet, m_ValorPropiedad
End Sub

