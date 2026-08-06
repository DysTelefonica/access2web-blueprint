VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SolicitudBusquedaViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO DE CLASE: SolicitudBusquedaViewModel (VERSIÓN PLANA PARA BÚSQUEDAS)
' ==========================================================================

' --- Propiedades aplanadas ---
Public idSolicitud As Long
Public tipoSolicitud As String
Public EstadoNombre As String
Public codigoSolicitud As String
Public idExpediente As Long
Public CodExp As String
Public Nemotecnico As String
Public FechaInicioContrato As Date
Public FechaFinContrato As Date
Public ResponsableTecnico As String
Public EmailResponsableTecnico As String
Public ResponsableCalidad As String
Public EmailResponsableCalidad As String
Public RacAsignado As String

Private m_objColCampos As Collection

Public Property Get ColCampos() As Collection
    If m_objColCampos Is Nothing Then
        Set m_objColCampos = New Collection
        With m_objColCampos
            .Add "idSolicitud"
            .Add "tipoSolicitud"
            .Add "EstadoNombre"
            .Add "codigoSolicitud"
            .Add "IDExpediente"
            .Add "CodExp"
            .Add "Nemotecnico"
            .Add "FechaInicioContrato"
            .Add "FechaFinContrato"
            .Add "ResponsableTecnico"
            .Add "EmailResponsableTecnico"
            .Add "ResponsableCalidad"
            .Add "EmailResponsableCalidad"
            .Add "RacAsignado"
        End With
    End If
    Set ColCampos = m_objColCampos
End Property

Public Function getPropiedad(ByVal m_NombrePropiedad As String) As Variant
    getPropiedad = CallByName(Me, m_NombrePropiedad, VbGet)
End Function

Public Sub SetPropiedad(ByVal m_NombrePropiedad As String, ByVal m_ValorPropiedad As Variant)
    CallByName Me, m_NombrePropiedad, VbLet, m_ValorPropiedad
End Sub
