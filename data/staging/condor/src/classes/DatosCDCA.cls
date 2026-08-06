VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DatosCDCA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

Public idDatosCDCA As Long
Public idSolicitud As Long

Public numContrato As String
Public refSuministrador As String
Public SuministradorNombreDir As String
Public refDesviacionesPrevias As String
Public requiereModificacionContrato As Variant

Public identificacionMaterial As String
Public numPlanoEspecificacion As String
Public cantidadPeriodo As String
Public numSerieLote As String
Public causaNC As String
Public descripcionImpactoNC As String
Public descripcionImpactoNCCont As String

Public afectaPrestaciones As Boolean
Public afectaSeguridad As Boolean
Public afectaFiabilidad As Boolean
Public afectaVidaUtil As Boolean
Public afectaMedioambiente As Boolean
Public afectaMantenibilidad As Boolean
Public afectaIntercambiabilidad As Boolean
Public afectaApariencia As Boolean
Public afectaOtros As Boolean
Public impactoCoste As String
Public clasificacionNC As String
Public esSuministradorAD As Variant
Public identificacionAutoridadDiseno As String
Public efectoFechaEntrega As String


Public firmaAprobacionRespIngenieriaNombre As String
Public firmaAprobacionRespProduccionNombre As String
Public firmaAprobacionRespCalidadNombre As String
Public firmaAprobacionRespDisenioNombre As String
Public firmaAprobacionRepresentanteSumNombre As String


Public racCodigo As String
Public observacionesRAC As String
Public racNombre As String
Public racDecision As String

Public decisionFinal As String
Public observacionesFinales As String
Public NombreFirmanteFinal As String



Private m_objColCampos As Collection


Public Property Get ColCampos() As Collection
    If m_objColCampos Is Nothing Then
        Set m_objColCampos = New Collection
        With m_objColCampos
            .Add "idDatosCDCA"
            .Add "idSolicitud"
            .Add "refSuministrador"
            .Add "SuministradorNombreDir"
            .Add "numContrato"
            .Add "identificacionMaterial"
            .Add "numPlanoEspecificacion"
            .Add "cantidadPeriodo"
            .Add "numSerieLote"
            .Add "descripcionImpactoNC"
            .Add "descripcionImpactoNCCont"
            .Add "refDesviacionesPrevias"
            .Add "causaNC"
            .Add "impactoCoste"
            .Add "afectaPrestaciones"
            .Add "afectaSeguridad"
            .Add "afectaFiabilidad"
            .Add "afectaVidaUtil"
            .Add "afectaMedioambiente"
            .Add "afectaMantenibilidad"
            .Add "afectaIntercambiabilidad"
            .Add "afectaApariencia"
            .Add "afectaOtros"
            .Add "clasificacionNC"
            .Add "requiereModificacionContrato"
            .Add "efectoFechaEntrega"
            .Add "identificacionAutoridadDiseno"
            .Add "firmaAprobacionRespIngenieriaNombre"
            .Add "firmaAprobacionRespProduccionNombre"
            .Add "firmaAprobacionRespCalidadNombre"
            .Add "firmaAprobacionRespDisenioNombre"
            .Add "firmaAprobacionRepresentanteSumNombre"
            .Add "esSuministradorAD"
            .Add "racCodigo"
            .Add "racNombre"
            .Add "observacionesRAC"
            .Add "racDecision"
            .Add "decisionFinal"
            .Add "observacionesFinales"
            .Add "NombreFirmanteFinal"
            
            
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


