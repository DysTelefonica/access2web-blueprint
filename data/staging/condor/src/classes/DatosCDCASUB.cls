VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DatosCDCASUB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: DatosCDCASUB.cls
' RESPONSABILIDAD: Entidad pura. Las propiedades COINCIDEN EXACTAMENTE con la tabla.
' ==========================================================================

Public idDatosCDCASUB As Long
Public idSolicitud As Long

' --- Sección Generales ---
Public refSuministrador As String
Public refSubSuministrador As String
Public suministradorPrincipalNombreDir As String
Public subSuministradorNombreDir As String
Public refDesviacionesPrevias As String
Public requiereModificacionContrato As Variant

' --- Sección Propuesta ---
Public identificacionMaterial As String
Public numPlanoEspecificacion As String
Public cantidadPeriodo As String
Public numSerieLote As String
Public causaNC As String
Public descripcionImpactoNC As String
Public descripcionImpactoNCCont As String

' --- Sección Impacto ---
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
Public efectoFechaEntrega As String
Public identificacionAutoridadDiseno As String
Public esSubSuministradorAD As Variant

' --- Sección Aprobación (Nombres coinciden con capturas parciales, asumo estándar) ---
Public firmaAprobacionRespIngenieriaNombre As String
Public firmaAprobacionRespProduccionNombre As String
Public firmaAprobacionRespCalidadNombre As String
Public firmaAprobacionRespDisenioNombre As String
Public firmaAprobacionRepresentanteSumNombre As String

' --- Sección RAC ---
Public racCodigo As String
Public observacionesRAC As String
Public racDecision As String
Public racRechazoMotivos As String
Public racNombre As String

' --- CAMPOS CORREGIDOS SEGÚN TABLA (RAC DELEGADOR) ---
Public observacionesRACDelegador As String
Public racNombreDelegador As String

' --- CAMPOS CORREGIDOS SEGÚN TABLA (DECISIÓN FINAL) ---
Public decisionFinal As String
Public observacionesFinales As String
Public NombreFirmanteFinal As String
Private m_objColCampos As Collection

Public Property Get ColCampos() As Collection
    If m_objColCampos Is Nothing Then
        Set m_objColCampos = New Collection
        With m_objColCampos
            .Add "idDatosCDCASUB"
            .Add "idSolicitud"
            .Add "refSuministrador"
            .Add "refSubSuministrador"
            .Add "suministradorPrincipalNombreDir"
            .Add "subSuministradorNombreDir"
            .Add "refDesviacionesPrevias"
            .Add "requiereModificacionContrato"
            
            .Add "identificacionMaterial"
            .Add "numPlanoEspecificacion"
            .Add "cantidadPeriodo"
            .Add "numSerieLote"
            .Add "causaNC"
            .Add "descripcionImpactoNC"
            .Add "descripcionImpactoNCCont"
            
            .Add "afectaPrestaciones"
            .Add "afectaSeguridad"
            .Add "afectaFiabilidad"
            .Add "afectaVidaUtil"
            .Add "afectaMedioambiente"
            .Add "afectaIntercambiabilidad"
            .Add "afectaMantenibilidad"
            .Add "afectaApariencia"
            .Add "afectaOtros"
            .Add "impactoCoste"
            .Add "clasificacionNC"
            .Add "efectoFechaEntrega"
            .Add "identificacionAutoridadDiseno"
            .Add "esSubSuministradorAD"
            
            .Add "firmaAprobacionRespIngenieriaNombre"
            .Add "firmaAprobacionRespProduccionNombre"
            .Add "firmaAprobacionRespCalidadNombre"
            .Add "firmaAprobacionRespDisenioNombre"
            .Add "firmaAprobacionRepresentanteSumNombre"
            
            .Add "racCodigo"
            .Add "observacionesRAC"
            .Add "racDecision"
            .Add "racRechazoMotivos"
            .Add "racNombre"
            
            ' CORREGIDOS EN LA COLECCIÓN TAMBIÉN
            .Add "observacionesRACDelegador"
            .Add "racNombreDelegador"
            
            .Add "decisionFinal"
            .Add "observacionesFinales"
            .Add "NombreFirmanteFinal"
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


