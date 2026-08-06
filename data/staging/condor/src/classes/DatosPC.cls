VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DatosPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

Public idDatosPC As Long
Public idSolicitud As Long
Public refContratoInspeccionOficial As String
Public refSuministrador As String
Public denominacionContrato As String
Public SuministradorNombreDir As String
Public objetoContrato As String

Public descripcionMaterialAfectado As String
Public numPlanoEspecificacion As String
Public descripcionPropuestaCambio As String
Public descripcionPropuestaCambioCont As String

Public motivoCorregirDeficiencias As Boolean
Public motivoMejorarCapacidad As Boolean
Public motivoAumentarNacionalizacion As Boolean
Public motivoMejorarSeguridad As Boolean
Public motivoMejorarFiabilidad As Boolean
Public motivoMejorarCosteEficacia As Boolean
Public motivoOtros As Boolean
Public motivoOtrosDetalle As String

Public incidenciaCoste As String

Public incidenciaPlazo As String

Public incidenciaSeguridad As Boolean
Public incidenciaFiabilidad As Boolean
Public incidenciaMantenibilidad As Boolean
Public incidenciaIntercambiabilidad As Boolean
Public incidenciaVidaUtilAlmacen As Boolean
Public incidenciaFuncionamientoFuncion As Boolean

Public impactoClasificacion As String 'MAYOR, MENOR
Public CambioAfectaAMaterial As String 'Material ya entregado Material por entregar



Public firmaOficinaTecnicaNombre As String
Public firmaRepSuministradorNombre As String

Public racCodigo As String
Public observacionesRAC As String
Public racNombre  As String
Public racDecision As String
Public racRechazoMotivos As String
Public obsAprobacionAutoridadDiseno As String
Public NombreAutoridadDiseno  As String

Public decisionFinal As String
Public obsDecisionFinal As String
Public NombreFirmanteFinal As String

Private m_objColCampos As Collection

Public Property Get ColCampos() As Collection

    If m_objColCampos Is Nothing Then
        Set m_objColCampos = New Collection
        With m_objColCampos
            .Add "idDatosPC"
            .Add "idSolicitud"
            .Add "refContratoInspeccionOficial"
            .Add "refSuministrador"
            .Add "denominacionContrato"
            .Add "suministradorNombreDir"
            .Add "objetoContrato"
            .Add "descripcionMaterialAfectado"
            .Add "numPlanoEspecificacion"
            .Add "descripcionPropuestaCambio"
            .Add "descripcionPropuestaCambioCont"
            .Add "motivoCorregirDeficiencias"
            .Add "motivoMejorarCapacidad"
            .Add "motivoAumentarNacionalizacion"
            .Add "motivoMejorarSeguridad"
            .Add "motivoMejorarFiabilidad"
            .Add "motivoMejorarCosteEficacia"
            .Add "motivoOtros"
            .Add "motivoOtrosDetalle"
            .Add "incidenciaCoste"
            .Add "incidenciaPlazo"
            .Add "incidenciaSeguridad"
            .Add "incidenciaFiabilidad"
            .Add "incidenciaMantenibilidad"
            .Add "incidenciaIntercambiabilidad"
            .Add "incidenciaVidaUtilAlmacen"
            .Add "incidenciaFuncionamientoFuncion"
            .Add "impactoClasificacion"
            .Add "CambioAfectaAMaterial"
            .Add "firmaOficinaTecnicaNombre"
            .Add "firmaRepSuministradorNombre"
            .Add "racCodigo"
            .Add "observacionesRAC"
            .Add "racNombre"
            .Add "racDecision"
            .Add "racRechazoMotivos"
            .Add "obsAprobacionAutoridadDiseno"
            .Add "NombreAutoridadDiseno"
            .Add "decisionFinal"
            .Add "obsDecisionFinal"
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

