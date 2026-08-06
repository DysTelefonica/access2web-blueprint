VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExpedienteViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO: ExpedienteViewModel.cls
' RESPONSABILIDAD: Contenedor de datos para el formulario frmDetalleExpediente.
' ==========================================================================

' --- Entidad principal ---
Public Expediente As Expediente

' --- Propiedades pre-calculadas para la vista ---
Public NombresContratistas As String
Public NombresSubcontratistas As String

' --- API uniforme de entidades CONDOR (opcional pero recomendado) ---
Public Function getPropiedad(ByVal nombre As String) As Variant
    On Error Resume Next
    getPropiedad = CallByName(Me, nombre, VbGet)
End Function

Public Sub SetPropiedad(ByVal nombre As String, ByVal valor As Variant)
    On Error Resume Next
    CallByName Me, nombre, VbLet, valor
End Sub
