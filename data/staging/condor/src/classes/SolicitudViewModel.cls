VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SolicitudViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'=== Class Module: SolicitudViewModel.cls ===
Option Compare Database
Option Explicit

' Contenedor de datos (sin lógica)
Public Solicitud As Solicitud
Public Expediente As Expediente

Public EstadoNombre As String
Public EstadoDescripcion As String
Public nombreContratistaPrincipal As String
Public nombreSubcontratista As String
Public JefeProyecto As String
Public JefeProyectoEmail As String
Public TecnicoCalidad As String
Public TecnicoCalidadEmail As String
Public TiempoCiclo As String
Public TipoSolicitudNombre As String
Public CodigoNCAsociada As String

Public vm_revisionCalidadEstado As String
Public vm_revisionCalidadComentarios As String

' Nuevas propiedades para Notificaciones Inteligentes (Spec-099)
Public vm_tecnicoSubsanacionRealizada As Boolean
Public vm_tecnicoComentariosSubsanacion As String
Public vm_racValidacionEstado As String
Public vm_racValidacionComentarios As String

' --- API uniforme de entidades CONDOR ---
Public Function getPropiedad(ByVal nombre As String) As Variant
    On Error GoTo Fin
    getPropiedad = CallByName(Me, nombre, VbGet)
    Exit Function
Fin:
    ' Si no existe la propiedad, devolver Empty (no lanzar)
End Function

Public Sub SetPropiedad(ByVal nombre As String, ByVal valor As Variant)
    On Error Resume Next
    CallByName Me, nombre, VbLet, valor
End Sub


