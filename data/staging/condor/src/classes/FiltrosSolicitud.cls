VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FiltrosSolicitud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ==========================================================================
' MÓDULO DE CLASE: FiltrosSolicitud.cls
' RESPONSABILIDAD: Contenedor de datos para los filtros de búsqueda avanzados.
' ==========================================================================

Public idResponsableTecnico As Long
Public idResponsableCalidad As Long
Public idSuministrador As Long
Public tipoSolicitud As String
Public idEstado As String ' Acepta múltiples valores (ej: "1,3")

Public Function HayFiltrosActivos() As Boolean
    ' Responsabilidad: Comprueba si alguno de los filtros tiene un valor aplicado.
    HayFiltrosActivos = False
    
    If idResponsableTecnico > 0 Then HayFiltrosActivos = True: Exit Function
    If idResponsableCalidad > 0 Then HayFiltrosActivos = True: Exit Function
    If idSuministrador > 0 Then HayFiltrosActivos = True: Exit Function
    If Nz(tipoSolicitud, "") <> "" Then HayFiltrosActivos = True: Exit Function
    If Nz(idEstado, "") <> "" And idEstado <> "0" Then HayFiltrosActivos = True: Exit Function
    
End Function
