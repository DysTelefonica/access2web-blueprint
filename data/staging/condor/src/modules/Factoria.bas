Attribute VB_Name = "Factoria"

' ==========================================================================
' MÓDULO: Factoria.bas
' RESPONSABILIDAD: Centralizar la creación de instancias de todas las
'                  clases de entidad del proyecto.
' ==========================================================================
Option Compare Database
Option Explicit

Public Function CreateEntity(ByVal entityName As String) As Object
    ' Esta función actúa como nuestra "fábrica" de objetos.
    ' Recibe el nombre de una clase y devuelve una nueva instancia.
    On Error Resume Next ' Manejo de errores si el nombre no existe
    
    Select Case entityName
        Case "Expediente"
            Set CreateEntity = New Expediente
        Case "Usuario"
            Set CreateEntity = New usuario
        Case "Solicitud"
            Set CreateEntity = New Solicitud
        Case "Adjunto"
            Set CreateEntity = New Adjunto
        Case "Estado"
            Set CreateEntity = New estado
        Case "Suministrador"
            Set CreateEntity = New suministrador
        ' --- IMPORTANTE: Añadir aquí una línea por cada nueva clase de entidad ---
         Case "SolicitudBusquedaViewModel"
            Set CreateEntity = New SolicitudBusquedaViewModel
        Case "DatosPC"
            Set CreateEntity = New DatosPC
        Case "DatosCDCA"
            Set CreateEntity = New DatosCDCA
        Case "DatosCDCASUB"
            Set CreateEntity = New DatosCDCASUB
        Case "DatosPCSUB"
            Set CreateEntity = New DatosPCSUB
        Case "LogCambio"
            Set CreateEntity = New LogCambio
        Case "LogError"
            Set CreateEntity = New LogError
        
        Case "Correo"
            Set CreateEntity = New Correo
        Case "ExpedienteSuministrador"
            Set CreateEntity = New ExpedienteSuministrador
        Case "LogEstado"
            Set CreateEntity = New LogEstado
        Case "MapeoCampos"
            Set CreateEntity = New MapeoCampos
        Case "ExpedienteViewModel"
            Set CreateEntity = New ExpedienteViewModel
        Case "AdjuntoViewModel"
            Set CreateEntity = New AdjuntoViewModel
        Case "ValidacionRevision"
            Set CreateEntity = New ValidacionRevision
        Case "NoConformidad"
            Set CreateEntity = New NoConformidad
        Case Else
            ' Si el nombre no se encuentra, la función devolverá Nothing
            Set CreateEntity = Nothing
    End Select
    
End Function

