Attribute VB_Name = "Factoria"
' Módulo: Factoria
' IMPORTANTE: Guarda este módulo con el nombre "Factoria"
Option Compare Database
Option Explicit

Public Function CreateEntity(className As String) As Object
    Select Case className
        Case "Justificacion"
            Set CreateEntity = New Justificacion
        
        ' Cuando añadas más entidades (ej. Expediente), añádelas aquí:
        ' Case "Expediente"
        '     Set CreateEntity = New Expediente
            
        Case Else
            Err.Raise 51000, "Factoria", "Clase no definida en Factoria: " & className
    End Select
End Function
