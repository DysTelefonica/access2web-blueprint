Attribute VB_Name = "modActualizaciones"

' Archivo: modActualizaciones.bas
Option Compare Database
Option Explicit

Public Sub EjecutarActualizacion_Ambito()
    Dim db As DAO.Database
    Dim sql As String
    
    On Error GoTo Errores
    Set db = getdb() ' Usamos la conexión global definida en FUNCIONES UTILES
    
    ' Añadimos el campo a la tabla tbSolicitudes
    ' Si ya existe, dará error y saltará al manejador, que es lo esperado.
    sql = "ALTER TABLE tbSolicitudes ADD COLUMN revisionCalidadAmbito TEXT(50);"
    
    db.Execute sql, dbFailOnError
    
    MsgBox "Base de datos actualizada: Campo 'revisionCalidadAmbito' añadido.", vbInformation, "CONDOR"
    
LimpiarYSalir:
    If Not db Is Nothing Then db.Close
    Exit Sub
    
Errores:
    If Err.Number = 3381 Or Err.Number = 3293 Then ' Campo ya existe / Error sintaxis si existe
        MsgBox "El campo ya existía. No se requieren cambios.", vbInformation, "CONDOR"
        Resume LimpiarYSalir
    Else
        MsgBox "Error: " & Err.description, vbCritical, "CONDOR"
        Resume LimpiarYSalir
    End If
End Sub



Public Sub EjecutarActualizacion_Spec080_MapeoCDCASUB()
    Dim db As DAO.Database
    
    On Error GoTo Errores
    Set db = getdb()
    
    db.Execute "UPDATE tbMapeoCampos SET nombrePlantilla = 'CDCASUB' WHERE nombrePlantilla = 'CD_CA_SUB';", dbFailOnError
    db.Execute "UPDATE tbMapeoCampos SET nombreCampoTabla = 'suministradorPrincipalNombreDir' WHERE nombrePlantilla = 'CDCASUB' AND nombreCampoTabla = 'suministradorNombreDir';", dbFailOnError
    
    MsgBox "Base de datos actualizada para Spec-080.", vbInformation, "CONDOR"
    
LimpiarYSalir:
    If Not db Is Nothing Then db.Close
    Exit Sub
    
Errores:
    MsgBox "Error en actualizacion Spec-080: " & Err.description, vbCritical, "CONDOR"
    Resume LimpiarYSalir
End Sub


