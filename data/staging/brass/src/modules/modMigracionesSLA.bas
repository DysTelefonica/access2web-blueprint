Attribute VB_Name = "modMigracionesSLA"

Option Compare Database
Option Explicit

'===========================================================
' MÓDULO DE MIGRACIÓN DE CAMPOS SLA - Spec-001
' Objetivo: Añadir 12 campos nuevos a TbEventos para SLAs
'===========================================================

' Constantes
Private Const TABLA_DESTINO As String = "TbEventos"

'===========================================================
' Función pública: MigrarCamposSLA
' Crea los 12 campos SLA en TbEventos si no existen
' Es idempotente: se puede ejecutar varias veces
'===========================================================
Public Function MigrarCamposSLA(Optional ByRef p_Error As String) As Boolean
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim i As Integer
    Dim arrCampos As Variant
    Dim nombreCampo As String
    Dim tipoCampo As Integer
    Dim tamanoCampo As Integer
    Dim creado As Boolean
    
    On Error GoTo errores
    
    p_Error = ""
    creado = False
    
    Set db = getdb()
    
    ' Verificar que la tabla existe
    If Not TablaExiste(db, TABLA_DESTINO) Then
        p_Error = "La tabla " & TABLA_DESTINO & " no existe"
        MigrarCamposSLA = False
        Exit Function
    End If
    
    Set tdf = db.TableDefs(TABLA_DESTINO)
    arrCampos = GetCamposSLA()
    
    For i = LBound(arrCampos) To UBound(arrCampos)
        nombreCampo = arrCampos(i)(0)
        tipoCampo = arrCampos(i)(1)
        
        ' Si es campo de texto, obtener tamaño
        If tipoCampo = dbText Then
            If UBound(arrCampos(i)) >= 2 Then
                tamanoCampo = arrCampos(i)(2)
            Else
                tamanoCampo = 50
            End If
        End If
        
        ' Verificar si el campo no existe
        If Not CampoExisteEnTabla(tdf, nombreCampo) Then
            ' Crear el campo
            If tipoCampo = dbText Then
                Set fld = tdf.CreateField(nombreCampo, tipoCampo, tamanoCampo)
            Else
                Set fld = tdf.CreateField(nombreCampo, tipoCampo)
            End If
            
            ' Añadir a la colección
            tdf.Fields.Append fld
            creado = True
            
            Debug.Print "Campo " & nombreCampo & " creado correctamente"
        Else
            Debug.Print "Campo " & nombreCampo & " ya existe"
        End If
        
        Set fld = Nothing
    Next i
    
    ' Refrescar tablas
    db.TableDefs.Refresh
    
    If creado Then
        Debug.Print "Migración de campos SLA completada - Nuevos campos añadidos"
    Else
        Debug.Print "Migración de campos SLA completada - Todos los campos ya existían"
    End If
    
    MigrarCamposSLA = True
    Exit Function
    
errores:
    p_Error = "Error en MigrarCamposSLA: " & Err.Description
    MigrarCamposSLA = False
    Debug.Print p_Error
End Function

'===========================================================
' GetCamposSLA
' Devuelve array con definición de campos SLA
'===========================================================
Private Function GetCamposSLA() As Variant
    ' Formato: Array(nombre, tipo [, tamaño])
    ' Tipos DAO: dbText=10, dbDate=8, dbBoolean=1
    GetCamposSLA = Array( _
        Array("FechaRecepcionNotificacion", dbDate), _
        Array("FechaInicioContactoCliente", dbDate), _
        Array("IncidenciaAveriaOReparacion", dbBoolean), _
        Array("FechaInicioTiempoAdquisicion", dbDate), _
        Array("FechaFinTiempoAdquisicion", dbDate), _
        Array("TipoReparacion", dbText, 50), _
        Array("Urgente", dbBoolean), _
        Array("EventoConServicioAfectado", dbBoolean), _
        Array("FechaRestablecimientoServicio", dbDate), _
        Array("TipoRepInsitu", dbBoolean), _
        Array("TipoRepNoSMT", dbBoolean), _
        Array("TipoRepValvulas", dbBoolean) _
    )
End Function

'===========================================================
' TablaExiste
' Verifica si una tabla existe en la base de datos
'===========================================================
Private Function TablaExiste(db As DAO.Database, p_NombreTabla As String) As Boolean
    Dim tdf As DAO.TableDef
    
    On Error Resume Next
    Set tdf = db.TableDefs(p_NombreTabla)
    TablaExiste = (Err.Number = 0)
    On Error GoTo 0
End Function

'===========================================================
' CampoExisteEnTabla
' Verifica si un campo existe en una tabla
'===========================================================
Private Function CampoExisteEnTabla(tdf As DAO.TableDef, p_NombreCampo As String) As Boolean
    Dim fld As DAO.Field
    
    On Error Resume Next
    Set fld = tdf.Fields(p_NombreCampo)
    CampoExisteEnTabla = (Err.Number = 0)
    On Error GoTo 0
End Function

'===========================================================
' VerificarCamposSLA
' Verifica que los 12 campos existen (para testing)
' Devuelve: True si todos existen, False si falta alguno
'===========================================================
Public Function VerificarCamposSLA(Optional ByRef p_CamposFaltantes As String) As Boolean
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim tdf As DAO.TableDef
    Dim arrCampos As Variant
    Dim i As Integer
    Dim nombreCampo As String
    Dim fld As DAO.Field
    Dim faltantes As String
    
    On Error GoTo errores
    
    p_CamposFaltantes = ""
    faltantes = ""
    
    Set db = getdb()
    Set tdf = db.TableDefs(TABLA_DESTINO)
    arrCampos = GetCamposSLA()
    
    For i = LBound(arrCampos) To UBound(arrCampos)
        nombreCampo = arrCampos(i)(0)
        
        If Not CampoExisteEnTabla(tdf, nombreCampo) Then
            If faltantes <> "" Then
                faltantes = faltantes & ", "
            End If
            faltantes = faltantes & nombreCampo
        End If
    Next i
    
    If faltantes <> "" Then
        p_CamposFaltantes = faltantes
        VerificarCamposSLA = False
    Else
        VerificarCamposSLA = True
    End If
    
    Exit Function
    
errores:
    p_CamposFaltantes = "Error al verificar: " & Err.Description
    VerificarCamposSLA = False
End Function




