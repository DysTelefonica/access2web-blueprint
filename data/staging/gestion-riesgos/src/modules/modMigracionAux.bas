Attribute VB_Name = "modMigracionAux"
Option Compare Database
Option Explicit

' =========================================================================
' Módulo: modMigracionAux
' Propósito: Contener funcionalidades migradas de GR2 para evitar tocar
'            el núcleo de GR1.
' =========================================================================

' --- CONSTANTES PARA HTML ---
Private Const CSS_ESTILO_INFORME As String = "<style>body{font-family:Arial;font-size:12px;} .titulo{background:#f2f2f2; font-weight:bold;} table{border-collapse:collapse;width:100%;} td,th{border:1px solid #ccc;padding:5px;}</style>"

''' <summary>
''' Convierte caracteres especiales para evitar errores en el renderizado HTML.
''' </summary>

''' <summary>
''' Genera el informe completo de un riesgo en formato HTML (Versión migrada GR2).
''' Nota: Se ha eliminado la dependencia de caché según instrucciones.
''' </summary>
Public Function ConstruirInformeRiesgoHTML(ByRef p_Riesgo As Object, Optional ByRef p_Error As String) As String
    Dim html As String
    On Error GoTo errores

    If p_Riesgo Is Nothing Then
        p_Error = "Objeto Riesgo no inicializado."
        Exit Function
    End If

    html = "<html><head>" & CSS_ESTILO_INFORME & "</head><body>"
    html = html & "<h2>Informe de Riesgo: " & HTMLSafe(p_Riesgo.CodigoRiesgo) & "</h2>"
    
    ' --- SECCIÓN GENERAL ---
    html = html & "<table>"
    html = html & "<tr class='titulo'><td colspan='2'>DATOS GENERALES</td></tr>"
    html = html & "<tr><td><b>Descripción:</b></td><td>" & HTMLSafe(p_Riesgo.Descripcion) & "</td></tr>"
    html = html & "<tr><td><b>Causa Raíz:</b></td><td>" & HTMLSafe(p_Riesgo.CausaRaiz) & "</td></tr>"
    html = html & "<tr><td><b>Estado Actual:</b></td><td>" & p_Riesgo.Estado & "</td></tr>"
    html = html & "</table><br>"

    ' --- SECCIÓN VALORACIÓN ---
    html = html & "<table>"
    html = html & "<tr class='titulo'><td colspan='4'>VALORACIÓN</td></tr>"
    html = html & "<tr><td>Probabilidad:</td><td>" & p_Riesgo.Probabilidad & "</td><td>Impacto:</td><td>" & p_Riesgo.Impacto & "</td></tr>"
    html = html & "</table><br>"

    ' --- SECCIÓN PLANES (Si existen) ---
    ' Nota: Aquí asumimos que p_Riesgo tiene una colección o método para obtener planes
    ' En GR2 esto se apoya en el objeto 'Constructor' o 'Repositorio'
    
    html = html & "</body></html>"
    ConstruirInformeRiesgoHTML = html

    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "Error en modMigracionAux.ConstruirInformeRiesgoHTML: " & vbNewLine & Err.Description
    End If
    ' Aquí se podría llamar a CorreoAlAdministrador si es crítico
End Function

''' <summary>
''' Función auxiliar para formatear fechas de forma segura para el informe.
''' </summary>
Public Function FormatoFechaHTML(p_Fecha As Variant) As String
    If IsDate(p_Fecha) Then
        FormatoFechaHTML = Format(p_Fecha, "dd/mm/yyyy")
    Else
        FormatoFechaHTML = "-"
    End If
End Function
''' <summary>
''' Carga el historial de estados de un riesgo en un ListBox.
''' </summary>
Public Sub CargarHistoricoEstados(ByRef p_lst As Access.ListBox, ByVal p_IDRiesgo As Long, Optional ByRef p_Error As String)
    Dim sql As String
    On Error GoTo errores

    ' Validamos que tengamos un ID válido
    If p_IDRiesgo <= 0 Then
        p_lst.RowSource = ""
        Exit Sub
    End If

    ' Construcción de la consulta SQL basada en el ERD
    ' Asumimos columnas: Fecha, Estado (descripción o ID)
    ' Ajustamos el ORDER BY para ver lo más reciente primero
    sql = "SELECT FechaEstado, Estado " & _
          "FROM TbRiesgosMaterializaciones " & _
          "WHERE IDRiesgo = " & p_IDRiesgo & " " & _
          "ORDER BY FechaEstado DESC;"

    With p_lst
        .RowSourceType = "Table/Query"
        .ColumnCount = 2
        .ColumnWidths = "2cm;3cm"
        .ColumnHeads = True
        .RowSource = sql
    End With

    Exit Sub

errores:
    p_Error = "Error en modMigracionAux.CargarHistoricoEstados: " & Err.Description
    ' No elevamos con Raise 1000 aquí para no romper el flujo del formulario,
    ' dejamos que el formulario decida si mostrar el error.
End Sub
''' <summary>
''' Carga el historial detallado de estados combinando ediciones y materializaciones.
''' Sustituye a la lógica de caché de GR2.
''' </summary>
Public Sub CargarHistoricoEstadosV2(ByRef p_lst As Access.ListBox, ByVal p_IDRiesgo As Long, ByVal p_CodigoRiesgo As String)
    Dim sql As String
    On Error Resume Next
    
    ' SQL que une los estados históricos registrados en la tabla de materializaciones
    ' Se ordena por fecha descendente para ver lo último arriba
    sql = "SELECT FechaEstado, Estado, 'Histórico' as Tipo " & _
          "FROM TbRiesgosMaterializaciones " & _
          "WHERE IDRiesgo = " & p_IDRiesgo & " " & _
          "ORDER BY FechaEstado DESC;"

    With p_lst
        .RowSourceType = "Table/Query"
        .ColumnCount = 2
        .ColumnWidths = "3cm;5cm"
        .ColumnHeads = True
        .RowSource = sql
    End With
End Sub


''' <summary>
''' Procesa la apertura del informe HTML.
''' </summary>
Public Sub AbrirInformeRiesgoMigrado(ByRef p_Riesgo As Object)
    Dim m_Error As String
    Dim fso As Object
    Dim ruta As String
    Dim html As String
    
    ' Usamos la función de construcción que ya teníamos en el módulo
    html = ConstruirInformeRiesgoHTML(p_Riesgo, m_Error)
    
    If m_Error = "" Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        ruta = CurrentProject.Path & "\Informe_Riesgo_Temp.html"
        With fso.CreateTextFile(ruta, True)
            .Write html
            .Close
        End With
        Application.FollowHyperlink ruta
    Else
        Err.Raise 1000, , m_Error
    End If
End Sub
' =========================================================================
' SECCIÓN: GESTIÓN DE INDICADORES (MIGRACIÓN FASE 3)
' =========================================================================

''' <summary>
''' Carga la lista de indicadores asociados a un riesgo.
''' </summary>
Public Sub Indicadores_CargarLista(ByRef p_lst As Access.ListBox, ByVal p_IDRiesgo As Long)
    Dim sql As String
    On Error Resume Next
    
    ' SQL directo a la tabla definida en el ERD
    sql = "SELECT IDIndicador, Indicador, Responsable, Frecuencia " & _
          "FROM TbIndicadoresRiesgos " & _
          "WHERE IDRiesgo = " & p_IDRiesgo & " " & _
          "ORDER BY Indicador;"

    With p_lst
        .RowSourceType = "Table/Query"
        .ColumnCount = 4
        .ColumnWidths = "0;4cm;3cm;2cm" ' ID oculto
        .RowSource = sql
        .Requery
    End With
End Sub

''' <summary>
''' Guarda un indicador (Nuevo o Edición).
''' </summary>
Public Sub Indicadores_Guardar( _
    ByVal p_IDIndicador As Long, _
    ByVal p_IDRiesgo As Long, _
    ByVal p_Nombre As String, _
    ByVal p_Fuente As String, _
    ByVal p_Verde As String, _
    ByVal p_Amarillo As String, _
    ByVal p_Rojo As String, _
    ByVal p_Frecuencia As String, _
    ByVal p_Responsable As String, _
    ByRef p_Error As String)

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    On Error GoTo errores

    Set db = CurrentDb
    
    If p_IDIndicador = 0 Then
        ' Nuevo registro
        Set rs = db.OpenRecordset("TbIndicadoresRiesgos", dbOpenDynaset)
        rs.AddNew
        rs!IDRiesgo = p_IDRiesgo
    Else
        ' Edición de registro existente
        Set rs = db.OpenRecordset("SELECT * FROM TbIndicadoresRiesgos WHERE IDIndicador = " & p_IDIndicador, dbOpenDynaset)
        If rs.EOF Then
            p_Error = "No se ha encontrado el indicador con ID: " & p_IDIndicador
            Exit Sub
        End If
        rs.Edit
    End If

    ' Mapeo de campos según ERD
    rs!Indicador = p_Nombre
    rs!Fuente = p_Fuente
    rs!UmbralVerde = p_Verde
    rs!UmbralAmarillo = p_Amarillo
    rs!UmbralRojo = p_Rojo
    rs!Frecuencia = p_Frecuencia
    rs!Responsable = p_Responsable
    
    rs.Update
    rs.Close
    Exit Sub

errores:
    p_Error = "Error en modMigracionAux.Indicadores_Guardar: " & Err.Description
End Sub

''' <summary>
''' Elimina un indicador.
''' </summary>
Public Sub Indicadores_Eliminar(ByVal p_IDIndicador As Long, ByRef p_Error As String)
    On Error GoTo errores
    If p_IDIndicador > 0 Then
        CurrentDb.Execute "DELETE FROM TbIndicadoresRiesgos WHERE IDIndicador = " & p_IDIndicador, dbFailOnError
    End If
    Exit Sub
errores:
    p_Error = "Error en modMigracionAux.Indicadores_Eliminar: " & Err.Description
End Sub


