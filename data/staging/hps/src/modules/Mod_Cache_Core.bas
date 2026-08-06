Attribute VB_Name = "Mod_Cache_Core"
Option Compare Database
Option Explicit

' =============================================================================================
' NOMBRE:       Mod_Cache_Core
' DESCRIPCIÓN:  Reescritura robusta de la gestión de caché de usuarios.
'               Soluciona el cruce de datos (ID 815/850) mediante borrado preventivo.
' =============================================================================================

''' <summary>
''' Regenera la caché de TODOS los usuarios listados en TbUsuarios.
''' Usa Transacción Global: Si falla uno, no se rompe nada.
''' </summary>
Public Function cache_usuarios_regenerar1(Optional ByRef p_Error As String) As String
    Dim ws As DAO.Workspace
    Dim rcdUsuarios As DAO.Recordset
    Dim m_ID As String
    Dim m_DatosLocal As DatosLocal
    Dim lContador As Long
    
    On Error GoTo errores
    
    Set ws = DBEngine.Workspaces(0)
    
    ' 1. INICIAR TRANSACCIÓN GLOBAL
    ' Esto evita que si falla el proceso al 50%, el usuario se quede con la base vacía.
    ws.BeginTrans
    
    ' Nota: Ya no hacemos un DELETE * masivo al principio.
    ' Es mejor ir usuario por usuario reemplazando, o si prefieres velocidad extrema,
    ' descomenta las siguientes líneas, pero es más seguro hacerlo fila a fila por si cancelas.
    ' CurrentDb.Execute "DELETE FROM TbDatosLocal", dbFailOnError
    ' CurrentDb.Execute "DELETE FROM TbDatosLocalParaIndicadores", dbFailOnError
    ' getdb().Execute "DELETE FROM TbUsuariosEntidades", dbFailOnError
    
    m_SQL = "SELECT ID FROM TbUsuarios WHERE F_Baja IS NULL" ' F_Baja IS NULL = usuario activo (sin baja)
    Set rcdUsuarios = getdb().OpenRecordset(m_SQL, dbOpenSnapshot)
    
    If Not rcdUsuarios.EOF Then
        rcdUsuarios.MoveLast
        Dim lTotal As Long: lTotal = rcdUsuarios.RecordCount
        rcdUsuarios.MoveFirst
        
        Do While Not rcdUsuarios.EOF
            lContador = lContador + 1
            m_ID = rcdUsuarios!ID
            
            ' Feedback visual
            Avance "Sincronizando ID " & m_ID & " (" & lContador & "/" & lTotal & ")"
            DoEvents
            
            ' 2. OBTENER DATOS CALCULADOS (Tu lógica de negocio compleja)
            Set m_DatosLocal = Nothing ' Limpieza explícita de puntero
            Set m_DatosLocal = getDatosLocalDeUsuario(p_ID:=m_ID, p_Error:=p_Error)
            
            If p_Error <> "" Then
                ' Si falla la lógica de negocio de un usuario, decidimos:
                ' ¿Paramos todo? O ¿Saltamos al siguiente?
                ' Aquí paro para proteger la integridad.
                Err.Raise 1000, , "Fallo calculando ID " & m_ID & ": " & p_Error
            End If
            
            ' 3. PLANCHAR DATOS EN TABLAS CACHÉ
            ' Pasamos el objeto ws para (opcionalmente) gestionar transacciones anidadas,
            ' pero como ya estamos en una, basta con llamar a la función.
            cache_usuario_regenerar1 p_DatosLocal:=m_DatosLocal, p_Error:=p_Error
            
    If p_Error <> "" Then
        Err.Raise 1000, , p_Error  ' propagar con descripción
    End If
            
            rcdUsuarios.MoveNext
        Loop
    End If
    
    ' 4. CONFIRMAR CAMBIOS
    ws.CommitTrans
    cache_usuarios_regenerar1 = "OK"
    Debug.Print "Proceso finalizado. Usuarios procesados: " & lContador
    
SALIR:
    On Error Resume Next
    rcdUsuarios.Close: Set rcdUsuarios = Nothing
    Set ws = Nothing
    Exit Function
    
errores:
    ws.Rollback ' DESHACER TODO SI FALLA
    If Err.Number <> 1000 Then
        p_Error = "Error CRÍTICO en cache_usuarios_regenerar1: " & Err.Description
    ElseIf p_Error = "" Then
        ' Error 1000 propagado desde cache_usuario_regenerar1 sin descripción en p_Error
        ' Usar Err.Description directamente ya que Err.Raise 1000,,p_Error preservó la descripción
        p_Error = "Error en cache_usuarios_regenerar1: " & Err.Description
    End If
    Debug.Print p_Error
    Resume SALIR
End Function

''' <summary>
''' Regenera la caché de UN solo usuario.
''' ESTRATEGIA: Borrar primero -> Insertar Nuevo. (Evita el .Edit corrupto)
''' </summary>
Public Function cache_usuario_regenerar1(Optional p_ID As String, _
                                        Optional p_DatosLocal As DatosLocal, _
                                        Optional ByRef p_Error As String) As String
                                        
    Dim m_DatosLocal As DatosLocal
    Dim m_ID As String
    
    On Error GoTo errores
    
    ' 1. Obtención del objeto de datos (Si no viene dado)
    If p_DatosLocal Is Nothing Then
        If p_ID = "" Then Err.Raise 1001, , "Se requiere ID o Objeto DatosLocal"
        Set m_DatosLocal = getDatosLocalDeUsuario(p_ID:=p_ID, p_Error:=p_Error)
        If p_Error <> "" Then Err.Raise 1000, , p_Error
    Else
        Set m_DatosLocal = p_DatosLocal
    End If
    
    If m_DatosLocal Is Nothing Then
        cache_usuario_regenerar1 = vbNullString
        Exit Function
    End If
    m_ID = m_DatosLocal.ID
    
    ' 2. LIMPIEZA PREVENTIVA (El secreto para que no se mezclen datos)
    ' Borramos de las 3 tablas antes de escribir nada.
    cache_usuario_borrar1 p_ID:=m_ID, p_Error:=p_Error
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    ' 3. INSERCIÓN EN LAS 3 TABLAS
    ' Usamos una función auxiliar para no repetir código 3 veces (DRY)

    ' Tabla 1: TbUsuariosEntidades (Remota/Compartida)
    ActualizarTablaGenerica getdb(), "TbUsuariosEntidades", m_DatosLocal, p_Error
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    ' Tabla 2: TbDatosLocal (Local)
    ActualizarTablaGenerica CurrentDb(), "TbDatosLocal", m_DatosLocal, p_Error
    If p_Error <> "" Then Err.Raise 1000, , p_Error

    ' Tabla 3: TbDatosLocalParaIndicadores (Local)
    ActualizarTablaGenerica CurrentDb(), "TbDatosLocalParaIndicadores", m_DatosLocal, p_Error
    If p_Error <> "" Then Err.Raise 1000, , p_Error
    
    cache_usuario_regenerar1 = "OK"
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "Error en cache_usuario_regenerar1 (ID " & m_ID & "): " & Err.Description
    ElseIf p_Error = "" Then
        ' Error 1000 propagado desde getDatosLocalDeUsuario sin descripción explícita
        p_Error = "Error en cache_usuario_regenerar1 (ID " & m_ID & "): " & Err.Description
    End If
    cache_usuario_regenerar1 = vbNullString
End Function

''' <summary>
''' Función auxiliar privada para escribir un objeto DatosLocal en cualquier tabla.
''' Solo hace INSERT (AddNew) porque asumimos que el registro ya se borró.
''' </summary>
Private Sub ActualizarTablaGenerica(db As DAO.Database, sTabla As String, oDatos As DatosLocal, ByRef sError As String)
    Dim rs As DAO.Recordset
    Dim vCampo As Variant
    Dim sValor As String

    On Error GoTo err_gen

    Set rs = db.OpenRecordset("SELECT * FROM " & sTabla & " WHERE 1=0", dbOpenDynaset) ' Abrimos vacío para añadir

    rs.AddNew
    rs!ID = oDatos.ID ' Aseguramos el ID primero

    For Each vCampo In oDatos.ColCampos
        ' Omitimos el ID si está en la colección para no duplicar error
        If UCase(vCampo) <> "ID" Then
            ' Obtenemos valor
            sValor = oDatos.getPropiedad(vCampo, sError)
            If sError <> "" Then
                Err.Raise 1000, , sError  ' propagar con descripción
            End If

            ' Intentamos asignar. Si el campo no existe en la tabla destino, lo ignoramos.
            On Error Resume Next
            If sValor <> "" Then
                rs.Fields(vCampo).value = sValor
            Else
                rs.Fields(vCampo).value = Null
            End If

            If Err.Number <> 0 And Err.Number <> 3265 Then ' 3265 = Elemento no encontrado en colección (campo no existe en tabla)
                ' Si es otro error (ej: tipo de dato), lo reportamos
                sError = "Error campo " & vCampo & ": " & Err.Description
                On Error GoTo err_gen
                Err.Raise 1000, , sError  ' propagar con descripción
            End If
            On Error GoTo err_gen
        End If
    Next vCampo

    rs.Update
    rs.Close
    Set rs = Nothing
    Exit Sub
    
err_gen:
    If Err.Number <> 1000 Then sError = Err.Description
    If Not rs Is Nothing Then rs.Close
End Sub

''' <summary>
''' Borra físicamente al usuario de las tablas caché.
''' </summary>
Public Function cache_usuario_borrar1(p_ID As String, Optional ByRef p_Error As String) As String
    On Error GoTo errores
    
    ' Usamos dbFailOnError para asegurarnos de que si la tabla está bloqueada, nos enteremos.
    CurrentDb.Execute "DELETE FROM TbDatosLocal WHERE ID=" & p_ID, dbFailOnError
    CurrentDb.Execute "DELETE FROM TbDatosLocalParaIndicadores WHERE ID=" & p_ID, dbFailOnError
    
    ' Nota: getdb() suele referirse a la vinculada. Asegúrate que tienes permisos de borrado allí.
    getdb().Execute "DELETE FROM TbUsuariosEntidades WHERE ID=" & p_ID, dbFailOnError
    
    cache_usuario_borrar1 = "OK"
    Exit Function

errores:
    p_Error = "Error borrando caché ID " & p_ID & ": " & Err.Description
End Function


