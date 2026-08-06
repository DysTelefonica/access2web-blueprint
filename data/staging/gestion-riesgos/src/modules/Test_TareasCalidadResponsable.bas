Attribute VB_Name = "Test_TareasCalidadResponsable"
Option Compare Database
Option Explicit

Private Function BuildOk(ByVal value As Variant, ByRef logs() As String) As String
    BuildOk = Test_Helper.BuildJsonOk(value, logs)
End Function

Private Function BuildFail(ByVal msg As String, ByRef logs() As String) As String
    BuildFail = Test_Helper.BuildJsonFail(msg, logs)
End Function

Private Function NormalizeSqlForAssert(ByVal sql As String) As String
    sql = Replace(sql, vbCrLf, " ")
    sql = Replace(sql, vbCr, " ")
    sql = Replace(sql, vbLf, " ")
    sql = Replace(sql, vbTab, " ")
    Do While InStr(1, sql, "  ", vbBinaryCompare) > 0
        sql = Replace(sql, "  ", " ")
    Loop
    NormalizeSqlForAssert = Trim$(sql)
End Function

Private Sub SaveCurrentRoleState(ByRef p_EsAdministrador As EnumSiNo, ByRef p_EsCalidad As EnumSiNo, ByRef p_EsTecnico As EnumSiNo)
    p_EsAdministrador = EsAdministrador
    p_EsCalidad = EsCalidad
    p_EsTecnico = EsTecnico
End Sub

Private Sub RestoreCurrentRoleState(ByVal p_EsAdministrador As EnumSiNo, ByVal p_EsCalidad As EnumSiNo, ByVal p_EsTecnico As EnumSiNo)
    EsAdministrador = p_EsAdministrador
    EsCalidad = p_EsCalidad
    EsTecnico = p_EsTecnico
End Sub

Public Function Test_TareasCalidadResponsable_Caducadas_AdminVacioSinFiltroResponsable() As String
    On Error GoTo EH

    Dim logs(0 To 3) As String
    logs(0) = "1. Arrange: admin con responsable vacio"
    logs(1) = "2. Act: getSQLEdicionesCaducadasSinPropuesta"
    logs(2) = "3. Assert: no filtra por NombreUsuarioCalidad"
    logs(3) = "4. Assert: no usa CadenaNombreAutorizados del conectado"

    Dim prevAdmin As EnumSiNo
    Dim prevCalidad As EnumSiNo
    Dim prevTecnico As EnumSiNo
    SaveCurrentRoleState prevAdmin, prevCalidad, prevTecnico

    Dim prevObjUsuarioParaTareas As Usuario
    Set prevObjUsuarioParaTareas = m_ObjUsuarioParaTareas

    Dim prevObjUsuarioConectado As Usuario
    Set prevObjUsuarioConectado = m_ObjUsuarioConectado

    EsAdministrador = EnumSiNo.Sí
    EsCalidad = EnumSiNo.No
    EsTecnico = EnumSiNo.No

    Set m_ObjUsuarioParaTareas = Nothing
    Set m_ObjUsuarioConectado = New Usuario
    m_ObjUsuarioConectado.Nombre = "USUARIO_CONECTADO_NO_USAR"

    Dim obj As TareasCalidad
    Set obj = New TareasCalidad

    Dim errMsg As String
    Dim sql As String
    sql = NormalizeSqlForAssert(obj.getSQLEdicionesCaducadasSinPropuesta(errMsg))
    If errMsg <> "" Then
        Test_TareasCalidadResponsable_Caducadas_AdminVacioSinFiltroResponsable = BuildFail(errMsg, logs)
        GoTo Cleanup
    End If

    If InStr(1, sql, "TbProyectos.NombreUsuarioCalidad", vbTextCompare) > 0 Then
        Test_TareasCalidadResponsable_Caducadas_AdminVacioSinFiltroResponsable = BuildFail("Admin sin responsable no debe filtrar por NombreUsuarioCalidad", logs)
        GoTo Cleanup
    End If

    If InStr(1, sql, "TbProyectos.CadenaNombreAutorizados", vbTextCompare) > 0 Then
        Test_TareasCalidadResponsable_Caducadas_AdminVacioSinFiltroResponsable = BuildFail("Admin sin responsable no debe filtrar por CadenaNombreAutorizados", logs)
        GoTo Cleanup
    End If

    Test_TareasCalidadResponsable_Caducadas_AdminVacioSinFiltroResponsable = BuildOk("caducadas_admin_vacio_sin_filtro_responsable", logs)

Cleanup:
    Set m_ObjUsuarioParaTareas = prevObjUsuarioParaTareas
    Set m_ObjUsuarioConectado = prevObjUsuarioConectado
    RestoreCurrentRoleState prevAdmin, prevCalidad, prevTecnico
    Exit Function
EH:
    Test_TareasCalidadResponsable_Caducadas_AdminVacioSinFiltroResponsable = BuildFail("Test_TareasCalidadResponsable_Caducadas_AdminVacioSinFiltroResponsable: " & Err.Description, logs)
    Resume Cleanup
End Function

Public Function Test_TareasCalidadResponsable_Caducadas_AdminSeleccionadoFiltraNombreUsuarioCalidad() As String
    On Error GoTo EH

    Dim logs(0 To 3) As String
    logs(0) = "1. Arrange: admin con responsable seleccionado"
    logs(1) = "2. Act: getSQLEdicionesCaducadasSinPropuesta"
    logs(2) = "3. Assert: filtra por NombreUsuarioCalidad del seleccionado"
    logs(3) = "4. Assert: no usa CadenaNombreAutorizados del conectado"

    Dim prevAdmin As EnumSiNo
    Dim prevCalidad As EnumSiNo
    Dim prevTecnico As EnumSiNo
    SaveCurrentRoleState prevAdmin, prevCalidad, prevTecnico

    Dim prevObjUsuarioParaTareas As Usuario
    Set prevObjUsuarioParaTareas = m_ObjUsuarioParaTareas

    Dim prevObjUsuarioConectado As Usuario
    Set prevObjUsuarioConectado = m_ObjUsuarioConectado

    EsAdministrador = EnumSiNo.Sí
    EsCalidad = EnumSiNo.No
    EsTecnico = EnumSiNo.No

    Set m_ObjUsuarioParaTareas = New Usuario
    m_ObjUsuarioParaTareas.Nombre = "RESPONSABLE_OBJETIVO"

    Set m_ObjUsuarioConectado = New Usuario
    m_ObjUsuarioConectado.Nombre = "USUARIO_CONECTADO_NO_USAR"

    Dim obj As TareasCalidad
    Set obj = New TareasCalidad

    Dim errMsg As String
    Dim sql As String
    sql = NormalizeSqlForAssert(obj.getSQLEdicionesCaducadasSinPropuesta(errMsg))
    If errMsg <> "" Then
        Test_TareasCalidadResponsable_Caducadas_AdminSeleccionadoFiltraNombreUsuarioCalidad = BuildFail(errMsg, logs)
        GoTo Cleanup
    End If

    If InStr(1, sql, "TbProyectos.NombreUsuarioCalidad)='RESPONSABLE_OBJETIVO'", vbTextCompare) = 0 Then
        Test_TareasCalidadResponsable_Caducadas_AdminSeleccionadoFiltraNombreUsuarioCalidad = BuildFail("Admin con responsable debe filtrar por NombreUsuarioCalidad seleccionado", logs)
        GoTo Cleanup
    End If

    If InStr(1, sql, "TbProyectos.CadenaNombreAutorizados", vbTextCompare) > 0 Then
        Test_TareasCalidadResponsable_Caducadas_AdminSeleccionadoFiltraNombreUsuarioCalidad = BuildFail("Admin con responsable no debe usar CadenaNombreAutorizados del conectado", logs)
        GoTo Cleanup
    End If

    Test_TareasCalidadResponsable_Caducadas_AdminSeleccionadoFiltraNombreUsuarioCalidad = BuildOk("caducadas_admin_responsable_filtra_nombre_qual", logs)

Cleanup:
    Set m_ObjUsuarioParaTareas = prevObjUsuarioParaTareas
    Set m_ObjUsuarioConectado = prevObjUsuarioConectado
    RestoreCurrentRoleState prevAdmin, prevCalidad, prevTecnico
    Exit Function
EH:
    Test_TareasCalidadResponsable_Caducadas_AdminSeleccionadoFiltraNombreUsuarioCalidad = BuildFail("Test_TareasCalidadResponsable_Caducadas_AdminSeleccionadoFiltraNombreUsuarioCalidad: " & Err.Description, logs)
    Resume Cleanup
End Function
