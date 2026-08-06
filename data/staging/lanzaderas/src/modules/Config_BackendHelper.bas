Attribute VB_Name = "Config_BackendHelper"
Option Compare Database
Option Explicit

Private Const LANZADERA_BACKEND_LOCAL As String = "C:\00repos\datos\Lanzadera_Datos.accdb"
Private Const LANZADERA_BACKEND_REMOTO As String = "\\datoste\aplicaciones_dys\Aplicaciones PpD\0Lanzadera\Lanzadera_Datos.accdb"
Private Const LANZADERA_RUTA_APP_REMOTA As String = "\\datoste\aplicaciones_dys\Aplicaciones PpD\0Lanzadera"
Private Const LANZADERA_RUTA_APP_LOCAL As String = "C:\Users\adm1\Telefonica\Aplicaciones_dys.TMETF - Aplicaciones PpD\0Lanzadera"
Private Const LANZADERA_ID_APLICACION As String = "12"
Private Const LANZADERA_EN_PRUEBAS As String = "No"
Private Const LANZADERA_PASSWORD_BACKEND As String = ""

Public Function ConfigurarBackendLanzadera_LocalConReposLocalYRemoto( _
                                                    Optional ByRef p_Error As String = "" _
                                                    ) As String
    ConfigurarBackendLanzadera_LocalConReposLocalYRemoto = ConfigurarTbConfiguracionBackends( _
                                                        p_BackendActivo:="LOCAL", _
                                                        p_BackendRemoto:=LANZADERA_BACKEND_REMOTO, _
                                                        p_BackendLocal:=LANZADERA_BACKEND_LOCAL, _
                                                        p_RutaTrabajoRemota:=LANZADERA_RUTA_APP_REMOTA, _
                                                        p_RutaTrabajoLocal:=LANZADERA_RUTA_APP_LOCAL, _
                                                        p_IDAplicacionDefault:=LANZADERA_ID_APLICACION, _
                                                        p_EnPruebasDefault:=LANZADERA_EN_PRUEBAS, _
                                                        p_PasswordBackendDefault:=LANZADERA_PASSWORD_BACKEND, _
                                                        p_Error:=p_Error)
End Function

Public Function ConfigurarTbConfiguracionBackends( _
                                        ByVal p_BackendActivo As String, _
                                        ByVal p_BackendRemoto As String, _
                                        ByVal p_BackendLocal As String, _
                                        ByVal p_RutaTrabajoRemota As String, _
                                        ByVal p_RutaTrabajoLocal As String, _
                                        Optional ByVal p_IDAplicacionDefault As String = "", _
                                        Optional ByVal p_EnPruebasDefault As String = "No", _
                                        Optional ByVal p_PasswordBackendDefault As String = "", _
                                        Optional ByRef p_Error As String = "" _
                                        ) As String
    On Error GoTo errores
    p_Error = ""

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim backendActivoNormalizado As String
    Dim idAplicacionPersistido As String
    Dim enPruebasPersistido As String
    Dim passwordBackendPersistido As String

    backendActivoNormalizado = UCase$(Trim$(p_BackendActivo))
    Select Case backendActivoNormalizado
        Case "LOCAL", "PROD"
        Case Else
            p_Error = "BackendActivo inválido para TbConfiguracionBackends: " & p_BackendActivo
            Err.Raise 1000
    End Select

    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT TOP 2 * FROM TbConfiguracionBackends", dbOpenDynaset)

    If rs.EOF Then
        idAplicacionPersistido = Trim$(p_IDAplicacionDefault)
        enPruebasPersistido = Trim$(p_EnPruebasDefault)
        passwordBackendPersistido = p_PasswordBackendDefault
        rs.AddNew
    Else
        rs.MoveNext
        If Not rs.EOF Then
            p_Error = "TbConfiguracionBackends tiene más de un registro; el helper no puede decidir cuál actualizar"
            Err.Raise 1000
        End If
        rs.MoveFirst
        idAplicacionPersistido = CStr(Nz(rs.Fields("IDAplicacion").value, p_IDAplicacionDefault))
        enPruebasPersistido = CStr(Nz(rs.Fields("EnPruebas").value, p_EnPruebasDefault))
        passwordBackendPersistido = CStr(Nz(rs.Fields("PasswordBackend").value, p_PasswordBackendDefault))
        rs.Edit
    End If

    If Trim$(idAplicacionPersistido) = "" Then idAplicacionPersistido = Trim$(p_IDAplicacionDefault)
    If Trim$(enPruebasPersistido) = "" Then enPruebasPersistido = Trim$(p_EnPruebasDefault)

    rs.Fields("BackendActivo").value = backendActivoNormalizado
    rs.Fields("BackendProduccion").value = Trim$(p_BackendRemoto)
    rs.Fields("BackendSandbox").value = Trim$(p_BackendLocal)
    rs.Fields("RutaDirectorioAplicacion_PROD").value = NormalizarRuta(Trim$(p_RutaTrabajoRemota))
    rs.Fields("RutaDirectorioAplicacion_LOCAL").value = NormalizarRuta(Trim$(p_RutaTrabajoLocal))
    rs.Fields("IDAplicacion").value = idAplicacionPersistido
    rs.Fields("EnPruebas").value = enPruebasPersistido
    rs.Fields("PasswordBackend").value = passwordBackendPersistido
    rs.Update

    ConfigurarTbConfiguracionBackends = "OK"
    Exit Function
errores:
    If Err.Number = 1000 Then
        If p_Error = "" Then p_Error = Err.Description
    Else
        p_Error = "El método ConfigurarTbConfiguracionBackends ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function


