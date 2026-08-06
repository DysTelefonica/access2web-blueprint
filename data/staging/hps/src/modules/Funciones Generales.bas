Attribute VB_Name = "Funciones Generales"
Option Compare Database
Option Explicit

Public Function AbiertoParaEscritura(p_CorreoUsuarioConectado As String, Optional ByRef p_Error As String) As EnumSiNo
    
    Dim m_Correo As String
    Dim m_SSID As Variant
    
    On Error GoTo errores
    If InStr(1, p_CorreoUsuarioConectado, "@") = 0 Then
        AbiertoParaEscritura = EnumSiNo.No
        Exit Function
    End If
    For Each m_SSID In m_ObjEntorno.colUsuariosPermitidosEnLocal
        m_Correo = m_ObjEntorno.colUsuariosPermitidosEnLocal(m_SSID)
        If p_CorreoUsuarioConectado = m_Correo Then
            AbiertoParaEscritura = EnumSiNo.Sí
            Exit Function
        End If
    Next
    AbiertoParaEscritura = EnumSiNo.No
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método AbiertoParaEscritura ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function
Public Function EstablecerSoloLectura(p_SoloLectura As EnumSiNo, Optional ByRef p_Error As String) As String
    
   
    
    On Error GoTo errores
    
    If p_SoloLectura = EnumSiNo.No Then
        
        If m_EnOficina = EnumSiNo.No Then
            p_Error = "No estamos en la oficina"
            Err.Raise 1000
        End If
       
        
        Application.TempVars("SoloLectura") = "No"
        m_EstadoConexion = "Lectura/Escritura"
        
    Else
        If m_EnOficina = EnumSiNo.No Then
            If m_ObjUsuarioConectado.PermisoEntradaLocalEnCasa <> EnumSiNo.Sí Then
                p_Error = "No está autorizado a entrar fuera de la oficina"
                Err.Raise 1000
            End If
        End If
        
        
        Application.TempVars("SoloLectura") = "Sí"
        m_EstadoConexion = "Solo lectura"
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerSoloLectura ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function


Public Function CorreoAlAdministrador( _
                                        P_MensajeError As String, _
                                        Optional ByRef p_Error As String _
                                        ) As String
                                        
   
    Dim m_mensaje As String
    Dim m_Asunto As String
    Dim m_Nombre As String
    Dim m_NombreFormulario As String
    Dim m_TextoEnOficina As String
    Dim m_Version As String
    
    Dim m_ObjCorreo As Correo
    On Error GoTo errores
    
    If P_MensajeError = "" Then
        p_Error = "No hay mensaje que enviar"
        Err.Raise 1000
    End If
    If Not m_ObjEntorno Is Nothing Then
        m_Version = "Ver: " & m_ObjEntorno.VersionAplicacion
    Else
        m_Version = "Ver: Desconocida"
    End If
    If m_EnOficina = Empty Then
        m_EnOficina = EnOficina(p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    If m_EnOficina = Empty Then
        m_TextoEnOficina = "En Oficina Desconocido"
    Else
        If m_EnOficina = EnumSiNo.Sí Then
            m_TextoEnOficina = "En Oficina"
        Else
            m_TextoEnOficina = "Fuera de Oficina"
        End If
    End If
    On Error Resume Next
    m_NombreFormulario = Application.Screen.ActiveForm.Name
    If Err.Number <> 0 Then
        Err.Clear
        m_NombreFormulario = "Desconocido"
    End If
    On Error GoTo errores
    m_Nombre = getNombreUsuarioConectado()
    
    m_Asunto = "Error en HPS " & m_Version & " " & m_Nombre & " " & m_TextoEnOficina
    m_mensaje = "FORMULARIO del ERROR: " & m_NombreFormulario & vbNewLine
    m_mensaje = m_mensaje & "<BR> </BR>" & vbNewLine
    On Error Resume Next
    m_mensaje = m_mensaje & "NOMBRE EQUIPO: " & VBA.Environ("COMPUTERNAME")
    m_mensaje = m_mensaje & "<BR> </BR>" & vbNewLine
    On Error GoTo errores
    
    m_mensaje = m_mensaje & "DETALLE: " & P_MensajeError
    
    Set m_ObjCorreo = New Correo
    With m_ObjCorreo
        .Asunto = m_Asunto
        .Cuerpo = m_mensaje
        .Destinatarios = "ardelperal@gmail.com;andres.romandelperal@telefonica.com"
        .FechaGrabacion = Now()
        .EnviarCorreo p_Error
    End With
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CorreoAlAdministrador ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function







Public Function GenerarConsultasDeSubFormulario( _
                                                    p_form As Form, _
                                                    Optional p_URLDirectorio As String, _
                                                    Optional p_NombreExcel As String, _
                                                    Optional p_TodosLosCampos As EnumSiNo = EnumSiNo.Sí, _
                                                    Optional ByRef p_Error As String _
                                                    ) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim ctl As Control
    Dim m_Valor As String
    Dim WbLibro As Object
    Dim WbHoja As Object
    Dim intFila As Integer
    Dim Fila As Integer
    Dim Columna As Integer
    Dim m_URLExcel As String
    Dim m_URLCarpeta As String
    Dim m_NombreExcel As String
    
    Dim AppExcel As Object
    Dim m_ObjColCamposCalculados As Scripting.Dictionary
    Dim m_NombreCampo As String
    Dim m_Filtro As String
    Dim m_Where As String
    
    On Error GoTo errores
    p_Error = ""
    
    If p_TodosLosCampos = EnumSiNo.No Then
        Set m_ObjColCamposCalculados = New Scripting.Dictionary
        m_ObjColCamposCalculados.CompareMode = TextCompare
    
        With m_ObjColCamposCalculados
            .Add "Curso_Realizado", "Curso_Realizado"
            .Add "HPS_NAC_SIN_DATOS", "HPS_NAC_SIN_DATOS"
            .Add "HPS_NAC_Activo", "HPS_NAC_Activo"
            .Add "HPS_NAC_ApuntoDeCaducar", "HPS_NAC_ApuntoDeCaducar"
            .Add "HPS_NAC_Caducado", "HPS_NAC_Caducado"
            .Add "HPS_NAC_Baja", "HPS_NAC_Baja"
            
            .Add "HPS_NAC_PendienteRenovacion", "HPS_NAC_PendienteRenovacion"
            .Add "HPS_NAC_Solicitado", "HPS_NAC_Solicitado"
            .Add "HPS_NAC_MESES_PARA_RENOVAR", "HPS_NAC_MESES_PARA_RENOVAR"
            
            .Add "HPS_OTAN_SIN_DATOS", "HPS_OTAN_SIN_DATOS"
            .Add "HPS_OTAN_Activo", "HPS_OTAN_Activo"
            .Add "HPS_OTAN_ApuntoDeCaducar", "HPS_OTAN_ApuntoDeCaducar"
            .Add "HPS_OTAN_Caducado", "HPS_OTAN_Caducado"
            .Add "HPS_OTAN_Baja", "HPS_OTAN_Baja"
            
            .Add "HPS_OTAN_PendienteRenovacion", "HPS_OTAN_PendienteRenovacion"
            .Add "HPS_OTAN_Solicitado", "HPS_OTAN_Solicitado"
            .Add "HPS_OTAN_MESES_PARA_RENOVAR", "HPS_OTAN_MESES_PARA_RENOVAR"
            
            .Add "HPS_ESA_SIN_DATOS", "HPS_ESA_SIN_DATOS"
            .Add "HPS_ESA_Activo", "HPS_ESA_Activo"
            .Add "HPS_ESA_ApuntoDeCaducar", "HPS_ESA_ApuntoDeCaducar"
            .Add "HPS_ESA_Caducado", "HPS_ESA_Caducado"
            .Add "HPS_ESA_Baja", "HPS_ESA_Baja"
            
            .Add "HPS_ESA_PendienteRenovacion", "HPS_ESA_PendienteRenovacion"
            .Add "HPS_ESA_Solicitado", "HPS_ESA_Solicitado"
            .Add "HPS_ESA_MESES_PARA_RENOVAR", "HPS_ESA_MESES_PARA_RENOVAR"
            
            .Add "HPS_UE_SIN_DATOS", "HPS_UE_SIN_DATOS"
            .Add "HPS_UE_Activo", "HPS_UE_Activo"
            .Add "HPS_UE_ApuntoDeCaducar", "HPS_UE_ApuntoDeCaducar"
            .Add "HPS_UE_Caducado", "HPS_UE_Caducado"
            .Add "HPS_UE_Baja", "HPS_UE_Baja"
            
            .Add "HPS_UE_PendienteRenovacion", "HPS_UE_PendienteRenovacion"
            .Add "HPS_UE_Solicitado", "HPS_UE_Solicitado"
            .Add "HPS_UE_MESES_PARA_RENOVAR", "HPS_UE_MESES_PARA_RENOVAR"
        End With
    End If
    If p_URLDirectorio = "" Then
        m_URLCarpeta = m_ObjEntorno.URLDirectorioLocal
    Else
        m_URLCarpeta = p_URLDirectorio
    End If
    If Right(m_URLCarpeta, 1) = "\" Then
        m_URLCarpeta = Left(m_URLCarpeta, Len(m_URLCarpeta) - 1)
    End If
    If p_NombreExcel = "" Then
        m_NombreExcel = "ConsultaUsuario"
    Else
        m_NombreExcel = p_NombreExcel
    End If
    m_NombreExcel = fso.GetBaseName(m_NombreExcel)
    
    m_URLExcel = m_URLCarpeta & "\" & m_NombreExcel & ".xlsx"
    If fso.FileExists(m_URLExcel) Then
        If FicheroAbierto(m_URLExcel) Then
            p_Error = "Cierre la consulta anterior"
            Err.Raise 1000
        End If
        fso.DeleteFile m_URLExcel, True
    End If
    m_Filtro = p_form.Filter
    If m_Filtro <> "" Then
        m_SQL = p_form.RecordSource
        dato = Split(m_SQL, ";")
        m_SQL = dato(0) & " " & "WHERE " & m_Filtro & ";"
        If InStr(1, m_SQL, "[" & p_form.Name & "].") <> 0 Then
            m_SQL = Replace(m_SQL, "[" & p_form.Name & "].", "")
        End If
        If InStr(1, m_SQL, "(") <> 0 Then
            m_SQL = Replace(m_SQL, "(", "")
        End If
        If InStr(1, m_SQL, ")") <> 0 Then
            m_SQL = Replace(m_SQL, ")", "")
        End If
        If InStr(1, m_SQL, "[") <> 0 Then
            m_SQL = Replace(m_SQL, "[", "")
        End If
        If InStr(1, m_SQL, "]") <> 0 Then
            m_SQL = Replace(m_SQL, "]", "")
        End If
    Else
        m_SQL = p_form.RecordSource
    End If
    On Error Resume Next
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    If Err.Number <> 0 Then
        Err.Clear
        m_SQL = p_form.RecordSource
        On Error GoTo errores
        Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    End If
    On Error GoTo errores
    If rcdDatos.EOF Then
        rcdDatos.Close
        Set rcdDatos = Nothing
        p_Error = "No hay registros"
        Err.Raise 1000
    End If
    Set AppExcel = CreateObject("Excel.Application")
    AppExcel.Visible = False
    Set WbLibro = AppExcel.Workbooks.Add
    WbLibro.SaveAs m_URLExcel
    Set WbHoja = WbLibro.Worksheets(1)
    With WbHoja
        intFila = 1
        rcdDatos.MoveFirst
        Columna = 1
        For Each ctl In p_form.Controls
            If ctl.ControlType = 109 Then
                If ctl.ColumnHidden = False Then
                    .Cells(intFila, Columna).value = ctl.Name
                    .Range(.Cells(intFila, Columna), .Cells(intFila, Columna)).Font.Bold = True
                    Columna = Columna + 1
                End If
            End If
        Next
        intFila = intFila + 1
        Do While Not rcdDatos.EOF
            Columna = 1
            For Each ctl In p_form.Controls
                m_NombreCampo = ctl.Name
                If Not m_ObjColCamposCalculados Is Nothing Then
                    If m_ObjColCamposCalculados.Exists(m_NombreCampo) Then
                        GoTo siguienteColumna
                    End If
                End If
                If ctl.ControlType = 109 Then
                    If ctl.ColumnHidden = False Then
                        'If ctl.Name = "F_Inicio" Then Stop
                        'Debug.Print m_NombreCampo
                        m_Valor = Nz(rcdDatos.Fields(m_NombreCampo), "")
                        If ctl.Name = "IDExpediente" Then
                            m_Valor = "" & m_Valor
                        End If
                        If IsDate(m_Valor) Then
                            .Range(.Cells(intFila, Columna), .Cells(intFila, Columna)).NumberFormat = "dd/mm/yyyy"
                            .Cells(intFila, Columna).value = CDate(m_Valor)
                        Else
                            .Cells(intFila, Columna).value = m_Valor
                        End If
                        
                        Columna = Columna + 1
                    End If
                End If
siguienteColumna:
            Next
            intFila = intFila + 1
            rcdDatos.MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    With WbHoja.Range("A1").CurrentRegion
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    WbHoja.Cells.EntireColumn.AutoFit
    WbLibro.Close True
    Set WbLibro = Nothing
    AppExcel.Quit
    Set AppExcel = Nothing
    
    
    GenerarConsultasDeSubFormulario = m_URLExcel
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GenerarConsultasDeSubFormulario ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not WbLibro Is Nothing Then
        WbLibro.Close False
        Set WbLibro = Nothing
    End If
    If Not AppExcel Is Nothing Then
        AppExcel.Quit
        Set AppExcel = Nothing
    End If
    
End Function



Public Function ConsultaAExcel( _
                                p_form As Form, _
                                Optional p_URLDirectorio As String, _
                                Optional p_NombreExcel As String, _
                                Optional p_TodosLosCampos As EnumSiNo = EnumSiNo.Sí, _
                                Optional ByRef p_Error As String _
                                ) As String
   
    Dim db As DAO.Database
    Dim m_SQL As String
    Dim m_URLExcel As String
    Dim m_URLCarpeta As String
    Dim m_NombreExcel As String
    Dim m_NombreConsulta As String
    Dim m_Consulta As QueryDef
        
    On Error GoTo errores
    
    
    If p_URLDirectorio = "" Then
        m_URLCarpeta = m_ObjEntorno.URLDirectorioLocal
    Else
        m_URLCarpeta = p_URLDirectorio
    End If
    If Right(m_URLCarpeta, 1) = "\" Then
        m_URLCarpeta = Left(m_URLCarpeta, Len(m_URLCarpeta) - 1)
    End If
    If p_NombreExcel = "" Then
        m_NombreExcel = "ConsultaUsuario"
    Else
        m_NombreExcel = p_NombreExcel
    End If
    m_NombreExcel = fso.GetBaseName(m_NombreExcel)
    
    m_URLExcel = m_URLCarpeta & "\" & m_NombreExcel & ".xlsx"
    If fso.FileExists(m_URLExcel) Then
        If FicheroAbierto(m_URLExcel) Then
            p_Error = "Cierre la consulta anterior"
            Err.Raise 1000
        End If
        fso.DeleteFile m_URLExcel, True
    End If
    m_SQL = p_form.RecordSource
    
    Set db = CurrentDb()
    m_NombreConsulta = fso.GetBaseName(fso.GetTempName)
    Set m_Consulta = db.CreateQueryDef(m_NombreConsulta, m_SQL)
    DoCmd.OutputTo ObjectType:=acOutputQuery, ObjectName:=m_Consulta.Name, OutputFormat:=acFormatXLSX, Outputfile:=m_URLExcel
    db.QueryDefs.Delete m_NombreConsulta
    ConsultaAExcel = m_URLExcel
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ConsultaAExcel ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
    
End Function

Public Function GenerarConsultaPorSQL( _
                                        p_SQL As String, _
                                        Optional p_URLDirectorio As String, _
                                        Optional p_NombreExcel As String, _
                                        Optional ByRef p_Error As String _
                                        ) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Valor As String
    Dim WbLibro As Object
    Dim WbHoja As Object
    Dim intFila As Integer
    Dim Fila As Integer
    Dim Columna As Integer
    Dim m_URLExcel As String
    Dim m_URLCarpeta As String
    Dim m_NombreExcel As String
    
    Dim AppExcel As Object
    Dim fld As Object
    Dim m_NombreCampo As Variant
    
    On Error GoTo errores
    
        
    If p_URLDirectorio = "" Then
        m_URLCarpeta = m_ObjEntorno.URLDirectorioLocal
    Else
        m_URLCarpeta = p_URLDirectorio
    End If
    If Right(m_URLCarpeta, 1) = "\" Then
        m_URLCarpeta = Left(m_URLCarpeta, Len(m_URLCarpeta) - 1)
    End If
    If p_NombreExcel = "" Then
        m_NombreExcel = "ConsultaUsuario"
    Else
        m_NombreExcel = p_NombreExcel
    End If
    m_NombreExcel = fso.GetBaseName(m_NombreExcel)
    
    m_URLExcel = m_URLCarpeta & "\" & m_NombreExcel & ".xlsx"
    If fso.FileExists(m_URLExcel) Then
        If FicheroAbierto(m_URLExcel) Then
            p_Error = "Cierre la consulta anterior"
            Err.Raise 1000
        End If
        fso.DeleteFile m_URLExcel, True
    End If
    Set rcdDatos = CurrentDb().OpenRecordset(p_SQL)
    If rcdDatos.EOF Then
        rcdDatos.Close
        Set rcdDatos = Nothing
        p_Error = "No hay registros"
        Err.Raise 1000
    End If
    Set AppExcel = CreateObject("Excel.Application")
    AppExcel.Visible = False
    Set WbLibro = AppExcel.Workbooks.Add
    WbLibro.SaveAs m_URLExcel
    Set WbHoja = WbLibro.Worksheets(1)
    With WbHoja
        intFila = 1
        rcdDatos.MoveFirst
        Columna = 1
        For Each fld In rcdDatos.Fields
            .Cells(intFila, Columna).value = fld.Name
            .Range(.Cells(intFila, Columna), .Cells(intFila, Columna)).Font.Bold = True
            Columna = Columna + 1
        Next
        intFila = intFila + 1
        Do While Not rcdDatos.EOF
            Columna = 1
            For Each fld In rcdDatos.Fields
                m_Valor = Nz(fld.value, "")
                m_NombreCampo = fld.Name
                If CStr(m_NombreCampo) = "IDExpediente" Then
                    m_Valor = "" & m_Valor
                End If
                If IsDate(m_Valor) Then
                    .Range(.Cells(intFila, Columna), .Cells(intFila, Columna)).NumberFormat = "dd/mm/yyyy"
                    .Cells(intFila, Columna).value = CDate(m_Valor)
                Else
                    .Cells(intFila, Columna).value = m_Valor
                End If
                
                Columna = Columna + 1
                    
            Next
            intFila = intFila + 1
            rcdDatos.MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    With WbHoja.Range("A1").CurrentRegion
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    WbHoja.Cells.EntireColumn.AutoFit
    WbLibro.Close True
    Set WbLibro = Nothing
    AppExcel.Quit
    Set AppExcel = Nothing
    
    
    GenerarConsultaPorSQL = m_URLExcel
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GenerarConsultaPorSQL ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not WbLibro Is Nothing Then
        WbLibro.Close False
        Set WbLibro = Nothing
    End If
    If Not AppExcel Is Nothing Then
        AppExcel.Quit
        Set AppExcel = Nothing
    End If
    
End Function


Public Function GenerarConsultaConObservaciones( _
                                                    p_form As Form, _
                                                    Optional p_URLDirectorio As String, _
                                                    Optional p_NombreExcel As String, _
                                                    Optional p_TodosLosCampos As EnumSiNo = EnumSiNo.Sí, _
                                                    Optional ByRef p_Error As String _
                                                    ) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día x
    '   -Modificaciones:
    '
    '   -Funcionamiento:
    
    '   -Llamada por:
   
    '   -Devuelve:
    '       GenerarConsultasDeSubFormulario = m_URLExcel
    '-------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset
    Dim ctl As Control
    Dim m_Valor As String
    Dim WbLibro As Object
    Dim WbHoja As Object
    Dim intFila As Integer
    Dim Fila As Integer
    Dim Columna As Integer
    Dim m_URLExcel As String
    Dim m_URLCarpeta As String
    Dim m_NombreExcel As String
    
    Dim AppExcel As Object
    Dim m_ObjColCamposCalculados As Scripting.Dictionary
    Dim m_NombreCampo As String
    Dim m_CadenaObservaciones As String
    Dim m_ID As String
    On Error GoTo errores
    
    
    If p_TodosLosCampos = EnumSiNo.No Then
        Set m_ObjColCamposCalculados = New Scripting.Dictionary
        m_ObjColCamposCalculados.CompareMode = TextCompare
    
        With m_ObjColCamposCalculados
            .Add "Curso_Realizado", "Curso_Realizado"
            .Add "HPS_NAC_SIN_DATOS", "HPS_NAC_SIN_DATOS"
            .Add "HPS_NAC_Activo", "HPS_NAC_Activo"
            .Add "HPS_NAC_ApuntoDeCaducar", "HPS_NAC_ApuntoDeCaducar"
            .Add "HPS_NAC_Caducado", "HPS_NAC_Caducado"
            .Add "HPS_NAC_Baja", "HPS_NAC_Baja"
            
            .Add "HPS_NAC_PendienteRenovacion", "HPS_NAC_PendienteRenovacion"
            .Add "HPS_NAC_Solicitado", "HPS_NAC_Solicitado"
            .Add "HPS_NAC_MESES_PARA_RENOVAR", "HPS_NAC_MESES_PARA_RENOVAR"
            
            .Add "HPS_OTAN_SIN_DATOS", "HPS_OTAN_SIN_DATOS"
            .Add "HPS_OTAN_Activo", "HPS_OTAN_Activo"
            .Add "HPS_OTAN_ApuntoDeCaducar", "HPS_OTAN_ApuntoDeCaducar"
            .Add "HPS_OTAN_Caducado", "HPS_OTAN_Caducado"
            .Add "HPS_OTAN_Baja", "HPS_OTAN_Baja"
            
            .Add "HPS_OTAN_PendienteRenovacion", "HPS_OTAN_PendienteRenovacion"
            .Add "HPS_OTAN_Solicitado", "HPS_OTAN_Solicitado"
            .Add "HPS_OTAN_MESES_PARA_RENOVAR", "HPS_OTAN_MESES_PARA_RENOVAR"
            
            .Add "HPS_ESA_SIN_DATOS", "HPS_ESA_SIN_DATOS"
            .Add "HPS_ESA_Activo", "HPS_ESA_Activo"
            .Add "HPS_ESA_ApuntoDeCaducar", "HPS_ESA_ApuntoDeCaducar"
            .Add "HPS_ESA_Caducado", "HPS_ESA_Caducado"
            .Add "HPS_ESA_Baja", "HPS_ESA_Baja"
            
            .Add "HPS_ESA_PendienteRenovacion", "HPS_ESA_PendienteRenovacion"
            .Add "HPS_ESA_Solicitado", "HPS_ESA_Solicitado"
            .Add "HPS_ESA_MESES_PARA_RENOVAR", "HPS_ESA_MESES_PARA_RENOVAR"
            
            .Add "HPS_UE_SIN_DATOS", "HPS_UE_SIN_DATOS"
            .Add "HPS_UE_Activo", "HPS_UE_Activo"
            .Add "HPS_UE_ApuntoDeCaducar", "HPS_UE_ApuntoDeCaducar"
            .Add "HPS_UE_Caducado", "HPS_UE_Caducado"
            .Add "HPS_UE_Baja", "HPS_UE_Baja"
            
            .Add "HPS_UE_PendienteRenovacion", "HPS_UE_PendienteRenovacion"
            .Add "HPS_UE_Solicitado", "HPS_UE_Solicitado"
            .Add "HPS_UE_MESES_PARA_RENOVAR", "HPS_UE_MESES_PARA_RENOVAR"
        End With
    End If
    If p_URLDirectorio = "" Then
        m_URLCarpeta = m_ObjEntorno.URLDirectorioLocal
    Else
        m_URLCarpeta = p_URLDirectorio
    End If
    If Right(m_URLCarpeta, 1) = "\" Then
        m_URLCarpeta = Left(m_URLCarpeta, Len(m_URLCarpeta) - 1)
    End If
    If p_NombreExcel = "" Then
        m_NombreExcel = "ConsultaUsuario"
    Else
        m_NombreExcel = p_NombreExcel
    End If
    m_NombreExcel = fso.GetBaseName(m_NombreExcel)
    
    m_URLExcel = m_URLCarpeta & "\" & m_NombreExcel & ".xlsx"
    If fso.FileExists(m_URLExcel) Then
        If FicheroAbierto(m_URLExcel) Then
            p_Error = "Cierre la consulta anterior"
            Err.Raise 1000
        End If
        fso.DeleteFile m_URLExcel, True
    End If
    
    m_SQL = p_form.RecordSource
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    If rcdDatos.EOF Then
        rcdDatos.Close
        Set rcdDatos = Nothing
        p_Error = "No hay registros"
        Err.Raise 1000
    End If
    Set AppExcel = CreateObject("Excel.Application")
    AppExcel.Visible = False
    Set WbLibro = AppExcel.Workbooks.Add
    WbLibro.SaveAs m_URLExcel
    Set WbHoja = WbLibro.Worksheets(1)
    With WbHoja
        intFila = 1
        rcdDatos.MoveFirst
        Columna = 1
        For Each ctl In p_form.Controls
            If ctl.ControlType = 109 Then
                If ctl.ColumnHidden = False Then
                    .Cells(intFila, Columna).value = ctl.Name
                    .Range(.Cells(intFila, Columna), .Cells(intFila, Columna)).Font.Bold = True
                    Columna = Columna + 1
                End If
            End If
        Next
'        'aquí van las observaciones
'        .Cells(intFila, columna).value = "OBSERVACIONES"
        .Range(.Cells(intFila, Columna), .Cells(intFila, Columna)).Font.Bold = True
        intFila = intFila + 1
        Do While Not rcdDatos.EOF
            Columna = 1
            For Each ctl In p_form.Controls
                m_NombreCampo = ctl.Name
                If Not m_ObjColCamposCalculados Is Nothing Then
                    If m_ObjColCamposCalculados.Exists(m_NombreCampo) Then
                        GoTo siguienteColumna
                    End If
                End If
                If ctl.ControlType = 109 Then
                    If ctl.ColumnHidden = False Then
                        'If ctl.Name = "HSP_NAC_" Then Stop
                        'Debug.Print ctl.Name
                        m_Valor = Nz(rcdDatos.Fields(m_NombreCampo), "")
                        If ctl.Name = "IDExpediente" Then
                            m_Valor = "" & m_Valor
                        End If
                        If IsDate(m_Valor) Then
                            .Range(.Cells(intFila, Columna), .Cells(intFila, Columna)).NumberFormat = "dd/mm/yyyy"
                            .Cells(intFila, Columna).value = CDate(m_Valor)
                        Else
                            .Cells(intFila, Columna).value = m_Valor
                        End If
                        
                        Columna = Columna + 1
                    End If
                End If
siguienteColumna:
            Next
            'aquí van las observaciones
            m_ID = Nz(rcdDatos.Fields("ID"), "")
            'If m_ID = "243" Then Stop
            m_CadenaObservaciones = getListaObservacionesParaExcel(m_ID, p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If m_CadenaObservaciones <> "" Then
                .Cells(intFila, Columna).value = m_CadenaObservaciones
            End If
            intFila = intFila + 1
            rcdDatos.MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    With WbHoja.Range("A1").CurrentRegion
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    WbHoja.Cells.EntireColumn.AutoFit
    WbLibro.Close True
    Set WbLibro = Nothing
    AppExcel.Quit
    Set AppExcel = Nothing
    
    
    GenerarConsultaConObservaciones = m_URLExcel
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GenerarConsultaConObservaciones ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not WbLibro Is Nothing Then
        WbLibro.Close False
        Set WbLibro = Nothing
    End If
    If Not AppExcel Is Nothing Then
        AppExcel.Quit
        Set AppExcel = Nothing
    End If
    
End Function


Public Function FormularioAbierto(Nombre As String) As Boolean
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día X
    '   -Modificaciones:
    
    '   -Funcionamiento:
    
    '   -Llamada desde
    '   -Devuelve
    '       FormularioAbierto = true or false
    '       FormularioAbierto =False-->Error
    '-------------------------------------------------------------------
    Dim strEstado As String
    On Error GoTo errores
    strEstado = SysCmd(acSysCmdGetObjectState, acForm, Nombre)
    If strEstado = "0" Then
        FormularioAbierto = False
    Else
        FormularioAbierto = True
    End If
    Exit Function
errores:
    FormularioAbierto = False
End Function
Public Function Seleccionar( _
                                p_EsArchivo As Boolean, _
                                p_Titulo As String, _
                                ByRef p_URLArchivo As String, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    Dim fDialog As Object
    Dim varFile As Variant
    
    On Error GoTo errores
    If p_EsArchivo = True Then
        Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    Else
        Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
    End If
    With fDialog
        .Show
        If p_EsArchivo Then
            .AllowMultiSelect = True
            .Title = p_Titulo
            .Filters.Clear
            .Filters.Add "All Files", "*.*"
        End If
        For Each varFile In .SelectedItems
            If p_URLArchivo = "" Then
                p_URLArchivo = varFile
            Else
                p_URLArchivo = p_URLArchivo & ";" & varFile
            End If
            
        Next
    End With
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Seleccionar ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function
Public Function Seleccionar1( _
                                p_EsArchivo As Boolean, _
                                Optional p_Titulo As String, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    Dim m_ObjfDialog As Object
    Dim varFile As Variant
    
    On Error GoTo errores
    
    If p_Titulo = "" Then
        p_Titulo = "Seleccione el archivo"
    End If
    If p_EsArchivo = True Then
        Set m_ObjfDialog = Application.FileDialog(msoFileDialogFilePicker)
    Else
        Set m_ObjfDialog = Application.FileDialog(msoFileDialogFolderPicker)
    End If
    With m_ObjfDialog
        .Show
        If p_EsArchivo Then
            .AllowMultiSelect = False
            .InitialFileName = m_ObjEntorno.URLArchivoUltimo
            .Title = p_Titulo
            .Filters.Clear
            .Filters.Add "All Files", "*.*"
        End If
        For Each varFile In .SelectedItems
            Seleccionar1 = CStr(varFile)
        Next
    End With
    If p_EsArchivo Then
        m_ObjEntorno.URLArchivoUltimo = CStr(varFile)
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Seleccionar1 ha producido el error : " & vbNewLine & Err.Description
    End If
    
End Function

Public Sub AjustarTamaño(frmFormulario As Form)
    Dim i As Long
    On Error Resume Next
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    With frmFormulario
        ' ajusto el ancho del formulario teniendo en cuenta si tiene o no selector de registros
        If Not .RecordSelectors Then
            .InsideWidth = frmFormulario.Width
        Else
            .InsideWidth = frmFormulario.Width + 250
        End If
        ' si se abre en vista formulario simple
        If .DefaultView = 0 Then
           ' ajusto el alto incluyendo las distintas secciones, encabezado, pie, grupos...
           ' como no sé el número de secciones del formulario, me salgo al producirse un error
           .InsideHeight = 0
           For i = 0 To 100
              .InsideHeight = .InsideHeight + .Section(i).Height
           Next
        End If
    End With
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
End Sub
Function FicheroAbierto(strURLArchivo As String) As Boolean
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día X
    '   -Modificaciones:
    '   -Funcionamiento:
    
    '   -Llamada desde:
    
    '   -Devuelve:
    '       FicheroAbierto = True o False
    '       FicheroAbierto = True -->Error
    '-------------------------------------------------------------------
    Dim infNumeroFichero As Integer, intNumeroError As Integer
    On Error Resume Next
    infNumeroFichero = FreeFile()
    Open strURLArchivo For Input Lock Read As #infNumeroFichero
    Close infNumeroFichero          ' Close the file.
    intNumeroError = Err.Number
    On Error GoTo 0        ' Turn error checking back on.
    Select Case intNumeroError
        Case 0
         FicheroAbierto = False
    
        ' Error number for "Permission Denied."
        ' File is already opened by another user.
        Case Else
            FicheroAbierto = True
    End Select
End Function





Public Function EstablecerControlesformlateral( _
                                                p_frm As Form, _
                                                Optional ByRef p_Error As String _
                                                ) As String

    Dim ctl As Control
    Dim m_Tag As String
    Dim m_ParteSeleccionado As String
    On Error GoTo errores
    'Debug.Print p_frm.Name
    p_frm.Detalle.BackColor = m_ObjEntorno.CabeceraBackColor
    For Each ctl In p_frm.Controls
        If TypeOf ctl Is Label Then
            'Debug.Print ctl.Name
            ctl.BorderColor = m_ObjEntorno.EtiquetaBorderColor
            m_Tag = Nz(ctl.Tag, "")
            If InStr(1, m_Tag, ";") = 0 Then
                ctl.BackStyle = 0
                ctl.ForeColor = m_ObjEntorno.EtiquetaLinkForeColor
            Else
                dato = Split(m_Tag, ";")
                m_ParteSeleccionado = dato(1)
                If m_ParteSeleccionado = "NoSeleccionado" Then
                    ctl.BackStyle = 0
                    ctl.ForeColor = m_ObjEntorno.EtiquetaLinkForeColor
                Else
                    ctl.BackStyle = 1
                    ctl.BackColor = m_ObjEntorno.EtiquetaLinkBackColorSeleccionado
                    ctl.ForeColor = m_ObjEntorno.EtiquetaForeColorSeleccionado
                End If
                
            End If
        End If
    Next
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerControlesformlateral ha producido el error: " & Err.Description
    End If
End Function

Public Function EstablecerControlesDatos( _
                                            p_frm As Form, _
                                            Optional ByRef p_Error As String _
                                            ) As String

    Dim ctl As Control
    Dim m_Tag As String
    Dim m_ParteSeleccionado As String
    On Error GoTo errores
    For Each ctl In p_frm.Controls
        If TypeOf ctl Is Label Then
            If Nz(ctl.Tag, "") = "NOCAMBIABLE" Then
                GoTo siguienteControl
            End If
            ctl.BorderColor = m_ObjEntorno.EtiquetaBorderColor
            If InStr(1, ctl.Name, "titulo") <> 0 Then
                ctl.BackStyle = 0
                ctl.ForeColor = m_ObjEntorno.EtiquetaLinkForeColor
                If InStr(1, ctl.Caption, "Consulta") <> 0 Then
                    ctl.Width = 3059
                Else
                    ctl.Width = 10258
                End If
                ctl.Left = (p_frm.InsideWidth / 2) - (ctl.Width / 2)
            Else
                ctl.BackStyle = 1
                ctl.ForeColor = m_ObjEntorno.CabeceraBackColor
                ctl.BackColor = m_ObjEntorno.EtiquetaBackColor
            End If
            
            
        End If
        If TypeOf ctl Is TextBox Or TypeOf ctl Is ComboBox Then
            ctl.BorderColor = m_ObjEntorno.CampoBorderColor
            ctl.ForeColor = m_ObjEntorno.CampoForeColor
            ctl.BackColor = m_ObjEntorno.CampoBackColor
        End If
        If TypeOf ctl Is ListBox Then
            ctl.BorderColor = m_ObjEntorno.ListaBorderColor
            ctl.ForeColor = m_ObjEntorno.ListaForeColor
            ctl.BackColor = m_ObjEntorno.ListaBackColor
        End If
siguienteControl:
    Next
    On Error Resume Next
    p_frm.titulo1.ForeColor = m_ObjEntorno.EtiquetaLinkForeColor
    p_frm.Titulo.ForeColor = m_ObjEntorno.EtiquetaLinkBackColorSeleccionado
    p_frm.CuadroGeneral.BackColor = m_ObjEntorno.ListaBackColor
    p_frm.EncabezadoDelFormulario.BackColor = m_ObjEntorno.CabeceraBackColor
    p_frm.PieDelFormulario.BackColor = m_ObjEntorno.CabeceraBackColor
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerControlesformlateral ha producido el error: " & Err.Description
    End If
End Function

Public Function getColeDeCommand(Optional ByRef p_Error As String) As Scripting.Dictionary

    Dim m_Command As String
    Dim m_Campo As String
    Dim m_Valor As String
    Dim m_VarItem As Variant
    Dim dato1 As Variant
    On Error GoTo errores
    
    'Campo1=Valor1|Campo2=Valor2|......|CampoN=ValorN
    m_Command = Application.TempVars("Command")
    If m_Command = "" Then
        Exit Function
    End If
    dato = Split(m_Command, "|")
    For Each m_VarItem In dato
        If InStr(1, m_VarItem, "=") <> 0 Then
            dato1 = Split(m_VarItem, "=")
            m_Campo = Trim(dato1(0))
            m_Valor = Trim(dato1(1))
            If getColeDeCommand Is Nothing Then
                Set getColeDeCommand = New Scripting.Dictionary
                getColeDeCommand.CompareMode = TextCompare
            End If
            If getColeDeCommand.Exists(m_Campo) Then
                getColeDeCommand.Remove (m_Campo)
            End If
            getColeDeCommand.Add m_Campo, m_Valor
            
        End If
    Next
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getColeDeCommand ha producido el error: " & Err.Description
    End If
End Function

Public Function RellenarComboEmpresas( _
                                        ByRef p_Cbo As ComboBox, _
                                        Optional p_SoloTramitadoras As EnumSiNo = EnumSiNo.No, _
                                        Optional p_ConCIF As EnumSiNo = EnumSiNo.Sí, _
                                        Optional ByRef p_Error As String _
                                        ) As String

    
    Dim m_Col As Scripting.Dictionary
    
    Dim m_ID As Variant
    Dim m_Suministrador As Suministrador
    
    
    On Error GoTo errores
    If p_SoloTramitadoras = Empty Then
        p_SoloTramitadoras = EnumSiNo.No
    End If
    If p_SoloTramitadoras = EnumSiNo.Sí Then
        Set m_Col = m_ObjEntorno.ColEmpresasTramitadoras
    Else
        Set m_Col = m_ObjEntorno.ColEmpresas
    End If
    p_Error = m_ObjEntorno.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    p_Cbo.RowSource = ""
    If m_Col Is Nothing Then
        Exit Function
    End If
    For Each m_ID In m_Col
        Set m_Suministrador = m_Col(m_ID)
        If p_ConCIF = EnumSiNo.Sí Then
            p_Cbo.AddItem m_Suministrador.IDSuministrador & ";" & Replace(m_Suministrador.Nemotecnico, ";", ":") & " (" & m_Suministrador.CIF & ")"
        Else
            p_Cbo.AddItem m_Suministrador.IDSuministrador & ";" & Replace(m_Suministrador.Nemotecnico, ";", ":")
        End If
        
        Set m_Suministrador = Nothing
    Next
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarComboEmpresas ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function RellenarCadenaContratistas( _
                                            ByRef p_Cbo As ComboBox, _
                                            Optional ByRef p_Error As String _
                                            ) As String

    
    Dim m_Col As Scripting.Dictionary
    Dim m_CadenaContratistas As Variant
    
    
    
    On Error GoTo errores
    Set m_Col = m_ObjEntorno.ColCadenaContratistas
    p_Error = m_ObjEntorno.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    p_Cbo.RowSource = ""
    If m_Col Is Nothing Then
        Exit Function
    End If
    For Each m_CadenaContratistas In m_Col
       
        p_Cbo.AddItem m_CadenaContratistas
        
        
    Next
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarCadenaContratistas ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function RellenarComboMotivos( _
                                    ByRef p_Cbo As ComboBox, _
                                    Optional ByRef p_Error As String _
                                    ) As String

    Dim m_Col As Scripting.Dictionary
    Dim m_Motivo As Variant
    On Error GoTo errores
    
    p_Cbo.RowSource = ""
    Set m_Col = m_ObjEntorno.Motivos
    If m_Col Is Nothing Then
        Exit Function
    End If
   
    For Each m_Motivo In m_Col
        p_Cbo.AddItem m_Motivo
    Next
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarComboMotivos ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function RellenarComboGrado( _
                                    p_Tipo As String, _
                                    ByRef p_Cbo As ComboBox, _
                                    Optional ByRef p_Error As String _
                                    ) As String

    Dim m_ObjColGrados As Scripting.Dictionary
    Dim m_Grado As Variant
    On Error GoTo errores
    
    p_Cbo.RowSource = ""
    If Not m_ObjEntorno.colGradosHPS.Exists(p_Tipo) Then
        p_Error = "No hay ningún listado de grados para el tipo " & p_Tipo
        Err.Raise 1000
    End If
    Set m_ObjColGrados = m_ObjEntorno.colGradosHPS(p_Tipo)
    If Not m_ObjColGrados Is Nothing Then
        For Each m_Grado In m_ObjColGrados
            p_Cbo.AddItem m_Grado
        Next
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarComboGrado ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function RellenarComboEstado( _
                                    ByRef p_Cbo As ComboBox, _
                                    Optional ByRef p_Error As String _
                                    ) As String

    
    Dim m_Estado As Variant
    On Error GoTo errores
    
    p_Cbo.RowSource = ""
    
    
    For Each m_Estado In m_ObjEntorno.ColEstadosVisiblesHPS.Keys
        p_Cbo.AddItem m_ObjEntorno.ColEstadosVisiblesHPS(m_Estado)
    Next
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarComboEstado ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function RellenarComboExpediente( _
                                        ByRef p_Cbo As ComboBox, _
                                        Optional ByRef p_Error As String _
                                        ) As String

    
    Dim m_ObjCol As Collection
    Dim m_ID  As Variant
    Dim m_Expediente As Expediente
    
    On Error GoTo errores
    
    p_Cbo.RowSource = ""
    Set m_ObjCol = m_ObjEntorno.ExpedientesUsados
    If m_ObjCol Is Nothing Then
        Exit Function
    End If
    For Each m_ID In m_ObjCol
        Set m_Expediente = m_ObjCol(m_ID)
        p_Cbo.AddItem m_Expediente.IDExpediente & ";" & m_Expediente.TextoExpediente
        Set m_Expediente = Nothing
    Next
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarComboExpediente ha devuelto el error: " & Err.Description
    End If
                                                
End Function
Public Function RellenarCombo( _
                                    p_NombreCampo As String, _
                                    ByRef p_Cbo As ComboBox, _
                                    Optional p_NombreTabla As String, _
                                    Optional ByRef p_Error As String _
                                    ) As String
                                                    
    
    Dim m_ObjCol As Collection
    Dim m_VarItem As Variant
    On Error GoTo errores
    
    
    Set m_ObjCol = getListaParaCombo(p_NombreCampo, p_NombreTabla, p_Error)
    
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    p_Cbo.RowSource = ""
    If Not m_ObjCol Is Nothing Then
        For Each m_VarItem In m_ObjCol
            p_Cbo.AddItem m_VarItem
        Next
    End If
    Set p_Cbo = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarCombo ha devuelto el error: " & Err.Description
    End If
                                                
End Function

Public Function GradoExiste( _
                                p_Grado As String, _
                                p_TipoHPS As String, _
                                Optional ByRef p_Error As String _
                                ) As EnumSiNo

    
    Dim m_ColGrados As Scripting.Dictionary
    Dim m_Grado As Variant
    On Error GoTo errores
    
    If p_Grado = "" Then
        GradoExiste = EnumSiNo.No
        Exit Function
    End If
    If p_TipoHPS = "" Then
        p_Error = "No se ha indicado el Tipo HPS"
        Err.Raise 1000
    End If
    If m_ObjEntorno.colGradosHPS Is Nothing Then
        p_Error = "No se ha inicializado la variable de entorno Entorno.ColEstadosVisiblesHPS"
        Err.Raise 1000
    End If
    If m_ObjEntorno.colGradosHPS.Exists(p_TipoHPS) Then
        Set m_ColGrados = m_ObjEntorno.colGradosHPS(p_TipoHPS)
        For Each m_Grado In m_ColGrados
            If CStr(m_Grado) = p_Grado Then
                GradoExiste = EnumSiNo.Sí
                Exit Function
            End If
        Next
    End If
    
    GradoExiste = EnumSiNo.No
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GradoExiste ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function EstablecerConn( _
                                ByRef p_URL As String, _
                                Optional p_Pass As String, _
                                Optional ByRef p_Error As String _
                                ) As ADODB.Connection
    
    Dim m_ConnectionString As String
    Dim m_Provider As String
    On Error GoTo errores
    
    m_Provider = "Microsoft.ACE.OLEDB.12.0"
    m_ConnectionString = "Data Source=" & p_URL & ";"
    If p_Pass <> "" Then
        m_ConnectionString = m_ConnectionString & "Jet OLEDB:Database Password=" & p_Pass & ";"
    End If
    Set EstablecerConn = New ADODB.Connection
    With EstablecerConn
        .Provider = m_Provider
        .ConnectionString = m_ConnectionString
        .Open
    End With
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerConn ha devuelto el error: " & vbNewLine & Err.Description
    End If
    
End Function

Public Function EjecutarShell( _
                                strComando As String, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    Dim wsh As Object
    Dim codigoSalida As Long
    
    On Error GoTo errores
    
    If strComando = "" Then
        p_Error = "No se ha indicado el comando"
        Err.Raise 1000
    End If
    
    ' Creamos el objeto Shell (Late Binding, no necesita referencias extra)
    Set wsh = CreateObject("WScript.Shell")
    
    ' Ejecutamos el comando.
    ' El segundo parámetro (0) oculta la ventana negra de MS-DOS.
    ' El tercer parámetro (True) hace que Access ESPERE a que termine la copia antes de seguir.
    codigoSalida = wsh.Run(strComando, 0, True)
    
    ' --- VALIDACIÓN ROBOCOPY ---
    ' Recordatorio: Robocopy usa códigos < 8 como éxito (0, 1, 2, 3...)
    If codigoSalida >= 8 Then
        p_Error = "Robocopy falló con código de salida: " & codigoSalida & vbNewLine & "Comando: " & strComando
        Err.Raise 1000
    End If
    
    EjecutarShell = "OK"
    
    ' Limpieza de memoria
    Set wsh = Nothing
    Exit Function
    
errores:
    Set wsh = Nothing
    If Err.Number <> 1000 Then
        p_Error = "El método EjecutarShell ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function
Private Function QuitarBarraFinal(ByVal ruta As String) As String
    ' Si la ruta acaba en \, se la quitamos
    If Right(ruta, 1) = "\" Then
        QuitarBarraFinal = Left(ruta, Len(ruta) - 1)
    Else
        QuitarBarraFinal = ruta
    End If
End Function
Public Function EstablecerContadores(Optional ByRef p_Error As String) As String
    
    Dim frm As Form
    Dim sFrm As Form
    Dim Irregulares As Integer
    Dim APuntoDeCaducar As Integer
    Dim Caducados As Integer
    Dim m_UsuariosHPSRequierenCurso As Integer
    Dim m_PrimeraConvocatoria As String
    Dim m_SegundaConvocatoria As String
    Dim m_CorreoJefeSeguridad As String
    Dim m_Solicitandose As Integer
    
    On Error GoTo errores
    If Not FormularioAbierto("FormInicial00Principal") Then
        Exit Function
    End If
    Set frm = Forms("FormInicial00Principal")
    If frm.SubFormLateral.SourceObject <> "FormInicial01Lateral" Then
        Exit Function
    End If
    Set sFrm = frm.SubFormLateral.Form
    
    Set m_ObjEntorno.ColUsuariosEstadoIrregular = Nothing
    If Not m_ObjEntorno.ColUsuariosEstadoIrregular Is Nothing Then
        Irregulares = m_ObjEntorno.ColUsuariosEstadoIrregular.Count
    End If
    Set m_ObjEntorno.ColUsuariosAPuntoCaducar = Nothing
    If Not m_ObjEntorno.ColUsuariosAPuntoCaducar Is Nothing Then
        APuntoDeCaducar = m_ObjEntorno.ColUsuariosAPuntoCaducar.Count
    Else
        APuntoDeCaducar = 0
    End If
    Set m_ObjEntorno.ColUsuariosCaducados = Nothing
    If Not m_ObjEntorno.ColUsuariosCaducados Is Nothing Then
        Caducados = m_ObjEntorno.ColUsuariosCaducados.Count
    Else
        Caducados = 0
    End If
    Set m_ObjEntorno.ColUsuariosHPSRequierenCurso = Nothing
    If Not m_ObjEntorno.ColUsuariosHPSRequierenCurso Is Nothing Then
        m_UsuariosHPSRequierenCurso = m_ObjEntorno.ColUsuariosHPSRequierenCurso.Count
    Else
        m_UsuariosHPSRequierenCurso = 0
    End If
    
    Set m_ObjEntorno.ColRequiereCursoFaltaPrimeraConvocatoria = Nothing
    If Not m_ObjEntorno.ColRequiereCursoFaltaPrimeraConvocatoria Is Nothing Then
        m_PrimeraConvocatoria = m_ObjEntorno.ColRequiereCursoFaltaPrimeraConvocatoria.Count
    Else
        m_PrimeraConvocatoria = 0
    End If
    
    Set m_ObjEntorno.ColRequiereCursoFaltaSegundaConvocatoria = Nothing
    If Not m_ObjEntorno.ColRequiereCursoFaltaSegundaConvocatoria Is Nothing Then
        m_SegundaConvocatoria = m_ObjEntorno.ColRequiereCursoFaltaSegundaConvocatoria.Count
    Else
        m_SegundaConvocatoria = 0
    End If
    Set m_ObjEntorno.ColRequiereCursoFaltaCorreoJefeSeguridad = Nothing
    If Not m_ObjEntorno.ColRequiereCursoFaltaCorreoJefeSeguridad Is Nothing Then
        m_CorreoJefeSeguridad = m_ObjEntorno.ColRequiereCursoFaltaCorreoJefeSeguridad.Count
    Else
        m_CorreoJefeSeguridad = 0
    End If
    Set m_ObjEntorno.ColUsuariosEnSolicitud = Nothing
    If Not m_ObjEntorno.ColUsuariosEnSolicitud Is Nothing Then
        m_Solicitandose = m_ObjEntorno.ColUsuariosEnSolicitud.Count
    Else
        m_Solicitandose = 0
    End If
    sFrm.lblHPSEstadoIrregular.Caption = "Irregular (" & Irregulares & ")"
    sFrm.lblHPSAPuntoDeCaducar.Caption = "A.P.Caducar (" & APuntoDeCaducar & ")"
    sFrm.lblHPSCaducados.Caption = "Caducados (" & Caducados & ")"
    sFrm.lblRequiereCurso.Caption = "Req. Curso ( " & m_UsuariosHPSRequierenCurso & " )"
    sFrm.lblRequiereCursoPtePrimeraConvocatoria.Caption = m_PrimeraConvocatoria
    sFrm.lblRequiereCursoPteSegundaConvocatoria.Caption = m_SegundaConvocatoria
    sFrm.lblRequiereCursoPteCorreoJefeSeguridad.Caption = m_CorreoJefeSeguridad
    
    sFrm.lblSolicitados.Caption = "Solicitados ( " & m_Solicitandose & " )"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerContadores ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getColDatosHPS(p_HPS As HPS, Optional ByRef p_Error As String) As Scripting.Dictionary
    
    Dim m_objColCampos As Collection
    Dim m_NombreCampo As Variant
    Dim m_Valor As String
    
    
    On Error GoTo errores
    p_Error = ""
    
    Set m_objColCampos = p_HPS.ColCampos
    For Each m_NombreCampo In m_objColCampos
        m_Valor = p_HPS.getPropiedad(CStr(m_NombreCampo), p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If getColDatosHPS Is Nothing Then
            Set getColDatosHPS = New Scripting.Dictionary
            getColDatosHPS.CompareMode = TextCompare
        End If
        getColDatosHPS.Add m_NombreCampo, m_Valor
    Next
    
        
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getColDatosHPS ha devuelto el error: " & Err.Description
    End If
End Function

' =============================================================================
' ParseJsonDate: converts a JSON /Date(ms)/ string to a VBA Date.
' Also handles plain date strings transparently.
' Returns 0 (1899-12-30) for empty/null/invalid input — callers must guard.
' =============================================================================
Public Function ParseJsonDate(ByVal p_Fecha As String) As Date
    Dim lPos1 As Long, lPos2 As Long
    Dim sNumero As String
    Dim vNumero As Variant
    
    On Error GoTo errores
    If Len(Trim$(p_Fecha)) = 0 Then
        Exit Function
    End If
    
    lPos1 = InStr(p_Fecha, "/Date(")
    lPos2 = InStr(p_Fecha, ")")
    If lPos1 > 0 And lPos2 > lPos1 Then
        ' JSON .NET serialized DateTime — extract Unix milliseconds
        sNumero = Mid$(p_Fecha, lPos1 + 7, lPos2 - lPos1 - 7)
        vNumero = CDec(sNumero)
        ' ms ? seconds from Unix epoch, then add to 1970-01-01
        ParseJsonDate = DateAdd("s", vNumero / 1000, #1/1/1970#)
    ElseIf IsDate(p_Fecha) Then
        ParseJsonDate = CDate(p_Fecha)
    End If
    Exit Function
errores:
    ' Invalid format — leave as 0 / 1899-12-30
End Function

Private Sub EvaluarHPSCaducidadMinima( _
                                    ByVal p_HPS As HPS, _
                                    ByVal p_TipoHPS As String, _
                                    ByRef p_FechaMinima As Date, _
                                    ByRef p_TipoMinimo As String)
    Dim m_FechaBaja As Date
    Dim m_FechaCaducidad As Date
    
    If p_HPS Is Nothing Then
        Exit Sub
    End If
    
    m_FechaBaja = ParseJsonDate(p_HPS.F_Baja)
    If m_FechaBaja <> 0 Then
        Exit Sub
    End If
    
    m_FechaCaducidad = ParseJsonDate(p_HPS.F_Caducidad)
    If m_FechaCaducidad = 0 Then
        Exit Sub
    End If
    
    If p_FechaMinima = 0 Or m_FechaCaducidad < p_FechaMinima Then
        p_FechaMinima = m_FechaCaducidad
        p_TipoMinimo = p_TipoHPS
    End If
End Sub

Private Sub SetIndicadorCaducidadPorTipo( _
                                    ByRef p_DatosLocal As DatosLocal, _
                                    ByVal p_TipoHPS As String, _
                                    ByVal p_EsAPunto As Boolean, _
                                    ByVal p_EsCaducado As Boolean)
    Select Case p_TipoHPS
        Case "NAC"
            p_DatosLocal.HPS_NAC_ApuntoDeCaducar = IIf(p_EsAPunto, "Sí", "No")
            p_DatosLocal.HPS_NAC_Caducado = IIf(p_EsCaducado, "Sí", "No")
        Case "OTAN"
            p_DatosLocal.HPS_OTAN_ApuntoDeCaducar = IIf(p_EsAPunto, "Sí", "No")
            p_DatosLocal.HPS_OTAN_Caducado = IIf(p_EsCaducado, "Sí", "No")
        Case "ESA"
            p_DatosLocal.HPS_ESA_ApuntoDeCaducar = IIf(p_EsAPunto, "Sí", "No")
            p_DatosLocal.HPS_ESA_Caducado = IIf(p_EsCaducado, "Sí", "No")
        Case "UE"
            p_DatosLocal.HPS_UE_ApuntoDeCaducar = IIf(p_EsAPunto, "Sí", "No")
            p_DatosLocal.HPS_UE_Caducado = IIf(p_EsCaducado, "Sí", "No")
    End Select
End Sub

Private Sub AplicarCaducidadGlobalEnDatosLocal( _
                                    ByRef p_DatosLocal As DatosLocal, _
                                    ByVal p_UsuarioHPS As UsuarioHPS)
    Dim m_FechaMinima As Date
    Dim m_TipoMinimo As String
    Dim m_HoyMasMesesRenovacion As Date
    
    p_DatosLocal.HPS_NAC_ApuntoDeCaducar = "No"
    p_DatosLocal.HPS_OTAN_ApuntoDeCaducar = "No"
    p_DatosLocal.HPS_ESA_ApuntoDeCaducar = "No"
    p_DatosLocal.HPS_UE_ApuntoDeCaducar = "No"
    p_DatosLocal.HPS_NAC_Caducado = "No"
    p_DatosLocal.HPS_OTAN_Caducado = "No"
    p_DatosLocal.HPS_ESA_Caducado = "No"
    p_DatosLocal.HPS_UE_Caducado = "No"
    
    If p_UsuarioHPS Is Nothing Then
        Exit Sub
    End If
    If p_UsuarioHPS.F_BajaCalculada <> "" Then
        Exit Sub
    End If
    
    EvaluarHPSCaducidadMinima p_UsuarioHPS.HPSNAC, "NAC", m_FechaMinima, m_TipoMinimo
    EvaluarHPSCaducidadMinima p_UsuarioHPS.HPSOTAN, "OTAN", m_FechaMinima, m_TipoMinimo
    EvaluarHPSCaducidadMinima p_UsuarioHPS.HPSESA, "ESA", m_FechaMinima, m_TipoMinimo
    EvaluarHPSCaducidadMinima p_UsuarioHPS.HPSUE, "UE", m_FechaMinima, m_TipoMinimo
    
    If m_FechaMinima = 0 Then
        Exit Sub
    End If
    
    If m_FechaMinima < Date Then
        SetIndicadorCaducidadPorTipo p_DatosLocal, m_TipoMinimo, False, True
        Exit Sub
    End If
    
    m_HoyMasMesesRenovacion = DateAdd("m", CInt(Application.TempVars("MesesParaRenovacion")), Date)
    If m_FechaMinima <= m_HoyMasMesesRenovacion Then
        SetIndicadorCaducidadPorTipo p_DatosLocal, m_TipoMinimo, True, False
    End If
End Sub

Public Function DatosHPSCalculados( _
                                    p_TipoDato As String, _
                                    p_ObjHPS As HPS, _
                                    Optional ByRef p_Error As String _
                                    ) As String
    
    Dim m_APuntoDeCaducar As String
    Dim m_MarcadoNoRenovacion As String
    Dim m_Solicitada As String
    Dim m_Caducado As String
    Dim m_Baja As String
    
    Dim m_Activa As String
    Dim m_MOTIVO_IRREGULAR As String
    Dim m_Irregular As String
    Dim m_FechaCaducidad As String
    Dim m_HoyMasNueveMeses As String
    
    Dim m_GradoExiste As EnumSiNo
    Dim m_FechaConcesion As Date
    Dim m_FechaBaja As Date
    Dim m_FechaSolicitud As Date
    Dim m_FechaCaduc As Date
    On Error GoTo errores
    p_Error = ""
    If p_ObjHPS Is Nothing Then
        p_Error = "Sin datos"
        Err.Raise 1000
    End If
    
    ' Pre-parse the three date fields so every branch can use them
    m_FechaBaja = ParseJsonDate(p_ObjHPS.F_Baja)
    m_FechaConcesion = ParseJsonDate(p_ObjHPS.F_Concesion)
    m_FechaSolicitud = ParseJsonDate(p_ObjHPS.F_Solicitud)
    m_FechaCaduc = ParseJsonDate(p_ObjHPS.F_Caducidad)
    
    If p_TipoDato = "Activo" Then
        If m_FechaBaja <> 0 Then
            DatosHPSCalculados = "No"
            Exit Function
        End If
        
        If m_FechaConcesion = 0 Or m_FechaCaduc = 0 Then
            DatosHPSCalculados = "No"
        Else
            If m_FechaCaduc >= Date Then
                DatosHPSCalculados = "Sí"
            Else
                DatosHPSCalculados = "No"
            End If
        End If
    ElseIf p_TipoDato = "Baja" Then
        If m_FechaBaja <> 0 Then
            DatosHPSCalculados = "Sí"
        Else
            DatosHPSCalculados = "No"
        End If
    
    ElseIf p_TipoDato = "Solicitada" Then
        
        If m_FechaBaja <> 0 Then
            DatosHPSCalculados = "No"
            Exit Function
        End If
        
        If m_FechaSolicitud <> 0 Then
            DatosHPSCalculados = "Sí"
        Else
            DatosHPSCalculados = "No"
        End If
    ElseIf p_TipoDato = "APuntoDeCaducar" Then
        
        If m_FechaBaja <> 0 Then
            DatosHPSCalculados = "No"
            Exit Function
        End If
        
 '        m_MarcadoNoRenovacion = DatosHPSCalculados("MarcadoNoRenovacion", p_ObjHPS, p_Error)
 '        If p_Error <> "" Then
 '            Err.Raise 1000
 '        End If
 '        If m_MarcadoNoRenovacion = "Sí" Then
 '            DatosHPSCalculados = "No"
 '            Exit Function
 '        End If
        If m_FechaCaduc <> 0 Then
            If Date <= m_FechaCaduc Then
                m_HoyMasNueveMeses = DateAdd("m", CInt(Application.TempVars("MesesParaRenovacion")), Date)
                If m_FechaCaduc <= m_HoyMasNueveMeses Then
                    DatosHPSCalculados = "Sí"
                Else
                    DatosHPSCalculados = "No"
                End If
            Else
                DatosHPSCalculados = "No"
            End If
        End If
        
        
    ElseIf p_TipoDato = "Caducado" Then
        
        If m_FechaBaja <> 0 Then
            DatosHPSCalculados = "No"
            Exit Function
        End If
        
 '        m_MarcadoNoRenovacion = DatosHPSCalculados("MarcadoNoRenovacion", p_ObjHPS, p_Error)
 '        If p_Error <> "" Then
 '            Err.Raise 1000
 '        End If
 '        If m_MarcadoNoRenovacion = "Sí" Then
 '            DatosHPSCalculados = "No"
 '            Exit Function
 '        End If
 '
        If m_FechaConcesion <> 0 And m_FechaCaduc <> 0 Then
            If m_FechaCaduc < Date Then
                DatosHPSCalculados = "Sí"
            Else
                DatosHPSCalculados = "No"
            End If
        End If
    ElseIf p_TipoDato = "PendienteRenovacion" Then
        
        If m_FechaBaja <> 0 Then
            DatosHPSCalculados = "No"
            Exit Function
        End If
        
        m_MarcadoNoRenovacion = DatosHPSCalculados("MarcadoNoRenovacion", p_ObjHPS, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If m_MarcadoNoRenovacion = "Sí" Then
            DatosHPSCalculados = "No"
            Exit Function
        End If
        m_Solicitada = DatosHPSCalculados("Solicitada", p_ObjHPS, p_Error)
        If m_Solicitada = "Sí" Then
            If m_FechaCaduc <> 0 Then
                DatosHPSCalculados = "Sí"
            Else
                DatosHPSCalculados = "No"
            End If
        Else
            DatosHPSCalculados = "No"
        End If
        
        
                
        
    ElseIf p_TipoDato = "Estado" Then
        
        If m_FechaBaja <> 0 Then
            DatosHPSCalculados = m_ObjEntorno.ColEstadosHPS(CStr(EnumEstadoHPS.Baja))
            Exit Function
        End If
        m_Irregular = DatosHPSCalculados("MOTIVO_IRREGULAR", p_ObjHPS, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If m_Irregular = "Sí" Then
            DatosHPSCalculados = "Irregular"
            Exit Function
        End If
        m_APuntoDeCaducar = DatosHPSCalculados("APuntoDeCaducar", p_ObjHPS, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        m_MarcadoNoRenovacion = DatosHPSCalculados("MarcadoNoRenovacion", p_ObjHPS, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        
        m_Solicitada = DatosHPSCalculados("Solicitada", p_ObjHPS, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        m_Caducado = DatosHPSCalculados("Caducado", p_ObjHPS, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        m_Activa = DatosHPSCalculados("Activo", p_ObjHPS, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        m_Baja = DatosHPSCalculados("Baja", p_ObjHPS, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        
        If m_Baja = "Sí" Then
            DatosHPSCalculados = m_ObjEntorno.ColEstadosHPS(CStr(EnumEstadoHPS.Baja))
            Exit Function
        End If
        
        
        If m_APuntoDeCaducar = "Sí" And m_MarcadoNoRenovacion = "Sí" Then
            DatosHPSCalculados = m_ObjEntorno.ColEstadosHPS(CStr(EnumEstadoHPS.APuntoCaducarRenovacionNo))
            Exit Function
        End If
        If m_APuntoDeCaducar = "Sí" And m_MarcadoNoRenovacion = "No" And m_Solicitada = "No" Then
            DatosHPSCalculados = m_ObjEntorno.ColEstadosHPS(CStr(EnumEstadoHPS.APuntoCaducarRenovacionSiSinSolicitud))
            Exit Function
        End If
        If m_APuntoDeCaducar = "Sí" And m_MarcadoNoRenovacion = "No" And m_Solicitada = "Sí" Then
            DatosHPSCalculados = m_ObjEntorno.ColEstadosHPS(CStr(EnumEstadoHPS.APuntoCaducarRenovacionSiConSolicitud))
            Exit Function
        End If
        If m_Caducado = "Sí" And m_MarcadoNoRenovacion = "Sí" Then
            DatosHPSCalculados = m_ObjEntorno.ColEstadosHPS(CStr(EnumEstadoHPS.CaducadaRenovacionNo))
            Exit Function
        End If
        If m_Caducado = "Sí" And m_MarcadoNoRenovacion = "No" And m_Solicitada = "No" Then
            DatosHPSCalculados = m_ObjEntorno.ColEstadosHPS(CStr(EnumEstadoHPS.CaducadaRenovacionSiSinSolicitud))
            Exit Function
        End If
        If m_Caducado = "Sí" And m_MarcadoNoRenovacion = "No" And m_Solicitada = "Sí" Then
            DatosHPSCalculados = m_ObjEntorno.ColEstadosHPS(CStr(EnumEstadoHPS.CaducadaRenovacionSiConSolicitud))
            Exit Function
        End If
        If m_Activa = "Sí" And m_Solicitada = "Sí" Then
            DatosHPSCalculados = m_ObjEntorno.ColEstadosHPS(CStr(EnumEstadoHPS.ActivoSolicitada))
            Exit Function
        End If
        If m_Activa = "Sí" And m_Solicitada = "No" Then
            DatosHPSCalculados = m_ObjEntorno.ColEstadosHPS(CStr(EnumEstadoHPS.ActivoSinSolicitar))
            Exit Function
        End If
        If m_Activa = "Sí" And m_MarcadoNoRenovacion = "Sí" Then
            DatosHPSCalculados = m_ObjEntorno.ColEstadosHPS(CStr(EnumEstadoHPS.ActivoRenovacionNo))
            Exit Function
        End If
        If m_Activa = "No" And m_MarcadoNoRenovacion = "Sí" Then
            DatosHPSCalculados = m_ObjEntorno.ColEstadosHPS(CStr(EnumEstadoHPS.CaducadaRenovacionNo))
            Exit Function
        End If
        
        If m_Solicitada = "Sí" And m_Activa = "No" Then
            DatosHPSCalculados = m_ObjEntorno.ColEstadosHPS(CStr(EnumEstadoHPS.SolicitudNueva))
            Exit Function
        End If
        If m_MarcadoNoRenovacion = "No" Then
            DatosHPSCalculados = m_ObjEntorno.ColEstadosHPS(CStr(EnumEstadoHPS.PendienteRenovacion))
            Exit Function
        End If
    ElseIf p_TipoDato = "Irregular" Then
        m_MOTIVO_IRREGULAR = DatosHPSCalculados("MOTIVO_IRREGULAR", p_ObjHPS, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If m_MOTIVO_IRREGULAR <> "" Then
            DatosHPSCalculados = "Sí"
        Else
            DatosHPSCalculados = "No"
        End If
    ElseIf p_TipoDato = "MOTIVO_IRREGULAR" Then
        
        If p_ObjHPS.F_Concesion = "" And p_ObjHPS.F_Caducidad = "" And p_ObjHPS.Grado = "" And p_ObjHPS.F_Solicitud = "" And p_ObjHPS.Renovacion = "" And p_ObjHPS.F_Baja = "" Then
            p_Error = "Se necesita al menos uno de los siguientes tados F_Concesion,F_Caducidad,Grado,F_Solicitud,Renovacion,F_Baja"
            Err.Raise 1000
        End If
        
        
        m_Solicitada = DatosHPSCalculados("Solicitada", p_ObjHPS, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        m_MarcadoNoRenovacion = DatosHPSCalculados("MarcadoNoRenovacion", p_ObjHPS, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        m_Baja = DatosHPSCalculados("Baja", p_ObjHPS, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        
        If m_MarcadoNoRenovacion = "Sí" Then
           If m_Solicitada = "Sí" Then
                DatosHPSCalculados = "HPS Solicitada con Renovación No"
                Exit Function
            End If
        End If
        If m_FechaConcesion <> 0 Or m_FechaCaduc <> 0 Then
           If m_FechaCaduc = 0 And m_FechaConcesion <> 0 Then
                DatosHPSCalculados = "HPS Con Fecha de Caducidad y no de Concesión"
                Exit Function
            
                
            End If
            If m_FechaCaduc <> 0 And m_FechaConcesion <> 0 Then
                If m_FechaConcesion > m_FechaCaduc Then
                    DatosHPSCalculados = "HPS Con Fecha de Caducidad anterior a la de Concesión"
                    Exit Function
                End If
                
            
                
            End If
            If m_FechaConcesion = 0 And m_FechaCaduc <> 0 Then
                DatosHPSCalculados = "HPS Con Fecha de Concesión y no de Caducidad"
                Exit Function
            End If
            m_GradoExiste = GradoExiste(p_ObjHPS.Grado, p_ObjHPS.tipoHps, p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If m_GradoExiste = EnumSiNo.No Then
                DatosHPSCalculados = "HPS Con Fecha de Concesión y de Caducidad sin Grado reconocido"
                Exit Function
            End If
            If m_FechaSolicitud <> 0 Then
               If m_FechaSolicitud < m_FechaConcesion Then
                    DatosHPSCalculados = "HPS Activo con Fecha de solicitud anterior a la de concesión"
                    Exit Function
                End If
            End If
        End If
        If m_Solicitada = "Sí" Then
            If m_Baja = "Sí" Then
                DatosHPSCalculados = "HPS Con Fecha de Solicitud y con fecha de baja"
                Exit Function
            End If
            
            m_GradoExiste = GradoExiste(p_ObjHPS.Grado, p_ObjHPS.tipoHps, p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If m_GradoExiste = EnumSiNo.No Then
                DatosHPSCalculados = "HPS Con Fecha de Solicitud sin Grado reconocido"
                Exit Function
            End If
        End If
        
    ElseIf p_TipoDato = "MESES_PARA_RENOVAR" Then
        
        If m_FechaCaduc <> 0 Then
            DatosHPSCalculados = DateDiff("m", Date, m_FechaCaduc)
        End If
    ElseIf p_TipoDato = "MarcadoNoRenovacion" Then
        
        If p_ObjHPS.Renovacion = "No" Then
            DatosHPSCalculados = "Sí"
        Else
            DatosHPSCalculados = "No"
        End If
        
    Else
        p_Error = "Tipo no reconocido"
        Err.Raise 1000
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método DatosHPSCalculados ha devuelto el error: " & Err.Description
    End If
    
End Function
Public Function EstaEnCole(ByRef p_Col As Collection, p_Elemento As Variant, Optional ByRef p_Error As String) As Boolean

    Dim m_Elemento As Variant
    On Error GoTo errores
    If p_Col Is Nothing Then
        Exit Function
    End If
    For Each m_Elemento In p_Col
        If CStr(m_Elemento) = p_Elemento Then
            EstaEnCole = True
            Exit Function
        End If
    Next
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstaEnCole ha devuelto el error: " & Err.Description
    End If
End Function
Public Function ActualizarListaDatosLocal(Optional ByRef p_Error As String) As String
    
    Dim frm As Form
    Dim m_ID As String
    
    On Error GoTo errores
    If FormularioAbierto("FormInicial00Principal") Then
        Set frm = Forms("FormInicial00Principal")
        If frm.Controls("SubFormCentral").SourceObject = "FormInicial02ConsultasPrincipal" Then
            Set frm = frm.Controls("SubFormCentral").Form
            If frm.Controls("SubFormResultados").SourceObject = "FormInicial03ConsultasDatos" Then
                Set frm = frm.Controls("SubFormResultados").Form
                If Not m_ObjUsuarioActivo Is Nothing Then
                    m_ID = m_ObjUsuarioActivo.ID
                End If
                frm.Requery
                If m_ID <> "" Then
                    FiltrarUsuarioActivo m_ID, p_Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                End If
            End If
        End If
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ActualizarListaDatosLocal ha devuelto el error: " & Err.Description
    End If
End Function

Public Function EsLetra(p_Letra As String, Optional ByRef p_Error As String) As Boolean

    On Error GoTo errores
    If Len(p_Letra) <> 1 Then
        p_Error = "Sólo se admite un caracter"
        Err.Raise 1000
    End If
    If p_Letra = "a" Or p_Letra = "A" Or p_Letra = "á" Or p_Letra = "Á" Or _
        p_Letra = "e" Or p_Letra = "E" Or p_Letra = "é" Or p_Letra = "É" Or _
        p_Letra = "i" Or p_Letra = "I" Or p_Letra = "í" Or p_Letra = "Í" Or _
        p_Letra = "o" Or p_Letra = "O" Or p_Letra = "ó" Or p_Letra = "Ó" Or _
        p_Letra = "u" Or p_Letra = "U" Or p_Letra = "ú" Or p_Letra = "Ú" Then
        EsLetra = True
    End If
        
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EsLetra ha devuelto el error: " & Err.Description
    End If
End Function

Public Function ParsearDatoParaConsulta(p_DatoOriginal As String, Optional ByRef p_Error As String) As String
    Dim i As Integer
    Dim m_Letra As String
    
    On Error GoTo errores
    
    If p_DatoOriginal = "" Then
        Err.Raise 1000
    End If
    For i = 1 To Len(p_DatoOriginal)
        m_Letra = Mid(p_DatoOriginal, i, 1)
        If EsLetra(m_Letra) Then
            m_Letra = "?"
        End If
        If ParsearDatoParaConsulta = "" Then
            ParsearDatoParaConsulta = m_Letra
        Else
            ParsearDatoParaConsulta = ParsearDatoParaConsulta & m_Letra
        End If
    Next
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ParsearDatoParaConsulta ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function



Public Function CambiarEstadoAFormulariosAbiertos( _
                                                    Optional ByRef p_Error As String _
                                                    ) As String
                                        
    
    On Error GoTo errores
    
    If Application.TempVars("SoloLectura") = "Sí" Then
        Application.TempVars("SoloLectura") = "No"
    Else
        Application.TempVars("SoloLectura") = "Sí"
    End If
    If Application.TempVars("SoloLectura") = "Sí" Then
        m_EstadoConexion = "Solo lectura"
    Else
        m_EstadoConexion = "Lectura/Escritura"
    End If
    
    
    If FormularioAbierto("FormInicial00Principal") = True Then
        DoCmd.Close acForm, "FormInicial00Principal", acSaveNo
    End If
    DoCmd.OpenForm "FormInicial00Principal"

    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CambiarEstadoAFormulariosAbiertos ha devuelto el error: " & Err.Description
    End If
End Function
Public Function RellenarComboUsuarioHPS( _
                                        ByRef p_Cbo As ComboBox, _
                                        Optional ByRef p_Error As String _
                                        ) As String

    
    Dim m_ObjCol As Collection
    Dim m_VarItem As Variant
    On Error GoTo errores
    
    
    Set m_ObjCol = Constructor.getListaUsuariosHPSParaCombo(p_Error)
    
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    p_Cbo.RowSource = ""
    If Not m_ObjCol Is Nothing Then
        For Each m_VarItem In m_ObjCol
            p_Cbo.AddItem m_VarItem
        Next
    End If
    Set p_Cbo = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarComboUsuarioHPS ha devuelto el error: " & Err.Description
    End If
                                                
End Function

Public Function RellenarComboUsuariosHistoricos( _
                                                ByRef p_Cbo As ComboBox, _
                                                Optional ByRef p_Error As String _
                                                ) As String

    
    Dim m_ObjCol As Collection
    Dim m_VarItem As Variant
    On Error GoTo errores
    
    
    Set m_ObjCol = Constructor.getListaUsuariosHistoricosParaCombo(p_Error)
    
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    p_Cbo.RowSource = ""
    If Not m_ObjCol Is Nothing Then
        For Each m_VarItem In m_ObjCol
            p_Cbo.AddItem m_VarItem
        Next
    End If
    Set p_Cbo = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarComboUsuariosHistoricos ha devuelto el error: " & Err.Description
    End If
                                                
End Function
Public Function FormatoDatoValido(dato As String, TipoDato As String, Optional ByRef p_Error As String) As Boolean
    
    Dim mObjExp As Object
    Dim Patron As String
    
    On Error GoTo errores
    If TipoDato = "email" Then
        Patron = "^[a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9-]+(?:\.[a-zA-Z0-9-]+)*$"
    ElseIf TipoDato = "DNI" Then
        Patron = "((([X-Z])|([LM])){1}([-]?)((\d){7})([-]?)([A-Z]{1}))|((\d{8})([-]?)([A-Z]))"
    Else
        p_Error = "Tipo no reconocido"
        Err.Raise 1000
    End If
    
    Set mObjExp = CreateObject("VBSCRIPT.RegExp")
    mObjExp.Pattern = Patron
    FormatoDatoValido = mObjExp.Test(dato)
    Set mObjExp = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método FormatoDatoValido ha devuelto el error: " & Err.Description
    End If
    Debug.Print p_Error
End Function
Public Function FiltrarUsuarioActivo(p_ID As String, Optional ByRef p_Error As String) As String

    Dim frm As Form
    Dim rcdDatos As DAO.Recordset
    
    On Error GoTo errores
    
    If Not FormularioAbierto("FormInicial00Principal") Then
        Exit Function
    End If
    
    Set frm = Forms("FormInicial00Principal")
    If frm.SubFormCentral.SourceObject = "FormInicial02ConsultasPrincipal" Then
        Set frm = frm.SubFormCentral.Form
        If frm.SubFormResultados.SourceObject = "FormInicial03ConsultasDatos" Then
            Set frm = frm.SubFormResultados.Form
            frm.Requery
            Set rcdDatos = frm.RecordsetClone 'Clonamos el recordset
            
            rcdDatos.FindFirst "[ID] =" & p_ID  'Buscamos
            If Not rcdDatos.NoMatch Then 'Si  se encuentra
                frm.Bookmark = rcdDatos.Bookmark 'Nos vamos al registro encontrado
                frm.ID.SetFocus 'Activamos el campo ID (ahora es cuando se produce el scroll)
            End If
            rcdDatos.Close
            Set rcdDatos = Nothing
        End If
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método FiltrarUsuarioActivo ha devuelto el error: " & Err.Description
    End If
End Function



Public Function CopiarDatosDeOficinaACarpetaLocal( _
                                                p_URLCarpeta As String, _
                                                Optional ByRef p_Error As String _
                                                ) As String
                                
    Dim URLAnexosOficina As String
    Dim URLAnexosLocal As String
    Dim URLAnexosSICAOficina As String
    Dim URLAnexosSICALocal As String
    Dim URLAnexosHistoricosOficina As String
    Dim URLAnexosHistoricosLocal As String
    Dim URLRecursosOficina As String
    Dim URLRecursosLocal As String
    
    Dim strCmd As String
    Dim resultadoShell As String
    
    On Error GoTo errores
    
    ' --- Validaciones Previas ---
    If m_EnOficina = EnumSiNo.No Then
        p_Error = "Esta Operación sólo se puede hacer estando en la oficina"
        Err.Raise 1000
    End If
    If Not fso.FolderExists(p_URLCarpeta) Then
        p_Error = "Se ha de indicar una carpeta local alcanzable" & vbNewLine & p_URLCarpeta
        Err.Raise 1000
    End If
    
    ' --- Definición de Rutas ---
    
    ' 1. ANEXOS HPS
    URLAnexosOficina = m_ObjEntorno.URLCarpetaAnexos
    If Not fso.FolderExists(URLAnexosOficina) Then
        p_Error = "No es Alcanzable la carpeta de anexos de la Oficina" & vbNewLine & URLAnexosOficina
        Err.Raise 1000
    End If
    URLAnexosLocal = p_URLCarpeta & "\ANEXOS\HPS"
    
    ' 2. ANEXOS SICA
    URLAnexosSICAOficina = m_ObjEntorno.URLCarpetaAnexosSICA
    If Not fso.FolderExists(URLAnexosSICAOficina) Then
        p_Error = "No es Alcanzable la carpeta de anexos SICA de la Oficina" & vbNewLine & URLAnexosSICAOficina
        Err.Raise 1000
    End If
    URLAnexosSICALocal = p_URLCarpeta & "\ANEXOS\SICA"
    
    ' 3. ANEXOS HISTORICO
    URLAnexosHistoricosOficina = m_ObjEntorno.URLCarpetaAnexosHistoricos
    If Not fso.FolderExists(URLAnexosHistoricosOficina) Then
        p_Error = "No es Alcanzable la carpeta de anexos HISTÓRICOS de la Oficina" & vbNewLine & URLAnexosHistoricosOficina
        Err.Raise 1000
    End If
    URLAnexosHistoricosLocal = p_URLCarpeta & "\ANEXOS\HISTORICO"
    
    ' 4. RECURSOS
    URLRecursosOficina = m_ObjEntorno.URLCarpetaRecursos
    URLRecursosLocal = p_URLCarpeta & "\recursos"

    ' --- Creación de carpetas locales si no existen ---
    CrearCarpetaSiNoExiste p_URLCarpeta & "\ANEXOS"
    CrearCarpetaSiNoExiste URLAnexosLocal
    CrearCarpetaSiNoExiste URLAnexosSICALocal
    CrearCarpetaSiNoExiste URLAnexosHistoricosLocal
    CrearCarpetaSiNoExiste URLRecursosLocal

'     --- EJECUCIÓN DE ROBOCOPY ---
'     Nota: Usamos /MIR (Espejo) /R:0 (Reintentos 0) /W:0 (Espera 0) /NP (No progress bar para no ensuciar logs)
    
    ' 1. Copiar RECURSOS
    strCmd = "ROBOCOPY " & Chr(34) & QuitarBarraFinal(URLRecursosOficina) & Chr(34) & " " & _
             Chr(34) & QuitarBarraFinal(URLRecursosLocal) & Chr(34) & " /MIR /R:0 /W:0 /NP"
    EjecutarShell strCmd, p_Error
    If p_Error <> "" Then Err.Raise 1000

    ' 2. Copiar ANEXOS HPS
    strCmd = "ROBOCOPY " & Chr(34) & QuitarBarraFinal(URLAnexosOficina) & Chr(34) & " " & _
             Chr(34) & QuitarBarraFinal(URLAnexosLocal) & Chr(34) & " /MIR /R:0 /W:0 /NP"
    EjecutarShell strCmd, p_Error
    If p_Error <> "" Then Err.Raise 1000
    
    ' 3. Copiar ANEXOS SICA
    strCmd = "ROBOCOPY " & Chr(34) & QuitarBarraFinal(URLAnexosSICAOficina) & Chr(34) & " " & _
             Chr(34) & QuitarBarraFinal(URLAnexosSICALocal) & Chr(34) & " /MIR /R:0 /W:0 /NP"
    EjecutarShell strCmd, p_Error
    If p_Error <> "" Then Err.Raise 1000
    
    ' 4. Copiar ANEXOS HISTORICO
    strCmd = "ROBOCOPY " & Chr(34) & QuitarBarraFinal(URLAnexosHistoricosOficina) & Chr(34) & " " & _
             Chr(34) & QuitarBarraFinal(URLAnexosHistoricosLocal) & Chr(34) & " /MIR /R:0 /W:0 /NP"
    EjecutarShell strCmd, p_Error
    If p_Error <> "" Then Err.Raise 1000
    
    CopiarDatosDeOficinaACarpetaLocal = "OK"
    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "Error en CopiarDatosDeOficinaACarpetaLocal: " & Err.Description
    End If
End Function

' Pequeña función auxiliar para limpiar el código principal
Private Sub CrearCarpetaSiNoExiste(ruta As String)
    If Not fso.FolderExists(ruta) Then
        fso.CreateFolder ruta
    End If
End Sub


Public Function CerrarFormulariosAbiertos(Optional ByRef p_Error As String) As String
    Dim obj As AccessObject, dbs As Object
    Dim m_FormularioActual As String
    On Error Resume Next
    m_FormularioActual = Application.Screen.ActiveForm.Name
    If Err.Number <> 0 Then
        Err.Clear
        
    End If
    Set dbs = Application.CurrentProject
    ' Search for open AccessObject objects in AllForms collection.
    For Each obj In dbs.AllForms
        If obj.IsLoaded = True Then
            ' Print name of obj.
            If m_FormularioActual <> "" Then
                If m_FormularioActual = obj.Name Then
                    GoTo siguiente
                End If
            End If
            DoCmd.Close acForm, obj.Name, acSaveNo
        End If
siguiente:
    Next obj
End Function
Public Function CopiarEstiloBoton( _
                                    Optional ByRef p_Boton As CommandButton, _
                                    Optional ByRef p_Error As String _
                                    ) As String

    Dim btnOrigen As CommandButton
    Dim ColPropiedades As Scripting.Dictionary
    On Error GoTo errores
    
    If p_Boton Is Nothing Then
        Exit Function
    End If
    If Application.TempVars("SoloLectura") = "Sí" Then
        Set ColPropiedades = m_ObjEntorno.ColPropiedadesBotonLectura
    Else
        Set ColPropiedades = m_ObjEntorno.ColPropiedadesBotonEscritura
    End If
    With btnOrigen
        p_Boton.BackColor = ColPropiedades("BackColor")
        p_Boton.BackShade = ColPropiedades("BackShade")
        p_Boton.BackStyle = ColPropiedades("BackStyle")
        p_Boton.BackTint = ColPropiedades("BackTint")
        p_Boton.BorderThemeColorIndex = ColPropiedades("BorderThemeColorIndex")
        p_Boton.ThemeFontIndex = ColPropiedades("ThemeFontIndex")
        p_Boton.BorderColor = ColPropiedades("BorderColor")
        p_Boton.BorderShade = ColPropiedades("BorderShade")
        p_Boton.BorderStyle = ColPropiedades("BorderStyle")
        p_Boton.BorderTint = ColPropiedades("BorderTint")
        p_Boton.CursorOnHover = ColPropiedades("CursorOnHover")
        p_Boton.FontSize = ColPropiedades("FontSize")
        p_Boton.FontWeight = ColPropiedades("FontWeight")
        p_Boton.FontBold = ColPropiedades("FontBold")
        p_Boton.ForeColor = ColPropiedades("ForeColor")
        p_Boton.ForeShade = ColPropiedades("ForeShade")
        p_Boton.ForeThemeColorIndex = ColPropiedades("ForeThemeColorIndex")
        p_Boton.ForeTint = ColPropiedades("ForeTint")
        p_Boton.Glow = ColPropiedades("Glow")
        p_Boton.Gradient = ColPropiedades("Gradient")
        p_Boton.HoverColor = ColPropiedades("HoverColor")
        p_Boton.HoverForeColor = ColPropiedades("HoverForeColor")
        p_Boton.HoverForeShade = ColPropiedades("HoverForeShade")
        p_Boton.HoverForeThemeColorIndex = ColPropiedades("HoverForeThemeColorIndex")
        p_Boton.HoverForeTint = ColPropiedades("HoverForeTint")
        p_Boton.HoverShade = ColPropiedades("HoverShade")
        p_Boton.HoverTint = ColPropiedades("HoverTint")
        p_Boton.PressedColor = ColPropiedades("PressedColor")
        p_Boton.PressedForeColor = ColPropiedades("PressedForeColor")
        p_Boton.PressedForeShade = ColPropiedades("PressedForeShade")
        p_Boton.PressedForeThemeColorIndex = ColPropiedades("PressedForeThemeColorIndex")
        p_Boton.ForeTint = ColPropiedades("ForeTint")
        p_Boton.PressedShade = ColPropiedades("PressedShade")
        p_Boton.PressedThemeColorIndex = ColPropiedades("PressedThemeColorIndex")
        p_Boton.PressedTint = ColPropiedades("PressedTint")
        p_Boton.Shadow = ColPropiedades("Shadow")
        p_Boton.Shape = ColPropiedades("Shape")
        p_Boton.ThemeFontIndex = ColPropiedades("ThemeFontIndex")
        p_Boton.Transparent = ColPropiedades("Transparent")
        p_Boton.UseTheme = ColPropiedades("UseTheme")
        
    End With
    CopiarEstiloBoton = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CopiarEstiloBoton ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getCadenaCorreosUsuarios( _
                                            p_ColUsuarios As Scripting.Dictionary, _
                                            Optional ByRef p_Error As String _
                                            ) As String

    Dim m_ID As Variant
    Dim m_Correo As String
    Dim m_ObjUsuario As Object
    
    Dim m_Cadena As String
    On Error GoTo errores
    For Each m_ID In p_ColUsuarios.Keys
        Set m_ObjUsuario = p_ColUsuarios(m_ID)
        m_Correo = m_ObjUsuario.Correo_e
        If InStr(1, m_Correo, "@") <> 0 Then
            If m_Cadena = "" Then
                m_Cadena = m_Correo
            Else
                m_Cadena = m_Cadena & ";" & m_Correo
            End If
        End If
        Set m_ObjUsuario = Nothing
    Next
    getCadenaCorreosUsuarios = m_Cadena
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getCadenaCorreosUsuarios ha producido el error: " & Err.Description
    End If
End Function

Public Function getUsuarioSICAActivo( _
                                        Optional p_Error As String _
                                    ) As UsuarioSICA
    Dim m_ID As String
    Dim frm As Form
    
    On Error GoTo errores
    Set frm = Forms("FormInicial00Principal").SubFormCentral.Form
    If frm.Name <> "FormInicial07UsuariosSICA" Then
        
        Exit Function
    End If
    Set frm = frm.SubFormResultados.Form
    m_ID = Nz(frm.ID, "")
    If m_ID = "" Then
       
        Exit Function
    End If
    
    Set m_ObjUsuarioSICAActivo = Constructor.getUsuarioSICA(m_ID, , , p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getUsuarioSICAActivo ha devuelto el error: " & Err.Description
    End If
    
End Function
Public Function getIDUsuario( _
                                Optional frm As Form, _
                                Optional p_NombreConID As String, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    Dim m_Resultado As String
    Dim ctlUsuario As ComboBox
    On Error GoTo errores
    
    If Not frm Is Nothing Then
        Set ctlUsuario = frm.UsuarioHPS
        m_Resultado = Nz(ctlUsuario.value, "")
    Else
        m_Resultado = p_NombreConID
    End If
    If InStr(1, m_Resultado, "(") = 0 Or InStr(1, m_Resultado, ")") = 0 Then
        Exit Function
    End If
    dato = Split(m_Resultado, "(")
    m_Resultado = dato(1)
    dato = Split(m_Resultado, ")")
    getIDUsuario = dato(0)
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getIDUsuario ha producido el error: " & Err.Description
    End If
End Function

Public Function getIDUsuarioHistorico( _
                                        Optional frm As Form, _
                                        Optional p_NombreConID As String, _
                                        Optional ByRef p_Error As String _
                                        ) As String
    
    Dim m_Resultado As String
    Dim ctlUsuario As ComboBox
    On Error GoTo errores
    
    If Not frm Is Nothing Then
        Set ctlUsuario = frm.UsuarioHistorico
        m_Resultado = Nz(ctlUsuario.value, "")
    Else
        m_Resultado = p_NombreConID
    End If
    If InStr(1, m_Resultado, "(") = 0 Or InStr(1, m_Resultado, ")") = 0 Then
        Exit Function
    End If
    dato = Split(m_Resultado, "(")
    m_Resultado = dato(1)
    dato = Split(m_Resultado, ")")
    getIDUsuarioHistorico = dato(0)
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getIDUsuarioHistorico ha producido el error: " & Err.Description
    End If
End Function


Public Sub SetText(Text As String)
    #If Win64 = 1 Then
        Dim hGlobalMemory As LongPtr
        Dim lpGlobalMemory As LongPtr
        Dim hClipMemory As LongPtr
    #Else

        Dim hGlobalMemory As Long
        Dim lpGlobalMemory As Long
        Dim hClipMemory As Long

    #End If

    Const GHND = &H42
    Const CF_TEXT = 1

   ' Allocate moveable global memory.
   '-------------------------------------------
   hGlobalMemory = GlobalAlloc(GHND, Len(Text) + 1)

   ' Lock the block to get a far pointer
   ' to this memory.
   lpGlobalMemory = GlobalLock(hGlobalMemory)

   ' Copy the string to this global memory.
   lpGlobalMemory = lstrcpy(lpGlobalMemory, Text)

   ' Unlock the memory.
   If GlobalUnlock(hGlobalMemory) <> 0 Then
      MsgBox "Could not unlock memory location. Copy aborted."
      GoTo CloseClipboard
   End If

   ' Open the Clipboard to copy data to.
   If OpenClipboard(0&) = 0 Then
      MsgBox "Could not open the Clipboard. Copy aborted."
      Exit Sub
   End If

   ' Clear the Clipboard.
   Call EmptyClipboard

   ' Copy the data to the Clipboard.
   hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory)

CloseClipboard:

   If CloseClipboard() = 0 Then
      MsgBox "Could not close Clipboard."
   End If
End Sub
Public Function GetText()
    #If VBA7 Then
        Dim hClipMemory As LongPtr
        Dim lpClipMemory As LongPtr
    #Else
    
        Dim hClipMemory As Long
        Dim lpClipMemory As Long
    #End If

    Dim MaximumSize As Long
    Dim ClipText As String

    Const CF_TEXT = 1

   If OpenClipboard(0&) = 0 Then
      MsgBox "Cannot open Clipboard. Another app. may have it open"
      Exit Function
   End If
          
   ' Obtain the handle to the global memory block that is referencing the text.
   hClipMemory = GetClipboardData(CF_TEXT)
   If IsNull(hClipMemory) Then
      MsgBox "Could not allocate memory"
      GoTo CloseClipboard
   End If
 
   ' Lock Clipboard memory so we can reference the actual data string.
   lpClipMemory = GlobalLock(hClipMemory)
 
   If Not IsNull(lpClipMemory) Then
      MaximumSize = 64
      
      Do
        MaximumSize = MaximumSize * 2
        
        ClipText = Space$(MaximumSize)
        Call lstrcpy(ClipText, lpClipMemory)
        Call GlobalUnlock(hClipMemory)
             
      Loop Until ClipText Like "*" & vbNullChar & "*"
      
      ' Peel off the null terminating character.
      ClipText = Left$(ClipText, InStrRev(ClipText, vbNullChar) - 1)
      
   Else
      MsgBox "Could not lock memory to copy string from."
   End If
 
CloseClipboard:
 
   Call CloseClipboard
   GetText = ClipText
 
End Function



Public Function DNIEnActivo( _
                            p_DNI As String, _
                            Optional ByRef p_Error As String _
                            ) As String
    
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    
    On Error GoTo errores
    If p_DNI = "" Then
        p_Error = "No se ha indicado p_DNI"
        Err.Raise 1000
    End If
     m_SQL = "SELECT TbUsuarios.Nombre,TbUsuarios.Apellido_1,TbUsuarios.Apellido_2 " & _
            "FROM TbUsuarios " & _
            "WHERE DNI='" & p_DNI & "';"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    
    With rcdDatos
        If Not .EOF Then
            DNIEnActivo = .Fields("Nombre") & " " & Nz(.Fields("Apellido_1"), "") & " " & Nz(.Fields("Apellido_2"), "")
        End If
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método DNIEnActivo ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function DNIEnHistorico( _
                                p_DNI As String, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    
    On Error GoTo errores
    If p_DNI = "" Then
        p_Error = "No se ha indicado p_DNI"
        Err.Raise 1000
    End If
     m_SQL = "SELECT TbUsuariosHistoricos.Nombre,TbUsuariosHistoricos.Apellido_1,TbUsuariosHistoricos.Apellido_2 " & _
            "FROM TbUsuariosHistoricos " & _
            "WHERE DNI='" & p_DNI & "';"
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            DNIEnHistorico = .Fields("Nombre") & " " & Nz(.Fields("Apellido_1"), "") & " " & Nz(.Fields("Apellido_2"), "")
        End If
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método DNIEnHistorico ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function


Public Function EnOficina(Optional ByRef p_Error As String) As EnumSiNo
    
    Dim strIPS As String
    On Error GoTo errores
    strIPS = GetIPAddresses
    If InStr(1, strIPS, SubRedOficina) = 0 Then
        EnOficina = EnumSiNo.No
    Else
        EnOficina = EnumSiNo.Sí
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EnOficina ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function

Public Function DameID( _
                        p_NombreTabla As String, _
                        p_NombreCampoID As String, _
                        Optional ByRef p_Db As DAO.Database, _
                        Optional ByRef p_Error As String _
                        ) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim lngIDMax As Long
    On Error GoTo errores
    
    If p_NombreTabla = "" Or p_NombreCampoID = "" Then
        p_Error = "Se ha de indicar el nombre de la tabla y de su campo ID"
        Err.Raise 1000
    End If
    If p_Db Is Nothing Then
        Set p_Db = getdb()
    End If
    m_SQL = "SELECT Max(" & p_NombreTabla & "." & p_NombreCampoID & ") AS MaxID " & _
            "FROM " & p_NombreTabla & ";"
    Set rcdDatos = p_Db.OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            If IsNumeric(Nz(.Fields("MaxID"), "")) Then
                lngIDMax = .Fields("MaxID")
            End If
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    DameID = CStr(lngIDMax + 1)
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método DameID ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function


Public Function EstablecerSemaforo(Optional p_ID As String, Optional ByRef p_Error As String) As String
    
    Dim frm As Form
    Dim m_URLSemaforo As String
    Dim m_ObjDatosLocal As DatosLocal
    
    On Error GoTo errores
    
    
    If Not FormularioAbierto("FormInicial00Principal") Then
        Exit Function
    End If
    Set frm = Forms("FormInicial00Principal")
    If frm.SubFormLateral.SourceObject <> "FormInicial01Lateral" Then
        Exit Function
    End If
    Set frm = frm.SubFormLateral.Form
    If p_ID = "" Then
        frm.Controls("ImagenSemaforo").Visible = False
        Set frm = Nothing
        Exit Function
    End If
    Set m_ObjDatosLocal = Constructor.getDatosLocal(p_ID, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_URLSemaforo = m_ObjDatosLocal.URLImagenSemaforo
    p_Error = m_ObjDatosLocal.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    frm.Controls("ImagenSemaforo").Visible = True
    frm.Controls("ImagenSemaforo").Picture = m_URLSemaforo
    Set frm = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerSemaforo ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function


Public Function EstablecerIndicadores(Optional ByRef p_Error As String) As String
    
    Dim frm As Form
    Dim m_Titulo As String
    Dim intNumero As Integer
    Dim m_NombreBoton As String
    On Error GoTo errores
    
    CopiarDatosAIndicadores p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_ObjIndicadores = New Indicador
    
    Set frm = Forms("FormInicial00Principal")
    If frm.SubFormCentral.SourceObject <> "FormInicial09Indicadores" Then
        Exit Function
    End If
    Set frm = frm.SubFormCentral.Form
    m_NombreBoton = "ComandoPendientesCursoHPS"
    m_Titulo = "HPS Pte. Curso ( # )"
    intNumero = m_ObjIndicadores.NPendientesCursoHPS
    p_Error = m_ObjIndicadores.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_Titulo = Replace(m_Titulo, "#", intNumero)
    frm.Controls(m_NombreBoton).Caption = m_Titulo
    
     
    
    m_NombreBoton = "ComandoAPuntoDeCaducar"
    m_Titulo = "A Punto de Caducar ( # )"
    intNumero = m_ObjIndicadores.NHPSAPuntoDeCaducar
    p_Error = m_ObjIndicadores.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_Titulo = Replace(m_Titulo, "#", intNumero)
    frm.Controls(m_NombreBoton).Caption = m_Titulo
    
    m_NombreBoton = "ComandoCaducados"
    m_Titulo = "Caducados ( # )"
    intNumero = m_ObjIndicadores.NHPSACaducadas
    p_Error = m_ObjIndicadores.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_Titulo = Replace(m_Titulo, "#", intNumero)
    frm.Controls(m_NombreBoton).Caption = m_Titulo
    
    m_NombreBoton = "ComandoEnSolicitud"
    m_Titulo = "En Solicitud ( # )"
    intNumero = m_ObjIndicadores.NHPSASolicitandose
    p_Error = m_ObjIndicadores.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_Titulo = Replace(m_Titulo, "#", intNumero)
    frm.Controls(m_NombreBoton).Caption = m_Titulo
    
    
    m_NombreBoton = "ComandoPendientesConvocatoria1"
    m_Titulo = "Pte.1ª Convocat. ( # )"
    intNumero = m_ObjIndicadores.NUsuariosPtesPrimeraConvocatoriaCurso
    p_Error = m_ObjIndicadores.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_Titulo = Replace(m_Titulo, "#", intNumero)
    frm.Controls(m_NombreBoton).Caption = m_Titulo
    
    
    m_NombreBoton = "ComandoPendientesConvocatoria2"
    m_Titulo = "Pte.2ª Convocat. ( # )"
    intNumero = m_ObjIndicadores.NUsuariosPtesSegundaConvocatoriaCurso
    p_Error = m_ObjIndicadores.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_Titulo = Replace(m_Titulo, "#", intNumero)
    frm.Controls(m_NombreBoton).Caption = m_Titulo
    
    m_NombreBoton = "ComandoPendientesCorreoJefeSeguridad"
    m_Titulo = "Pte. Correo JS ( # )"
    intNumero = m_ObjIndicadores.NUsuariosPtesEnvioCorreoJefeSeguridadCurso
    p_Error = m_ObjIndicadores.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_Titulo = Replace(m_Titulo, "#", intNumero)
    frm.Controls(m_NombreBoton).Caption = m_Titulo
    Set frm = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerIndicadores ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function
Public Function CopiarDatosAIndicadores( _
                            Optional ByRef p_Error As String, _
                            Optional ByVal p_CacheAcabDeRegenerar As Boolean = False) As String
    
    Dim m_SQL As String
    Static s_YaCopiadoEnEstaCarga As Boolean  ' Evita copias multiples durante la cascada Form_Load
    
    On Error GoTo errores
    
    ' Durante la carga de FormInicial09Indicadores, el formulario principal y sus
    ' subformularios continuos llaman a esta funcion casi a la vez. Las tablas
    ' fisicas son necesarias como RecordSource, por eso la primera llamada SIEMPRE
    ' sincroniza TbDatosLocalParaIndicadores y TbUsuariosSICALocalParaIndicadores.
    ' Las llamadas posteriores de la misma cascada se omiten para no repetir el
    ' DELETE+INSERT masivo 3-4 veces.
    
    If p_CacheAcabDeRegenerar Then
        s_YaCopiadoEnEstaCarga = False
        Exit Function
    End If
    
    If s_YaCopiadoEnEstaCarga Then
        Debug.Print "CopiarDatosAIndicadores: skip (ya sincronizado en esta carga)"
        Exit Function
    End If
    
    m_SQL = "DELETE * FROM TbDatosLocalParaIndicadores;"
    CurrentDb.Execute m_SQL, dbFailOnError
    
    m_SQL = "INSERT INTO TbDatosLocalParaIndicadores " & _
            "SELECT TbDatosLocal.* " & _
            "FROM TbDatosLocal;"
    CurrentDb.Execute m_SQL, dbFailOnError
    
    m_SQL = "DELETE * FROM TbUsuariosSICALocalParaIndicadores;"
    CurrentDb.Execute m_SQL, dbFailOnError
    
    m_SQL = "INSERT INTO TbUsuariosSICALocalParaIndicadores " & _
            "SELECT TbUsuariosSICALocal.* " & _
            "FROM TbUsuariosSICALocal;"
    CurrentDb.Execute m_SQL, dbFailOnError
    
    s_YaCopiadoEnEstaCarga = True
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CopiarDatosAIndicadores ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function
Public Function EstablecerFuenteDatosIndicadores( _
                                                    p_TipoIndicador As EnumTipoIndicador, _
                                                    Optional ByRef p_Error As String _
                                                    ) As String
    
    Dim frm As Form
    Dim m_SQL As String
    Dim m_SQLInicial As String
    Dim m_SQLInicialSICA As String
    On Error GoTo errores
   
    Set frm = Forms("FormInicial00Principal")
    If frm.SubFormCentral.SourceObject <> "FormInicial09Indicadores" Then
        Exit Function
    End If
    Set frm = frm.SubFormCentral.Form
    
    If frm Is Nothing Then
        Exit Function
    End If
    m_SQLInicial = "SELECT * " & _
                    "FROM TbDatosLocalParaIndicadores "
                    
'    m_SQLInicial = "SELECT * " & _
'                    "FROM TbUsuariosSICALocalParaIndicadores "
                    
    
    If p_TipoIndicador = EnumTipoIndicador.PendientesCursoHPS Then
        m_SQL = m_SQLInicial & _
                "WHERE Requiere_Curso='Sí';"
        If frm.SubFormIndicadores.SourceObject <> "FormInicial09IndicadoresDatosLocalDatos" Then
            frm.SubFormIndicadores.SourceObject = "FormInicial09IndicadoresDatosLocalDatos"
        End If
    ElseIf p_TipoIndicador = EnumTipoIndicador.PtesPrimeraConvocatoriaCurso Then
        m_SQL = m_SQLInicial & _
                "WHERE (((TbDatosLocalParaIndicadores.Requiere_PrimeraConvocatoriaCurso)='Sí'));"
        If frm.SubFormIndicadores.SourceObject <> "FormInicial09IndicadoresDatosLocalDatos" Then
            frm.SubFormIndicadores.SourceObject = "FormInicial09IndicadoresDatosLocalDatos"
        End If
    ElseIf p_TipoIndicador = EnumTipoIndicador.PtesSegundaConvocatoriaCurso Then
        m_SQL = m_SQLInicial & _
                "WHERE Requiere_SegundaConvocatoriaCurso='Sí';"
        If frm.SubFormIndicadores.SourceObject <> "FormInicial09IndicadoresDatosLocalDatos" Then
            frm.SubFormIndicadores.SourceObject = "FormInicial09IndicadoresDatosLocalDatos"
        End If
    ElseIf p_TipoIndicador = EnumTipoIndicador.PtesEnvioCorreoJefeSeguridadCurso Then
        m_SQL = m_SQLInicial & _
                "WHERE Requiere_CorreoJefeSeguridadCurso='Sí';"
        If frm.SubFormIndicadores.SourceObject <> "FormInicial09IndicadoresDatosLocalDatos" Then
            frm.SubFormIndicadores.SourceObject = "FormInicial09IndicadoresDatosLocalDatos"
        End If
    
    ElseIf p_TipoIndicador = EnumTipoIndicador.HPSAPuntoDeCaducar Then
        m_SQL = m_SQLInicial & _
                "WHERE HPS_NAC_ApuntoDeCaducar='Sí' OR " & _
                "HPS_OTAN_ApuntoDeCaducar='Sí' or " & _
                "HPS_ESA_ApuntoDeCaducar='Sí' OR " & _
                "HPS_UE_ApuntoDeCaducar='Sí';"
        If frm.SubFormIndicadores.SourceObject <> "FormInicial09IndicadoresDatosLocalDatos" Then
            frm.SubFormIndicadores.SourceObject = "FormInicial09IndicadoresDatosLocalDatos"
        End If
    ElseIf p_TipoIndicador = EnumTipoIndicador.HPSACaducadas Then
        m_SQL = m_SQLInicial & _
                "WHERE HPS_NAC_Caducado='Sí' OR " & _
                "HPS_OTAN_Caducado='Sí' or " & _
                "HPS_ESA_Caducado='Sí' OR " & _
                "HPS_UE_Caducado='Sí';"
        If frm.SubFormIndicadores.SourceObject <> "FormInicial09IndicadoresDatosLocalDatos" Then
            frm.SubFormIndicadores.SourceObject = "FormInicial09IndicadoresDatosLocalDatos"
        End If
    ElseIf p_TipoIndicador = EnumTipoIndicador.HPSASolicitandose Then
        m_SQL = m_SQLInicial & _
                "WHERE Not HPS_UE_F_Solicitud Is Null OR " & _
                "Not HPS_OTAN_F_Solicitud Is Null OR " & _
                "NOt HPS_ESA_F_Solicitud Is Null OR " & _
                "Not HPS_UE_F_Solicitud Is Null;"
        If frm.SubFormIndicadores.SourceObject <> "FormInicial09IndicadoresDatosLocalDatos" Then
            frm.SubFormIndicadores.SourceObject = "FormInicial09IndicadoresDatosLocalDatos"
        End If
    End If
    Set frm = frm.SubFormIndicadores.Form
    
    frm.RecordSource = m_SQL
    frm.Requery
    Set frm = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerFuenteDatosIndicadores ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function
Public Function EstablecerFuenteDatosIndicadoresConEmpresas( _
                                                                p_TipoEmpresa As String, _
                                                                p_NombreEmpresa As String, _
                                                                Optional ByRef p_Error As String _
                                                                ) As String
        
    Dim frm As Form
    Dim m_SQL As String
    Dim m_SQLFinal As String
    
    On Error GoTo errores
    Set frm = Forms("FormInicial00Principal")
    If frm.SubFormCentral.SourceObject <> "FormInicial09Indicadores" Then
        Exit Function
    End If
    Set frm = frm.SubFormCentral.Form
    If frm Is Nothing Then
        Exit Function
    End If
    If frm.SubFormIndicadores.SourceObject <> "FormInicial09IndicadoresDatosLocalDatos" Then
        Exit Function
    End If
    Set frm = frm.SubFormIndicadores.Form
    
    m_SQL = frm.RecordSource
    m_SQLFinal = getSQLConEmpresaSeleccionada(m_SQL, p_TipoEmpresa, p_NombreEmpresa, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    frm.RecordSource = m_SQLFinal
    frm.Requery
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GenerarWherePorColCamposSinEmpresas ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function
Private Function getSQLConEmpresaSeleccionada( _
                                                p_SQL As String, _
                                                p_TipoEmpresa As String, _
                                                p_NombreEmpresa As String, _
                                                Optional ByRef p_Error As String _
                                                ) As String
        
    Dim m_WhereSinEmpresas As String
    Dim m_whereFinal As String
    Dim m_Col As Scripting.Dictionary
    Dim m_ParteSQLNoWhere As String
    On Error GoTo errores
    Set m_Col = getColCamposNoEmpresa(p_SQL, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    m_WhereSinEmpresas = ConstruirWhereConColeccion(m_Col, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_whereFinal = getWhereAdicionandoEmpresaYValor(m_WhereSinEmpresas, p_TipoEmpresa, p_NombreEmpresa, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If InStr(1, p_SQL, "WHERE") <> 0 Then
        dato = Split(p_SQL, "WHERE")
        m_ParteSQLNoWhere = dato(0)
    Else
        m_ParteSQLNoWhere = p_SQL
        m_ParteSQLNoWhere = Replace(m_ParteSQLNoWhere, ";", "")
        
    End If
    m_ParteSQLNoWhere = Trim(m_ParteSQLNoWhere)
    getSQLConEmpresaSeleccionada = m_ParteSQLNoWhere & " " & m_whereFinal
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getSQLConEmpresaSeleccionada ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function

Private Function getColCamposNoEmpresa( _
                                    p_SQL As String, _
                                    Optional ByRef p_Error As String _
                                    ) As Scripting.Dictionary
    Dim dato As Variant
    Dim dato1 As Variant
    Dim m_trozoDespuesWhere As String
    Dim m_TrozoAnd As Variant
    Dim m_TrozoIgual As String
    Dim m_TrozoAndAntesDeIgual As String
    Dim m_TrozoAndAntesDeIgualSinParentesis As String
    Dim m_Valor As String
    
    On Error GoTo errores
    'WHERE (((TbDatosLocalParaIndicadores.DNI)='50722531X') AND ((TbDatosLocalParaIndicadores.Nombre)='Fernando') " & _
        "AND ((TbDatosLocalParaIndicadores.Apellido_1)='Lázaro') AND ((TbSuministradoresLocal.Nombre)='TSOL S.A.U.'));
    If InStr(1, p_SQL, "WHERE") = 0 Then
        Exit Function
    End If
    dato = Split(p_SQL, "WHERE")
    m_trozoDespuesWhere = dato(1)
    m_trozoDespuesWhere = Trim(m_trozoDespuesWhere)
    If InStr(1, m_trozoDespuesWhere, ";") <> 0 Then
        dato = Split(m_trozoDespuesWhere, ";")
        m_trozoDespuesWhere = dato(0)
        m_trozoDespuesWhere = Trim(m_trozoDespuesWhere)
    End If
    
    'ahora tenemos el where sin where y sin ;
    'lo vamos a separar por AND
    If InStr(1, m_trozoDespuesWhere, ") AND (") <> 0 Then
        dato = Split(m_trozoDespuesWhere, ") AND (")
        For Each m_TrozoAnd In dato
            m_TrozoAnd = Trim(m_TrozoAnd)
            If InStr(1, m_TrozoAnd, "=") <> 0 Then
                dato1 = Split(m_TrozoAnd, "=")
                m_TrozoAndAntesDeIgual = Trim(dato1(0))
                m_TrozoAndAntesDeIgualSinParentesis = Trim(Replace(m_TrozoAndAntesDeIgual, "(", ""))
                m_TrozoAndAntesDeIgualSinParentesis = Trim(Replace(m_TrozoAndAntesDeIgualSinParentesis, ")", ""))
                m_Valor = Trim(dato1(1))
                m_Valor = Trim(Replace(m_Valor, "(", ""))
                m_Valor = Trim(Replace(m_Valor, ")", ""))
                
                If InStr(1, m_TrozoAndAntesDeIgualSinParentesis, "TbSuministradoresLocal") = 0 Then
                    If getColCamposNoEmpresa Is Nothing Then
                        Set getColCamposNoEmpresa = New Scripting.Dictionary
                        getColCamposNoEmpresa.CompareMode = TextCompare
                    End If
                    If Not getColCamposNoEmpresa.Exists(m_TrozoAndAntesDeIgualSinParentesis) Then
                        getColCamposNoEmpresa.Add m_TrozoAndAntesDeIgualSinParentesis, m_Valor
                    End If
                End If
            End If
            
            m_TrozoAnd = ""
            m_TrozoAndAntesDeIgual = ""
            m_TrozoAndAntesDeIgualSinParentesis = ""
            m_Valor = ""
        Next
    Else
        m_TrozoAnd = m_trozoDespuesWhere
        m_TrozoAnd = Trim(m_TrozoAnd)
        If InStr(1, m_TrozoAnd, "=") <> 0 Then
            dato1 = Split(m_TrozoAnd, "=")
            m_TrozoAndAntesDeIgual = Trim(dato1(0))
            m_TrozoAndAntesDeIgualSinParentesis = Trim(Replace(m_TrozoAndAntesDeIgual, "(", ""))
            m_TrozoAndAntesDeIgualSinParentesis = Trim(Replace(m_TrozoAndAntesDeIgualSinParentesis, ")", ""))
            m_Valor = Trim(dato1(1))
            m_Valor = Trim(Replace(m_Valor, "(", ""))
            m_Valor = Trim(Replace(m_Valor, ")", ""))
            
            If InStr(1, m_TrozoAndAntesDeIgualSinParentesis, "TbSuministradoresLocal") = 0 Then
                If getColCamposNoEmpresa Is Nothing Then
                    Set getColCamposNoEmpresa = New Scripting.Dictionary
                    getColCamposNoEmpresa.CompareMode = TextCompare
                End If
                If Not getColCamposNoEmpresa.Exists(m_TrozoAndAntesDeIgualSinParentesis) Then
                    getColCamposNoEmpresa.Add m_TrozoAndAntesDeIgualSinParentesis, m_Valor
                End If
            End If
        End If
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getColCamposNoEmpresa ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function

Private Function ConstruirWhereConColeccion( _
                                            p_Col As Scripting.Dictionary, _
                                            Optional ByRef p_Error As String _
                                            ) As String
    Dim m_Col As Scripting.Dictionary
    Dim m_Campo As Variant
    Dim m_Valor As String
    Dim m_Where As String
    Dim m_CampoCompuesto As String
    Dim m_CampoConValor As String
    Dim m_Cadena As String
    
    On Error GoTo errores
    'WHERE (((TbDatosLocalParaIndicadores.DNI)='50722531X') AND ((TbDatosLocalParaIndicadores.Nombre)='Fernando') " & _
        "AND ((TbDatosLocalParaIndicadores.Apellido_1)='Lázaro') AND ((TbSuministradoresLocal.Nombre)='TSOL S.A.U.'));
    If p_Col Is Nothing Then
        Exit Function
    End If
    
    For Each m_Campo In p_Col
        m_CampoConValor = m_Campo & "=" & p_Col(m_Campo)
        If m_Cadena = "" Then
            m_Cadena = m_CampoConValor
        Else
            m_Cadena = m_Cadena & " AND " & m_CampoConValor
        End If
    Next
    m_Where = "WHERE " & m_Cadena & ";"
    ConstruirWhereConColeccion = m_Where
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ConstruirWhereConColeccion ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function
Private Function getWhereAdicionandoEmpresaYValor( _
                                                p_Where As String, _
                                                p_TipoEmpresa As String, _
                                                p_NombreEmpresa As String, _
                                                Optional ByRef p_Error As String _
                                                ) As String
        
    
    
    On Error GoTo errores
    
    'Empresa Usuario;Empresa Tramitadora;Juridica Contratación
    If Not m_ObjEntorno.ColCamposSuministradores.Exists(p_TipoEmpresa) Then
        p_Error = "No se reconoce el tipo de empresa"
        Err.Raise 1000
    End If
    If p_NombreEmpresa = "" Then
        p_Error = "No se ha indicado el nombre de la empresa"
        Err.Raise 1000
    End If
    If InStr(1, p_Where, "WHERE") = 0 Then
        p_Error = "No se ha indicado el WHERE"
        Err.Raise 1000
    End If
    p_Where = Replace(p_Where, ";", "")
    p_Where = Trim(p_Where)
    
    getWhereAdicionandoEmpresaYValor = p_Where & " AND " & m_ObjEntorno.ColCamposSuministradores(p_TipoEmpresa) & "='" & p_NombreEmpresa & "';"
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getWhereAdicionandoEmpresaYValor ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function

Public Function ActualizarUsuarioSICA( _
                                        p_ID As String, _
                                        Optional ByRef p_Eliminando As EnumSiNo = EnumSiNo.No, _
                                        Optional ByRef p_Error As String _
                                        ) As String
                                
   
    Dim frm As Form
    
    On Error GoTo errores
    
    p_Error = ""
    
    If p_Eliminando = EnumSiNo.Sí Then
        Call DeleteSicaLocalCaches(p_ID, p_Error:=p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        Call RefreshSicaLocalCaches(p_ID, p_Error:=p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    Set frm = Forms("FormInicial00Principal")
    If frm.SubFormCentral.SourceObject = "Forminicial07UsuariosSICA" Then
        Set frm = frm.SubFormCentral.Form
        frm.Refresh
    End If
    
    
    ActualizarUsuarioSICA = "OK"
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ActualizarUsuarioSICA ha devuelto el error: " & Err.Description
    End If
    

End Function

Public Function AbrirEnLocal( _
                                p_URLFinal As String, _
                                lngHwnd As Long, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    Dim m_NombreArchivo As String
    Dim m_URLCompletaLocal As String
    
    On Error GoTo errores
    
    
    If Not fso.FileExists(p_URLFinal) Then
        p_Error = "No es accesible la ruta del archivo que se pretende abrir" & vbNewLine & p_URLFinal
        Err.Raise 1000
    End If
    m_NombreArchivo = fso.GetFile(p_URLFinal).Name
    m_URLCompletaLocal = m_ObjEntorno.URLDirectorioLocal & m_NombreArchivo
    If fso.FileExists(m_URLCompletaLocal) Then
        If FicheroAbierto(m_URLCompletaLocal) Then
            p_Error = "Tiene el archivo abierto"
            Err.Raise 1000
        End If
    End If
    fso.CopyFile p_URLFinal, m_URLCompletaLocal, True
    Ejecutar lngHwnd, "open", m_URLCompletaLocal, "", "", 1
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        If Err.Number = 70 Then
            p_Error = "No se puede completar la operación, pruebe a cerrar el documento que tiene abierto"
        Else
            p_Error = "La función AbrirEnLocal ha dado el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
        End If
        
    End If
    
End Function

Public Function getFechaHPSConcesionMinima(p_ID As String, Optional ByRef p_Error As String) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    
    On Error GoTo errores
    ' CORREGIDO: Buscamos F_Concesion y ordenamos ascendente (la más antigua primero)
    m_SQL = "SELECT TbHPS.F_Concesion " & _
            "FROM TbHPS " & _
            "WHERE ((Not (TbHPS.F_Concesion) Is Null) And ((TbHPS.IDUsuario) =" & p_ID & ")) " & _
            "ORDER BY TbHPS.F_Concesion ASC;"
            
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    If Not rcdDatos.EOF Then
        getFechaHPSConcesionMinima = Nz(rcdDatos.Fields("F_Concesion"), "")
    End If
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getFechaHPSConcesionMinima ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getDatosLocalDeUsuario( _
                                        Optional p_ID As String, _
                                        Optional p_UsuarioHPS As UsuarioHPS, _
                                        Optional ByRef p_Error As String _
                                        ) As DatosLocal

    Dim m_DatosLocal As DatosLocal
    On Error GoTo errores
    If p_UsuarioHPS Is Nothing Then
        Set p_UsuarioHPS = Constructor.getUsuarioHPS(p_IDUsuario:=p_ID, p_Error:=p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If p_UsuarioHPS Is Nothing Then
            Exit Function
        End If
    End If
    'If p_ID = "338" Then Stop
    Avance "Obteniendo datos ... " & p_ID

    Set m_DatosLocal = New DatosLocal
    With m_DatosLocal
        .ID = p_UsuarioHPS.ID
        .DNI = p_UsuarioHPS.DNI
        .Nombre = p_UsuarioHPS.Nombre
        .Apellido_1 = p_UsuarioHPS.Apellido_1
        .Apellido_2 = p_UsuarioHPS.Apellido_2
        .Telefono = p_UsuarioHPS.Telefono
        If Not p_UsuarioHPS.EmpresaUsuario Is Nothing Then
            .EmpresaUsuario = p_UsuarioHPS.EmpresaUsuario.Nombre
        End If
        If Not p_UsuarioHPS.EmpresaHPS Is Nothing Then
            .EmpresaTramitadora = p_UsuarioHPS.EmpresaHPS.Nombre
        End If
        If p_UsuarioHPS.CadenaContratistas <> "" Then
            .CadenaContratistas = p_UsuarioHPS.CadenaContratistas
        End If
        If Not p_UsuarioHPS.Expediente Is Nothing Then
            
            .CodExp = .CodExp
        End If
        .Correo_e = p_UsuarioHPS.Correo_e
        .F_Nacimiento = p_UsuarioHPS.F_Nacimiento
        .LugarNacimiento = p_UsuarioHPS.LugarNacimiento
        .Motivo_HPS = p_UsuarioHPS.Motivo_HPS
        .IDExpediente = p_UsuarioHPS.IDExpediente
        .F_Curso = p_UsuarioHPS.F_Curso
        .CursoEnVigor = p_UsuarioHPS.CursoEnVigor
        .F_Baja = p_UsuarioHPS.F_Baja
        .FAvisoConcesion = p_UsuarioHPS.FAvisoConcesion
        .Requiere_CorreoJefeSeguridadCurso = p_UsuarioHPS.RequiereComunicacionConcesionCalculadoTexto
        .FechaHPSConcesionMinima = p_UsuarioHPS.FechaHPSConcesionMinimaCalculada
        .Curso_Realizado = p_UsuarioHPS.Curso_Realizado
        .Requiere_Curso = p_UsuarioHPS.RequiereCursoCalculadoTexto
        .FechaPrimeraConvocatoria = p_UsuarioHPS.FechaPrimeraConvocatoria
        .FechaSegundaConvocatoria = p_UsuarioHPS.FechaSegundaConvocatoria
        .FechaCorreoNoCurso = p_UsuarioHPS.FechaCorreoNoCurso
        .Requiere_PrimeraConvocatoriaCurso = p_UsuarioHPS.Requiere_PrimeraConvocatoriaCurso
        .Requiere_SegundaConvocatoriaCurso = p_UsuarioHPS.Requiere_SegundaConvocatoriaCurso
        .Requiere_CorreoJefeSeguridadCurso = p_UsuarioHPS.Requiere_CorreoJefeSeguridadCurso
        .Observaciones = p_UsuarioHPS.Observaciones
        RellenaHPSEnDatosLocal m_DatosLocal, , p_UsuarioHPS, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    Set getDatosLocalDeUsuario = m_DatosLocal
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getDatosLocalDeUsuario ha devuelto el error: " & Err.Description
    End If
End Function

Public Function getColDatosLocalesDeUsuarios( _
                                            Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary
    Dim m_Col As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_UsuarioHPS As UsuarioHPS
    Dim m_DatosLocal As DatosLocal
    On Error GoTo errores
    Set m_Col = m_ObjEntorno.UsuariosHPS
    p_Error = m_ObjEntorno.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Exit Function
    End If
    For Each m_ID In m_Col
        Set m_UsuarioHPS = m_Col(m_ID)
        Set m_DatosLocal = getDatosLocalDeUsuario(p_UsuarioHPS:=m_UsuarioHPS, p_Error:=p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If getColDatosLocalesDeUsuarios Is Nothing Then
            Set getColDatosLocalesDeUsuarios = New Scripting.Dictionary
            getColDatosLocalesDeUsuarios.CompareMode = TextCompare
        End If
        If Not getColDatosLocalesDeUsuarios.Exists(m_DatosLocal.ID) Then
            getColDatosLocalesDeUsuarios.Add m_DatosLocal.ID, m_DatosLocal
        End If
        Set m_DatosLocal = Nothing
        Set m_UsuarioHPS = Nothing
    Next
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getColDatosLocalesDeUsuarios ha devuelto el error: " & Err.Description
    End If
End Function









Public Function ActualizaUsuariosSICALocal( _
                                            Optional ByRef p_Error As String _
                                            ) As String
                                
    Dim rcdUsuariosSICALocal As DAO.Recordset
    Dim rcdUsuariosSICA As DAO.Recordset
    Dim dbsOrigen As DAO.Database
    Dim dbsDestino As DAO.Database
    
    Dim fld As DAO.Field
    
    On Error GoTo errores
    
    p_Error = ""
    
    
    
    Set dbsOrigen = getdb()
    Set dbsDestino = CurrentDb()
    
    
    m_SQL = "DELETE TbUsuariosSICALocal.* " & _
            "FROM TbUsuariosSICALocal ;"
    CurrentDb().Execute (m_SQL)
    
    m_SQL = "SELECT TbUsuariosSICA.*, TbSuministradores.Nombre AS EmpresaTramitadora " & _
            "FROM TbSuministradores RIGHT JOIN TbUsuariosSICA ON TbSuministradores.IDSuministrador = TbUsuariosSICA.IDEmpresaTramitadora;"
    Set rcdUsuariosSICA = dbsOrigen.OpenRecordset(m_SQL)
    If rcdUsuariosSICA.EOF Then
        rcdUsuariosSICA.Close
        Set rcdUsuariosSICA = Nothing
        dbsOrigen.Close
        Set dbsOrigen = Nothing
       
        Exit Function
    End If
    
    
    m_SQL = "SELECT TbUsuariosSICALocal.* " & _
            "FROM TbUsuariosSICALocal ;"
    Set rcdUsuariosSICALocal = dbsDestino.OpenRecordset(m_SQL)
    rcdUsuariosSICA.MoveFirst
    Do While Not rcdUsuariosSICA.EOF
        rcdUsuariosSICALocal.AddNew
            For Each fld In rcdUsuariosSICALocal.Fields
                If fld.Name = "IdEmpresaTramitadora" Then GoTo siguiente
                rcdUsuariosSICALocal(fld.Name).value = rcdUsuariosSICA(fld.Name).value
siguiente:
            Next
        rcdUsuariosSICALocal.Update
        rcdUsuariosSICA.MoveNext
    Loop
    rcdUsuariosSICALocal.Close
    Set rcdUsuariosSICALocal = Nothing
    rcdUsuariosSICA.Close
     Set rcdUsuariosSICA = Nothing
     dbsOrigen.Close
     Set dbsOrigen = Nothing
     
    m_SQL = "DELETE * FROM TbUsuariosSICALocalParaIndicadores;"
    CurrentDb.Execute m_SQL
    
    m_SQL = "INSERT INTO TbUsuariosSICALocalParaIndicadores " & _
            "SELECT TbUsuariosSICALocal.* " & _
            "FROM TbUsuariosSICALocal;"
    CurrentDb.Execute m_SQL


    ActualizaUsuariosSICALocal = "OK"
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ActualizaUsuariosSICALocal ha devuelto el error: " & Err.Description
    End If
    
    
End Function

Public Function ActualizaUsuarioSICALocal( _
                                            p_ID As String, _
                                            Optional ByRef p_Error As String _
                                            ) As String
                                
    Dim rcdUsuariosSICALocal As DAO.Recordset
    Dim rcdUsuariosSICA As DAO.Recordset
    Dim dbsOrigen As DAO.Database
    Dim dbsDestino As DAO.Database
    
    Dim fld As DAO.Field
    
    On Error GoTo errores
    
    p_Error = ""
    
   
    
    Set dbsOrigen = getdb()
    Set dbsDestino = CurrentDb()
    
    
    m_SQL = "DELETE * " & _
            "FROM TbUsuariosSICALocal " & _
            "WHERE ID='" & p_ID & "';"
    CurrentDb().Execute (m_SQL)
     
    m_SQL = "SELECT TbUsuariosSICA.*, TbSuministradores.Nombre AS EmpresaTramitadora " & _
            "FROM TbSuministradores RIGHT JOIN TbUsuariosSICA ON TbSuministradores.IDSuministrador = TbUsuariosSICA.IDEmpresaTramitadora " & _
            "WHERE ID='" & p_ID & "';"

    Set rcdUsuariosSICA = dbsOrigen.OpenRecordset(m_SQL)
    If rcdUsuariosSICA.EOF Then
        rcdUsuariosSICA.Close
        Set rcdUsuariosSICA = Nothing
        dbsOrigen.Close
        Set dbsOrigen = Nothing
        
        Exit Function
    End If
    
    
    m_SQL = "SELECT TbUsuariosSICALocal.* " & _
            "FROM TbUsuariosSICALocal " & _
            "WHERE ID='" & p_ID & "';"
    Set rcdUsuariosSICALocal = dbsDestino.OpenRecordset(m_SQL)
    rcdUsuariosSICA.MoveFirst
    Do While Not rcdUsuariosSICA.EOF
        rcdUsuariosSICALocal.AddNew
            For Each fld In rcdUsuariosSICALocal.Fields
                If fld.Name = "IDEmpresaTramitadora" Then GoTo siguiente
                rcdUsuariosSICALocal(fld.Name).value = rcdUsuariosSICA(fld.Name).value
siguiente:
            Next
        rcdUsuariosSICALocal.Update
        rcdUsuariosSICA.MoveNext
    Loop
    rcdUsuariosSICALocal.Close
    Set rcdUsuariosSICALocal = Nothing
    rcdUsuariosSICA.Close
     Set rcdUsuariosSICA = Nothing
     dbsOrigen.Close
     Set dbsOrigen = Nothing
     
    m_SQL = "DELETE * FROM TbUsuariosSICALocalParaIndicadores " & _
            "WHERE ID='" & p_ID & "';"
    CurrentDb.Execute m_SQL
    
    m_SQL = "INSERT INTO TbUsuariosSICALocalParaIndicadores " & _
            "SELECT TbUsuariosSICALocal.* " & _
            "FROM TbUsuariosSICALocal " & _
            "WHERE ID='" & p_ID & "';"
    CurrentDb.Execute m_SQL


    ActualizaUsuarioSICALocal = "OK"
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ActualizaUsuarioSICALocal ha devuelto el error: " & Err.Description
    End If
    
    
    
End Function
Public Function RellenarEmpresasLocal( _
                                            Optional ByRef p_Error As String _
                                            ) As String
                                
    Dim rcdEmpresasLocal As DAO.Recordset
    Dim rcdEmpresas As DAO.Recordset
    Dim dbsOrigen As DAO.Database
    Dim dbsDestino As DAO.Database
    
    Dim fld As Object
    
    On Error GoTo errores
    
    p_Error = ""
    
    
    
    Set dbsOrigen = getdb()
    Set dbsDestino = CurrentDb()
    
    
    m_SQL = "DELETE TbSuministradoresLocal.* " & _
            "FROM TbSuministradoresLocal ;"
    DoCmd.SetWarnings False
    DoCmd.RunSQL m_SQL
    DoCmd.SetWarnings True
    
    m_SQL = "SELECT TbSuministradores.* " & _
            "FROM TbSuministradores ;"
    Set rcdEmpresas = dbsOrigen.OpenRecordset(m_SQL)
    If rcdEmpresas.EOF Then
        rcdEmpresas.Close
        Set rcdEmpresas = Nothing
        dbsOrigen.Close
        Set dbsOrigen = Nothing
       
        Exit Function
    End If
    
    
    m_SQL = "SELECT TbSuministradoresLocal.* " & _
            "FROM TbSuministradoresLocal ;"
    Set rcdEmpresasLocal = dbsDestino.OpenRecordset(m_SQL)
    rcdEmpresas.MoveFirst
    Do While Not rcdEmpresas.EOF
        rcdEmpresasLocal.AddNew
            For Each fld In rcdEmpresasLocal.Fields
                rcdEmpresasLocal(fld.Name).value = rcdEmpresas(fld.Name).value
            Next
        rcdEmpresasLocal.Update
        rcdEmpresas.MoveNext
    Loop
    rcdEmpresasLocal.Close
    Set rcdEmpresasLocal = Nothing
    rcdEmpresas.Close
     Set rcdEmpresas = Nothing
     dbsOrigen.Close
     Set dbsOrigen = Nothing
     
     Exit Function
    RellenarEmpresasLocal = "OK"
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarEmpresasLocal ha devuelto el error: " & Err.Description
    End If
    
    
End Function
Public Function RellenarEmpresaLocal( _
                                        p_IDSuministrador As String, _
                                        Optional ByRef p_Error As String _
                                        ) As String
                                
    Dim rcdSuministradoresLocal As DAO.Recordset
    Dim rcdSuministradores As DAO.Recordset
    Dim dbsOrigen As DAO.Database
    Dim dbsDestino As DAO.Database
    
    Dim fld As Object
    
    On Error GoTo errores
    
    p_Error = ""
    
   
    
    Set dbsOrigen = getdb()
    Set dbsDestino = CurrentDb()
    
    
    m_SQL = "DELETE TbSuministradoresLocal.* " & _
            "FROM TbSuministradoresLocal " & _
            "WHERE IDSuministrador=" & p_IDSuministrador & ";"
    CurrentDb().Execute m_SQL
   
    
    m_SQL = "SELECT TbSuministradoresLocal.* " & _
            "FROM TbSuministradoresLocal " & _
            "WHERE IDSuministrador=" & p_IDSuministrador & ";"
    Set rcdSuministradores = dbsOrigen.OpenRecordset(m_SQL)
    If rcdSuministradores.EOF Then
        rcdSuministradores.Close
        Set rcdSuministradores = Nothing
        dbsOrigen.Close
        Set dbsOrigen = Nothing
        
        Exit Function
    End If
    
    
    m_SQL = "SELECT TbSuministradoresLocal.* " & _
            "FROM TbSuministradoresLocal " & _
            "WHERE IDSuministrador=" & p_IDSuministrador & ";"
    Set rcdSuministradoresLocal = dbsDestino.OpenRecordset(m_SQL)
    rcdSuministradores.MoveFirst
    Do While Not rcdSuministradores.EOF
        rcdSuministradoresLocal.AddNew
            For Each fld In rcdSuministradoresLocal.Fields
                rcdSuministradoresLocal(fld.Name).value = rcdSuministradores(fld.Name).value
            Next
        rcdSuministradoresLocal.Update
        rcdSuministradores.MoveNext
    Loop
    rcdSuministradoresLocal.Close
    Set rcdSuministradoresLocal = Nothing
    rcdSuministradores.Close
     Set rcdSuministradores = Nothing
     dbsOrigen.Close
     Set dbsOrigen = Nothing
     
     Exit Function
    RellenarEmpresaLocal = "OK"
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarEmpresaLocal ha devuelto el error: " & Err.Description
    End If
   
    
End Function
Public Function RellenarUsuariosHistoricosLocal( _
                                            Optional ByRef p_Error As String _
                                            ) As String
                                
    Dim rcdUsuariosHistoricosLocal As DAO.Recordset
    Dim rcdUsuariosHistoricos As DAO.Recordset
    Dim dbsOrigen As DAO.Database
    Dim dbsDestino As DAO.Database
    
    Dim fld As DAO.Field
    
    On Error GoTo errores
    
    p_Error = ""
    
   
    
    Set dbsOrigen = getdb()
    Set dbsDestino = CurrentDb()
    
    
    m_SQL = "DELETE TbUsuariosHistoricosLocal.* " & _
            "FROM TbUsuariosHistoricosLocal ;"
    DoCmd.SetWarnings False
    DoCmd.RunSQL m_SQL
    DoCmd.SetWarnings True
    
    m_SQL = "SELECT DISTINCT TbUsuariosHistoricos.*, TbSuministradores.Nombre AS EmpresaUsuario, " & _
            "TbSuministradores_1.Nombre AS EmpresaTramitadora,CadenaContratistas AS JuridicaContrato, TbExpedientes.CodExp " & _
            "FROM ((TbUsuariosHistoricos INNER JOIN TbSuministradores " & _
            "ON TbUsuariosHistoricos.IDEmpresaUsuario = TbSuministradores.IDSuministrador) " & _
            "INNER JOIN TbSuministradores AS TbSuministradores_1 " & _
            "ON TbUsuariosHistoricos.IDEmpresaHPS = TbSuministradores_1.IDSuministrador) " & _
            "INNER JOIN TbExpedientes ON TbUsuariosHistoricos.IDExpediente = TbExpedientes.IDExpediente;"
    Set rcdUsuariosHistoricos = dbsOrigen.OpenRecordset(m_SQL)
    If rcdUsuariosHistoricos.EOF Then
        rcdUsuariosHistoricos.Close
        Set rcdUsuariosHistoricos = Nothing
        dbsOrigen.Close
        Set dbsOrigen = Nothing
        
        Exit Function
    End If
    
    m_SQL = "SELECT TbUsuariosHistoricosLocal.* " & _
            "FROM TbUsuariosHistoricosLocal ;"
    Set rcdUsuariosHistoricosLocal = dbsDestino.OpenRecordset(m_SQL)
    
    rcdUsuariosHistoricos.MoveFirst
    Do While Not rcdUsuariosHistoricos.EOF
        rcdUsuariosHistoricosLocal.AddNew
            For Each fld In rcdUsuariosHistoricosLocal.Fields
                'Debug.Print fld.Name
                rcdUsuariosHistoricosLocal(fld.Name).value = rcdUsuariosHistoricos(fld.Name).value
siguiente:
            Next
        rcdUsuariosHistoricosLocal.Update
        rcdUsuariosHistoricos.MoveNext
    Loop
    
    rcdUsuariosHistoricosLocal.Close
    Set rcdUsuariosHistoricosLocal = Nothing
    rcdUsuariosHistoricos.Close
     Set rcdUsuariosHistoricos = Nothing
     dbsOrigen.Close
     Set dbsOrigen = Nothing
     
     Exit Function
    RellenarUsuariosHistoricosLocal = "OK"
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarUsuariosHistoricosLocal ha devuelto el error: " & Err.Description
    End If
    
   
    
End Function




Public Function ValorRepetido( _
                                p_NombreTabla As String, _
                                p_NombreCampo As String, _
                                p_ValorCampo As String, _
                                Optional ByRef db As DAO.Database, _
                                Optional ByRef p_Error As String _
                                ) As Boolean
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    
    On Error GoTo errores
    
    If p_NombreTabla = "" Then
        p_Error = "Se ha de indicar el Nombre de la tabla"
        Err.Raise 1000
    End If
    If p_NombreCampo = "" Then
        p_Error = "Se ha de indicar el Nombre del campo"
        Err.Raise 1000
    End If
    If p_ValorCampo = "" Then
        p_Error = "Se ha de indicar el Valor del campo"
        Err.Raise 1000
    End If
    If db Is Nothing Then
        Set db = getdb()
    End If
    
    m_SQL = "SELECT " & p_NombreTabla & "." & p_NombreCampo & " " & _
            "FROM " & p_NombreTabla & " " & _
            "WHERE " & p_NombreCampo & "='" & p_ValorCampo & "';"
    On Error Resume Next
    Set rcdDatos = db.OpenRecordset(m_SQL)
    If Err.Number <> 0 Then
        m_SQL = "SELECT " & p_NombreTabla & "." & p_NombreCampo & " " & _
            "FROM " & p_NombreTabla & " " & _
            "WHERE " & p_NombreCampo & "=" & p_ValorCampo & ";"
        Err.Clear
        Set rcdDatos = db.OpenRecordset(m_SQL)
        If Err.Number <> 0 Then
            Err.Raise 1000
        End If
    End If
    On Error GoTo errores
    With rcdDatos
        If Not .EOF Then
            ValorRepetido = True
        Else
            ValorRepetido = False
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ValorRepetido ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Private Function getFBajaHPS( _
                            p_ID As String, _
                            Optional ByRef p_Error As String _
                            ) As String
    
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_getFBajaHPS As String
    
    On Error GoTo errores
    If Not IsNumeric(p_ID) Then
        p_Error = "No se ha indicado un id usuario numérico válido"
        Err.Raise 1000
    End If
   
     m_SQL = "SELECT TbHPS.F_Baja " & _
            "FROM TbHPS " & _
            "WHERE ((Not (TbHPS.F_Baja) Is Null) And ((TbHPS.IDUsuario) =" & p_ID & ")) " & _
            "ORDER BY TbHPS.F_Baja;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            m_getFBajaHPS = .Fields("F_Baja")
        End If
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    If Not IsDate(m_getFBajaHPS) Then
        m_SQL = "SELECT TbHPS.F_Caducidad " & _
                "FROM TbHPS " & _
                "WHERE (((TbHPS.Renovacion)='No') AND ((TbHPS.IDUsuario)=" & p_ID & ") AND ((TbHPS.F_Caducidad)<Now()));"
        Set rcdDatos = getdb().OpenRecordset(m_SQL)
        With rcdDatos
            If Not .EOF Then
                m_getFBajaHPS = .Fields("F_Caducidad")
            End If
            
        End With
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getFBajaHPS ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function RellenarObjetoConDatos(p_Objeto As Object, frm As Form, Optional ByRef p_Error As String) As String
    
    
    Dim ctl As Control
    On Error GoTo errores
    
    
    For Each ctl In frm.Controls
        If InStr(1, ctl.Tag, "DATOS") <> 0 Then
            p_Objeto.SetPropiedad ctl.Name, Nz(ctl.value, ""), p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        
    Next
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarObjetoConDatos ha producido el error: " & Err.Description
    End If
End Function

Public Function HaHabidoCambios(p_ObjetoAlInicio As Object, frm As Form, Optional ByRef p_Error As String) As Boolean
    
    Dim m_DatoInicial As Variant
    Dim m_DatoFinal As String
    Dim ctl As Control
    Dim intResultado As Integer
    On Error GoTo errores
    
        
    For Each ctl In frm.Controls
        'If ctl.Name = "Observaciones" Then Stop
        If InStr(1, ctl.Tag, "DATOS") <> 0 Then
            m_DatoInicial = p_ObjetoAlInicio.getPropiedad(ctl.Name, p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            
            m_DatoFinal = Nz(ctl.value, "")
            intResultado = StrComp(m_DatoInicial, m_DatoFinal, vbBinaryCompare)
            If intResultado = 1 Or intResultado = -1 Then
                HaHabidoCambios = True
                Exit Function
            End If
        End If
        
    Next
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HaHabidoCambios ha producido el error: " & Err.Description
    End If

End Function
Public Function DatosDelObjetoAlFormulario(p_Objeto As Object, frm As Form, Optional ByRef p_Error As String) As String
    
    Dim m_Valor As String
    Dim ctl As Control
    On Error GoTo errores
    
        
    For Each ctl In frm.Controls
        If InStr(1, ctl.Tag, "DATOS") <> 0 Then
            'Debug.Print ctl.Name
            
            m_Valor = p_Objeto.getPropiedad(ctl.Name, p_Error)
            If p_Error <> "" Then
                m_Valor = ""
            End If
            If m_Valor <> "" Then
                ctl.value = m_Valor
            End If
            
        End If
        
    Next
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método DatosDelObjetoAlFormulario ha producido el error: " & Err.Description
    End If

End Function
Public Function getNombreUsuarioConectado(Optional ByRef p_Error As String) As String
    
    Dim m_UsuarioMaquina As Usuario
    
    On Error GoTo errores
    
    If Not m_ObjUsuarioConectado Is Nothing Then
        getNombreUsuarioConectado = m_ObjUsuarioConectado.Nombre
        Exit Function
    End If
    
    Set m_UsuarioMaquina = Constructor.getUsuarioConectadoPorMaquina(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_UsuarioMaquina Is Nothing Then
        getNombreUsuarioConectado = "Desconocido"
        Exit Function
    End If
    getNombreUsuarioConectado = m_UsuarioMaquina.Nombre
    Exit Function
errores:
    getNombreUsuarioConectado = "Desconocido"
End Function

Public Function MeterUnMotivoEnListaMOtivos( _
                                            ByRef p_Motivo As String, _
                                            Optional ByRef p_Error As String _
                                            ) As String

    
    Dim m_SQL As String
    Dim rcdDatos As DAO.Recordset
    
    On Error GoTo errores
    If p_Motivo = "" Then
        Exit Function
    End If
    p_Motivo = Trim(p_Motivo)
    
    m_SQL = "SELECT * " & _
            "FROM TbMotivoHPS " & _
            "WHERE Motivo_HPS='" & p_Motivo & "';"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            .AddNew
                .Fields("Motivo_HPS") = p_Motivo
            .Update
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing

    Exit Function

errores:
    If Err.Number <> 1000 Then
        p_Error = "El método MeterUnMotivoEnListaMOtivos ha devuelto el error: " & Err.Description
    End If

End Function

Private Function getUsuarioMaquina( _
                            Optional ByRef p_Error As String _
                            ) As String
    Dim objNetwork As Object
    On Error GoTo errores
    Set objNetwork = CreateObject("Wscript.Network")
    With objNetwork
        getUsuarioMaquina = .UserName & "|" & .computername
    End With
   
    Set objNetwork = Nothing
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getUsuarioMaquina ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function getMaquina( _
                            Optional ByRef p_Error As String _
                            ) As String
    Dim flag As String
    Dim dato As Variant
    On Error GoTo errores
    flag = getUsuarioMaquina
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        getMaquina = dato(1)
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getMaquina ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function getUsuariodeMaquina( _
                                Optional ByRef p_Error As String _
                                ) As String
    Dim flag As String
    Dim dato As Variant
    On Error GoTo errores
    flag = getUsuarioMaquina
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        getUsuariodeMaquina = dato(0)
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getUsuariodeMaquina ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function


Public Function getHora() As String
    Dim xmlHttp As Object
    Dim url As String
    Dim response As String
    Dim json As Object
    Dim hora As String
    Dim dato As Variant
    Dim flag As String
    
    ' URL de la API para obtener la hora en Madrid, España
    url = "http://worldtimeapi.org/api/timezone/Europe/Madrid"
    
    ' Crear el objeto XMLHTTP
    Set xmlHttp = CreateObject("MSXML2.XMLHTTP")
    
    ' Hacer la solicitud a la API
    xmlHttp.Open "GET", url, False
    xmlHttp.send
    
    ' Obtener la respuesta
    response = xmlHttp.responseText
    
    ' Analizar la respuesta JSON
    Set json = JsonConverter.ParseJson(response)
    
    ' Extraer la hora de la respuesta JSON
    hora = json("datetime")
    dato = Split(hora, "T")
    flag = dato(1)
    dato = Split(flag, ".")
    hora = dato(0)
    ' Retornar la hora
    getHora = hora
End Function
Public Function Avance( _
                        p_Linea As Variant, _
                        Optional ByRef p_Error As String _
                        ) As String
    
    
    Dim m_Form As Form
    
    On Error Resume Next
    If lbl Is Nothing Then
        Set m_Form = Screen.ActiveForm
        Set lbl = m_Form.lblEstado
        If Err.Number <> 0 Then
            Err.Clear
            Exit Function
        End If
   
    End If
    On Error GoTo errores
    
    If lbl.Visible = False Then
        lbl.Visible = True
    End If
    VBA.DoEvents
    lbl.Caption = p_Linea
    VBA.DoEvents
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Avance ha devuelto el error: " & Err.Description
    End If
End Function
Public Sub AvanceNuevo(ByVal texto As String)
    ' Esta función actualiza el Label del formulario de inicio si está abierto
    
    On Error Resume Next ' Evitamos errores si el formulario no está abierto
    
    If CurrentProject.AllForms("FormInicial").IsLoaded Then
        ' 1. Actualizamos el texto
        Forms("FormInicial").lblEstado.Caption = texto
        
        ' 2. ESTO ES LO MÁS IMPORTANTE:
        ' DoEvents obliga a Windows a repintar la pantalla inmediatamente.
        ' Sin esto, el label no cambiaría hasta acabar todo el proceso.
        DoEvents
    End If
    
    ' Opcional: También actualizar la barra de estado de Access (abajo a la izquierda)
    SysCmd acSysCmdSetStatus, texto
End Sub
Public Function AvanceCerrar() As String
    
    
    
    
    On Error Resume Next
    Application.Screen.ActiveForm.Controls("lblEstado").Visible = False
    
    Exit Function
    
End Function


Private Function getDirectorioOneDrive(Optional ByRef p_Error As String) As String
    Dim fso As Object
    Dim carpetaRaiz As Object
    Dim subCarpeta As Object
    Dim rutaEncontrada As String
    Dim encontrado As Boolean
    Dim dato As Variant
    On Error GoTo errores
    
    rutaEncontrada = Environ("OneDrive")
    If InStr(1, rutaEncontrada, "OneDrive") <> 0 Then
        dato = Split(rutaEncontrada, "OneDrive")
        getDirectorioOneDrive = dato(0) & "OneDrive"
        Exit Function
    End If
    
    
    ' Crear objeto FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Obtener la carpeta raíz de C:\
    Set carpetaRaiz = fso.GetFolder("C:\")
    
    ' Inicializar variables
    encontrado = False
    rutaEncontrada = ""
    
    ' Recorrer las subcarpetas en la raíz de C:\
    For Each subCarpeta In carpetaRaiz.SubFolders
        If InStr(1, subCarpeta.Name, "OneDrive", vbTextCompare) > 0 Then
            rutaEncontrada = subCarpeta.path
            encontrado = True
            Exit For
        
        End If
    Next subCarpeta
    
    ' Mostrar el resultado
    If encontrado Then
        getDirectorioOneDrive = rutaEncontrada
   
    End If
    
    ' Liberar objetos
    Set subCarpeta = Nothing
    Set carpetaRaiz = Nothing
    Set fso = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getDirectorioOneDrive ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Debug.Print p_Error
End Function
Private Function getDirectorioOneDriveTelefonicaApps(Optional ByRef p_Error As String) As String
    Dim fso As Object
    Dim carpeta As String
    Dim m_RutaOneDrive As String
    
    On Error GoTo errores
    ' Crear objeto FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    m_RutaOneDrive = getDirectorioOneDrive(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_RutaOneDrive = "" Then
        Exit Function
    End If
    carpeta = m_RutaOneDrive & "\Telefonica\Aplicaciones_dys.TMETF - Aplicaciones PpD\"
    If Not fso.FolderExists(carpeta) Then
        Exit Function
    End If
    getDirectorioOneDriveTelefonicaApps = carpeta
    
    ' Liberar objetos
   
    
    Set fso = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getDirectorioOneDriveTelefonicaApps ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Debug.Print p_Error
End Function
Private Function getDirectorioOneDriveApps(Optional ByRef p_Error As String) As String
    Dim fso As Object
    Dim carpeta As String
    Dim m_RutaOneDrive As String
    
    On Error GoTo errores
    ' Crear objeto FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    m_RutaOneDrive = getDirectorioOneDrive(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_RutaOneDrive = "" Then
        Exit Function
    End If
    'C:\OneDrive\OneDrive - Telefonica\00LABORAL\Aplicaciones PpD
    carpeta = m_RutaOneDrive & "\OneDrive - Telefonica\00LABORAL\Aplicaciones PpD\"
    If Not fso.FolderExists(carpeta) Then
        Exit Function
    End If
    getDirectorioOneDriveApps = carpeta
    
    ' Liberar objetos
   
    
    Set fso = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getDirectorioOneDriveApps ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Debug.Print p_Error
End Function


Public Function getRutaAplicacionesLocal(Optional ByRef p_Error As String) As String
    Dim fso As Object
    Dim m_RutaOneDrive As String
    Dim m_RutaOneDriveTelefonica As String
    
    On Error GoTo errores
    ' Crear objeto FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    m_RutaOneDriveTelefonica = getDirectorioOneDriveTelefonicaApps(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    If fso.FolderExists(m_RutaOneDriveTelefonica) Then
        Set fso = Nothing
        getRutaAplicacionesLocal = m_RutaOneDriveTelefonica
        Exit Function
    End If
    
    m_RutaOneDrive = getDirectorioOneDrive(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_RutaOneDrive = "" Then
        Set fso = Nothing
        Exit Function
    End If
   
    getRutaAplicacionesLocal = m_RutaOneDrive
    
    ' Liberar objetos
   
    
    Set fso = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getRutaAplicacionesLocal ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Debug.Print p_Error
End Function



Public Function CopiarDatosIndicadorUsuariosTotales( _
                                                    Optional ByRef p_Error As String _
                                                    ) As String
    
    Dim rcdDatosDatos As DAO.Recordset
    Dim rcdDatosUI As DAO.Recordset
    Dim fld As DAO.Field
    Dim m_SQL As String
    
    On Error GoTo errores
    
    m_SQL = "DELETE * FROM TbDatosLocalParaIndicadores;"
    CurrentDb().Execute m_SQL
    
    m_SQL = "TbUsuariosEntidades"
    Set rcdDatosDatos = getdb().OpenRecordset(m_SQL)
    If rcdDatosDatos.EOF Then
        rcdDatosDatos.Close
        Set rcdDatosDatos = Nothing
        Exit Function
    End If
    m_SQL = "TbDatosLocalParaIndicadores"
    Set rcdDatosUI = CurrentDb().OpenRecordset(m_SQL)
    rcdDatosDatos.MoveFirst
    Do While Not rcdDatosDatos.EOF
        rcdDatosUI.AddNew
            For Each fld In rcdDatosDatos.Fields
                'Debug.Print fld.Name
                rcdDatosUI.Fields(fld.Name).value = rcdDatosDatos.Fields(fld.Name).value
            Next
        rcdDatosUI.Update
        VBA.DoEvents
        
        rcdDatosDatos.MoveNext
    Loop
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CopiarDatosIndicadorUsuariosTotales ha devuelto el error: " & Err.Description
    End If
    Debug.Print p_Error
End Function

Public Function RellenaTodosLosDatosEntidadConTbUsuarioHPS( _
                                                        Optional ByRef p_Error As String _
                                                        ) As String
        
    'va a tomar los datos remotos de TbUsuarios y HPS y regeneran la tabla remota de TbUsuariosEntidades que es la misma que
    ' la tabla local de TbDatosLocal y TbDatosLocalParaIndicadores
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Valor As String
    Dim m_DatosLocal As DatosLocal
    Dim m_Col As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_UsuarioHPS As UsuarioHPS
    
    
    On Error GoTo errores
    
    If m_ObjEntorno Is Nothing Then
        EVE
    End If
    m_SQL = "DELETE * FROM TbUsuariosEntidades;"
    getdb().Execute m_SQL
    Set m_Col = m_ObjEntorno.UsuariosHPS
    If m_Col Is Nothing Then
        Exit Function
    End If
    m_SQL = "TbUsuariosEntidades"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    For Each m_ID In m_Col
        Set m_UsuarioHPS = m_Col(m_ID)
        Avance "Actualizando Datos Locales ..." & m_UsuarioHPS.NombreCompleto
        VBA.DoEvents
        Debug.Print m_UsuarioHPS.NombreCompleto
        VBA.DoEvents
        Set m_DatosLocal = getDatosLocalDeUsuario(p_UsuarioHPS:=m_UsuarioHPS, p_Error:=p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        rcdDatos.AddNew
            For Each m_Campo In m_DatosLocal.ColCampos
                
                'Debug.Print m_Campo
                
                m_Valor = m_DatosLocal.getPropiedad(m_Campo, p_Error)
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
                If m_Valor <> "" Then
                    rcdDatos.Fields(m_Campo).value = m_Valor
                End If
            Next
        rcdDatos.Update
        Set m_UsuarioHPS = Nothing
    Next
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenaTodosLosDatosEntidadConTbUsuarioHPS ha devuelto el error: " & Err.Description
    End If
    Debug.Print p_Error
End Function
Public Function RegenerarDatosEntidadFaltantes( _
                                                Optional ByRef p_Error As String _
                                                ) As String
        
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Campo As Variant
    Dim m_Valor As String
    Dim m_DatosLocal As DatosLocal
    Dim m_Col As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_UsuarioHPS As UsuarioHPS
    
    
    On Error GoTo errores
    
    
    
    If m_ObjEntorno Is Nothing Then
        EVE
    End If
    
    Set m_Col = m_ObjEntorno.UsuariosNoEnEntidad
    If m_Col Is Nothing Then
        Exit Function
    End If
    m_SQL = "TbUsuariosEntidades"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    For Each m_ID In m_Col
        
        
        Set m_DatosLocal = getDatosLocalDeUsuario(p_ID:=CStr(m_ID), p_Error:=p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If m_DatosLocal Is Nothing Then
            Set m_DatosLocal = New DatosLocal
            rcdDatos.AddNew
                For Each m_Campo In m_DatosLocal.ColCampos
                    
                    'Debug.Print m_Campo
                   ' If CStr(m_Campo) = "HPS_UE_MESES_PARA_RENOVAR" Then Stop
                    m_Valor = m_DatosLocal.getPropiedad(m_Campo, p_Error)
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                    If m_Valor <> "" Then
                        rcdDatos.Fields(m_Campo).value = m_Valor
                    End If
                Next
            rcdDatos.Update
        End If
        
        Set m_UsuarioHPS = Nothing
    Next
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RegenerarDatosEntidadFaltantes ha devuelto el error: " & Err.Description
    End If
    Debug.Print p_Error
End Function

Public Function QuitarDatosEntidadSobrantes( _
                                            Optional ByRef p_Error As String _
                                            ) As String
        
    
    
    Dim m_SQL As String
    
    Dim m_Col As Scripting.Dictionary
    Dim m_ID As Variant
   
    
    
    On Error GoTo errores
    
    
    
    If m_ObjEntorno Is Nothing Then
        EVE
    End If
    
    Set m_Col = m_ObjEntorno.UsuariosEnEntidadNoExistentes
    If m_Col Is Nothing Then
        Exit Function
    End If
    
    
    For Each m_ID In m_Col
        m_SQL = "DELETE * FROM TbUsuariosEntidades WHERE id=" & m_ID & ";"
        getdb().Execute m_SQL
        m_SQL = "DELETE * FROM TbDatosLocal WHERE id=" & m_ID & ";"
        CurrentDb().Execute m_SQL
        
    Next
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método QuitarDatosEntidadSobrantes ha devuelto el error: " & Err.Description
    End If
    Debug.Print p_Error
End Function
Function FormatearNombreCompleto(p_NombreCompleto As String) As String
    ' Capitalizar cada palabra
    Dim NombreCompleto As String
    Dim dato As Variant
    Dim m_Trozo As Variant
    Dim m_TrozoCapitalizado As String
    
    dato = Split(p_NombreCompleto, " ")
    For Each m_Trozo In dato
        m_TrozoCapitalizado = Capitalizar(CStr(m_Trozo))
        If NombreCompleto = "" Then
            NombreCompleto = m_TrozoCapitalizado
        Else
            NombreCompleto = NombreCompleto & " " & m_TrozoCapitalizado
        End If
    Next
    
    
    
    FormatearNombreCompleto = NombreCompleto
End Function

Function Capitalizar(texto As String) As String
    ' Convierte solo la primera letra en mayúscula y el resto en minúscula
    Dim dato As Variant
    Dim m_Texto As Variant
    Dim m_TextoCompleto As String
    dato = Split(texto, " ")
    For Each m_Texto In dato
        If m_TextoCompleto = "" Then
            m_TextoCompleto = UCase(Left(m_Texto, 1)) & LCase(Mid(m_Texto, 2))
        Else
            m_TextoCompleto = m_TextoCompleto & " " & UCase(Left(m_Texto, 1)) & LCase(Mid(m_Texto, 2))
        End If
    Next
    Capitalizar = m_TextoCompleto
End Function

Public Function RellenaHPSEnDatosLocal( _
                                        ByRef p_DatosLocal As DatosLocal, _
                                        Optional p_ID As String, _
                                        Optional p_UsuarioHPS As UsuarioHPS, _
                                        Optional ByRef p_Error As String _
                                        ) As String
    
    
   
    Dim m_HPS As HPS
    Dim m_FechaHPSConcesionMinima As String
    
   
    On Error GoTo errores
    
    If p_UsuarioHPS Is Nothing Then
        Set p_UsuarioHPS = Constructor.getUsuarioHPS(p_IDUsuario:=p_ID, p_Error:=p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If p_UsuarioHPS Is Nothing Then
            Exit Function
        End If
    End If
    p_ID = p_UsuarioHPS.ID
    'Nacional
    Set m_HPS = p_UsuarioHPS.HPSNAC
    If Not m_HPS Is Nothing Then
       
        With m_HPS
            p_DatosLocal.HPS_NAC_SIN_DATOS = "No"
            p_DatosLocal.HPS_NAC_F_Concesion = .F_Concesion
            p_DatosLocal.HPS_NAC_F_Caducidad = .F_Caducidad
            p_DatosLocal.HPS_NAC_Grado = .Grado
            p_DatosLocal.HPS_NAC_F_Solicitud = .F_Solicitud
            p_DatosLocal.HPS_NAC_Especialidad = .Especialidad
            p_DatosLocal.HPS_NAC_Renovacion = .Renovacion
            p_DatosLocal.HPS_NAC_F_Baja = .F_Baja
            
            If p_UsuarioHPS.F_BajaCalculada = "" Then
                p_DatosLocal.HPS_NAC_Activo = DatosHPSCalculados("Activo", m_HPS, p_Error)
            Else
                 p_DatosLocal.HPS_NAC_Activo = "No"
            End If
            If p_UsuarioHPS.F_BajaCalculada = "" Then
               p_DatosLocal.HPS_NAC_ApuntoDeCaducar = DatosHPSCalculados("APuntoDeCaducar", m_HPS, p_Error)
            Else
                p_DatosLocal.HPS_NAC_ApuntoDeCaducar = "No"
            End If
            p_DatosLocal.HPS_NAC_Baja = DatosHPSCalculados("Baja", m_HPS, p_Error)
            If p_UsuarioHPS.F_BajaCalculada = "" Then
                p_DatosLocal.HPS_NAC_Caducado = DatosHPSCalculados("Caducado", m_HPS, p_Error)
                   
            Else
                p_DatosLocal.HPS_NAC_Caducado = "No"
            End If
            If p_UsuarioHPS.F_BajaCalculada = "" Then
                p_DatosLocal.HPS_NAC_PendienteRenovacion = DatosHPSCalculados("PendienteRenovacion", m_HPS, p_Error)
            Else
                p_DatosLocal.HPS_NAC_PendienteRenovacion = "No"
            End If
            If p_UsuarioHPS.F_BajaCalculada = "" Then
                p_DatosLocal.HPS_NAC_Solicitado = DatosHPSCalculados("Solicitada", m_HPS, p_Error)
            Else
                p_DatosLocal.HPS_NAC_Solicitado = "No"
            End If
            If p_UsuarioHPS.F_BajaCalculada = "" Then
                p_DatosLocal.HPS_NAC_ESTADO = DatosHPSCalculados("Estado", m_HPS, p_Error)
                p_DatosLocal.HPS_NAC_Irregular = DatosHPSCalculados("Irregular", m_HPS, p_Error)
                If p_DatosLocal.HPS_NAC_Irregular = "Sí" Then
                    p_DatosLocal.HPS_NAC_MOTIVO_IRREGULAR = DatosHPSCalculados("MOTIVO_IRREGULAR", m_HPS, p_Error)
                End If
                p_DatosLocal.HPS_NAC_MESES_PARA_RENOVAR = DatosHPSCalculados("MESES_PARA_RENOVAR", m_HPS, p_Error)
            End If
            
            
        End With
    Else
        p_DatosLocal.HPS_NAC_SIN_DATOS = "Sí"
    End If
    'OTAN
    Set m_HPS = p_UsuarioHPS.HPSOTAN
    If Not m_HPS Is Nothing Then
        p_DatosLocal.HPS_OTAN_SIN_DATOS = "No"
        With m_HPS
            p_DatosLocal.HPS_OTAN_F_Concesion = .F_Concesion
            p_DatosLocal.HPS_OTAN_F_Caducidad = .F_Caducidad
            p_DatosLocal.HPS_OTAN_Grado = .Grado
            p_DatosLocal.HPS_OTAN_F_Solicitud = .F_Solicitud
            p_DatosLocal.HPS_OTAN_Especialidad = .Especialidad
            p_DatosLocal.HPS_OTAN_Renovacion = .Renovacion
            p_DatosLocal.HPS_OTAN_F_Baja = .F_Baja
            
            If p_UsuarioHPS.F_BajaCalculada = "" Then
                p_DatosLocal.HPS_OTAN_Activo = DatosHPSCalculados("Activo", m_HPS, p_Error)
            Else
                 p_DatosLocal.HPS_OTAN_Activo = "No"
            End If
            If p_UsuarioHPS.F_BajaCalculada = "" Then
               p_DatosLocal.HPS_OTAN_ApuntoDeCaducar = DatosHPSCalculados("APuntoDeCaducar", m_HPS, p_Error)
            Else
                p_DatosLocal.HPS_OTAN_ApuntoDeCaducar = "No"
            End If
            p_DatosLocal.HPS_OTAN_Baja = DatosHPSCalculados("Baja", m_HPS, p_Error)
            If p_UsuarioHPS.F_BajaCalculada = "" Then
                p_DatosLocal.HPS_OTAN_Caducado = DatosHPSCalculados("Caducado", m_HPS, p_Error)
                   
            Else
                p_DatosLocal.HPS_OTAN_Caducado = "No"
            End If
            If p_UsuarioHPS.F_BajaCalculada = "" Then
                p_DatosLocal.HPS_OTAN_PendienteRenovacion = DatosHPSCalculados("PendienteRenovacion", m_HPS, p_Error)
            Else
                p_DatosLocal.HPS_OTAN_PendienteRenovacion = "No"
            End If
            If p_UsuarioHPS.F_BajaCalculada = "" Then
                p_DatosLocal.HPS_OTAN_Solicitado = DatosHPSCalculados("Solicitada", m_HPS, p_Error)
            Else
                p_DatosLocal.HPS_OTAN_Solicitado = "No"
            End If
            If p_UsuarioHPS.F_BajaCalculada = "" Then
                p_DatosLocal.HPS_OTAN_ESTADO = DatosHPSCalculados("Estado", m_HPS, p_Error)
                p_DatosLocal.HPS_OTAN_Irregular = DatosHPSCalculados("Irregular", m_HPS, p_Error)
                If p_DatosLocal.HPS_OTAN_Irregular = "Sí" Then
                    p_DatosLocal.HPS_OTAN_MOTIVO_IRREGULAR = DatosHPSCalculados("MOTIVO_IRREGULAR", m_HPS, p_Error)
                End If
                p_DatosLocal.HPS_OTAN_MESES_PARA_RENOVAR = DatosHPSCalculados("MESES_PARA_RENOVAR", m_HPS, p_Error)
            End If
            
            
        End With
    Else
        p_DatosLocal.HPS_OTAN_SIN_DATOS = "Sí"
    End If
    'ESA
    Set m_HPS = p_UsuarioHPS.HPSESA
    If Not m_HPS Is Nothing Then
        With m_HPS
            p_DatosLocal.HPS_ESA_SIN_DATOS = "No"
            p_DatosLocal.HPS_ESA_F_Concesion = .F_Concesion
            p_DatosLocal.HPS_ESA_F_Caducidad = .F_Caducidad
            p_DatosLocal.HPS_ESA_Grado = .Grado
            p_DatosLocal.HPS_ESA_F_Solicitud = .F_Solicitud
            p_DatosLocal.HPS_ESA_Especialidad = .Especialidad
            p_DatosLocal.HPS_ESA_Renovacion = .Renovacion
            p_DatosLocal.HPS_ESA_F_Baja = .F_Baja
            
            If p_UsuarioHPS.F_BajaCalculada = "" Then
                p_DatosLocal.HPS_ESA_Activo = DatosHPSCalculados("Activo", m_HPS, p_Error)
            Else
                 p_DatosLocal.HPS_ESA_Activo = "No"
            End If
            If p_UsuarioHPS.F_BajaCalculada = "" Then
               p_DatosLocal.HPS_ESA_ApuntoDeCaducar = DatosHPSCalculados("APuntoDeCaducar", m_HPS, p_Error)
            Else
                p_DatosLocal.HPS_ESA_ApuntoDeCaducar = "No"
            End If
            p_DatosLocal.HPS_ESA_Baja = DatosHPSCalculados("Baja", m_HPS, p_Error)
            If p_UsuarioHPS.F_BajaCalculada = "" Then
                p_DatosLocal.HPS_ESA_Caducado = DatosHPSCalculados("Caducado", m_HPS, p_Error)
                   
            Else
                p_DatosLocal.HPS_ESA_Caducado = "No"
            End If
            If p_UsuarioHPS.F_BajaCalculada = "" Then
                p_DatosLocal.HPS_ESA_PendienteRenovacion = DatosHPSCalculados("PendienteRenovacion", m_HPS, p_Error)
            Else
                p_DatosLocal.HPS_ESA_PendienteRenovacion = "No"
            End If
            If p_UsuarioHPS.F_BajaCalculada = "" Then
                p_DatosLocal.HPS_ESA_Solicitado = DatosHPSCalculados("Solicitada", m_HPS, p_Error)
            Else
                p_DatosLocal.HPS_ESA_Solicitado = "No"
            End If
            If p_UsuarioHPS.F_BajaCalculada = "" Then
                p_DatosLocal.HPS_ESA_ESTADO = DatosHPSCalculados("Estado", m_HPS, p_Error)
                p_DatosLocal.HPS_ESA_Irregular = DatosHPSCalculados("Irregular", m_HPS, p_Error)
                If p_DatosLocal.HPS_ESA_Irregular = "Sí" Then
                    p_DatosLocal.HPS_ESA_MOTIVO_IRREGULAR = DatosHPSCalculados("MOTIVO_IRREGULAR", m_HPS, p_Error)
                End If
                p_DatosLocal.HPS_ESA_MESES_PARA_RENOVAR = DatosHPSCalculados("MESES_PARA_RENOVAR", m_HPS, p_Error)
            End If
            
            
        End With
    Else
        p_DatosLocal.HPS_ESA_SIN_DATOS = ""
    End If
    'UE
    Set m_HPS = p_UsuarioHPS.HPSUE
    If Not m_HPS Is Nothing Then
        With m_HPS
            p_DatosLocal.HPS_UE_SIN_DATOS = "No"
            p_DatosLocal.HPS_UE_F_Concesion = .F_Concesion
            p_DatosLocal.HPS_UE_F_Caducidad = .F_Caducidad
            p_DatosLocal.HPS_UE_Grado = .Grado
            p_DatosLocal.HPS_UE_F_Solicitud = .F_Solicitud
            p_DatosLocal.HPS_UE_Especialidad = .Especialidad
            p_DatosLocal.HPS_UE_Renovacion = .Renovacion
            p_DatosLocal.HPS_UE_F_Baja = .F_Baja
            
            If p_UsuarioHPS.F_BajaCalculada = "" Then
                p_DatosLocal.HPS_UE_Activo = DatosHPSCalculados("Activo", m_HPS, p_Error)
            Else
                 p_DatosLocal.HPS_UE_Activo = "No"
            End If
            If p_UsuarioHPS.F_BajaCalculada = "" Then
               p_DatosLocal.HPS_UE_ApuntoDeCaducar = DatosHPSCalculados("APuntoDeCaducar", m_HPS, p_Error)
            Else
                p_DatosLocal.HPS_UE_ApuntoDeCaducar = "No"
            End If
            p_DatosLocal.HPS_UE_Baja = DatosHPSCalculados("Baja", m_HPS, p_Error)
            If p_UsuarioHPS.F_BajaCalculada = "" Then
                p_DatosLocal.HPS_UE_Caducado = DatosHPSCalculados("Caducado", m_HPS, p_Error)
                   
            Else
                p_DatosLocal.HPS_UE_Caducado = "No"
            End If
            If p_UsuarioHPS.F_BajaCalculada = "" Then
                p_DatosLocal.HPS_UE_PendienteRenovacion = DatosHPSCalculados("PendienteRenovacion", m_HPS, p_Error)
            Else
                p_DatosLocal.HPS_UE_PendienteRenovacion = "No"
            End If
            If p_UsuarioHPS.F_BajaCalculada = "" Then
                p_DatosLocal.HPS_UE_Solicitado = DatosHPSCalculados("Solicitada", m_HPS, p_Error)
            Else
                p_DatosLocal.HPS_UE_Solicitado = "No"
            End If
            If p_UsuarioHPS.F_BajaCalculada = "" Then
                p_DatosLocal.HPS_UE_ESTADO = DatosHPSCalculados("Estado", m_HPS, p_Error)
                p_DatosLocal.HPS_UE_Irregular = DatosHPSCalculados("Irregular", m_HPS, p_Error)
                If p_DatosLocal.HPS_UE_Irregular = "Sí" Then
                    p_DatosLocal.HPS_UE_MOTIVO_IRREGULAR = DatosHPSCalculados("MOTIVO_IRREGULAR", m_HPS, p_Error)
                End If
                p_DatosLocal.HPS_UE_MESES_PARA_RENOVAR = DatosHPSCalculados("MESES_PARA_RENOVAR", m_HPS, p_Error)
            End If
            
            
        End With
    Else
        p_DatosLocal.HPS_UE_SIN_DATOS = "Sí"
    End If

    ' La caducidad de indicadores se decide por el MIN(F_Caducidad) global del
    ' usuario entre los tipos HPS existentes. Se mantienen los campos fisicos por
    ' tipo porque los formularios continuos filtran sobre TbDatosLocalParaIndicadores.
    AplicarCaducidadGlobalEnDatosLocal p_DatosLocal, p_UsuarioHPS

    m_FechaHPSConcesionMinima = p_UsuarioHPS.FechaHPSConcesionMinimaCalculada
    
    p_DatosLocal.FechaHPSConcesionMinima = m_FechaHPSConcesionMinima
                
   
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método NumeroTelValido ha devuelto el error: " & Err.Description
    End If
    Debug.Print p_Error
End Function

Public Function InicializarTablasLocalesYEntidad_Optimizado( _
                                                            Optional p_ForzarRegeneracionTotal As EnumSiNo = EnumSiNo.No, _
                                                            Optional ByRef p_Error As String _
                                                        ) As String

    On Error GoTo errores

    Dim dbLocal As DAO.Database
    Dim dbSource As DAO.Database
    Dim rsSource As DAO.Recordset
    Dim rsLocal As DAO.Recordset
    Dim strSQL As String
    Dim dictLocal As Object
    Dim idLocal As String
    Dim idSource As String
    Dim hashLocal As String
    Dim hashSource As String
    Dim m_Contador As Long
    Dim vKey As Variant
    Dim m_linea  As String
    
    p_Error = vbNullString

    Set dbLocal = CurrentDb
    Set dbSource = getdb()

    DoCmd.SetWarnings False

    ' FASE PREVIA: LIMPIEZA TOTAL
    If p_ForzarRegeneracionTotal <> EnumSiNo.No Then
        m_linea = "Regeneracion Total..."
        AvanceNuevo m_linea
        cache_usuarios_regenerar1 p_Error
        If p_Error <> vbNullString Then Err.Raise 1000, , p_Error

    Else
        m_linea = "Analizando diferencias entre Servidor y Local..."
        AvanceNuevo m_linea

        ' 1) Cargar IDs y Hash de TbDatosLocal en diccionario
        Set dictLocal = CreateObject("Scripting.Dictionary")
        dictLocal.CompareMode = vbTextCompare

        ' IMPORTANTE: Evitar SELECT * para no depender de schema cambiante
        strSQL = _
            "SELECT ID, DNI, Nombre, Apellido_1, Apellido_2, Telefono,EmpresaUsuario,EmpresaTramitadora,  " & _
            "CadenaContratistas,Correo_e,F_Nacimiento,LugarNacimiento,IDExpediente,F_Curso ,CursoEnVigor , " & _
            "Requiere_Curso,F_Baja,Motivo_HPS " & _
            "FROM TbDatosLocal"

        Set rsLocal = dbLocal.OpenRecordset(strSQL, dbOpenSnapshot)

        Do While Not rsLocal.EOF
            idLocal = Trim$(CStr(Nz(rsLocal!ID, vbNullString)))
            If idLocal <> vbNullString Then
                hashLocal = GenerarHashComparacion_Normalizado(rsLocal)
                If Not dictLocal.Exists(idLocal) Then
                    dictLocal.Add idLocal, hashLocal
                Else
                    ' Si hay duplicados de ID en local, nos quedamos con el primero (o puedes decidir el último)
                    ' dictLocal(idLocal) = hashLocal
                End If
            End If
            rsLocal.MoveNext
        Loop

        rsLocal.Close
        Set rsLocal = Nothing

        ' 2) Recorrer TbUsuarios (Origen)
        strSQL = _
            "SELECT U.ID, U.DNI, U.Nombre, U.Apellido_1, U.Apellido_2, U.Telefono, S1.Nombre AS EmpresaUsuario, " & _
            "S2.Nombre AS EmpresaTramitadora, U.CadenaContratistas, U.Correo_e, U.F_Nacimiento, " & _
            "U.LugarNacimiento, U.IDExpediente, U.F_Curso, U.CursoEnVigor, U.Requiere_Curso, U.F_Baja, " & _
            "Motivo_HPS " & _
            "FROM (TbUsuarios AS U LEFT JOIN TbSuministradores AS S1 ON U.IDEmpresaUsuario = S1.IDSuministrador) " & _
            "LEFT JOIN TbSuministradores AS S2 ON U.IDEmpresaHPS = S2.IDSuministrador;"

        Set rsSource = dbSource.OpenRecordset(strSQL, dbOpenSnapshot)

        m_Contador = 0
        Do While Not rsSource.EOF

            idSource = Trim$(CStr(Nz(rsSource!ID, vbNullString)))

            If idSource <> vbNullString Then

                hashSource = GenerarHashComparacion_Normalizado(rsSource)

                If dictLocal.Exists(idSource) Then
                    hashLocal = dictLocal(idSource)

                    ' Existe en ambos -> Comprobar cambios (comparacion binaria, sin rarezas de locale)
                    If StrComp(hashLocal, hashSource, vbBinaryCompare) <> 0 Then
                        'AvanceNuevo "Actualizando usuario ID: " & idSource
                        cache_usuario_regenerar p_ID:=idSource, p_Error:=p_Error
                        If p_Error <> vbNullString Then Err.Raise 1000, , p_Error
                    End If

                    ' Marcar como procesado
                    dictLocal.Remove idSource

                Else
                    ' No existe en local -> Crear
                    AvanceNuevo "Anadiendo usuario ID: " & idSource
                    cache_usuario_regenerar p_ID:=idSource, p_Error:=p_Error
                    If p_Error <> vbNullString Then Err.Raise 1000, , p_Error
                End If
            End If

            m_Contador = m_Contador + 1

            If (m_Contador Mod 100) = 0 Then
                DoEvents
                m_linea = "Procesados " & m_Contador & " usuarios..."
                AvanceNuevo m_linea
                Debug.Print m_linea
                DoEvents
            End If

            rsSource.MoveNext
        Loop

        rsSource.Close
        Set rsSource = Nothing

        ' 3) Borrar los que quedan en el diccionario (ya no existen en origen)
        For Each vKey In dictLocal.Keys
            AvanceNuevo "Borrando usuario obsoleto ID: " & CStr(vKey)
            cache_usuario_borrar p_ID:=CStr(vKey), p_Error:=p_Error
            If p_Error <> vbNullString Then Err.Raise 1000, , p_Error
        Next vKey
    End If

    ' Limpieza objetos activos
    Set m_ObjUsuarioActivo = Nothing
    Set m_ObjUsuarioSICAActivo = Nothing
    Set m_ObjObservacionActiva = Nothing

    DoCmd.SetWarnings True
    InicializarTablasLocalesYEntidad_Optimizado = "OK"
    Exit Function

errores:
    DoCmd.SetWarnings True

    If Err.Number = 1000 Then
        If p_Error = vbNullString Then p_Error = "Error (1000) en InicializarTablas (Optimizado)."
    Else
        p_Error = "Error en InicializarTablas (Optimizado): " & Err.Number & " - " & Err.Description
    End If

    InicializarTablasLocalesYEntidad_Optimizado = vbNullString
End Function
Private Function GenerarHashComparacion_Normalizado(rs As DAO.Recordset) As String
    On Error GoTo EH

    Dim s As String

    s = s & NormTxt(rs, "DNI") & "|"
    s = s & NormTxt(rs, "Nombre") & "|"
    s = s & NormTxt(rs, "Apellido_1") & "|"
    s = s & NormTxt(rs, "Apellido_2") & "|"
    s = s & NormTxt(rs, "Telefono") & "|"
    s = s & NormTxt(rs, "EmpresaUsuario") & "|"
    s = s & NormTxt(rs, "EmpresaTramitadora") & "|"
    s = s & NormTxt(rs, "CadenaContratistas") & "|"
    s = s & NormTxt(rs, "Correo_e") & "|"
    s = s & NormDate(rs, "F_Nacimiento") & "|"
    s = s & NormTxt(rs, "LugarNacimiento") & "|"
    s = s & NormNum(rs, "IDExpediente") & "|"
    s = s & NormDate(rs, "F_Curso") & "|"
    s = s & NormTxt(rs, "CursoEnVigor") & "|"
    s = s & NormTxt(rs, "Requiere_Curso") & "|"
    s = s & NormTxt(rs, "Motivo_HPS") & "|"
    

    GenerarHashComparacion_Normalizado = s
    Exit Function

EH:
    Debug.Print "GenerarHashComparacion_Normalizado ERROR: " & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, , Err.Description
End Function

Private Function NormTxt(rs As DAO.Recordset, ByVal fieldName As String) As String
    Dim v As Variant, s As String
    v = rs.Fields(fieldName).value

    If IsNull(v) Then
        NormTxt = vbNullString
        Exit Function
    End If

    s = CStr(v)
    s = Replace$(s, ChrW$(160), " ")
    s = Replace$(s, vbCrLf, vbLf)
    s = Replace$(s, vbCr, vbLf)
    NormTxt = Trim$(s)
End Function

Private Function NormDate(rs As DAO.Recordset, ByVal fieldName As String) As String
    Dim v As Variant, d As Date

    v = rs.Fields(fieldName).value
    If IsNull(v) Or v = vbNullString Then
        NormDate = vbNullString
        Exit Function
    End If

    ' Si ya es fecha
    If IsDate(v) Then
        d = CDate(v)
        NormDate = Format$(d, "yyyymmdd")  ' o "yyyymmddhhnnss" si quieres hora
        Exit Function
    End If

    ' Si no es convertible, devuelve el texto normalizado para que NO explote y puedas comparar
    NormDate = "INVALID_DATE:" & NormTxt(rs, fieldName)
End Function

Private Function NormNum(rs As DAO.Recordset, ByVal fieldName As String) As String
    Dim v As Variant
    Dim s As String
    Dim d As Double

    v = rs.Fields(fieldName).value
    If IsNull(v) Or v = vbNullString Then
        NormNum = vbNullString
        Exit Function
    End If

    ' Si es numérico (incluye strings numéricos)
    If IsNumeric(v) Then
        d = CDbl(v)
        ' Formato estable, sin separador regional
        NormNum = Replace$(Format$(d, "0.###############"), ",", ".")
        Exit Function
    End If

    ' Si no es convertible, NO rompas: marca y sigue
    s = CStr(v)
    s = Replace$(s, ChrW$(160), " ")
    s = Trim$(s)
    NormNum = "INVALID_NUM:" & s
End Function

Private Function NormBool(rs As DAO.Recordset, ByVal fieldName As String) As String
    Dim v As Variant
    v = rs.Fields(fieldName).value

    If IsNull(v) Or v = vbNullString Then
        NormBool = vbNullString
        Exit Function
    End If

    ' Access: True=-1, False=0. Strings: "True"/"False"/"1"/"0"
    If VarType(v) = vbBoolean Then
        NormBool = IIf(v, "1", "0")
    ElseIf IsNumeric(v) Then
        NormBool = IIf(CDbl(v) <> 0, "1", "0")
    Else
        Select Case UCase$(Trim$(CStr(v)))
            Case "TRUE", "VERDADERO", "SI", "S", "YES", "Y", "-1", "1"
                NormBool = "1"
            Case "FALSE", "FALSO", "NO", "N", "0"
                NormBool = "0"
            Case Else
                NormBool = "INVALID_BOOL:" & Trim$(CStr(v))
        End Select
    End If
End Function





