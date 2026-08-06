Attribute VB_Name = "Funciones Generales"
Option Compare Database
Option Explicit
Public Micoleccion As New Collection, strURLArchivoEncontrado As String

Private fso As New FileSystemObject
Public Function URLCompletaArchivoAyuda( _
                                        strNombreFormulario As String _
                                        ) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día X
    '   -Modificaciones:
    '       -12/03/2011.- Formateo la función a como está la mayoría
    '   -Funcionamiento:
    '       -devuelve una url mirando en la tabla TbHerramientaDocAyuda.NombreArchivoAyuda según el strNombreFormulario
    '   -Llamada desde:
   
    '   -Devuelve:
    '       URLCompletaArchivoAyuda = strURLAnexo
    '       URLCompletaArchivoAyuda = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim strNombreArchivo As String, strURLAnexo As String, strTextoError As String
    On Error GoTo errores
    
    '----------------------------------
    '   -Devuelve:
    '       Dame = strValorObtenido
    '       Dame = "#ERR" & "|" & strTextoError
    '--------------------------------------------
    flag = Dame("TbHerramientaDocAyuda", "NombreArchivoAyuda", "NombreFormulario", strNombreFormulario)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Dame ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strNombreArchivo = flag
    If strNombreArchivo = "" Then
        strTextoError = "No se ha podido encontrar el archivo de ayuda para este formulario"
        Err.Raise 1000
    End If
    strURLAnexo = m_ObjEntorno.URLDirectorioDocumentacionAyuda & strNombreArchivo
    
    URLCompletaArchivoAyuda = strURLAnexo
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método URLCompletaArchivoAyuda ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
    URLCompletaArchivoAyuda = "#ERR" & "|" & strTextoError
End Function

Public Function GrabarFechaInformeRACACole( _
                                            p_Col As Scripting.Dictionary, _
                                            Optional ByRef p_Error As String _
                                            ) As String
    
    Dim m_ID As Variant
    Dim m_Evento As Evento
    
    On Error GoTo errores
    If p_Col Is Nothing Then
        Exit Function
    End If
    
    For Each m_ID In p_Col
        Set m_Evento = p_Col(m_ID)
         m_Linea = "Registrando Fecha de Informe al RAC a " & m_Evento.IDEVENTO
        Avance m_Linea
        m_Evento.InformeRACRegistrarFecha CStr(Date), p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Set m_Evento = Nothing
    Next
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GrabarFechaInformeRACACole ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
    
End Function
Public Function DameNumeroRegistrosPorSQL( _
                                            strSQLLista As String, _
                                            Optional db As DAO.Database _
                                            ) As String
     '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día  5/10/2009
    '   -Modificaciones:
    '       -22/03/2011.- Formateo la función a como está la mayoría
    '   -Funcionamiento:
    '       -Va a devolver el número de registros que hay en un sql
    '   -Llamada desde
    '       -Material.MaterialSeguimientoConItervencionAbierta
    '       -Material.IntervencionDatosDefechasIntermedias
    '       -Material.IntervencionUltimosDatos
    '       -Material.IntervencionEdicion
    '       -Parte.SacarIDEventoDeIDParte
    '       -Planificacion.NumeroAnexosEnPlanificacion
    '       -Planificacion.NPlanificaciones
    '       -Salida.RellenoTablasAuxiliares
    '       -Salida.ObtenerHojaEventosDeExportacion
    '       -Salida.ObtenerHojaActividadesDeExportacion
    '       -Salida.ObtenerHojaMaterialesDeExportacion
    '       -Salida.ObtenerHojaEquiposEnMantenimientoRealizado
    '       -Salida.ObtenerHojaEquiposEnMantenimientoProgramado
    '       -Salida.ObtenerHojaEquiposEnMantenimientoProgramado
    '       -Form_Facturacion.DameTituloLista
    '       -Form_PartesGestion.SeleccionarParte
    '       -Form_CalibracionEquiposGestion.ComandoEliminarEquipo_Click
    '       -Form_CalibracionEquiposGestion.SeleccionarLista
    '       -Calibracion.EliminarEquipoACalibrar
    '       -Facturacion.RellenarTbAuxConCandidatosAFacturacion
    '       -Facturacion.GenerarFactura
    '       -Facturacion.PasarDeTbAuxFacturacionATablasParaExcel
    '       -Facturacion.CreaHojaEventos
    '       -Facturacion.CreaHojaActividades
    '       -Facturacion.CreaHojaMateriales
    '   -Devuelve:
    '       DameNumeroRegistrosPorSQL = CStr(lngRegistros)
    '       DameNumeroRegistrosPorSQL ="#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, lngRegistros As Long, strTextoError As String
    On Error GoTo errores
'    If InStr(1, strSQLLista, "SELECT") = 0 Then
'        strTextoError = "Introduzca un SQL que comience por SELECT"
'        Err.Raise 1000
'    End If
    If db Is Nothing Then
        Set db = CurrentDb()
    End If
    Set rcdDatos = db.OpenRecordset(strSQLLista)
    With rcdDatos
        If Not .EOF Then
            .MoveLast
            .MoveFirst
            lngRegistros = .RecordCount
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    DameNumeroRegistrosPorSQL = CStr(lngRegistros)
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método DameNumeroRegistrosPorSQL ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    DameNumeroRegistrosPorSQL = "#ERR" & "|" & strTextoError
End Function
Public Sub DesactivaLabel(formulario As Form)
    Dim ctrControl As Control, ctrForm As Form
    On Error Resume Next
    Set ctrForm = formulario
    For Each ctrControl In ctrForm
        With ctrControl
            If .ControlType = acLabel And Left(.Name, 3) = "lbl" Then
                If .ForeColor <> lngColorEtiquetaMenuSinPulsar Then
                    .ForeColor = lngColorEtiquetaMenuSinPulsar
                End If
             End If
        End With
    Next ctrControl
End Sub
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
Function FicheroAbierto(strFichero As String) As Boolean
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día X
    '   -Modificaciones:
    
    '   -Funcionamiento:
    
    '   -Llamada desde
    '       -BorrarHTMLENLocal
    '   -Devuelve:
    '       FicheroAbierto = true or false
    '       FicheroAbierto =False-->Error
    '-------------------------------------------------------------------
    Dim intfilenum As Integer
    On Error GoTo errores
    intfilenum = FreeFile()
    Open strFichero For Binary Access Read Write Lock Read Write As #intfilenum
    Close #intfilenum
    FicheroAbierto = False
    Exit Function
errores:
    FicheroAbierto = True
End Function
Public Function AbrirEnLocal( _
                                strURLAnexo As String, _
                                lngHwnd As Long _
                                ) As String
    
    Dim strNombreArchivo As String, strURLAnexoLocal As String
    Dim strTextoError As String
    On Error GoTo errores
    If Not fso.FileExists(strURLAnexo) Then
        strTextoError = "No es accesible la ruta del archivo que se pretende abrir" & vbCrLf & strURLAnexo
        Err.Raise 1000
    End If
    strNombreArchivo = fso.GetFileName(strURLAnexo)
    If FicheroAbierto(strURLAnexo) Then
        strTextoError = "El archivo debe estar abierto"
        Err.Raise 1000
    End If
    
    
    strURLAnexoLocal = m_ObjEntorno.URLDirectorioLocal & strNombreArchivo
    If fso.FileExists(strURLAnexoLocal) Then
        If FicheroAbierto(strURLAnexoLocal) Then
            strTextoError = "No se puede completar la operación, pruebe a cerrar el documento que tiene abierto"
            Err.Raise 1000
        End If
        fso.DeleteFile strURLAnexoLocal, True
    End If
    fso.CopyFile strURLAnexo, strURLAnexoLocal, True
    Ejecutar lngHwnd, "open", CStr(strURLAnexoLocal), "", "", 1
    
    AbrirEnLocal = strURLAnexoLocal
    Exit Function
errores:
    If Err.Number <> 1000 Then
        If Err.Number = 70 Then
            strTextoError = "No se puede completar la operación, pruebe a cerrar el documento que tiene abierto"
        Else
            strTextoError = "La función AbrirEnLocal ha dado el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
        End If
    End If
    
    AbrirEnLocal = "#ERR" & "|" & strTextoError
End Function
Public Sub AjustarTamaño(frmFormulario As Form)
    Dim i As Integer
    On Error GoTo AjustarTamaño_TratamientoErrores
    ' ajusto el ancho del formulario teniendo en cuenta si tiene o no selector de registros
    If Not frmFormulario.RecordSelectors Then
        frmFormulario.InsideWidth = frmFormulario.Width
    Else
        frmFormulario.InsideWidth = frmFormulario.Width + 250
    End If
    ' si se abre en vista formulario simple
    If frmFormulario.DefaultView = 0 Then
        'ajusto el alto incluyendo las distintas secciones, encabezado, pie, grupos...
        ' como no sé el número de secciones del formulario, me salgo al producirse un error
        frmFormulario.InsideHeight = 0
        For i = 0 To 100
            frmFormulario.InsideHeight = frmFormulario.InsideHeight + frmFormulario.Section(i).Height
        Next
    End If
AjustarTamaño_Salir:
   DoCmd.Restore
   On Error GoTo 0
   Exit Sub
AjustarTamaño_TratamientoErrores:
   If Not Err = 2462 Then  ' "El número de sección que introdujo no es válido."
      MsgBox "Error " & Err.Number & " en proc.: AjustarTamaño de Módulo: Módulo1 (" & Err.Description & ")"
   End If
   Resume AjustarTamaño_Salir
End Sub         ' AjustarTamaño

Public Function Dame( _
                    strTabla As String, _
                    strNombreCampoABuscar As String, _
                    strCampoID As String, _
                    strValor As String _
                    ) As String

    
    Dim rcdDatos As DAO.Recordset, strValorObtenido As String, strTextoError As String
    On Error GoTo errores
    m_SQL = "SELECT " & strTabla & "." & strNombreCampoABuscar & " " & _
                "FROM " & strTabla & " " & _
                "WHERE (((" & strTabla & "." & strCampoID & ")=" & strValor & "));"
    On Error Resume Next
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    If Err.Number <> 0 Then
        Err.Clear
        m_SQL = "SELECT " & strTabla & "." & strNombreCampoABuscar & " " & _
                    "FROM " & strTabla & " " & _
                    "WHERE (((" & strTabla & "." & strCampoID & ")='" & strValor & "'));"
        Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    End If
    With rcdDatos
        If Not .EOF Then
            strValorObtenido = Nz(rcdDatos(strNombreCampoABuscar), "")
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Dame = strValorObtenido
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Dame ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    Dame = "#ERR" & "|" & strTextoError
End Function

Public Function DameID1( _
                        p_NombreTabla As String, _
                        p_NombreCampoID As String, _
                        Optional ByRef p_db As DAO.Database, _
                        Optional ByRef p_Error As String _
                        ) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim lngIDMax As Long
    On Error GoTo errores
    
    If p_NombreTabla = "" Or p_NombreCampoID = "" Then
        p_Error = "Se ha de indicar el nombre de la tabla y de su campo ID"
        Err.Raise 1000
    End If
    If p_db Is Nothing Then
        Set p_db = getdb()
    End If
    m_SQL = "SELECT Max(" & p_NombreTabla & "." & p_NombreCampoID & ") AS MaxID " & _
            "FROM " & p_NombreTabla & ";"
    Set rcdDatos = p_db.OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            If IsNumeric(Nz(.Fields("MaxID"), "")) Then
                lngIDMax = .Fields("MaxID")
            End If

        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    DameID1 = CStr(lngIDMax + 1)
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método DameID1 ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function
Public Function DameID( _
                        strNombreTabla As String, _
                        strNombreCampoID As String _
                        ) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día x
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       -va a mirar la tabla strNombreTabla y nos va a dar un entero más que tenga el máximo dentro del campo llamado strNombreCampoID
    '   -Llamada por:
    '       -Gasto.Alta
    '       -MantenimientoProgramado.AltaIDEquipo
    '       -MantenimientoProgramado.AltaSeguimiento
    '       -Material.AltaIDEquipo
    '       -Planificacion.AltaPlanificacion
    '       -Planificacion.AnexarAPlanificacion
    '       -SubContratacion.Alta
    
    '       -Calibracion.AltaEquipoACalibrar
    '       -Calibracion.AltaCalibracion
    '       -Facturacion.GenerarFactura
    '       -Form_FormGastoAlta.Alta
    '   -Devuelve:
    '       DameID = CStr(lngMaxID + 1)
    '       DameID = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, lngMaxID As Long, lngID As Long, strTextoError As String
    On Error GoTo errores
    If strNombreTabla = "" Then
        strTextoError = "No se introducido un nombre de tabla adecuado"
        Err.Raise 1000
    End If
    If strNombreCampoID = "" Then
        strTextoError = "No se introducido un nombre del campo ID"
        Err.Raise 1000
    End If
    lngMaxID = 0
    m_SQL = "SELECT " & strNombreTabla & "." & strNombreCampoID & " " & _
            "FROM " & strNombreTabla & ";"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                If IsNumeric(Nz(.Fields(strNombreCampoID), "")) Then
                    lngID = CLng(.Fields(strNombreCampoID))
                    If lngID > lngMaxID Then
                        lngMaxID = lngID
                    End If
                End If
                .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    DameID = CStr(lngMaxID + 1)
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método DameID ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    DameID = "#ERR" & "|" & strTextoError
End Function
Public Function DameValorSiguienteAlMaximoDeUnCampo( _
                                                        strNombreTabla As String, _
                                                        strNombreCampo As String, _
                                                        strParteFijaInicial As String, _
                                                        intNumeroCaracteresDesdeLaDerecha As Integer _
                                                        ) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día x
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       -va a mirar la tabla strNombreTabla y nos va a dar un entero más que tenga el máximo dentro del campo llamado strNombreCampoID
    '   -Llamada por:
    '       -ImportarExcel.PasarDeTablasAuxiliaresARealesActividades
    '       -MaterialProgramado.MaterialDeActividadAlta
    '       -MaterialProgramado.MaterialNODeActividadAlta
    '       -Material.MaterialDeActividadAlta
    '       -Material.MaterialNODeActividadAlta
    '       -Actividad.Alta
    '   -Devuelve:
    '       DameValorSiguienteAlMaximoDeUnCampo = strParteFijaInicial & Format(CStr(lngValorMaximo + 1), String(intNumeroCaracteresDesdeLaDerecha, "0"))
    '       DameValorSiguienteAlMaximoDeUnCampo = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, strValorCampo As String, lngValor As Long, lngValorMaximo As Long, strTextoError As String
    On Error GoTo errores
    If strNombreTabla = "" Or strNombreCampo = "" Or intNumeroCaracteresDesdeLaDerecha = 0 Then
        strTextoError = "Se ha de introducir el nombre de la tabla el del campo y el número de caracteres a la derecha"
        Err.Raise 1000
    End If
    m_SQL = strNombreTabla
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
       If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                strValorCampo = Nz(.Fields(strNombreCampo), "")
                If Len(strValorCampo) >= intNumeroCaracteresDesdeLaDerecha Then
                    If Left(strValorCampo, Len(strParteFijaInicial)) = strParteFijaInicial Then
                        If IsNumeric(Right(strValorCampo, intNumeroCaracteresDesdeLaDerecha)) Then
                            lngValor = CLng(Right(strValorCampo, intNumeroCaracteresDesdeLaDerecha))
                            If lngValor > lngValorMaximo Then
                                lngValorMaximo = lngValor
                            End If
                        End If
                    End If
                End If
            .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If Len(CStr(lngValorMaximo + 1)) > intNumeroCaracteresDesdeLaDerecha Then
        strTextoError = "El máximo supera los caracteres reservados " & intNumeroCaracteresDesdeLaDerecha
        Err.Raise 1000
    End If
    DameValorSiguienteAlMaximoDeUnCampo = strParteFijaInicial & Format(CStr(lngValorMaximo + 1), String(intNumeroCaracteresDesdeLaDerecha, "0"))
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método DameValorSiguienteAlMaximoDeUnCampo ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    DameValorSiguienteAlMaximoDeUnCampo = "#ERR" & "|" & strTextoError
End Function

Public Function DameHorasTecnico( _
                                    strAliasTecnico As String, _
                                    strFecha As String, _
                                    Optional strIDActividadAExcluir As String, _
                                    Optional strLaborables As String _
                                    ) As String
     '-------------------------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 04/06/2013
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       -va a mirar en la tabla "TbActividades" cuántas horas lleva el usuario ese día
    '   -Llamada desde
    '       -Form_FormActividadAlta.ComandoAlta_Click
    '       -Form_FormActividadEdicion.ComandoRegistrar_Click
    '       -Actividad.Alta
    '       -Actividad.Edicion
    '       -Devuelve:
    '       DameHorasTecnico = cstr(dblHoras)
    '       DameHorasTecnico = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    Dim dblHoras As Double, rcdDatos As DAO.Recordset, strNombreCampo As String, strTextoError As String
    On Error GoTo errores
    If strAliasTecnico = "" Then
        strTextoError = "Se ha de indicar el técnico"
        Err.Raise 1000
    End If
    If Not IsDate(strFecha) Then
        strTextoError = "Se ha de indicar la fecha"
        Err.Raise 1000
    End If
    If strLaborables <> "Sí" And strLaborables <> "No" Then
        strLaborables = "Sí"
    End If
    If strLaborables = "Sí" Then
        strNombreCampo = "HorasLaborables"
    Else
       strNombreCampo = "HorasExtras"
    End If
    If strIDActividadAExcluir = "" Then
        m_SQL = "SELECT TbActividades." & strNombreCampo & " " & _
                "FROM TbActividades " & _
                "WHERE (((TbActividades.FechaAlta)=#" & Format(strFecha, "mm/dd/yyyy") & _
                "#) AND ((TbActividades.ALIASTECNICO)='" & strAliasTecnico & "'));"
        
    Else
        m_SQL = "SELECT TbActividades." & strNombreCampo & " " & _
                "FROM TbActividades " & _
                "WHERE (((TbActividades.FechaAlta)=#" & Format(strFecha, "mm/dd/yyyy") & _
                "#) AND ((TbActividades.ALIASTECNICO)='" & strAliasTecnico & _
                "') AND (Not (TbActividades.IDActividad)='" & strIDActividadAExcluir & "'));"
    End If
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                If IsNumeric(Nz(.Fields(strNombreCampo), "")) Then
                    dblHoras = dblHoras + CDbl(.Fields(strNombreCampo))
                End If
                .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    DameHorasTecnico = CStr(dblHoras)
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método DameHorasTecnico ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    DameHorasTecnico = "#ERR" & "|" & strTextoError
End Function


Public Function FechaViernes( _
                                intMes As Integer, _
                                intAño As Integer, _
                                Optional intSemana As Integer _
                                ) As String
     '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 17/06/2014
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       -
   
    '   -Llamada por:
    '       -Planificacion.AltaPlanificacion
    '       -Planificacion.CerrarPlanificacionSinTrabajos
    '       -Form_FormPlanificacionAlta.CerrarPlanificacionSinTrabajos
    '       -Form_FormPlanificacionReprogramacion.EstablecerListaSemana
    '   -Devuelve:
    '       FechaViernes =dteFechaViernesSemanax para intSemana=x <>0
    '       FechaViernes =strCadenaFechas para intSemana=0
    '       strCadenaFechas=01 & dteFechaViernesSemana1 & vbcrlf & _
                            02 & dteFechaViernesSemana2 & vbcrlf & _
    '                       ...
    '                       intNumeroSemanas & dteFechaViernesSemanax
    '       FechaViernes="#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    'si no tiene esa semana viernes FechaViernes="0"
    Dim dteFecha As Date, intDiaDeLaSemanaPrimeraSemana As Integer, intDiasTieneLaSemana As Integer, intDiaComienzaSegundaSemana As Integer
    Dim intNumeroSemanas As Integer, intDiaAcabalaUltimaSemana As Integer, dteFechaViernes As Date
    Dim dteFechaUltimoDiaMes As Date, intNumeroDiasMes As Integer, dteFechaSiguienteMes As Date
    Dim dteFechaViernesSemana1 As Date, dteFechaViernesSemana2 As Date, dteFechaViernesSemana3 As Date, dteFechaViernesSemana4 As Date, _
        dteFechaViernesSemana6 As Date, dteFechaViernesSemana5 As Date
    Dim strCadenaFechas As String, strTextoError As String
    On Error GoTo errores
    If intMes < 1 And intMes > 12 Then
        strTextoError = "Los meses son 12"
        Err.Raise 1000
    End If
    If intAño < 1 Then
        strTextoError = "El año ha de ser un número entero positivo"
        Err.Raise 1000
    End If
    intNumeroSemanas = NumeroSemanasTieneElMes(intMes, intAño)
    If intSemana > intNumeroSemanas Then
        strTextoError = "El mes " & intMes & " del año " & intAño & " no tiene la semana número " & intSemana
        Err.Raise 1000
    End If
    dteFecha = CDate("01/" & CStr(intMes) & "/" & CStr(intAño))
    intDiaDeLaSemanaPrimeraSemana = Weekday(dteFecha, vbMonday)
    intDiaComienzaSegundaSemana = 8 - intDiaDeLaSemanaPrimeraSemana
    dteFechaSiguienteMes = DateAdd("m", 1, dteFecha)
    intNumeroDiasMes = DateDiff("d", dteFecha, dteFechaSiguienteMes)
    dteFechaUltimoDiaMes = CDate(CStr(intNumeroDiasMes) & "/" & CStr(intMes) & "/" & CStr(intAño))
    intDiaAcabalaUltimaSemana = Weekday(dteFechaUltimoDiaMes, vbMonday)
    If intDiaDeLaSemanaPrimeraSemana <= 5 Then
        dteFechaViernesSemana1 = DateAdd("d", 5 - intDiaDeLaSemanaPrimeraSemana, dteFecha)
    End If
    dteFechaViernesSemana2 = DateAdd("d", 7 - intDiaDeLaSemanaPrimeraSemana + 5, dteFecha)
    dteFechaViernesSemana3 = DateAdd("d", 7, dteFechaViernesSemana2)
    dteFechaViernesSemana4 = DateAdd("d", 7, dteFechaViernesSemana3)
    If intNumeroSemanas = 5 Then
        If intDiaAcabalaUltimaSemana >= 5 Then
            dteFechaViernesSemana5 = DateAdd("d", 7, dteFechaViernesSemana4)
        End If
    ElseIf intNumeroSemanas = 6 Then
        dteFechaViernesSemana5 = DateAdd("d", 7, dteFechaViernesSemana4)
        If intDiaAcabalaUltimaSemana >= 5 Then
            dteFechaViernesSemana6 = DateAdd("d", 7, dteFechaViernesSemana5)
        End If
    End If
    If intSemana = 0 Then
        If dteFechaViernesSemana1 <> #12:00:00 AM# Then
            strCadenaFechas = "01 " & CStr(dteFechaViernesSemana1)
        End If
        If dteFechaViernesSemana2 <> #12:00:00 AM# Then
            If strCadenaFechas = "" Then
                strCadenaFechas = "02 " & CStr(dteFechaViernesSemana2)
            Else
                strCadenaFechas = strCadenaFechas & vbCrLf & _
                                "02 " & CStr(dteFechaViernesSemana2)
            End If
        End If
        If dteFechaViernesSemana3 <> #12:00:00 AM# Then
            If strCadenaFechas = "" Then
                strCadenaFechas = "03 " & CStr(dteFechaViernesSemana3)
            Else
                strCadenaFechas = strCadenaFechas & vbCrLf & _
                                "03 " & CStr(dteFechaViernesSemana3)
            End If
        End If
        If dteFechaViernesSemana4 <> #12:00:00 AM# Then
            If strCadenaFechas = "" Then
                strCadenaFechas = "04 " & CStr(dteFechaViernesSemana4)
            Else
                strCadenaFechas = strCadenaFechas & vbCrLf & _
                                "04 " & CStr(dteFechaViernesSemana4)
            End If
        End If
        If dteFechaViernesSemana5 <> #12:00:00 AM# Then
            If strCadenaFechas = "" Then
                strCadenaFechas = "05 " & CStr(dteFechaViernesSemana5)
            Else
                strCadenaFechas = strCadenaFechas & vbCrLf & _
                                "05 " & CStr(dteFechaViernesSemana5)
            End If
        End If
        If dteFechaViernesSemana6 <> #12:00:00 AM# Then
            If strCadenaFechas = "" Then
                strCadenaFechas = "06 " & CStr(dteFechaViernesSemana6)
            Else
                strCadenaFechas = strCadenaFechas & vbCrLf & _
                                "06 " & CStr(dteFechaViernesSemana6)
            End If
        End If
        FechaViernes = strCadenaFechas
        Exit Function
    ElseIf intSemana = 1 Then
        dteFechaViernes = dteFechaViernesSemana1
    ElseIf intSemana = 2 Then
        dteFechaViernes = dteFechaViernesSemana2
    ElseIf intSemana = 3 Then
        dteFechaViernes = dteFechaViernesSemana3
    ElseIf intSemana = 4 Then
        dteFechaViernes = dteFechaViernesSemana4
    ElseIf intSemana = 5 Then
        dteFechaViernes = dteFechaViernesSemana5
    ElseIf intSemana = 6 Then
        dteFechaViernes = dteFechaViernesSemana6
    End If
    If dteFechaViernes = #12:00:00 AM# Then
        FechaViernes = "0"
    Else
        FechaViernes = CStr(dteFechaViernes)
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método FechaViernes ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    FechaViernes = "#ERR" & "|" & strTextoError
End Function
Public Function EquipoSinPlanificaciones(strIDEquipo As String, Optional strNunca As String) As Integer
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 30/06/2014
    '   -Modificaciones:
    '
    '   -Funcionamiento:
    '       -si strNunca="Sí"-->Será True si el strIDEquipo nunca se ha programado una planificación
    '      -si strNunca="Sí"-->Será True si el strIDEquipo no tiene ninguna planificación activa (sin cerrar)
   
    '   -llamada por:
    
    '   -Devuelve:
    '       EquipoSinPlanificaciones ="Sí/No"
    '       EquipoSinPlanificaciones ="#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim strIDEquipoEnBD As String, intNumeroPlanificaciones As Integer, strTextoError As String
    On Error GoTo errores
    If Not IsNumeric(strIDEquipo) Then
        strTextoError = "Introduzca el ID del Equipo"
        Err.Raise 1000
    End If
    If strNunca <> "Sí" And strNunca <> "No" Then
        strNunca = "Sí"
    End If
    If strNunca = "Sí" Then
        '------------------------------------------
        '   -Devuelve:
        '       Dame = strValorObtenido
        '       Dame = "#ERR" & "|" & strTextoError
        '--------------------------------------------
        flag = Dame("TbPlanificacion", "IDEquipo", "IDEquipo", strIDEquipo)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método Dame ha devuelto el error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        strIDEquipoEnBD = flag
        If IsNumeric(strIDEquipoEnBD) Then
            EquipoSinPlanificaciones = 0
        Else
            EquipoSinPlanificaciones = 1
        End If
    Else
        '-----------------------------------------------
        '   -Devuelve:
        '       Dame = strDescriptivo
        '       Dame = "#ERR" & "|" & strTextoError
        '-----------------------------------------------
        flag = Dame("TbPlanificacionEquipos", "IDEquipo", "IDEquipo", strIDEquipo)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método Dame ha devuelto un error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        If flag = "" Then
            EquipoSinPlanificaciones = 0
        Else
            EquipoSinPlanificaciones = 1
        End If
       
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Planificacion.EquipoSinPlanificaciones ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    EquipoSinPlanificaciones = -1
End Function
Public Function DameFechaFinEvento( _
                                    strIDEvento As String _
                                    ) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 31/1/13
    '   -Modificaciones:
    '
    '   -Funcionamiento:
    '       -va a devolver la última fecha de la última actividad que tenga asociado el evento
    '
    '   -llamada por:
    '       -DameFechaFinEvento
    '   -Devuelve:
    '       DameFechaFinEvento = strFechaUltimaActividad
    '       DameFechaFinEvento = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, strFechaUltimaActividad As String, strTextoError As String
    On Error GoTo errores
    If strIDEvento = "" Then
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbActividades.FechaAlta " & _
            "FROM TbActividades " & _
            "WHERE (((TbActividades.IDEvento)='" & strIDEvento & "')) " & _
            "ORDER BY TbActividades.FechaAlta DESC;"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            strFechaUltimaActividad = Nz(.Fields("FechaAlta"), "")
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    DameFechaFinEvento = strFechaUltimaActividad
    Exit Function
errores:
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    DameFechaFinEvento = "#ERR" & "|" & strTextoError
End Function

Public Function DameHorasMaximasEnActividadesEventoPorDia( _
                                                            p_IDEvento As String, _
                                                            ByRef p_ColHorasDia As Scripting.Dictionary, _
                                                            Optional ByRef p_Error As String _
                                                            ) As String
    Dim m_objRcdDatos As DAO.Recordset
    Dim m_Horas As Double
    Dim m_Fecha As Variant
    
    m_SQL = "SELECT TbActividades.FechaAlta, Max(TbActividades.HorasLaborables) AS MáxDeHorasLaborables " & _
                "FROM TbActividades " & _
                "WHERE (((TbActividades.IDEvento)='" & p_IDEvento & "')) " & _
                "GROUP BY TbActividades.FechaAlta;"
    Set m_objRcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With m_objRcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                m_Fecha = .Fields("FechaAlta")
                If IsNumeric(Nz(.Fields("MáxDeHorasLaborables"), "")) Then
                    m_Horas = CDbl(.Fields("MáxDeHorasLaborables"))
                    If p_ColHorasDia Is Nothing Then
                        Set p_ColHorasDia = New Scripting.Dictionary
                    End If
                    p_ColHorasDia.Add m_Fecha, m_Horas
                End If
                
                .MoveNext
            Loop

        End If
    End With
    m_objRcdDatos.Close
    Set m_objRcdDatos = Nothing
    
    Exit Function
errores:
    If Not m_objRcdDatos Is Nothing Then
        m_objRcdDatos.Close
        Set m_objRcdDatos = Nothing
    End If
    If Err.Number <> 0 Then
        p_Error = "El método DameHorasMaximasEnActividadesEvento ha devuelto el error: " & Err.Description
    End If
    
End Function







Public Function EstaEnCombo(cbo As ComboBox, strValor As String) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día x
    '   -Modificaciones:
    '
    '   -Funcionamiento:
    '
    '   -llamada por:
    '       -Form_FormIDEquiposGestion.CargarIDEquipo
    '       -Form_FormIDEquiposGestion.CargarSubSistema
    '       -Form_FormIDEquiposGestion.CargarBUI
    '       -Form_FormEventoGestion.CargarIDEquipo
    '       -Form_FormEventoGestion.CargarSubSistema
    '   -Devuelve:
    '       EstaEnCombo = Sí/No
    '       EstaEnCombo = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim i As Integer, strTextoError As String
    On Error GoTo errores
    If strValor = "" Then
        strTextoError = "Se ha de indicar el valor"
        Err.Raise 1000
    End If
    For i = 0 To cbo.ListCount - 1
        If cbo.ItemData(i) = strValor Then
            EstaEnCombo = "Sí"
            Exit Function
        End If
    Next
    EstaEnCombo = "No"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método EstaEnCombo ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    EstaEnCombo = "#ERR" & "|" & strTextoError
End Function
Public Sub RecuadrarRango(Rango As excel.Range, Optional strTipoMarco = "Sencillo")
    On Error Resume Next
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    If strTipoMarco = "Sencillo" Then
        With Rango
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
                        
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).Weight = xlThin
            .Borders(xlEdgeLeft).ColorIndex = xlAutomatic
                        
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Weight = xlThin
            .Borders(xlEdgeTop).ColorIndex = xlAutomatic
                        
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Weight = xlThin
            .Borders(xlEdgeBottom).ColorIndex = xlAutomatic
                        
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeRight).Weight = xlThin
            .Borders(xlEdgeRight).ColorIndex = xlAutomatic
            
            .Borders(xlInsideVertical).LineStyle = xlContinuous
            .Borders(xlInsideVertical).Weight = xlThin
            .Borders(xlInsideVertical).ColorIndex = xlAutomatic
        End With
    ElseIf strTipoMarco = "Sencillo No Interior" Then
        Rango.Borders(xlDiagonalDown).LineStyle = xlNone
        Rango.Borders(xlDiagonalUp).LineStyle = xlNone
        With Rango.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With Rango.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With Rango.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With Rango.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        Rango.Borders(xlInsideVertical).LineStyle = xlNone
        Rango.Borders(xlInsideHorizontal).LineStyle = xlNone
    
    ElseIf strTipoMarco = "Doble" Then
        With Rango
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
            
        End With
        With Rango.Borders(xlEdgeLeft)
            .LineStyle = xlDouble
            .Weight = xlThick
            .ColorIndex = xlAutomatic
        End With
        With Rango.Borders(xlEdgeTop)
            .LineStyle = xlDouble
            .Weight = xlThick
            .ColorIndex = xlAutomatic
        End With
        With Rango.Borders(xlEdgeBottom)
            .LineStyle = xlDouble
            .Weight = xlThick
            .ColorIndex = xlAutomatic
        End With
        With Rango.Borders(xlEdgeRight)
            .LineStyle = xlDouble
            .Weight = xlThick
            .ColorIndex = xlAutomatic
        End With
        With Rango
            .Borders(xlInsideVertical).LineStyle = xlNone
            .Borders(xlInsideHorizontal).LineStyle = xlNone
            
        End With
    ElseIf strTipoMarco = "Multiple" Then
         With Rango
            .Borders(xlDiagonalDown).LineStyle = xlNone
            .Borders(xlDiagonalUp).LineStyle = xlNone
            
        End With
        With Rango.Borders(xlEdgeLeft)
            .LineStyle = xlDouble
            .Weight = xlThick
            .ColorIndex = xlAutomatic
        End With
        With Rango.Borders(xlEdgeTop)
            .LineStyle = xlDouble
            .Weight = xlThick
            .ColorIndex = xlAutomatic
        End With
        With Rango.Borders(xlEdgeBottom)
            .LineStyle = xlDouble
            .Weight = xlThick
            .ColorIndex = xlAutomatic
        End With
        With Rango.Borders(xlEdgeRight)
            .LineStyle = xlDouble
            .Weight = xlThick
            .ColorIndex = xlAutomatic
        End With
        With Rango.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        On Error Resume Next
        With Rango.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
    End If
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
End Sub

Public Sub ColocarTituloEnFiltro(lst As ListBox, lblEtiqueta As Label)
    Dim strTextoError As String, strTituloFiltro As String, intElementosEnLista
    On Error GoTo errores
    If lst.ColumnHeads = True Then
        intElementosEnLista = lst.ListCount - 1
    Else
        intElementosEnLista = lst.ListCount
    End If
    If intElementosEnLista = 1 Then
        strTituloFiltro = "Lista de filtrados ( " & CStr(intElementosEnLista) & " Elemento)"
    Else
        strTituloFiltro = "Lista de filtrados ( " & CStr(intElementosEnLista) & " Elementos)"
    End If
    lblEtiqueta.Caption = strTituloFiltro
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar obtener el título se ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    pregunta = MsgBox(strTextoError, vbCritical, "Error")

End Sub
Public Sub VerAyudaFormulario(frm As Form)
    Dim strURLAyuda As String, strTextoError As String
    On Error GoTo errores
    '------------------------------------------------------------------------------
    '   -Devuelve:
    '       URLCompletaArchivoAyuda = strURLAnexo
    '       URLCompletaArchivoAyuda = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = URLCompletaArchivoAyuda(frm.Name)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método URLCompletaArchivoAyuda ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strURLAyuda = flag
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       AbrirEnLocal = strURLAnexoLocal
    '       AbrirEnLocal = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = AbrirEnLocal(strURLAyuda, frm.hWnd)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método URLCompletaArchivoAyuda ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        strTextoError = "Al intentar abrir la ayuda se ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    pregunta = MsgBox(strTextoError, vbCritical, "Error")
End Sub
Public Function Seleccionar(EsArchivo As Boolean, strTitulo As String) As String
    
    Dim fDialog As Object, varFile As Variant, strTextoError As String
    On Error GoTo errores
    If EsArchivo = True Then
        Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    Else
        Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
    End If
    With fDialog
        .Show
        If EsArchivo Then
            .AllowMultiSelect = False
            .InitialFileName = m_ObjEntorno.URLArchivoUltimo
            .Title = strTitulo
            .Filters.Clear
            .Filters.Add "All Files", "*.*"
        End If
        For Each varFile In .SelectedItems
            Seleccionar = CStr(varFile)
        Next
    End With
    If EsArchivo Then
        m_ObjEntorno.URLArchivoUltimo = CStr(varFile)
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Seleccionar ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    Seleccionar = "#ERR" & "|" & strTextoError
End Function
Public Sub RellenaCombo(ByRef cbo As ComboBox, m_SQL As String, Optional strFilaTitulos As String)
    Dim rcdDatos As DAO.Recordset, fld As Field, i As Integer, strValor As String, strFila As String
    On Error GoTo errores
    VBA.DoEvents
    DoCmd.Hourglass True
    VBA.DoEvents
    cbo.RowSource = ""
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            If strFilaTitulos <> "" Then
               
                cbo.AddItem strFilaTitulos
            End If
            Do While Not .EOF
                strFila = ""
                For i = 0 To .Fields.count - 1
                    strValor = Nz(.Fields(i).Value, "")
                    If strFila = "" Then
                        strFila = strValor
                    Else
                        strFila = strFila & ";" & strValor
                    End If
                Next
                cbo.AddItem strFila
                .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    VBA.DoEvents
    DoCmd.Hourglass False
    VBA.DoEvents
    Exit Sub
errores:
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    DoCmd.Hourglass False
End Sub
Public Function EstablecerComboEquipo(Optional strBUI As String, Optional strSUBSISTEMA As String) As String
    
    Dim strTextoError As String
    On Error GoTo errores
    If strBUI = "" Then strBUI = "-1"
    If strSUBSISTEMA = "" Then strSUBSISTEMA = "-1"
    EstablecerComboEquipo = "SELECT TbEquipos.IDEquipo,TbEquipos.Equipo,TbEquipos.FechaObsoleto " & _
                "FROM TbEquipos " & _
                "WHERE BUI='" & strBUI & "' AND SUBSISTEMA='" & strSUBSISTEMA & "';"
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método EstablecerComboSubSistema ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    EstablecerComboEquipo = "ERR" & "|" & strTextoError
End Function
Public Function EstablecerComboSubSistema(Optional strBUI As String) As String
   
    Dim strTextoError As String, m_SQL As String
    On Error GoTo errores
    If strBUI = "" Then strBUI = "-1"
    m_SQL = "SELECT  TbSubsistemaBui.SUBSISTEMA " & _
            "FROM TbSubsistemaBui " & _
            "WHERE (((TbSubsistemaBui.BUI)='" & strBUI & "')) " & _
            "ORDER BY TbSubsistemaBui.SUBSISTEMA;"
    EstablecerComboSubSistema = m_SQL
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método EstablecerComboSubSistema ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    EstablecerComboSubSistema = "ERR" & "|" & strTextoError
End Function
Public Function EstablecerComboMaterial(p_IDEquipo As String, Optional ByRef p_Error As String) As String
    
    
    On Error GoTo errores
    
    EstablecerComboMaterial = "SELECT DISTINCT TbMaterial.Material " & _
                    "FROM ((TbMaterial INNER JOIN TbActividades ON TbMaterial.IDActividad = TbActividades.IDActividad) " & _
                    "INNER JOIN TbEventos ON TbActividades.IDEvento = TbEventos.IDEvento) " & _
                    "INNER JOIN TbEquipos ON TbEventos.IDEquipo = TbEquipos.IDEquipo " & _
                    "WHERE (((TbEquipos.IDEquipo)=" & p_IDEquipo & ")) " & _
                    "ORDER BY TbMaterial.Material;"
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EstablecerComboMaterial ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function


Public Function DameDiasVigenciaCalibracion( _
                                            strIDEquipo As String _
                                            ) As String
    '-------------------------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 15/11/13
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       -si no tiene calibracion -->DameDiasVigenciaCalibracion="SC"
    '       -si tiene alguna calibración-->DameDiasVigenciaCalibracion=cstr(intDias) si es negativa es que ya no está en vigencia
    '   -Llamada desde
    
    '   -Devuelve:
        
    '       DameDiasVigenciaCalibracion = strTexto
    '-----------------------------------------------
    Dim strTexto As String, rcdDatos As DAO.Recordset, strUltimaFechaFinCalibracion As String, intDias As Integer
    On Error GoTo errores
    If Not IsNumeric(strIDEquipo) Then
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbEquiposCalibrablesFechas.FechaFinCalibrado " & _
            "FROM TbEquiposCalibrablesFechas " & _
            "WHERE (((TbEquiposCalibrablesFechas.IDEquipoCalibrable) = " & strIDEquipo & ")) " & _
            "ORDER BY TbEquiposCalibrablesFechas.FechaFinCalibrado DESC;"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            strUltimaFechaFinCalibracion = Nz(.Fields("FechaFinCalibrado"), "")
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If Not IsDate(strUltimaFechaFinCalibracion) Then
        strTexto = "SC"
    Else
        intDias = DateDiff("d", Now(), strUltimaFechaFinCalibracion)
        strTexto = CStr(intDias)
       
    End If
    DameDiasVigenciaCalibracion = strTexto
    Exit Function
errores:
    DameDiasVigenciaCalibracion = "#"
End Function
Public Function ObtenerCodigoLI(strCodExp As String) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día
    '   -Modificaciones:
    '
    '   -Funcionamiento:
    '
    '   -llamada por:
    '       -Facturacion.ObtenerLI
    '   -Devuelve:
    '       ObtenerCodigoLI = Sí/No
    '       ObtenerCodigoLI = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, db As DAO.Database, strURLBBDD As String
    Dim strEdicionMax As String, strCodigo As String, strOrdinalActualMax As String, _
        strCodigoUltimoLI As String, strAño As String, strCodExpSinAño As String, strPassActual As String, strTextoError As String
    On Error GoTo errores
    If Len(strCodExp) <> 7 Then
        strTextoError = "El código de expediente no tiene la forma XXXX/YY"
        Err.Raise 1000
    End If
    If InStr(1, strCodExp, "/") = 0 Then
        strTextoError = "El código de expediente no tiene la forma XXXX/YY"
        Err.Raise 1000
    End If
    dato = Split(strCodExp, "/")
    strCodExpSinAño = CStr(dato(0))
    strAño = CStr(dato(1))
    If Len(strCodExpSinAño) <> 4 Then
        strTextoError = "El código de expediente no tiene la forma XXXX/YY"
        Err.Raise 1000
    End If
    If Not IsNumeric(strAño) Then
        strTextoError = "El código de expediente no tiene la forma XXXX/YY"
        Err.Raise 1000
    End If
    strURLBBDD = "\\10.14.7.52\utilidades\Aplicaciones PpD\AGEDO\recursos\Agedo20.accde"
    If Not fso.FileExists(strURLBBDD) Then
        strTextoError = "La base de dato de AGEDO no se encuentra en : " & vbNewLine & strURLBBDD
        Err.Raise 1000
    End If
    strPassActual = "dpddpd"
    Set db = OpenDatabase(strURLBBDD, False, False, "MS Access;PWD=" & strPassActual & "")
    m_SQL = "SELECT TbDocumentos.Codigo " & _
            "FROM TbDocumentos " & _
            "WHERE (((TbDocumentos.CodExp)='" & strCodExp & _
            "') AND ((TbDocumentos.Tipo)='LI') AND ((TbDocumentos.Area)='A')) " & _
            "ORDER BY TbDocumentos.Codigo DESC;"
    Set rcdDatos = db.OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            strCodigoUltimoLI = Nz(.Fields("Codigo"), "")
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    db.Close
    If strCodigoUltimoLI = "" Then
        strOrdinalActualMax = "0"
    Else
        strOrdinalActualMax = Right(strCodigoUltimoLI, 2)
    End If
    If Not IsNumeric(strOrdinalActualMax) Then
        strOrdinalActualMax = "0"
    End If
    
    If strCodigoUltimoLI = "" Then
        strCodigo = "LI" & strAño & strCodExpSinAño & "CCCA00"
    Else
        strCodigo = Left(strCodigoUltimoLI, Len(strCodigoUltimoLI) - 2) & Format(CInt(strOrdinalActualMax) + 1, "00")
    End If
    
    ObtenerCodigoLI = strCodigo
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Facturacion.ObtenerCodigoLI ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    
    If Not db Is Nothing Then
        Set db = Nothing
    End If
    ObtenerCodigoLI = "#ERR" & "|" & strTextoError
End Function


Public Function GenerarConsultas( _
                                p_NombreInforme As String, _
                                lngHwn As Long, _
                                Optional p_SQLConsulta As String, _
                                Optional lst As ListBox, _
                                Optional strRemota As String = "Sí" _
                                ) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día x
    '   -Modificaciones:
   
    '   -Funcionamiento:
    '       -pasa el sql a un excel
    '   -Llamada desde
    '       -Form_FormExportarExcelArchivosPartes.ComandoExportarExcelArchivosNoEnDD_Click
    '       -Form_FormExportarExcelArchivosPartes.ComandoExportarExcelArchivosPartes_Click
    '       -Form_FormGastoGestion.ComandoExportarExcel_Click
    '   -Devuelve:
    '       GenerarConsultas = Descriptivo
    '       GenerarConsultas = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, rcd As DAO.Recordset, WbLibro As excel.Workbook, wbHoja As excel.Worksheet, intFila As Integer, strTextoError As String
    Dim fila As Integer, columna As Integer, strNombreCampo As String, strValor As String, strURLDirectorioExcel As String, strURLExcel As String, strNombreArchivo As String, _
         MiCol As New Collection, strRegistro As String, VarItem As Variant, dato1 As Variant, varItem1 As Variant, _
        lngNumeroRegistros As Long
    On Error GoTo errores
    If strRemota <> "Sí" And strRemota <> "No" Then
        strRemota = "Sí"
    End If
    If p_SQLConsulta = "" And lst Is Nothing Then
        strTextoError = "El origen de la consulta no se ha establecido"
        Err.Raise 1000
    End If
    If p_SQLConsulta <> "" And Not lst Is Nothing Then
        strTextoError = "El origen de la consulta no se ha establecido"
        Err.Raise 1000
    End If
    If Not lst Is Nothing Then
        If lst.ColumnHeads = True Then
            lngNumeroRegistros = lst.ListCount - 1
        Else
            lngNumeroRegistros = lst.ListCount
        End If
        If lngNumeroRegistros = 0 Then
            strTextoError = "La consulta no contiene registros"
            Err.Raise 1000
        End If
    End If
    
    
    strNombreArchivo = fso.GetBaseName(p_NombreInforme) & "_" & Hour(Now()) & "_" & Minute(Now()) & "_" & Second(Time) & ".xlsx"
    strURLExcel = m_ObjEntorno.URLDirectorioLocal & strNombreArchivo
    If fso.FileExists(strURLExcel) Then
        If FicheroAbierto(strURLExcel) Then
            strTextoError = "Tiene una consulta abierta"
            Err.Raise 1000
        End If
        fso.DeleteFile strURLExcel, True
    End If
    If Not lst Is Nothing Then
        
        For fila = 0 To lngNumeroRegistros
            strRegistro = ""
            For columna = 0 To lst.ColumnCount - 1
                strValor = Nz(lst.Column(columna, fila), "null")
                If strValor = "" Then strValor = "Null"
                If strRegistro = "" Then
                    strRegistro = strValor
                Else
                    If IsNumeric(strValor) Then
                        strValor = Replace(strValor, ",", ".")
                    End If
                    strRegistro = strRegistro & ";" & strValor
                End If
            Next columna
            MiCol.Add strRegistro
        Next fila
    Else
        m_SQL = p_SQLConsulta
        If strRemota = "Sí" Then
            Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
            With rcdDatos
                If Not .EOF Then
                    .MoveFirst
                    For columna = 0 To .Fields.count - 1
                        strValor = Nz(.Fields(columna).Name, "")
                        If strRegistro = "" Then
                            strRegistro = strValor
                        Else
                            strRegistro = strRegistro & ";" & strValor
                        End If
                    Next columna
                    MiCol.Add strRegistro
                    Do While Not .EOF
                        strRegistro = ""
                        For columna = 0 To .Fields.count - 1
                            strValor = Nz(.Fields(columna).Value, "")
                            If strRegistro = "" Then
                                strRegistro = strValor
                            Else
                                strRegistro = strRegistro & ";" & strValor
                            End If
                        Next columna
                        MiCol.Add strRegistro
                        .MoveNext
                    Loop
                End If
            End With
            rcdDatos.Close
            Set rcdDatos = Nothing
        Else
            Set rcd = CurrentDb().OpenRecordset(m_SQL)
            With rcd
                 If Not .EOF Then
                    .MoveFirst
                    For columna = 0 To .Fields.count - 1
                        strValor = Nz(.Fields(columna).Name, "")
                        If strRegistro = "" Then
                            strRegistro = strValor
                        Else
                            strRegistro = strRegistro & ";" & strValor
                        End If
                    Next columna
                    MiCol.Add strRegistro
                    Do While Not .EOF
                        strRegistro = ""
                        For columna = 0 To .Fields.count - 1
                            strValor = Nz(.Fields(columna).Value, "")
                            If strRegistro = "" Then
                                strRegistro = strValor
                            Else
                                strRegistro = strRegistro & ";" & strValor
                            End If
                        Next columna
                        MiCol.Add strRegistro
                        .MoveNext
                    Loop
                End If
            
            End With
            rcd.Close
            Set rcd = Nothing
        End If
    End If
    If MiCol.count = 0 Then
         strTextoError = "La consulta no contiene registros"
        Err.Raise 1000
    End If
    Set appexcel = New excel.Application
    appexcel.Visible = False
    Set WbLibro = appexcel.Workbooks.Add
    WbLibro.SaveAs strURLExcel
    Set wbHoja = WbLibro.Worksheets(1)
    With wbHoja
        intFila = 1
        For Each VarItem In MiCol
            dato = Split(VarItem, ";")
            columna = 1
            For Each varItem1 In dato
                If varItem1 = "Null" Then varItem1 = ""
                If IsDate(varItem1) Then
                    varItem1 = Format(varItem1, "mm/dd/yyyy")
                End If
                .Cells(intFila, columna).Value = varItem1
                columna = columna + 1
            Next
            intFila = intFila + 1
        Next
    End With
    WbLibro.Close True
    Set WbLibro = Nothing
    appexcel.Quit
    Set appexcel = Nothing
    Ejecutar lngHwn, "open", strURLExcel, "", "", 1
    
    GenerarConsultas = strURLExcel
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Consultas.GenerarConsultas ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        Set rcdDatos = Nothing
    End If
    If Not rcd Is Nothing Then
        Set rcd = Nothing
    End If
    
     If Not WbLibro Is Nothing Then
        WbLibro.Close False
        Set WbLibro = Nothing
    End If
    If Not appexcel Is Nothing Then
        appexcel.Quit
        Set appexcel = Nothing
    End If
    GenerarConsultas = "#ERR" & "|" & strTextoError
End Function


Public Function GenerarConsultasSQL( _
                                    p_SQLConsulta As String, _
                                    Optional ByRef p_Error As String _
                                    ) As String
   
    
    Dim m_Col As Scripting.Dictionary
    
    
    
'    Dim strNombreCampo As String, strValor As String, strURLDirectorioExcel As String, m_URLExcel As String,  As String, _
'         MiCol As New Collection, strRegistro As String, VarItem As Variant, dato1 As Variant, varItem1 As Variant, _
'        lngNumeroRegistros As Long
    On Error GoTo errores
    
    If p_SQLConsulta = "" Then
        p_Error = "El origen de la consulta no se ha establecido"
        Err.Raise 1000
    End If
    
    Dim m_NombreArchivo As String
    
    m_NombreArchivo = fso.GetTempName & ".xlsx"
    Dim m_URLExcel As String
    m_URLExcel = m_ObjEntorno.URLDirectorioLocal & m_NombreArchivo
    If fso.FileExists(m_URLExcel) Then
        If FicheroAbierto(m_URLExcel) Then
            p_Error = "Tiene una consulta abierta"
            Err.Raise 1000
        End If
        fso.DeleteFile m_URLExcel, True
    End If
    Dim m_SQL As String
    Dim rcdDatos As DAO.Recordset
    m_SQL = p_SQLConsulta
    Set rcdDatos = CurrentDb.OpenRecordset(m_SQL)
    If rcdDatos.EOF Then
        p_Error = "La consulta no contiene registros"
        Err.Raise 1000
    End If
    Dim fld As DAO.Field
    Set appexcel = New excel.Application
    appexcel.Visible = False
    Dim WbLibro As excel.Workbook
    Dim wbHoja As excel.Worksheet
    Set WbLibro = appexcel.Workbooks.Add
    WbLibro.SaveAs m_URLExcel
    Set wbHoja = WbLibro.Worksheets(1)
    Dim m_Valor As String
    Dim intFila As Integer
    Dim fila As Integer
    Dim columna As Integer
    
    intFila = 1
    
    Do While Not rcdDatos.EOF
        columna = 1
        For Each fld In rcdDatos.Fields
            If intFila = 1 Then
                'CABECERAS
                wbHoja.Cells(intFila, columna).Value = fld.Name
                'Debug.Print fld.Name, fld.Type
            Else
                 m_Valor = Nz(rcdDatos.Fields(fld.Name), "")
                If m_Valor <> "" Then
                    If fld.Type = 7 Then 'es un número doble
                        m_Valor = Replace(m_Valor, ",", ".")
                    ElseIf fld.Type = 8 Then 'tipo fecha
                        m_Valor = Format(m_Valor, "mm/dd/yyyy")
                    End If
                    wbHoja.Cells(intFila, columna).Value = m_Valor
                End If
            End If
            
            columna = columna + 1
        Next
       
        
        intFila = intFila + 1
        rcdDatos.MoveNext
    Loop
    
    WbLibro.Close True
    Set WbLibro = Nothing
    appexcel.Quit
    Set appexcel = Nothing
    Ejecutar 1, "open", m_URLExcel, "", "", 1
    
    GenerarConsultasSQL = m_URLExcel
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Consultas.GenerarConsultasSQL ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
     If Not WbLibro Is Nothing Then
        WbLibro.Close False
        Set WbLibro = Nothing
    End If
    If Not appexcel Is Nothing Then
        appexcel.Quit
        Set appexcel = Nothing
    End If
    
End Function
Public Function GenerarConsultaEventos( _
                                        p_Col As Scripting.Dictionary, _
                                        Optional ByRef p_Error As String _
                                        ) As String
   
    
    Dim m_ID As Variant
    Dim m_Evento As Evento
    Dim m_Campo As Variant
    Dim m_HorasLabTotales As String
    Dim m_HorasExtTotales As String
    On Error GoTo errores
    
    If p_Col Is Nothing Then
        Exit Function
    End If
    
    Dim m_NombreArchivo As String
    
    m_NombreArchivo = fso.GetTempName & ".xlsx"
    Dim m_URLExcel As String
    m_URLExcel = m_ObjEntorno.URLDirectorioLocal & m_NombreArchivo
    If fso.FileExists(m_URLExcel) Then
        If FicheroAbierto(m_URLExcel) Then
            p_Error = "Tiene una consulta abierta"
            Err.Raise 1000
        End If
        fso.DeleteFile m_URLExcel, True
    End If
    
    Set appexcel = New excel.Application
    appexcel.Visible = False
    Dim WbLibro As excel.Workbook
    Dim wbHoja As excel.Worksheet
    Set WbLibro = appexcel.Workbooks.Add
    WbLibro.SaveAs m_URLExcel
    Set wbHoja = WbLibro.Worksheets(1)
    Dim m_Valor As String
    Dim intFila As Integer
    Dim fila As Integer
    Dim columna As Integer
    
    intFila = 1
    
    'CABECERA
    With wbHoja
        .Cells(intFila, 1).Value = "NODO"
        .Cells(intFila, 2).Value = "BUI"
        .Cells(intFila, 3).Value = "SUBSISTEMA"
        .Cells(intFila, 4).Value = "EQUIPO"
        .Cells(intFila, 5).Value = "PMPR"
        .Cells(intFila, 6).Value = "TIPOEVENTO"
        .Cells(intFila, 7).Value = "CRITICIDAD"
        .Cells(intFila, 8).Value = "TÉCNICO"
        .Cells(intFila, 9).Value = "ORIGINADOR"
        .Cells(intFila, 10).Value = "DESCRIPCIÓN"
        .Cells(intFila, 11).Value = "FECHAALTA"
        .Cells(intFila, 12).Value = "CONTACTO"
        .Cells(intFila, 13).Value = "FRANQUEADO"
        .Cells(intFila, 14).Value = "NOTAS"
        .Cells(intFila, 15).Value = "FECHA EN INFORME RAC"
        .Cells(intFila, 16).Value = "CAUSAFIN"
        .Cells(intFila, 17).Value = "FECHA ÚLTIMA ACTIVIDAD"
        .Cells(intFila, 18).Value = "HORAS"
    End With
    intFila = intFila + 1
    For Each m_ID In p_Col
        Debug.Print m_ID
        If CStr(m_ID) = "TEL2406001" Then Stop
        Set m_Evento = p_Col(m_ID)
        
         With wbHoja
            .Cells(intFila, 1).Value = m_Evento.NODO
            .Cells(intFila, 2).Value = m_Evento.BUI
            .Cells(intFila, 3).Value = m_Evento.SubSistema
            If Not m_Evento.Equipo Is Nothing Then
                .Cells(intFila, 4).Value = m_Evento.Equipo.Equipo
            End If
            
            .Cells(intFila, 5).Value = m_Evento.PMPR
            .Cells(intFila, 6).Value = m_Evento.TipoEvento
            .Cells(intFila, 7).Value = m_Evento.Criticidad
            If Not m_Evento.Tecnico Is Nothing Then
                .Cells(intFila, 8).Value = m_Evento.Tecnico.Nombre
            End If
            
            .Cells(intFila, 9).Value = m_Evento.Originador
            .Cells(intFila, 10).Value = m_Evento.Descripcion
            m_Valor = m_Evento.FECHAALTAEVENTO
            If IsDate(m_Valor) Then
                .Cells(intFila, 11).Value = Format(m_Valor, "mm/dd/yyyy")
            End If
            If m_Evento.Contacto <> "" Then
                .Cells(intFila, 12).Value = m_Evento.Contacto
            End If
            
            If m_Evento.Franqueado = True Then
                .Cells(intFila, 13).Value = "Sí"
            Else
                .Cells(intFila, 13).Value = "No"
            End If
            If m_Evento.Notas <> "" Then
                .Cells(intFila, 14).Value = m_Evento.Notas
            End If
            
            m_Valor = m_Evento.FechaEnInformeRAC
            If IsDate(m_Valor) Then
                .Cells(intFila, 15).Value = m_Valor
            End If
            If m_Evento.CausaFin <> "" Then
                .Cells(intFila, 16).Value = m_Evento.CausaFin
            End If
            
             m_Valor = m_Evento.FechaFinalCalculada
            If IsDate(m_Valor) Then
                .Cells(intFila, 17).Value = m_Valor
            End If
            If IsNumeric(m_Evento.HorasLabTotales) Then
                m_HorasLabTotales = m_Evento.HorasLabTotales
            Else
                m_HorasLabTotales = "0"
            End If
            If IsNumeric(m_Evento.HorasExtTotales) Then
                m_HorasExtTotales = m_Evento.HorasExtTotales
            Else
                m_HorasExtTotales = "0"
            End If
            
            m_Valor = CStr(CDbl(m_HorasLabTotales) + CDbl(m_HorasExtTotales))
            .Cells(intFila, 18).Value = Replace(m_Valor, ",", ".")
        End With
        Set m_Evento = Nothing
        intFila = intFila + 1
    Next
        
    WbLibro.Close True
    Set WbLibro = Nothing
    appexcel.Quit
    Set appexcel = Nothing
    Ejecutar 1, "open", m_URLExcel, "", "", 1
    
    GenerarConsultaEventos = m_URLExcel
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Consultas.GenerarConsultaEventos ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
     If Not WbLibro Is Nothing Then
        WbLibro.Close False
        Set WbLibro = Nothing
    End If
    If Not appexcel Is Nothing Then
        appexcel.Quit
        Set appexcel = Nothing
    End If
    
End Function
Public Function DameTipoCampo(ByRef fld As DAO.Field) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 06/11/2019
    '   -Modificaciones:
   
    '   -Funcionamiento:
    '       -se le pasa un campo que ha de venir de un recordset con al menos un registro
    '       - Nos devuelve el tipo del campo
    '   -Llamada desde
    
    '   -Devuelve:
    '       strTipo = "Boolean","Numérico","Texto Corto","Texto Largo","Fecha"
    '       DameTipoCampo = strTipo
    '       GenerarConsultas = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim strTipo As String, strTextoError As String
    On Error GoTo errores
    If fld.Type = 1 Then
        strTipo = "Boolean"
    ElseIf fld.Type = 3 Then
        strTipo = "Numérico"
    ElseIf fld.Type = 4 Then
        strTipo = "Numérico"
    ElseIf fld.Type = 10 Then
        strTipo = "Texto Corto"
    ElseIf fld.Type = 12 Then
        strTipo = "Texto Largo"
    ElseIf fld.Type = 8 Then
        strTipo = "Fecha"
    ElseIf fld.Type = 7 Then
        strTipo = "Numérico"
        'Stop
    Else
        strTextoError = "Tipo no reconocido: " & fld.Type
        Err.Raise 1000
    End If
    DameTipoCampo = strTipo
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método DameTipoCampo ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DameTipoCampo = "#ERR" & "|" & strTextoError
End Function
Public Function DameCamposQueCambian( _
                                        strCadenaCambios As String, _
                                        Optional ByRef col As Collection _
                                        ) As String
     '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 07/11/2019
    '   -Modificaciones:
   
    '   -Funcionamiento:
    '       -se le pasa la cadena de texto que contiene los cambios
    '           'NOMBRECAMPO1::VGestor||VCentro||TipoDato##NOMBRECAMPO2::VGestor||VCentro||TipoDato
    '       - Se va cortando la cadena por el nombre de los campos y se rellena la colección dada
    '   -Llamada desde
    
    '   -Devuelve:
    '       strCadenaCampos=strCampo1 & "#" &  strCampo2 & "#" & ...
    '       DameCamposQueCambian = strCadenaCampos
    '       DameCamposQueCambian = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim VarItem As Variant, dato1 As Variant, strNombreCampo As String, strCadenaCampos As String, strTextoError As String
    On Error GoTo errores
    If Not col Is Nothing Then
        Set col = New Collection
    End If
    '----------------------------------
    ' OBTENER LOS CAMPOS QUE CAMBIAN
    '----------------------------------
    'NOMBRECAMPO1::VGestor||VCentro||TipoDato##NOMBRECAMPO2::VGestor||VCentro||TipoDato
    dato = Split(strCadenaCambios, "##")
    For Each VarItem In dato
        dato1 = Split(VarItem, "::")
        strNombreCampo = dato1(0)
        If strCadenaCampos = "" Then
            strCadenaCampos = strNombreCampo
        Else
            strCadenaCampos = strCadenaCampos & "#" & strNombreCampo
        End If
        If Not col Is Nothing Then
            col.Add strNombreCampo
        End If
    Next
    DameCamposQueCambian = strCadenaCampos
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método DameCamposQueCambian ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DameCamposQueCambian = "#ERR" & "|" & strTextoError
End Function
Public Function DameIDTecnico(strTipoParaFacturacion As String) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 15/11/2019
    '   -Modificaciones:
   
    '   -Funcionamiento:
    '       -Con strTipoParaFacturacion vamos a la tabla de con el campo TipoParaFacturacion=strTipoParaFacturacion
    '           TbTIPOTECNICO y le ponemos uno más al ordinal
    '   -Llamada desde
    
    '   -Devuelve:
    '       DameIDTecnico = strIDTecnico
    '       DameIDTecnico = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, intOrdinalMaximo As Integer, intOrdinal As Integer, strIDTecnicoActual As String, _
        strIDTecnico As String, strParticula As String, strTextoError As String
    On Error GoTo errores
    If strTipoParaFacturacion = "" Then
        strTextoError = "Se ha de indicar un tipo"
        Err.Raise 1000
    End If
    m_SQL = "SELECT DISTINCT TbTipoTecnico.TIPOTECNICO " & _
            "FROM TbTipoTecnico " & _
            "WHERE (((TbTipoTecnico.TipoParaFacturacion)='" & strTipoParaFacturacion & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strParticula = Left(strTipoParaFacturacion, 3)
            strIDTecnico = strParticula & "-1"
            '-----------------------------------------------
            '   -Devuelve:
            '       Dame = strDescriptivo
            '       Dame = "#ERR" & "|" & strTextoError
            '-----------------------------------------------
            flag = Dame("TbTIPOTECNICO", "TIPOTECNICO", "TIPOTECNICO", strIDTecnico)
            If InStr(1, flag, "|") <> 0 Then
                dato = Split(flag, "|")
                strTextoError = "El método Dame ha devuelto un error: " & vbNewLine & dato(1)
                Err.Raise 1000
            End If
            If flag <> "" Then
                strTextoError = "El Sistema no ha podido generar un IDTécnico válido"
                Err.Raise 1000
            End If
            rcdDatos.Close
            Set rcdDatos = Nothing
            DameIDTecnico = strIDTecnico
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            strIDTecnicoActual = .Fields("TIPOTECNICO")
            If InStr(1, strIDTecnicoActual, "-") Then
                dato = Split(strIDTecnicoActual, "-")
                strParticula = dato(0)
                If IsNumeric(dato(1)) Then
                    intOrdinal = CInt(dato(1))
                    If intOrdinal > intOrdinalMaximo Then
                        intOrdinalMaximo = intOrdinal
                    End If
                End If
            End If
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If strParticula = "" Then
        strTextoError = "El Sistema no ha podido generar un IDTécnico válido"
        Err.Raise 1000
    End If
    strIDTecnico = strParticula & "-" & intOrdinalMaximo + 1
    DameIDTecnico = strIDTecnico
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método DameIDTecnico ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    DameIDTecnico = "#ERR" & "|" & strTextoError
End Function
Public Function EstadoPlanificacion( _
                                         strIDPlanificacion As String _
                                         ) As String

    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 17/06/2014
    '   -Modificaciones:
    '
    '   -Funcionamiento:
    '       -strIDPlanificacion existente y numérica
    '       -Obtenemos el valor de FechaCierre,MotivoCierre
    '           -si FechaCierre=NUll-->EstadoPlanificacion="Abierta"
    '           -si FechaCierre<>NUll
    '               -MotivoCierre="Trabajo Realizado" --->EstadoPlanificacion ="Cerrada por trabajo realizado"
    '               -MotivoCierre="Reprogramación" --->EstadoPlanificacion ="Cerrada por reprogramación"
    
    
    '   -llamada por:
    '       -CerrarPlanificacionSinTrabajos
    '       -Form_FormPlanificacionBusqueda.ComandoTrabajoQueLaCierra_Click
    '       -Form_FormPlanificacionAnexos.establecerbotonera
    '       -Form_FormPlanificacionAnexos.ListaDocumentos_Click
    '   -Devuelve:
    '       strEstado = "Abierta"
    '       strEstado="Cerrada por trabajo realizado"
    '       strEstado="Cerrada por reprogramación"
    '       EstadoPlanificacion=strEstado
    '       EstadoPlanificacion="#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, strFechaCierre, strMotivoCierre As String, strEstado As String, strTextoError As String
    On Error GoTo errores
    If Not IsNumeric(strIDPlanificacion) Then
        strTextoError = "Se ha de indicar el IDPlanificación"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbPlanificacion.* " & _
            "FROM TbPlanificacion " & _
            "WHERE (((TbPlanificacion.IDPlanificacion)=" & strIDPlanificacion & "));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strTextoError = "Planificación no registrada"
            Err.Raise 1000
        End If
        strFechaCierre = Nz(.Fields("FechaCierre"), "")
        strMotivoCierre = Nz(.Fields("MotivoCierre"), "")
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If Not IsDate(strFechaCierre) Then
        strEstado = "Abierta"
    Else
        If strMotivoCierre = "Trabajo Realizado" Then
            strEstado = "Cerrada por trabajo realizado"
        ElseIf strMotivoCierre = "Reprogramación" Then
            strEstado = "Cerrada por reprogramación"
        Else
            strTextoError = "Estado de la planificación desconocida"
            Err.Raise 1000
        End If
    End If
    EstadoPlanificacion = strEstado
    Exit Function
errores:
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    EstadoPlanificacion = "#ERR" & "|" & strTextoError
End Function
Public Function EquipoPlanificable(p_IDEquipo As String) As String
     '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 17/06/2014
    '   -Modificaciones:
    '
    '   -Funcionamiento:
    '       -va a mirar en la tabla TbPlanificacionEquipos para ver si está
    '
   
    '   -llamada por:
    '       -me.AltaPlanificacion
    '       -me.SacarEquipoDePlanificacion
    '       -me.AltaEquipo
    '       -Form_FormPlanificacionAltaEquipo.ComandoAltaMtoPrevNoDeEvento_Click
    '       -Form_FormIDEquiposGestion.ComandoEstablecerParaPlanificacion_Click
    '   -Devuelve:
    '       EquipoPlanificable = strEquipoPlanificable
    '       EquipoPlanificable = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim strIDEquipoEnPlanificable As String, strFechaBaja As String, strFechaObsoleto As String, _
        strEquipoPlanificable As String, strTextoError As String
    On Error GoTo errores
    '------------------------------------------
    '   -Devuelve:
    '       Dame = strValorObtenido
    '       Dame = "#ERR" & "|" & strTextoError
    '--------------------------------------------
    flag = Dame("TbEquipos", "FechaObsoleto", "IDEquipo", p_IDEquipo)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Dame ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strFechaObsoleto = flag
    If IsDate(strFechaObsoleto) Then
        strEquipoPlanificable = "No"
        EquipoPlanificable = strEquipoPlanificable
        Exit Function
    End If
    '------------------------------------------
    '   -Devuelve:
    '       Dame = strValorObtenido
    '       Dame = "#ERR" & "|" & strTextoError
    '--------------------------------------------
    flag = Dame("TbPlanificacionEquipos", "IDEquipo", "IDEquipo", p_IDEquipo)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Dame ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strIDEquipoEnPlanificable = flag
    If Not IsNumeric(strIDEquipoEnPlanificable) Then
        strEquipoPlanificable = "No"
        EquipoPlanificable = strEquipoPlanificable
        Exit Function
    End If
    '------------------------------------------
    '   -Devuelve:
    '       Dame = strValorObtenido
    '       Dame = "#ERR" & "|" & strTextoError
    '--------------------------------------------
    flag = Dame("TbPlanificacionEquipos", "FechabajaParaPlanificacion", "IDEquipo", p_IDEquipo)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Dame ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strFechaBaja = flag
    If IsDate(strFechaBaja) Then
        strEquipoPlanificable = "No"
    Else
        strEquipoPlanificable = "Sí"
    End If
    EquipoPlanificable = "#ERR" & "|" & strTextoError
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método EquipoPlanificable ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    EquipoPlanificable = "#ERR" & "|" & strTextoError
End Function
Public Function PlanificacionYaExistente( _
                                            strIDEquipo As String, _
                                            intSemana As Integer, _
                                            intMes As Integer, _
                                            intAño As Integer _
                                            ) As String
     '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 17/06/2014
    '   -Modificaciones:
    '
    '   -Funcionamiento:
    '       -va a mirar en la tabla TbPlanificacionEquipos para ver si está
    '
   
    '   -llamada por:
    '       -me.AltaPlanificacion
    '       -me.CerrarPlanificacionSinTrabajos
    '   -Devuelve:
    '       PlanificacionYaExistente = strPlanificacionYaExistente
    '       PlanificacionYaExistente = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, strTextoError As String
    On Error GoTo errores
    If Not IsNumeric(strIDEquipo) Then
        strTextoError = "No se ha indicado el IDEquipo"
        Err.Raise 1000
    End If
    If intSemana < 1 Or intSemana > 6 Then
        strTextoError = "La semana va entre 1 y 6"
        Err.Raise 1000
    End If
    If intMes < 1 Or intMes > 12 Then
        strTextoError = "El mes va entre 1 y 12"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbPlanificacion.IDPlanificacion " & _
                "FROM TbPlanificacion " & _
                "WHERE (((TbPlanificacion.IDEquipo)=" & strIDEquipo & _
                ") AND ((TbPlanificacion.Semana)=" & intSemana & _
                ") AND ((TbPlanificacion.Mes)=" & intMes & _
                ") AND ((TbPlanificacion.Anio)=" & intAño & _
                "));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            PlanificacionYaExistente = "No"
        Else
            PlanificacionYaExistente = "Sí"
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método PlanificacionYaExistente ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    PlanificacionYaExistente = "#ERR" & "|" & strTextoError
End Function
Public Function AltaPlanificacion( _
                                    strIDEquipo As String, _
                                    intSemana As Integer, _
                                    intMes As Integer, _
                                    intAño As Integer _
                                    ) As String
     '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 17/06/2014
    '   -Modificaciones:
    '
    '   -Funcionamiento:
    '       -va a dar de AltaPlanificacion un registro en la tabla TbPlanificacion
    '       1) VIABILIDAD DE LAS VARIABLES DE ENTRADA
    '           -Equipo Planificable--->EquipoPlanificable
    '           -intMes < 1 Or intMes > 12
    '           -intSemana < 1 Or intSemana > intSemanasTieneElMes----->NumeroSemanasTieneElMes
    '           -strFechaViernes que sea fecha--->FechaViernes---->strFechaPrevistaCierre
    '           -PlanificacionYaExistente
    '       2) ACTUACIÓN
    '           -strIDPlanificacion--->DameID
    '           -AltaPlanificacion en la tabla TbPlanificacion
    '   -llamada por:
    '       -me.CerrarPlanificacionSinTrabajos
    '       -Form_FormPlanificacionAlta.ComandoAlta_Click
    '   -Devuelve:
    '       AltaPlanificacion = strIDPlanificacion & ";" & strFechaPrevistaCierre
    '       AltaPlanificacion = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, strFechaPrevistaCierre As String, strFechaRegistro As String, _
        strIDPlanificacion As String, intSemanasTieneElMes As Integer, strTextoError As String
    On Error GoTo errores
    If Not IsNumeric(strIDEquipo) Then
        strTextoError = "No se ha introducido un equipo"
        Err.Raise 1000
    End If
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       EquipoPlanificable = strEquipoPlanificable
    '       EquipoPlanificable = "#ERR" & "|" & strTextoError
    '-----------------------------------------------------------
    flag = EquipoPlanificable(strIDEquipo)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método EquipoPlanificable ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If flag = "No" Then
        strTextoError = "Este equipo no está marcado como planificable"
        Err.Raise 1000
    End If
    If intMes < 1 Or intMes > 12 Then
        strTextoError = "El mes ha de estar entre 1 y 12"
        Err.Raise 1000
    End If
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       NumeroSemanasTieneElMes =intNumeroDiasTieneElMes
    '       NumeroSemanasTieneElMes=-1 si error
    '-------------------------------------------------------------------
    intSemanasTieneElMes = NumeroSemanasTieneElMes(intMes, intAño)
    If intSemanasTieneElMes = -1 Then
        strTextoError = "la función que calcula las semanas que tiene un mes ha dado un error desconocido"
        Err.Raise 1000
    End If
    If intSemana < 1 Or intSemana > intSemanasTieneElMes Then
        strTextoError = "La semana de planificación ha de estar entre 1 y el número de semanas que tiene este mes: " & CStr(intSemanasTieneElMes)
        Err.Raise 1000
    End If
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       FechaViernes =dteFechaViernesSemanax para intSemana=x <>0
    '       FechaViernes =strCadenaFechas para intSemana=0
    '       strCadenaFechas=01 & dteFechaViernesSemana1 & vbcrlf & _
                            02 & dteFechaViernesSemana2 & vbcrlf & _
    '                       ...
    '                       intNumeroSemanas & dteFechaViernesSemanax
    '       FechaViernes="#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = FechaViernes(intMes, intAño, intSemana)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método FechaViernes ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strFechaPrevistaCierre = flag
    If strFechaPrevistaCierre = "0" Or Not IsDate(strFechaPrevistaCierre) Then
        strTextoError = "La semana " & CStr(intSemana) & " del mes " & CStr(intMes) & " del año " & CStr(intAño) & " no tiene viernes."
        Err.Raise 1000
    End If
    '------------------------
    ' VERIFICACIÓN DE QUE NO HAY OTRA PLANIFICACIÓN PARA LA MISMA SEMANA
    '-----------------
     '   -Devuelve:
    '       PlanificacionYaExistente = strPlanificacionYaExistente
    '       PlanificacionYaExistente = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = PlanificacionYaExistente(strIDEquipo, intSemana, intMes, intAño)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método PlanificacionYaExistente ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If flag = "Sí" Then
        strTextoError = "La semana " & CStr(intSemana) & " del mes " & CStr(intMes) & " del año " & CStr(intAño) & " ya está planificada para este equipo."
        Err.Raise 1000
    End If
    '---------------------
    ' ACTUACIÓN
    '--------------
     '------------------------------------------
    '   -Devuelve:
    '       DameID = CStr(lngMaxID + 1)
    '       DameID = "#ERR" & "|" & strTextoError
    '--------------------------------------------
    flag = DameID("TbPlanificacion", "IDPlanificacion")
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método DameID ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strIDPlanificacion = flag
    If Not IsNumeric(strIDPlanificacion) Then
        strTextoError = "La función DAMEID no ha devuelto un valor con formato conocido"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbPlanificacion.* " & _
            "FROM TbPlanificacion " & _
            "WHERE (((TbPlanificacion.IDPlanificacion)=" & strIDPlanificacion & "));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            strTextoError = "La función DAMEID no ha devuelto un valor existente."
            Err.Raise 1000
        End If
        .AddNew
            .Fields("IDPlanificacion") = strIDPlanificacion
            .Fields("FechaRegistro") = Now()
            .Fields("IDEquipo") = strIDEquipo
            .Fields("Anio") = intAño
            .Fields("Mes") = intMes
            .Fields("Semana") = intSemana
            .Fields("FechaPrevistaCierre") = strFechaPrevistaCierre
            
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    AltaPlanificacion = strIDPlanificacion & ";" & strFechaPrevistaCierre
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Planificacion.AltaPlanificacion ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    AltaPlanificacion = "#ERR" & "|" & strTextoError
End Function
Public Function AnexarAPlanificacion( _
                                        strIDPlanificacion As String, _
                                        strURLAnexo As String, _
                                        Optional strNombreAnexo As String _
                                        ) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 17/06/2014
    '   -Modificaciones:
    '
    '   -Funcionamiento:
    '       -va a crear un registro en la tabla TbPlanificacionAnexos
    '           -strMotivoCierre="Reprogramación"
    '       1) VIABILIDAD DE LAS VARIABLES DE ENTRADA
    '           -strIDPlanificacion numérico y existente
    '           -strURLAnexo alcanzable
    
    '       2) ACTUACIÓN
    '           -obtención del nombre del anexo final
    '           strNombreAnexoResultante =strIDPlanificacionFormateada & "_" & cstr(cint(strNumeroAnexosDeEstaPlanificacion)+1 ) & "." & strExtension
    '   -llamada por:
    '       -me.CerrarPlanificacionSinTrabajos
    '       -Form_FormPlanificacionAnexos.AñadirDocumento
    '   -Devuelve:
    '       AnexarAPlanificacion = strURLRemoto
    '       AnexarAPlanificacion = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, strURLRemoto As String, strExtension As String, intNumeroAnexos As Integer, _
        strID As String, strTextoError As String
    On Error GoTo errores
    If Not IsNumeric(strIDPlanificacion) Then
        strTextoError = "El IDPlanificación ha de ser un número"
        Err.Raise 1000
    End If
    If Not fso.FileExists(strURLAnexo) Then
        strTextoError = "El Anexo debe ser alcanzable"
        Err.Raise 1000
    End If
    
     
     If strNombreAnexo = "" Then
        strNombreAnexo = fso.GetFileName(strURLAnexo)
     End If
     strExtension = fso.GetExtensionName(strURLAnexo)
     '-------------------------------------------------------------------
     '   -Devuelve:
     '       NumeroAnexosEnPlanificacion = intNumeroAnexos
     '       NumeroAnexosEnPlanificacion = -1 si Error
     '-------------------------------------------------------------------
     intNumeroAnexos = NumeroAnexosEnPlanificacion(strIDPlanificacion)
     If intNumeroAnexos = -1 Then
         strTextoError = "El método Planificacion.NumeroAnexosEnPlanificacion ha devuelto un error desconocido"
         Err.Raise 1000
     End If
     strNombreAnexo = fso.GetBaseName(strNombreAnexo) & "_" & Format("00", CStr(intNumeroAnexos + 1)) & "." & strExtension
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       YaExisteEseNombreParaEseAnexo ="Sí" o "No"
    '       YaExisteEseNombreParaEseAnexo ="#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = YaExisteEseNombreParaEseAnexo(strIDPlanificacion, strNombreAnexo)
    If InStr(1, flag, "|") <> 0 Then
       dato = Split(flag, "|")
       strTextoError = "El método Eliminar ha devuelto un error: " & vbNewLine & dato(1)
       Err.Raise 1000
    End If
    If flag = "Sí" Then
        strTextoError = "El nombre obtenido por el sistema: " & strNombreAnexo & " ya está registrado previamente"
        Err.Raise 1000
    End If
    strURLRemoto = m_ObjEntorno.URLDirectorioPlanificaciones & strNombreAnexo
    If fso.FileExists(strURLRemoto) Then
        '----------------------------------------
        '   -Devuelve:
        '       FicheroAbierto = true or false
        '       FicheroAbierto =False-->Error
        '----------------------------------------
         If FicheroAbierto(strURLRemoto) Then
             strTextoError = "Debe estar abierto el actual archivo de planificación"
             Err.Raise 1000
         End If
    End If
    fso.CopyFile strURLAnexo, strURLRemoto, True
    '------------------------------------------
    '   -Devuelve:
    '       DameID = CStr(lngMaxID + 1)
    '       DameID = "#ERR" & "|" & strTextoError
    '--------------------------------------------
    flag = DameID("TbPlanificacionAnexos", "IDAnexoPlanificacion")
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método DameID ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strID = flag
    If Not IsNumeric(strID) Then
        strTextoError = "El método DameID ha devuelto un ID con formato no reconocible"
        Err.Raise 1000
    End If
     m_SQL = "SELECT TbPlanificacionAnexos.* " & _
             "FROM TbPlanificacionAnexos " & _
             "WHERE (((TbPlanificacionAnexos.IDAnexoPlanificacion)=" & strID & "));"
     Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
     With rcdDatos
        .AddNew
            .Fields("IDAnexoPlanificacion") = strID
            .Fields("IDPlanificacion") = strIDPlanificacion
            .Fields("NombreAnexo") = strNombreAnexo
            .Fields("UsuarioAnexa") = m_ObjUsuarioConectado.UsuarioRed
            .Fields("FechaAnexo") = Now()
            .Fields("URLAnexo") = fso.GetFileName(strURLRemoto)
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    AnexarAPlanificacion = strURLRemoto
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método AnexarAPlanificacion ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    
    AnexarAPlanificacion = "#ERR" & "|" & strTextoError
End Function
Public Function YaExisteEseNombreParaEseAnexo( _
                                                strIDPlanificacion As String, _
                                                strNOMBRE As String _
                                                ) As String
    '-------------------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 17/06/2014
    '   -Modificaciones:
    '
  
    '   -llamada por:
    '       -Planificacion.AnexarAPlanificacion
    '       -Form_FormPlanificacionAnexos.AñadirDocumento
    '   -Devuelve:
    '       YaExisteEseNombreParaEseAnexo ="Sí" o "No"
    '       YaExisteEseNombreParaEseAnexo ="#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, strTextoError As String
    On Error GoTo errores
    If Not IsNumeric(strIDPlanificacion) Then
        strTextoError = "La Planificación ha de ser numérica"
        Err.Raise 1000
    End If
    If strNOMBRE = "" Then
        strTextoError = "Se ha de indicar el nombre"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbPlanificacionAnexos.IDPlanificacion, TbPlanificacionAnexos.NombreAnexo " & _
            "FROM TbPlanificacionAnexos " & _
            "WHERE (((TbPlanificacionAnexos.IDPlanificacion)=" & strIDPlanificacion & _
            ") AND ((TbPlanificacionAnexos.NombreAnexo)='" & strNOMBRE & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            YaExisteEseNombreParaEseAnexo = "No"
        Else
            YaExisteEseNombreParaEseAnexo = "Sí"
        End If
       
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método YaExisteEseNombreParaEseAnexo ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    YaExisteEseNombreParaEseAnexo = "#ERR" & "|" & strTextoError
End Function
Public Function NPlanificaciones( _
                                    Optional intAño As Integer, _
                                    Optional intIDEquipo As Integer, _
                                    Optional strEstadoPlanificacion As String, _
                                    Optional intDiasAviso As Integer, _
                                    Optional strMotivoCierre As String _
                                    ) As Integer
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 23/06/2014
    '   -Modificaciones:
    '
    '   -Funcionamiento:
    '       -Va a contar los elementos de una consulta con los parámetros introducidos
    '
   
    '   -llamada por:
    '       -Form_FormPlanificacionGestion.ComandoActualizar_Click
    '       -Form_FormPlanificacionEquipos.EstablecerBotonera
    '   -Devuelve:
    '       NPlanificaciones = clng(strNumeroPlanificaciones)
    '       NPlanificaciones =-1--->Si Error
    '-------------------------------------------------------------------
    Dim strParteWhere As String, strParteWhereAños As String, strParteWhereIDEquipo As String, strParteMotivoCierre As String, strParteWhereDifFechas As String
    Dim strSQLPartePrincipal As String, strSQLFinal As String, strNumeroPlanificaciones As String
    Dim intAño1 As Integer, intAño2 As Integer, strTextoError As String
    On Error GoTo errores
    
    If strEstadoPlanificacion <> "Realizadas" And strEstadoPlanificacion <> "No Realizadas" And strEstadoPlanificacion <> "Fuera de Plazo" And _
        strEstadoPlanificacion <> "Cumplen Próximamente" And strEstadoPlanificacion <> "En Plazo" And strEstadoPlanificacion <> "Totales" Then
        Err.Raise 1000
    End If
    strSQLPartePrincipal = "SELECT TbPlanificacion.IDPlanificacion " & _
                            "FROM TbPlanificacion "
    If intAño = 0 Then
        intAño1 = 1900
        intAño2 = 2100
    Else
        intAño1 = intAño
        intAño2 = intAño
    End If
    strParteWhereAños = "((TbPlanificacion.Anio) Between " & CStr(intAño1) & " And " & CStr(intAño2) & ")"
    If intIDEquipo = 0 Then
        strParteWhereIDEquipo = "((TbPlanificacion.IDEquipo) Like '*' Or (TbPlanificacion.IDEquipo) Is Null)"
    Else
        strParteWhereIDEquipo = "((TbPlanificacion.IDEquipo)=" & CStr(intIDEquipo) & ")"
    End If
    If strEstadoPlanificacion = "Realizadas" Then
        strParteMotivoCierre = "(Not (TbPlanificacion.MotivoCierre) Is Null)"
        strParteWhereDifFechas = "((DateDiff('d',Now(),[FechaPrevistaCierre])) Like '*')"
    ElseIf strEstadoPlanificacion = "No Realizadas" Then
        strParteMotivoCierre = "((TbPlanificacion.MotivoCierre) Is Null)"
        strParteWhereDifFechas = "((DateDiff('d',Now(),[FechaPrevistaCierre])) Like '*')"
    ElseIf strEstadoPlanificacion = "Fuera de Plazo" Then
        strParteMotivoCierre = "((TbPlanificacion.MotivoCierre) Is Null)"
        strParteWhereDifFechas = "((DateDiff('d',Now(),[FechaPrevistaCierre]))<0)"
    ElseIf strEstadoPlanificacion = "Cumplen Próximamente" Then
        strParteMotivoCierre = "((TbPlanificacion.MotivoCierre) Is Null)"
        strParteWhereDifFechas = "((IIf(IsNull([FechaCierre]),DateDiff('d',Now(),[fechaprevistacierre]),0)) Between 0 And " & CInt(m_ObjEntorno.DiasParaAvisoDePlanificaciones) & ")"
    ElseIf strEstadoPlanificacion = "En Plazo" Then
        strParteMotivoCierre = "((TbPlanificacion.MotivoCierre) Is Null)"
        strParteWhereDifFechas = "((DateDiff('d',Now(),[FechaPrevistaCierre]))>0)"
    ElseIf strEstadoPlanificacion = "Totales" Then
        strParteMotivoCierre = "((TbPlanificacion.MotivoCierre) Like '*' Or (TbPlanificacion.MotivoCierre) Is Null)"
        strParteWhereDifFechas = "((DateDiff('d',Now(),[FechaPrevistaCierre],)) Like '*')"
    ElseIf strEstadoPlanificacion = "" Then
        strParteMotivoCierre = "((TbPlanificacion.MotivoCierre) Like '*' Or (TbPlanificacion.MotivoCierre) Is Null)"
        strParteWhereDifFechas = "((DateDiff('d',Now(),[FechaPrevistaCierre],)) Like '*')"
    End If
    strSQLFinal = strSQLPartePrincipal & _
                "WHERE (" & strParteWhereAños & " AND " & strParteWhereAños & " AND " & strParteWhereIDEquipo & " AND " & strParteMotivoCierre & _
                " AND " & strParteWhereDifFechas & ");"
     '-------------------------------------------------------------
    '   -Devuelve:
    '       DameNumeroRegistrosPorSQL = CStr(lngRegistros)
    '       DameNumeroRegistrosPorSQL = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = DameNumeroRegistrosPorSQL(strSQLFinal)
    If Not IsNumeric(flag) Then
        Err.Raise 1000
    End If
    NPlanificaciones = CInt(flag)
    Exit Function
errores:
   NPlanificaciones = -1
End Function
Public Function NumeroSemanasTieneElMes(intMes As Integer, intAño As Integer) As Integer
     '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 17/06/2014
    '   -Modificaciones:
    
    '   -Funcionamiento:
    
   
    '   -Llamada por:
    '       -Planificacion.AltaPlanificacion
    '       -Planificacion.CerrarPlanificacionSinTrabajos
    '   -Devuelve:
    '       NumeroSemanasTieneElMes =intNumeroDiasTieneElMes
    '       NumeroSemanasTieneElMes=-1 si error
    '-------------------------------------------------------------------
    Dim dteFecha As Date, dteFechaSiguienteMes As Date, intSemana As Integer, intDiaPrimerDiasemana As Integer
    Dim intNumeroDiasTieneElMes As Integer, intDias As Integer
    Dim dteFechaPrimerDiaSemana1 As Date, dteFechaPrimerDiaSemana2 As Date, dteFechaPrimerDiaSemana3 As Date, dteFechaPrimerDiaSemana4 As Date, _
        dteFechaPrimerDiaSemana5 As Date, dteFechaPrimerDiaSemana6 As Date
    On Error GoTo errores
    dteFecha = CDate("01/" & CStr(intMes) & "/" & CStr(intAño))
    dteFechaSiguienteMes = DateAdd("m", 1, dteFecha)
    intNumeroDiasTieneElMes = DateDiff("d", dteFecha, dteFechaSiguienteMes, vbMonday)
    intDias = 0
    '-------------------
    ' SEMANA 1
    '-----------------
        dteFechaPrimerDiaSemana1 = dteFecha
        intSemana = 1
        intDiaPrimerDiasemana = Weekday(dteFechaPrimerDiaSemana1, vbMonday)
    '-------------------
    ' SEMANA 2
    '-----------------
        intDias = 8 - intDiaPrimerDiasemana
        dteFechaPrimerDiaSemana2 = DateAdd("d", intDias, dteFechaPrimerDiaSemana1)
        intSemana = intSemana + 1
    '-------------------
    ' SEMANA 3
    '-----------------
        intDias = Day(dteFechaPrimerDiaSemana2) + 7
        dteFechaPrimerDiaSemana3 = CDate(CStr(intDias) & "/" & CStr(intMes) & "/" & CStr(intAño))
        intSemana = intSemana + 1
    '-------------------
    ' SEMANA 4
    '-----------------
        intDias = Day(dteFechaPrimerDiaSemana3) + 7
        dteFechaPrimerDiaSemana4 = CDate(CStr(intDias) & "/" & CStr(intMes) & "/" & CStr(intAño))
        intSemana = intSemana + 1
    '-------------------
    ' SEMANA 5
    '-----------------
        If intNumeroDiasTieneElMes <= intDias + 7 Then
            NumeroSemanasTieneElMes = intSemana
            Exit Function
        End If
        If intNumeroDiasTieneElMes - intDias >= 7 Then
            intDias = Day(dteFechaPrimerDiaSemana4) + 7
            dteFechaPrimerDiaSemana5 = CDate(CStr(intDias) & "/" & CStr(intMes) & "/" & CStr(intAño))
            intSemana = intSemana + 1
        Else
            intDias = Day(dteFechaPrimerDiaSemana4) + intNumeroDiasTieneElMes - intDias
            dteFechaPrimerDiaSemana5 = CDate(CStr(intDias) & "/" & CStr(intMes) & "/" & CStr(intAño))
            intSemana = intSemana + 1
        End If
    '-------------------
    ' SEMANA 6
    '-----------------
        If intNumeroDiasTieneElMes < intDias + 7 Then
            NumeroSemanasTieneElMes = intSemana
            Exit Function
        End If
        If intNumeroDiasTieneElMes - intDias >= 7 Then
            intDias = Day(dteFechaPrimerDiaSemana5) + 7
            dteFechaPrimerDiaSemana6 = CDate(CStr(intDias) & "/" & CStr(intMes) & "/" & CStr(intAño))
            intSemana = intSemana + 1
        Else
            intDias = Day(dteFechaPrimerDiaSemana5) + intNumeroDiasTieneElMes - intDias
            dteFechaPrimerDiaSemana6 = CDate(CStr(intDias) & "/" & CStr(intMes) & "/" & CStr(intAño))
            intSemana = intSemana + 1
        End If
        NumeroSemanasTieneElMes = intSemana
    Exit Function
errores:
    NumeroSemanasTieneElMes = -1
End Function
Public Function NumeroAnexosEnPlanificacion(strIDPlanificacion As String) As Integer
     '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 17/06/2014
    '   -Modificaciones:
    '
    '   -Funcionamiento:
    '       -va a contar el  número de anexos que tiene para un determinado ID en la tabla TbPlanificacionAnexos
    '   -llamada por:
    '       -me.AnexarAPlanificacion
    '   -Devuelve:
    '       NumeroAnexosEnPlanificacion = intNumeroAnexos
    '       NumeroAnexosEnPlanificacion = -1 si Error
    '-------------------------------------------------------------------
    On Error GoTo errores
    If Not IsNumeric(strIDPlanificacion) Then
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbPlanificacionAnexos.IDPlanificacion " & _
            "FROM TbPlanificacionAnexos " & _
            "WHERE (((TbPlanificacionAnexos.IDPlanificacion)=" & strIDPlanificacion & "));"
    '-------------------------------------------------------------
    '   -Devuelve:
    '       DameNumeroRegistrosPorSQL = CStr(lngRegistros)
    '       DameNumeroRegistrosPorSQL = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = DameNumeroRegistrosPorSQL(m_SQL)
    If Not IsNumeric(flag) Then
        Err.Raise 1000
    End If
    NumeroAnexosEnPlanificacion = CInt(flag)
    Exit Function
errores:
    NumeroAnexosEnPlanificacion = -1
End Function
Public Function AltaEquipoEnPlanificacion( _
                                            strIDEquipo As String, _
                                            Optional strAlias As String, _
                                            Optional strFechaAlta As String, _
                                            Optional intNumeroMesesMto As Integer _
                                            ) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 12/06/2014
    '   -Modificaciones:
    '
    '   -Funcionamiento:
    '       -va marcar un equipo como planificable metiéndolo en la tabla TbPlanificacionEquipos
    '
    
    '   -llamada por:
    '       -Form_FormPlanificacionAltaEquipo.ComandoAlta_Click
    '   -Devuelve:
    '       AltaEquipoEnPlanificacion = Descriptivo
    '       AltaEquipoEnPlanificacion = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, strTextoError As String
    On Error GoTo errores
    If Not IsNumeric(strIDEquipo) Then
        strTextoError = "El Equipo lo identifica un identificador numérico"
        Err.Raise 1000
    End If
    If strFechaAlta <> "" Then
        If Not IsDate(strFechaAlta) Then
            strTextoError = "Fecha de Alta incorrecta"
            Err.Raise 1000
        End If
    Else
        strFechaAlta = Date
    End If
    m_SQL = "SELECT TbPlanificacionEquipos.* " & _
            "FROM TbPlanificacionEquipos " & _
            "WHERE (((TbPlanificacionEquipos.IDEquipo)=" & strIDEquipo & "));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            strTextoError = "No se puede dar de alta a un equipo que ya  forma parte de los planificables"
            Err.Raise 1000
        End If
        .AddNew
            .Fields("IDEquipo") = strIDEquipo
            .Fields("FechaAltaParaPlanificacion") = strFechaAlta
            If strAlias <> "" Then
                .Fields("Alias") = strAlias
            End If
            If intNumeroMesesMto <> 0 Then
                .Fields("PeriodicidadEnMesesRecomendada") = intNumeroMesesMto
            End If
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    AltaEquipoEnPlanificacion = "Equipo dado de alta para las planificaciones"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método AltaEquipoEnPlanificacion ha producido el error nº: " & Err.Number & vbCrLf & _
        "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    AltaEquipoEnPlanificacion = "#ERR" & "|" & strTextoError
End Function
Public Function DameCadenaAnexos(strIDPlanificacion As String) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 23/06/2014
    '   -Modificaciones:
    '
    '   -Funcionamiento:
    '       -Va a mirar en la tabla TbPlanificacionAnexos para el ID=strIDPlanificacion y los concatena con ;
    '   -llamada por:
    '       -Form_FormPlanificacionBusqueda.EstablecerBotonera
    '   -Devuelve:
    '       strURLCadenaAnexos=strURLAnexo1 & ";" & strURLAnexo2 & ";" & .... & strURLAnexoN
    '       DameCadenaAnexos = strURLCadenaAnexos
    '       DameCadenaAnexos = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, strNombreArchivo As String, _
        strURLCadenaAnexos As String, strURLAnexo As String, strTextoError As String
    On Error GoTo errores
    If Not IsNumeric(strIDPlanificacion) Then
        strTextoError = "No se ha introducido un valor numérico en el IDPlanificación"
        Err.Raise 1000
    End If
    
    
    m_SQL = "SELECT TbPlanificacionAnexos.URLAnexo " & _
                "FROM TbPlanificacionAnexos " & _
                "WHERE (((TbPlanificacionAnexos.IDPlanificacion)=" & strIDPlanificacion & "));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                strURLAnexo = Nz(.Fields("URLAnexo"), "")
                If strURLAnexo <> "" Then
                    If InStr(1, strURLAnexo, ":\") <> 0 Or InStr(1, strURLAnexo, "\\") <> 0 Then
                        strNombreArchivo = fso.GetFileName(strURLAnexo)
                    Else
                        strNombreArchivo = strURLAnexo
                    End If
                    strURLAnexo = m_ObjEntorno.URLDirectorioPlanificaciones & strNombreArchivo
                    If strURLCadenaAnexos = "" Then
                        strURLCadenaAnexos = strURLAnexo
                    Else
                        strURLCadenaAnexos = strURLCadenaAnexos & ";" & strURLAnexo
                    End If
                End If
                .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    DameCadenaAnexos = strURLCadenaAnexos
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método DameCadenaAnexos ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    
    DameCadenaAnexos = "#ERR" & "|" & strTextoError
End Function
Public Function EliminarPlanificacion(strIDPlanificacion As String) As String
    'sólo si no se ha completado el alta de planificación cuando reprogramamos otra
     '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 24/06/2014
    '   -Modificaciones:
    '
    '   -Funcionamiento:
    '       -Va a eliminar de  la tabla TbPlanificacionAnexos para el ID=strIDPlanificacion sólo por error en la función CerrarPlanificacionSinTrabajos
    '   -llamada por:
    '       -me.CerrarPlanificacionSinTrabajos
    '   -Devuelve:
    '       EliminarPlanificacion =strDescriptivo
    '       EliminarPlanificacion ="#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim strCadenaAnexos As String, strTextoError As String
    On Error GoTo errores
    If Not IsNumeric(strIDPlanificacion) Then
        strTextoError = "El IDPlanificación no es numérico"
        Err.Raise 1000
    End If
    '------------------------------------------
    '   -Devuelve:
    '       Dame = strValorObtenido
    '       Dame = "#ERR" & "|" & strTextoError
    '--------------------------------------------
    flag = Dame("TbPlanificacion", "FechaCierre", "IDPlanificacion", strIDPlanificacion)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Dame ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If IsDate(flag) Then
        strTextoError = "La planificación está cerrada"
        Err.Raise 1000
    End If
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       strURLCadenaAnexos=strURLAnexo1 & ";" & strURLAnexo2 & ";" & .... & strURLAnexoN
    '       DameCadenaAnexos = strURLCadenaAnexos
    '       DameCadenaAnexos = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = DameCadenaAnexos(strIDPlanificacion)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método DameCadenaAnexos ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If flag <> "" Then
        strTextoError = "La planificación tiene anexos"
        Err.Raise 1000
    End If
    m_SQL = "DELETE TbPlanificacion.IDPlanificacion " & _
                "FROM TbPlanificacion " & _
                "WHERE (((TbPlanificacion.IDPlanificacion)=" & strIDPlanificacion & "));"
    DoCmd.SetWarnings False
    DoCmd.RunSQL m_SQL
    DoCmd.SetWarnings True
    EliminarPlanificacion = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método EliminarPlanificacion ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DoCmd.SetWarnings True
    EliminarPlanificacion = "#ERR" & "|" & strTextoError
End Function

Public Function CerrarPlanificacionSinTrabajos( _
                                                strIDPlanificacion As String, _
                                                strMotivacionReprogramacion As String, _
                                                intSemana As Integer, _
                                                intMes As Integer, _
                                                intAño As Integer, _
                                                Optional strURLAnexo As String _
                                                ) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 17/06/2014
    '   -Modificaciones:
    '
    '   -Funcionamiento:
    '       -va a editar un registro en la tabla TbPlanificacion
    '           -strMotivoCierre="Reprogramación"
    '       1) VIABILIDAD DE LAS VARIABLES DE ENTRADA
    '           -strIDPlanificacion numérico y existente
    '           -strIDEvento="" y strIDActividad=""---->Mal una de las dos ha de estar rellena
    '           -strIDEvento existente y no usado en otro cierre
    '           -strIDActividad existente y no usada en otro cierre
    
    '       2) ACTUACIÓN
    
    '           -Editar en la tabla TbPlanificacion
    '               strMotivoCierre="Trabajo Realizado"
    '               strFechaCierre=Now()
    '               Registro de strIDEvento ó strIDActividad
    '   -llamada por:
    '       -Form_FormPlanificacionReprogramacion.ComandoReplanificar_Click
    '   -Devuelve:
    '       CerrarPlanificacionConTrabajos = Descriptivo
    '       CerrarPlanificacionConTrabajos ="#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, strIDEquipo As String, strIDPlanificacionNueva As String, _
        intSemanasTieneElMes As Integer, strURLAnexoRemoto As String, blnYaCreadaNueva As Boolean, blnYaCerradaVieja As Boolean, _
        strMotivoCierre As String, strFechaCierre As String, strTextoError As String
    On Error GoTo errores
    If Not IsNumeric(strIDPlanificacion) Then
        strTextoError = "La planificación ha de ser numérica"
        Err.Raise 1000
    End If
    If strMotivacionReprogramacion = "" Then
        strTextoError = "Se ha de indicar el motivo de la Reprogramación."
        Err.Raise 1000
    End If
    If strURLAnexo <> "" Then
        If Not fso.FileExists(strURLAnexo) Then
            strTextoError = "Si se dice que se introduce un anexo, éste ha de ser alcanzable "
            Err.Raise 1000
        End If
        
        
    End If
    '-------------------------------------------------------------------
    '   -Devuelve:
    '      EstadoPlanificacion=strEstado
    '      EstadoPlanificacion="-1" Si error
    '       strEstado="Abierta"
    '       strEstado="Cerrada por trabajo realizado"
    '       strEstado="Cerrada por reprogramación"
    '-------------------------------------------------------------------
    flag = EstadoPlanificacion(strIDPlanificacion)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método EstadoPlanificacion ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If flag <> "Abierta" Then
        strTextoError = "Esta planificación está en un estado de: " & flag & " y sólo puede estar Abierta"
        Err.Raise 1000
    End If
    '------------------------------------------
    '   -Devuelve:
    '       Dame = strValorObtenido
    '       Dame = "#ERR" & "|" & strTextoError
    '--------------------------------------------
    flag = Dame("TbPlanificacion", "IDEquipo", "IDPlanificacion", strIDPlanificacion)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Dame ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strIDEquipo = flag
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       NumeroSemanasTieneElMes =intNumeroDiasTieneElMes
    '       NumeroSemanasTieneElMes=-1 si error
    '-------------------------------------------------------------------
    intSemanasTieneElMes = NumeroSemanasTieneElMes(intMes, intAño)
    If intSemanasTieneElMes = -1 Then
        strTextoError = "la función que calcula las semanas que tiene un mes ha dado un error desconocido"
        Err.Raise 1000
    End If
    If intSemana < 1 Or intSemana > intSemanasTieneElMes Then
        strTextoError = "La semana de planificación ha de estar entre 1 y el número de semanas que tiene este mes: " & CStr(intSemanasTieneElMes)
        Err.Raise 1000
    End If
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       FechaViernes =dteFechaViernesSemanax para intSemana=x <>0
    '       FechaViernes =strCadenaFechas para intSemana=0
    '       strCadenaFechas=01 & dteFechaViernesSemana1 & vbcrlf & _
                            02 & dteFechaViernesSemana2 & vbcrlf & _
    '                       ...
    '                       intNumeroSemanas & dteFechaViernesSemanax
    '       FechaViernes="#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = FechaViernes(intMes, intAño, intSemana)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método FechaViernes ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If flag = "0" Or Not IsDate(flag) Then
        strTextoError = "La semana " & CStr(intSemana) & " del mes " & CStr(intMes) & " del año " & CStr(intAño) & " no tiene viernes."
        Err.Raise 1000
    End If
    '------------------------
    ' VERIFICACIÓN DE QUE NO HAY OTRA PLANIFICACIÓN PARA LA MISMA SEMANA
    '------------------------------------------------------------------
    '   -Devuelve:
    '       PlanificacionYaExistente = strPlanificacionYaExistente
    '       PlanificacionYaExistente = "#ERR" & "|" & strTextoError
    '------------------------------------------------------------------
    flag = PlanificacionYaExistente(strIDEquipo, intSemana, intMes, intAño)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método PlanificacionYaExistente ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If flag = "Sí" Then
        strTextoError = "La semana " & CStr(intSemana) & " del mes " & CStr(intMes) & " del año " & CStr(intAño) & " ya está planificada para este equipo."
        Err.Raise 1000
    End If
    '------------------------------------------------------------------
    '   -Devuelve:
    '       AltaPlanificacion = strIDPlanificacion & ";" & strFechaPrevistaCierre
    '       AltaPlanificacion = "#ERR" & "|" & strTextoError
    '------------------------------------------------------------------
    flag = AltaPlanificacion(strIDEquipo, intSemana, intMes, intAño)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método AltaPlanificacion ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If InStr(1, flag, ";") = 0 Then
        strTextoError = "El método AltaPlanificacion ha devuelto un valor con un formato desoconocido"
        Err.Raise 1000
    End If
    dato = Split(flag, ";")
    strIDPlanificacionNueva = Nz(dato(0), "")
    If Not IsNumeric(strIDPlanificacionNueva) Then
        strTextoError = "El método Planificación.AltaPlanificacion ha devuelto un valor con un formato desoconocido"
        Err.Raise 1000
    End If
    blnYaCreadaNueva = True
    strFechaCierre = Now()
    strMotivoCierre = "Reprogramación"
    m_SQL = "SELECT TbPlanificacion.* " & _
            "FROM TbPlanificacion " & _
            "WHERE (((TbPlanificacion.IDPlanificacion)=" & strIDPlanificacion & "));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        .Edit
            .Fields("MotivoCierre") = strMotivoCierre
            .Fields("FechaCierre") = strFechaCierre
            .Fields("MotivacionReprogramacion") = strMotivacionReprogramacion
            .Fields("IDNuevaPlanificacion") = strIDPlanificacionNueva
        .Update
        blnYaCerradaVieja = True
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If strURLAnexo <> "" Then
        '-------------------------------------------------------------------
        '   -Devuelve:
        '       AnexarAPlanificacion = strURLRemoto
        '       AnexarAPlanificacion = "#ERR" & "|" & strTextoError
        '-----------------------------------------------
        flag = AnexarAPlanificacion(strIDPlanificacion, strURLAnexo)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método Eliminar ha devuelto un error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        strURLAnexoRemoto = flag
    End If
    
    CerrarPlanificacionSinTrabajos = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Planificacion.CerrarPlanificacionSinTrabajos ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    If blnYaCreadaNueva = True And blnYaCerradaVieja = False And IsNumeric(strIDPlanificacionNueva) Then
        '-------------------------------------------------------------------
        '   -Devuelve:
        '       EliminarPlanificacion =strDescriptivo
        '       EliminarPlanificacion ="#ERR" & "|" & strTextoError
        '-------------------------------------------------------------------
        flag = EliminarPlanificacion(strIDPlanificacionNueva)
    End If
    
    CerrarPlanificacionSinTrabajos = "#ERR" & "|" & strTextoError
End Function
Public Function CerrarPlanificacionConTrabajos( _
                                                strIDPlanificacion As String, _
                                                Optional strIDEvento As String, _
                                                Optional strIDNOEvento As String _
                                                ) As String
     '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 17/06/2014
    '   -Modificaciones:
    '
    '   -Funcionamiento:
    '       -va a editar un registro en la tabla TbPlanificacion
    '       1) VIABILIDAD DE LAS VARIABLES DE ENTRADA
    '           -strIDPlanificacion numérico y existente
    '           -strIDEvento="" y strIDActividad=""---->Mal una de las dos ha de estar rellena
    '           -strIDEvento existente y no usado en otro cierre
    '           -strIDNoEvento existente y no usada en otro cierre
    
    '       2) ACTUACIÓN
    
    '           -Editar en la tabla TbPlanificacion
    '               strMotivoCierre="Trabajo Realizado"
    '               strFechaCierre=Now()
    '               Registro de strIDEvento ó strIDNoEvento
    '   -llamada por:
    '       -Form_FormPlanificacionCierre.ComandoCierre_Click
    '   -Devuelve:
    '       CerrarPlanificacionConTrabajos = Descriptivo
    '       CerrarPlanificacionConTrabajos = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, strMotivoCierre As String, strFechaCierre As String, strFechaFinalEvento As String, _
        strIDEquipo As String, strTextoError As String
    On Error GoTo errores
    If Not IsNumeric(strIDPlanificacion) Then
        strTextoError = "Se ha de introducir un valor correcto de la planificación a cerrar"
        Err.Raise 1000
    End If
    If strIDEvento = "" And strIDNOEvento = "" Then
        strTextoError = "Para cerrar la planificación por trabajos realizados, éstos se han de incluir, o bien, mediante el IDEvento, " & _
                        "o bien, mediante el IDActividad"
        Err.Raise 1000
    End If
    If strIDEvento <> "" And strIDNOEvento <> "" Then
         strTextoError = "Para cerrar la planificación por trabajos realizados, éstos se han de incluir, o bien, mediante el IDEvento, " & _
                        "o bien, mediante el IDActividad"
        Err.Raise 1000
    End If
    If strIDEvento <> "" Then
        '------------------------------------------
        '   -Devuelve:
        '       Dame = strValorObtenido
        '       Dame = "#ERR" & "|" & strTextoError
        '--------------------------------------------
        flag = Dame("TbPlanificacion", "IDEvento", "IDEvento", strIDEvento)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método Dame ha devuelto el error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        If flag <> "" Then
            strTextoError = "Ese evento o actividad ya ha sido usado para cerrar otra planificación anteriormente."
            Err.Raise 1000
        End If
        '------------------------------------------
        '   -Devuelve:
        '       Dame = strValorObtenido
        '       Dame = "#ERR" & "|" & strTextoError
        '--------------------------------------------
        flag = Dame("TbEventos", "TIPOEVENTO", "IDEvento", strIDEvento)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método Dame ha devuelto el error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        If flag <> "MANT. PROGRAMADO" Then
            strTextoError = "El evento ha de ser del tipo MANT. PROGRAMADO"
            Err.Raise 1000
        End If
        '----------------------------------------------------------------
        '   -Devuelve:
        '       DameFechaFinEvento = strFechaUltimaActividad
        '       DameFechaFinEvento = #ERR
        '----------------------------------------------------------------
        flag = DameFechaFinEvento(strIDEvento)
        If flag = "#ERR" Then
            strTextoError = "No se ha podido determinar la fecha fin del Evento"
            Err.Raise 1000
        End If
        If Not IsDate(flag) Then
            strTextoError = "El evento ha de estar cerrado"
            Err.Raise 1000
        End If
        strFechaCierre = flag
    Else
        If Not IsNumeric(strIDNOEvento) Then
            strTextoError = "El ID no de evento ha de ser numérico"
            Err.Raise 1000
        End If
        '------------------------------------------
        '   -Devuelve:
        '       Dame = strValorObtenido
        '       Dame = "#ERR" & "|" & strTextoError
        '--------------------------------------------
        flag = Dame("TbPlanificacion", "IDNoEvento", "IDNoEvento", strIDNOEvento)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método Dame ha devuelto el error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        If flag <> "" Then
            strTextoError = "Ese evento o actividad ya ha sido usado para cerrar otra planificación anteriormente."
            Err.Raise 1000
        End If
        '------------------------------------------
        '   -Devuelve:
        '       Dame = strValorObtenido
        '       Dame = "#ERR" & "|" & strTextoError
        '--------------------------------------------
        flag = Dame("TbPlanificacion", "IDEquipo", "IDPlanificacion", strIDPlanificacion)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método Dame ha devuelto el error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        strIDEquipo = flag
        If Not IsNumeric(strIDEquipo) Then
            strTextoError = "No se ha encontrado un IDEquipo numérico para el IDPlanificación: " & strIDPlanificacion
            Err.Raise 1000
        End If
        m_SQL = "SELECT TbPlanificacionRegistrada.ID,TbPlanificacionRegistrada.FechaMantoRealizado " & _
                "FROM TbPlanificacionRegistrada " & _
                "WHERE (((TbPlanificacionRegistrada.ID)=" & strIDNOEvento & _
                ") AND ((TbPlanificacionRegistrada.IDEquipo)=" & strIDEquipo & "));"
        Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
        With rcdDatos
            If .EOF Then
                strTextoError = "La actividad no de evento no está registrada en la tabla TbPlanificacionRegistrada para el equipo de la planificación"
                Err.Raise 1000
            End If
            strFechaCierre = Nz(.Fields("FechaMantoRealizado"), "")
            If Not IsDate(strFechaCierre) Then
                strTextoError = "No se ha podido determinar la fecha de cierre mirando en el campo TbPlanificacionRegistrada.FechaMantoRealizado"
                Err.Raise 1000
            End If
        End With
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    '---------------------
    ' ACTUACIÓN
    '--------------
    strMotivoCierre = "Trabajo Realizado"
    m_SQL = "SELECT TbPlanificacion.* " & _
            "FROM TbPlanificacion " & _
            "WHERE (((TbPlanificacion.IDPlanificacion)=" & strIDPlanificacion & "));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strTextoError = "La Planificacion con ID: " & strIDPlanificacion & " no ha sido grabada en el Sismeta"
            Err.Raise 1000
        End If
        .Edit
            .Fields("MotivoCierre") = strMotivoCierre
            .Fields("FechaCierre") = strFechaCierre
            If strIDEvento <> "" Then
                .Fields("IDEvento") = strIDEvento
            Else
                .Fields("IDNoEvento") = strIDNOEvento
            End If
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    CerrarPlanificacionConTrabajos = "Planificación cerrada correctamente a Fecha " & strFechaCierre
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método CerrarPlanificacionConTrabajos ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    CerrarPlanificacionConTrabajos = "#ERR" & "|" & strTextoError
End Function

Public Function ObtenerHojaActividadesDeExportacion( _
                                                    ByRef wbHoja As excel.Worksheet _
                                                    ) As String

                                                   
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 01/2/13
    '   -Modificaciones:
    '
    '   -Funcionamiento:
    '       -va a rellenar la hoja llamada "ACTIVIDAD REALIZADA" en la hoja que se le pasa wbHoja
    '       -Tomamos la tabla TbAuxActividad ya rellena con los datos que tenga
    
    
    '   -llamada por:
  
    '   -Devuelve:
    '       ObtenerHojaActividadesDeExportacion = Descriptivo
    '       ObtenerHojaActividadesDeExportacion = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, intFila As Integer, strNodo As String, strBUI As String, strSUBSISTEMA As String, strEquipo As String, strIDEvento As String, _
        strTIPOTECNICO As String, strAliasTecnico As String, strActividad As String, strFechaActividad As String, _
        strHorasLaborables As String, strHORASEXTRAS As String, strUbicacion As String, strDESCRIPCION As String, _
        lngNumeroRegistro As Long, lngNumeroRegistrosTotales As Long, strTextoError As String
    On Error GoTo errores
    
    '-------------------
    '  CABECERA DE LA HOJA
    '--------------------
    With wbHoja
        .Cells(1, 1).Value = "NODO"
        .Cells(1, 2).Value = "BUI"
        .Cells(1, 3).Value = "SUBSISTEMA"
        .Cells(1, 4).Value = "EQUIPO"
        .Cells(1, 5).Value = "ID_EVENTO (PT)"
        .Cells(1, 6).Value = "TÉCNICO"
        .Cells(1, 7).Value = "INICIALES_TÉCNICO"
        .Cells(1, 8).Value = "ACTIVIDAD"
        .Cells(1, 9).Value = "FECHA_ACTIVIDAD"
        .Cells(1, 10).Value = "HORAS_LABORABLES"
        .Cells(1, 11).Value = "HORAS_EXTRAS"
        .Cells(1, 12).Value = "UBICACIÓN"
        .Cells(1, 13).Value = "DESCRIPCIÓN"
        
    End With
    wbHoja.Range(wbHoja.Cells(1, 1), wbHoja.Cells(1, 13)).Interior.ColorIndex = 15
    wbHoja.Range(wbHoja.Cells(1, 1), wbHoja.Cells(1, 13)).Font.Bold = True
    intFila = 2
    '--------------
    ' recorrer la tabla auxiliar de los datos
    '--------------
    m_SQL = "SELECT TbAuxActividad.*, TbCodActividad.ACTIVIDADNOPROGRAMADA " & _
            "FROM TbAuxActividad LEFT JOIN TbCodActividad ON TbAuxActividad.TIPOACTIVIDAD = TbCodActividad.CODIGO " & _
            "ORDER BY TbAuxActividad.IDEVENTO DESC;"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveLast
            .MoveFirst
            lngNumeroRegistrosTotales = .RecordCount
            lngNumeroRegistro = 1
            Do While Not .EOF
                strNodo = Nz(.Fields("NODO"), "")
                strBUI = Nz(.Fields("BUI"), "")
                strSUBSISTEMA = Nz(.Fields("SUBSISTEMA"), "")
                strEquipo = Nz(.Fields("EQUIPO"), "")
                strIDEvento = Nz(.Fields("IDEVENTO"), "")
                strTIPOTECNICO = Nz(.Fields("TIPOTECNICO"), "")
                strAliasTecnico = Nz(.Fields("ALIASTecnico"), "")
                strActividad = Nz(.Fields("ACTIVIDADNOPROGRAMADA"), "")
                strFechaActividad = Nz(.Fields("FECHAACTIVIDAD"), "")
                strHorasLaborables = Nz(.Fields("HORASLABORABLES"), "")
                strHORASEXTRAS = Nz(.Fields("HORASEXTRAS"), "")
                strUbicacion = Nz(.Fields("UBICACION"), "")
                strDESCRIPCION = Nz(.Fields("DESCRIPCION"), "")
                wbHoja.Range(wbHoja.Cells(1, 1), wbHoja.Cells(1, 13)).Font.Bold = True
                
                wbHoja.Cells(intFila, 1) = strNodo
                wbHoja.Cells(intFila, 2) = strBUI
                wbHoja.Cells(intFila, 3) = strSUBSISTEMA
                wbHoja.Cells(intFila, 4) = strEquipo
                wbHoja.Cells(intFila, 5) = strIDEvento
                wbHoja.Cells(intFila, 6) = strTIPOTECNICO
                wbHoja.Cells(intFila, 7) = strAliasTecnico
                wbHoja.Cells(intFila, 8) = strActividad
                If IsDate(strFechaActividad) Then
                    wbHoja.Cells(intFila, 9) = CDate(Format(strFechaActividad, "DD/MM/YYYY"))
                End If
                If IsNumeric(strHorasLaborables) Then
                    wbHoja.Cells(intFila, 10) = CDbl(strHorasLaborables)
                End If
                If IsNumeric(strHORASEXTRAS) Then
                    wbHoja.Cells(intFila, 11) = CDbl(strHORASEXTRAS)
                End If
                
                wbHoja.Cells(intFila, 12) = strUbicacion
                wbHoja.Cells(intFila, 13) = strDESCRIPCION
                intFila = intFila + 1
                lngNumeroRegistro = lngNumeroRegistro + 1
                .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    RecuadrarRango wbHoja.Range(wbHoja.Cells(1, 1), wbHoja.Cells(lngNumeroRegistro + 1, 13)), "Multiple"
    With wbHoja
        .Columns("A:A").EntireColumn.AutoFit
        .Columns("B:B").EntireColumn.AutoFit
        .Columns("C:C").EntireColumn.AutoFit
        .Columns("D:D").EntireColumn.AutoFit
        .Columns("E:E").EntireColumn.AutoFit
        .Columns("F:F").EntireColumn.AutoFit
        .Columns("G:G").EntireColumn.AutoFit
        .Columns("H:H").EntireColumn.AutoFit
        .Columns("I:I").EntireColumn.AutoFit
        .Columns("J:J").EntireColumn.AutoFit
        .Columns("K:K").EntireColumn.AutoFit
        .Columns("L:L").EntireColumn.AutoFit
        .Columns("M:M").ColumnWidth = 93
        .Columns("M:M").WrapText = True
    End With
    ObtenerHojaActividadesDeExportacion = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método ObtenerHojaActividadesDeExportacion ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    ObtenerHojaActividadesDeExportacion = "#ERR" & "|" & strTextoError
 End Function
 Public Function ObtenerHojaMaterialesDeExportacion( _
                                                    ByRef wbHoja As excel.Worksheet _
                                                    ) As String

                                                   
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 01/2/13
    '   -Modificaciones:
    '       -13/06/2013: Se considera también las subcontrataciones
    '   -Funcionamiento:
    '       -va a rellenar la hoja llamada "MATERIAL UTILIZADO Y REPARADO" en la hoja que se le pasa wbHoja
    '       -Tomamos la tabla TbAuxMateriales ya rellena con los datos que tenga
    
    
    '   -llamada por:
  
    '   -Devuelve:
    '       ObtenerHojaMaterialesDeExportacion = Descriptivo
    '       ObtenerHojaMaterialesDeExportacion = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, intFila As Integer, strNodo As String, strBUI As String, strSUBSISTEMA As String, strEquipo As String, strIDEvento As String, _
         strTipoAccion As String, strPN As String, strNS As String, strMaterial As String, strFechaEntrega As String, strCOSTE As String, _
         strPrecio As String, strReparadoPor As String, strCodigoRegistroCambios As String
    Dim strFechaAltaSubContratacion As String, strFechaFinSubContratacion As String, strEmpresaSubContratacion As String, strDescripcionSubContratacion As String, _
        strImporteSubContratacion As String, lngNumeroRegistrosTotales As Long, lngNumeroRegistro As Long, strTextoError As String
    Dim intNFilaInicioMateriales As Integer, intNFilaFinMateriales As Integer
    On Error GoTo errores
    
    intFila = 1
    '-------------------
    '  CABECERA DE LOS MATERIALES
    '--------------------
    With wbHoja
        .Cells(intFila, 1).Value = "NODO"
        .Cells(intFila, 2).Value = "BUI"
        .Cells(intFila, 3).Value = "SUBSISTEMA"
        .Cells(intFila, 4).Value = "EQUIPO"
        .Cells(intFila, 5).Value = "ID_EVENTO (PT)"
        .Cells(intFila, 6).Value = "TIPO ACCIÓN"
        .Cells(intFila, 7).Value = "P/N"
        .Cells(intFila, 8).Value = "S/N"
        .Cells(intFila, 9).Value = "DESCRIPCIÓN EQUIPO/MATERIAL"
        .Cells(intFila, 10).Value = "FECHA INSTALACIÓN / DESINSTALACION / ENTREGA"
        .Cells(intFila, 11).Value = "COSTE"
        .Cells(intFila, 12).Value = "PRECIO (SIN IVA)"
        .Cells(intFila, 13).Value = "REPARADO POR"
    End With
    wbHoja.Range(wbHoja.Cells(intFila, 1), wbHoja.Cells(intFila, 13)).HorizontalAlignment = xlCenter
    wbHoja.Range(wbHoja.Cells(intFila, 1), wbHoja.Cells(intFila, 13)).Interior.ColorIndex = 15
    wbHoja.Range(wbHoja.Cells(intFila, 1), wbHoja.Cells(intFila, 13)).Font.Bold = True
    intNFilaInicioMateriales = intFila
    intFila = intFila + 1
    '--------------
    ' recorrer la tabla auxiliar de los datos
    '--------------
    m_SQL = "SELECT TbAuxMateriales.* " & _
            "FROM TbAuxMateriales " & _
            "ORDER BY TbAuxMateriales.IDEVENTO;"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    rcdDatos.MoveFirst
    With rcdDatos
        If Not .EOF Then
            .MoveLast
            .MoveFirst
            lngNumeroRegistro = 1
            lngNumeroRegistrosTotales = .RecordCount
             Do While Not .EOF
                strNodo = Nz(.Fields("NODO"), "")
                strBUI = Nz(.Fields("BUI"), "")
                strSUBSISTEMA = Nz(.Fields("SUBSISTEMA"), "")
                strEquipo = Nz(.Fields("EQUIPO"), "")
                strIDEvento = Nz(.Fields("IDEVENTO"), "")
                strTipoAccion = Nz(.Fields("TIPOACCION"), "")
                strPN = Nz(.Fields("PN"), "")
                strNS = Nz(.Fields("NS"), "")
                strMaterial = Nz(.Fields("MATERIAL"), "")
                strFechaEntrega = Nz(.Fields("FECHAENTREGA"), "")
                strCOSTE = Nz(.Fields("COSTE"), "")
                strPrecio = ""
                strReparadoPor = Nz(.Fields("REPARADOPOR"), "")
                
                wbHoja.Range(wbHoja.Cells(1, 1), wbHoja.Cells(1, 13)).Font.Bold = True
                wbHoja.Cells(intFila, 1) = strNodo
                wbHoja.Cells(intFila, 2) = strBUI
                wbHoja.Cells(intFila, 3) = strSUBSISTEMA
                wbHoja.Cells(intFila, 4) = strEquipo
                wbHoja.Cells(intFila, 5) = strIDEvento
                wbHoja.Cells(intFila, 6) = strTipoAccion
                wbHoja.Cells(intFila, 7) = strPN
                wbHoja.Cells(intFila, 8) = strNS
                wbHoja.Cells(intFila, 9) = strMaterial
                If IsDate(strFechaEntrega) Then
                    wbHoja.Cells(intFila, 10).Value = CDate(Format(strFechaEntrega, "DD/MM/YYYY"))
                End If
                If IsNumeric(strCOSTE) Then
                    wbHoja.Cells(intFila, 11) = CDbl(strCOSTE)
                End If
                If Not IsNumeric(strCOSTE) Then strCOSTE = "0"
                wbHoja.Cells(intFila, 11).NumberFormat = "#,##0.00 $"
                If IsNumeric(strPrecio) Then
                    wbHoja.Cells(intFila, 12) = CDbl(strPrecio)
                End If
                If Not IsNumeric(strPrecio) Then strPrecio = "0"
                wbHoja.Cells(intFila, 12).NumberFormat = "#,##0.00 $"
                wbHoja.Cells(intFila, 13) = strReparadoPor
                lngNumeroRegistro = lngNumeroRegistro + 1
                intFila = intFila + 1
                .MoveNext
            Loop
        End If
        intNFilaFinMateriales = intFila - 1
        RecuadrarRango wbHoja.Range(wbHoja.Cells(intNFilaInicioMateriales, 1), wbHoja.Cells(intNFilaFinMateriales, 13)), "Multiple"
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    With wbHoja
        .Columns("A:A").EntireColumn.AutoFit
        .Columns("B:B").EntireColumn.AutoFit
        .Columns("C:C").EntireColumn.AutoFit
        .Columns("D:D").EntireColumn.AutoFit
        .Columns("E:E").EntireColumn.AutoFit
        .Columns("F:F").EntireColumn.AutoFit
        .Columns("G:G").EntireColumn.AutoFit
        .Columns("H:H").EntireColumn.AutoFit
        .Columns("I:I").EntireColumn.AutoFit
        .Columns("J:J").EntireColumn.AutoFit
        .Columns("K:K").EntireColumn.AutoFit
        .Columns("L:L").EntireColumn.AutoFit
        .Columns("M:M").EntireColumn.AutoFit
    End With
    ObtenerHojaMaterialesDeExportacion = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método ObtenerHojaMaterialesDeExportacion ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    ObtenerHojaMaterialesDeExportacion = "#ERR" & "|" & strTextoError
 End Function


Public Function ObtenerHojaEventosDeExportacion( _
                                                    ByRef wbHoja As excel.Worksheet _
                                                    ) As String

                                                   
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 01/2/13
    '   -Modificaciones:
    '
    '   -Funcionamiento:
    '       -va a rellenar la hoja llamada "EVENTOS" en la hoja que se le pasa wbHoja
    '       -Tomamos la tabla TbAuxEventos ya rellena con los datos que tenga
    
    
    '   -llamada por:
    
    '   -Devuelve:
    '       ObtenerHojaEventosDeExportacion =Descriptivo
    '       ObtenerHojaEventosDeExportacion ="#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, intFila As Integer, strNodo As String, strBUI As String, strSUBSISTEMA As String, strEquipo As String, strIDEvento As String, strPMPR As String, _
        strTipoEvento As String, strCriticidad As String, strTIPOTECNICO As String, strAliasTecnico As String, strORIGINADOR As String, strFechaAlta As String, _
        strHoraAlta As String, strTiempoRespuesta As String, strFechaFin As String, strHoraFin As String, strCAUSAFIN As String, strContacto As String, _
        strCadenaIDEventosGenerados As String, strDESCRIPCION As String
    Dim lngNumeroRegistrosTotales As Long, lngNumeroRegistro As Long, strTextoError As String
    On Error GoTo errores
    
    m_SQL = "TbAuxEventos"
    '-------------------
    '  CABECERA DE LA HOJA
    '--------------------
    With wbHoja
        .Cells(1, 1).Value = "NODO"
        .Cells(1, 2).Value = "BUI"
        .Cells(1, 3).Value = "SUBSISTEMA"
        .Cells(1, 4).Value = "EQUIPO"
        .Cells(1, 5).Value = "ID_EVENTO (PT)"
        .Cells(1, 6).Value = "PM/PR"
        .Cells(1, 7).Value = "TIPO_EVENTO"
        .Cells(1, 8).Value = "CRITICIDAD"
        .Cells(1, 9).Value = "TÉCNICO"
        .Cells(1, 10).Value = "ORIGINADOR"
        .Cells(1, 11).Value = "FECHA_ALTA"
        .Cells(1, 12).Value = "HORA_ALTA"
        .Cells(1, 13).Value = "TIEMPO_RESPUESTA"
        .Cells(1, 14).Value = "FECHA_FIN"
        .Cells(1, 15).Value = "HORA_FIN"
        .Cells(1, 16).Value = "CAUSA_FIN"
        .Cells(1, 17).Value = "CONTACTO"
        .Cells(1, 18).Value = "EVENTO_GENERADO"
        .Cells(1, 19).Value = "DESCRIPCIÓN"
    End With
    wbHoja.Range(wbHoja.Cells(1, 1), wbHoja.Cells(1, 19)).Font.Bold = True
    wbHoja.Range(wbHoja.Cells(1, 1), wbHoja.Cells(1, 19)).Interior.ColorIndex = 15
    intFila = 2
    '--------------
    ' recorrer la tabla auxiliar de los datos
    '--------------
    m_SQL = "TbAuxEventos"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    rcdDatos.MoveFirst
    With rcdDatos
        lngNumeroRegistrosTotales = 0
        lngNumeroRegistro = 1
        If Not .EOF Then
            .MoveLast
            .MoveFirst
            Do While Not .EOF
                strNodo = Nz(.Fields("NODO"), "")
                strBUI = Nz(.Fields("BUI"), "")
                strSUBSISTEMA = Nz(.Fields("SUBSISTEMA"), "")
                strEquipo = Nz(.Fields("EQUIPO"), "")
                strIDEvento = Nz(.Fields("IDEVENTO"), "")
                strPMPR = Nz(.Fields("PMPR"), "")
                strTipoEvento = Nz(.Fields("TIPOEVENTO"), "")
                strCriticidad = Nz(.Fields("CRITICIDAD"), "")
                strTIPOTECNICO = Nz(.Fields("TIPOTECNICO"), "")
                strORIGINADOR = Nz(.Fields("ORIGINADOR"), "")
                strFechaAlta = Nz(.Fields("FECHAALTA"), "")
                strHoraAlta = Nz(.Fields("HORAALTA"), "")
                strTiempoRespuesta = Nz(.Fields("TIEMPORESPUESTA"), "")
                strFechaFin = Nz(.Fields("FECHAFIN"), "")
                strHoraFin = Nz(.Fields("HORAFIN"), "")
                strCAUSAFIN = Nz(.Fields("CAUSAFIN"), "")
                strContacto = Nz(.Fields("CONTACTO"), "")
                strCadenaIDEventosGenerados = Nz(.Fields("IDEVENTOGENERADO"), "")
                strDESCRIPCION = Nz(.Fields("DESCRIPCION"), "")
                
                wbHoja.Cells(intFila, 1) = strNodo
                wbHoja.Cells(intFila, 2) = strBUI
                wbHoja.Cells(intFila, 3) = strSUBSISTEMA
                wbHoja.Cells(intFila, 4) = strEquipo
                wbHoja.Cells(intFila, 5) = strIDEvento
                wbHoja.Cells(intFila, 6) = strPMPR
                wbHoja.Cells(intFila, 7) = strTipoEvento
                wbHoja.Cells(intFila, 8) = strCriticidad
                wbHoja.Cells(intFila, 9) = strTIPOTECNICO
                wbHoja.Cells(intFila, 10) = strORIGINADOR
                If IsDate(strFechaAlta) Then
                    wbHoja.Cells(intFila, 11) = CDate(Format(strFechaAlta, "dd/mm/yyyy"))
                End If
                If IsDate(strHoraAlta) Then
                    wbHoja.Cells(intFila, 12) = CDate(Format(strHoraAlta, "hh:mm"))
                End If
                If IsDate(strTiempoRespuesta) Then
                    wbHoja.Cells(intFila, 13) = CDate(Format(strTiempoRespuesta, "hh:mm"))
                End If
                If IsDate(strFechaFin) Then
                    wbHoja.Cells(intFila, 14) = CDate(Format(strFechaFin, "dd/mm/yyyy"))
                End If
                If IsDate(strHoraFin) Then
                    wbHoja.Cells(intFila, 15) = CDate(Format(strHoraFin, "hh:mm"))
                End If
                wbHoja.Cells(intFila, 16) = strCAUSAFIN
                wbHoja.Cells(intFila, 17) = strContacto
                wbHoja.Cells(intFila, 18) = strCadenaIDEventosGenerados
                wbHoja.Cells(intFila, 19) = strDESCRIPCION
                intFila = intFila + 1
                lngNumeroRegistro = lngNumeroRegistro + 1
                .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    RecuadrarRango wbHoja.Range(wbHoja.Cells(2, 1), wbHoja.Cells(intFila - 1, 19)), "Multiple"
    With wbHoja
        .Columns("A:A").EntireColumn.AutoFit
        .Columns("B:B").EntireColumn.AutoFit
        .Columns("C:C").EntireColumn.AutoFit
        .Columns("D:D").EntireColumn.AutoFit
        .Columns("E:E").EntireColumn.AutoFit
        .Columns("F:F").EntireColumn.AutoFit
        .Columns("G:G").EntireColumn.AutoFit
        .Columns("H:H").EntireColumn.AutoFit
        .Columns("I:I").EntireColumn.AutoFit
        .Columns("J:J").EntireColumn.AutoFit
        .Columns("K:K").EntireColumn.AutoFit
        .Columns("L:L").EntireColumn.AutoFit
        .Columns("M:M").EntireColumn.AutoFit
        .Columns("N:N").EntireColumn.AutoFit
        .Columns("O:O").EntireColumn.AutoFit
        .Columns("P:P").ColumnWidth = 95
        .Columns("P:P").MergeCells = False
        .Columns("Q:Q").EntireColumn.AutoFit
        .Columns("R:R").EntireColumn.AutoFit
        .Columns("S:S").ColumnWidth = 95
        .Columns("S:S").MergeCells = False
    End With
    ObtenerHojaEventosDeExportacion = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método ObtenerHojaEventosDeExportacion ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    ObtenerHojaEventosDeExportacion = "#ERR" & "|" & strTextoError
 End Function
 Public Function MaterialProvieneDeActividad(strIDMaterial As String) As String
    '-------------------------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 11/04/13
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       -dependerá de actividad si poroviene de un evento
   
    '   -Llamada desde
    '       -Form_FormMaterialesGestion.SeleccionarMaterial
    '       -Form_FormMaterialRepuestoEdicion.ComandoEdicionMaterialSinActividad_Click
    '       -Form_FormMaterialReparacionEdicion.ComandoEdicionMaterialSinActividad_Click
    '       -FormMaterialesGestion.ComandoEliminarMatSinActividad_Click
    '       -Me.EliminarMaterialSinActividad
    '       -Form_FormMaterialIntervencionAlta.EstablecerDatos
    '       -Form_FormMaterialSeguimiento.EstablecerDatos
    '   -Devuelve:
    '       MaterialProvieneDeActividad = "Sí" "No"
    '       MaterialProvieneDeActividad = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    Dim rcdDatos As DAO.Recordset, strProvieneDeActividad As String, strIDEvento As String, strTextoError As String
    On Error GoTo errores
    If strIDMaterial = "" Then
        strTextoError = "El Material se ha de indicar "
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbEventos.IDEvento " & _
            "FROM TbEventos RIGHT JOIN (TbMaterial " & _
            "LEFT JOIN TbActividades ON TbMaterial.IDActividad = TbActividades.IDActividad) ON TbEventos.IDEvento = TbActividades.IDEvento " & _
            "WHERE (((TbMaterial.IDMaterial)='" & strIDMaterial & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strProvieneDeActividad = "No"
        Else
            strIDEvento = Nz(.Fields("IDEvento"), "")
            If strIDEvento = "" Then
                strProvieneDeActividad = "No"
            Else
                strProvieneDeActividad = "Sí"
            End If
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    MaterialProvieneDeActividad = strProvieneDeActividad
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método MaterialProvieneDeActividad ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    MaterialProvieneDeActividad = "#ERR" & "|" & strTextoError
End Function
Public Function SeguimientoCerrado( _
                                    strIDMaterial As String _
                                    ) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 12/09/14
    '   -Modificaciones:
    '
    '   -Funcionamiento:
    '       -Para el strIDMaterial consulta la tabla TbMaterialSeguimiento  si algún registro tiene UltimaIntervencion="Sí"
    '   -Llamada desde
    '       -IntervencionUltimosDatos
    '   -Devuelve:
    '       SeguimientoCerrado ="Sí o No"
    '       SeguimientoCerrado ="#ERR" & "|" & strTextoError
    '-----------------------------------------------
    Dim rcdDatos As DAO.Recordset, strUltimaIntervencion As String, strTextoError As String
    On Error GoTo errores
    If strIDMaterial = "" Then
        strTextoError = "Se ha de indicar el IDMaterial"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbMaterialSeguimiento.UltimaIntervencion " & _
            "FROM TbMaterialSeguimiento " & _
            "WHERE (((TbMaterialSeguimiento.IDMaterial)='" & strIDMaterial & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            SeguimientoCerrado = "No"
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            strUltimaIntervencion = Nz(.Fields("UltimaIntervencion"), "No")
            If strUltimaIntervencion = "Sí" Then
                rcdDatos.Close
                Set rcdDatos = Nothing
                SeguimientoCerrado = "Sí"
                Exit Function
            End If
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    SeguimientoCerrado = "No"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método SeguimientoCerrado ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    SeguimientoCerrado = "#ERR" & "|" & strTextoError
End Function
Public Function UltimaFechaFinSeguimiento( _
                                                strIDMaterial As String _
                                                ) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 12/09/14
    '   -Modificaciones:
    '
    '   -Funcionamiento:
    '       -Para el strIDMaterial consulta la tabla TbMaterialSeguimiento  la última fecha, si no hubiera ninguna intervención daría la
    '       fecha de inicio del proceso dentro de TbMaterial.FechaIncio
    '   -Llamada desde
    
    '   -Devuelve:
    '       UltimaFechaFinSeguimiento =strFechaUltima
    '       UltimaFechaFinSeguimiento ="#ERR" & "|" & strTextoError
    '-----------------------------------------------
    Dim rcdDatos As DAO.Recordset, strFechaUltima As String, blnSinIntervenciones As Boolean, strTextoError As String
    On Error GoTo errores
    If strIDMaterial = "" Then
        strTextoError = "Se ha de indicar el IDMaterial"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbMaterialSeguimiento.IDMaterial " & _
            "FROM TbMaterialSeguimiento " & _
            "WHERE (((TbMaterialSeguimiento.IDMaterial)='" & strIDMaterial & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            blnSinIntervenciones = True
        Else
            blnSinIntervenciones = False
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If blnSinIntervenciones = True Then
        m_SQL = "SELECT TbMaterial.FechaIncio " & _
                "FROM TbMaterial " & _
                "WHERE (((TbMaterial.IDMaterial)='" & strIDMaterial & "'));"
        Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
        With rcdDatos
            If .EOF Then
                strTextoError = "No está registrado el seguimiento del material en la tabla TbMaterial"
                Err.Raise 1000
            End If
            strFechaUltima = Nz(.Fields("FechaIncio"), "")
            If Not IsDate(strFechaUltima) Then
                strTextoError = "No está registrado el seguimiento del material en la tabla TbMaterial"
                Err.Raise 1000
            End If
        End With
        rcdDatos.Close
        Set rcdDatos = Nothing
    Else
        m_SQL = "SELECT TbMaterialSeguimiento.FechaFinIntervencion " & _
                "FROM TbMaterialSeguimiento " & _
                "WHERE (((TbMaterialSeguimiento.IDMaterial)='" & strIDMaterial & _
                "') AND (Not (TbMaterialSeguimiento.FechaFinIntervencion) Is Null)) " & _
                "ORDER BY TbMaterialSeguimiento.FechaFinIntervencion DESC;"
        Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
        With rcdDatos
            .MoveFirst
            strFechaUltima = Nz(.Fields("FechaFinIntervencion"), "")
            If Not IsDate(strFechaUltima) Then
                strTextoError = "No está registrado el seguimiento del material en la tabla TbMaterial"
                Err.Raise 1000
            End If
        End With
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    UltimaFechaFinSeguimiento = strFechaUltima
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método UltimaFechaFinSeguimiento ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    UltimaFechaFinSeguimiento = "#ERR" & "|" & strTextoError
End Function
Public Function MaterialIntervencionAlta( _
                                        strIDMaterial As String, _
                                        strFechaInicioIntervencion As String, _
                                        strDestino As String, _
                                        strMotivo As String, _
                                        Optional strResultado As String, _
                                        Optional strFechaFinIntervencion As String, _
                                        Optional strUltimaIntervencion As String _
                                        ) As String
     '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 12/09/14
    '   -Modificaciones:
    '
    '   -Funcionamiento:
    '       -se verifica que el IDMaterial esté registrada en TbMaterial
    '       -Se verifica que el seguimiento no esté ya cerrado (alguna intervención marcada como última)
    '       -Se verifica que no haya otra intervención aún abierta (sin strFechaFinIntervencion)
    '       -La fecha inicial de intervención ha de ser igual o mayor que la ultima FechaFinIntervencion que pudiera haber
    '       -Grabar en la tabla TbMaterialSeguimiento
    '   -Llamada desde
    
    '   -Devuelve:
    '       MaterialIntervencionAlta = strIDSeguimiento
    '       MaterialIntervencionAlta = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    Dim rcdDatos As DAO.Recordset, strIDIntervencion As String, strUltimaFechaFinSeguimiento As String, strTextoError As String
    On Error GoTo errores
    If strIDMaterial = "" Then
        strTextoError = "El IDMaterial es obligatorio"
        Err.Raise 1000
    End If
    '-----------------------------------------------
    '   -Devuelve:
    '       EstaFranqueado = strValorFranqueado
    '       EstaFranqueado = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = EstaFranqueado("MA", strIDMaterial)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método EstaFranqueado ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If flag = "Sí" Then
        strTextoError = "Proviene de un evento franqueado"
        Err.Raise 1000
    End If
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       EstaFacturado = strEstaFacturado
    '       EstaFacturado = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = EstaFacturado("MA", strIDMaterial)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método EstaFacturado ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If flag = "Sí" Then
        strTextoError = "Proviene de un material ya facturado"
        Err.Raise 1000
    End If
    If Not IsDate(strFechaInicioIntervencion) Then
        strTextoError = "Es obligatoria la fecha de inicio de intervención"
        Err.Raise 1000
    End If
    If strUltimaIntervencion <> "Sí" And strUltimaIntervencion <> "No" Then
        strUltimaIntervencion = "No"
    End If
    If strDestino = "" Then
        strTextoError = "Destino obligatorio"
        Err.Raise 1000
    End If
     If strMotivo = "" Then
        strTextoError = "Motivo obligatorio"
        Err.Raise 1000
    End If
    If strFechaFinIntervencion <> "" Then
        If Not IsDate(strFechaFinIntervencion) Then
            strTextoError = "Fecha Fin de Intervención incorrecta"
            Err.Raise 1000
        End If
    End If
    '-----------------------------------------------
    '   -Devuelve:
    '       SeguimientoCerrado ="Sí o No"
    '       SeguimientoCerrado ="#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = SeguimientoCerrado(strIDMaterial)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método SeguimientoCerrado ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If flag = "Sí" Then
        strTextoError = "El proceso ya estaba finalizado previamente"
        Err.Raise 1000
    End If
    '------------------------------------------------------------------------
    '   -Devuelve:
    '       MaterialSeguimientoConItervencionAbierta ="Sí o No"
    '       MaterialSeguimientoConItervencionAbierta ="#ERR" & "|" & strTextoError
    '------------------------------------------------------------------------
    flag = MaterialSeguimientoConItervencionAbierta(strIDMaterial)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método MaterialSeguimientoConItervencionAbierta ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If flag = "Sí" Then
        strTextoError = "Hay una intervención previa sin cerrar"
        Err.Raise 1000
    End If
    '-----------------------------------------------
    '   -Devuelve:
    '       UltimaFechaFinSeguimiento =strFechaUltima
    '       UltimaFechaFinSeguimiento ="#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = UltimaFechaFinSeguimiento(strIDMaterial)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método UltimaFechaFinSeguimiento ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strUltimaFechaFinSeguimiento = flag
    If Not IsDate(strUltimaFechaFinSeguimiento) Then
        strTextoError = "La función UltimaFechaFinSeguimiento ha devuelto un error desconocido."
        Err.Raise 1000
    End If
    If CDate(strFechaInicioIntervencion) < CDate(strUltimaFechaFinSeguimiento) Then
        strTextoError = "Fecha inicio de intervención anterior a la del cierre de la última o a la del alta del material. "
        Err.Raise 1000
    End If
    '-------------------------------------------------------------------
    '   -Devuelve:
    '       DameIDIntervencion = strParteFija & Format(cstr(lngIDMaximo+1), "00")
    '       DameIDIntervencion = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    flag = DameIDIntervencion(strIDMaterial)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método DameIDIntervencion ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strIDIntervencion = flag
    If strIDIntervencion = "" Then
        strTextoError = "La función DameIDIntervencion ha devuelto un IDSeguimiento con formato desconocido"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbMaterialSeguimiento.* " & _
            "FROM TbMaterialSeguimiento " & _
            "WHERE (((TbMaterialSeguimiento.IDSeguimiento)='" & strIDIntervencion & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            strTextoError = "La función DameIDIntervencion ha devuelto un IDSeguimiento ya registrado previamente"
            Err.Raise 1000
        End If
        .AddNew
            .Fields("IDSeguimiento") = strIDIntervencion
            .Fields("IDMaterial") = strIDMaterial
            .Fields("FechaInicioIntervencion") = strFechaInicioIntervencion
            .Fields("Destino") = strDestino
            .Fields("Motivo") = strMotivo
            If IsDate(strFechaFinIntervencion) Then
                .Fields("FechaFinIntervencion") = strFechaFinIntervencion
            End If
            If strResultado <> "" Then
                .Fields("Resultado") = strResultado
            End If
            .Fields("UltimaIntervencion") = strUltimaIntervencion
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    '-----------------------------------------------
    '   -Devuelve:
    '       SeguimientoCerrado ="Sí o No"
    '       SeguimientoCerrado ="#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = SeguimientoCerrado(strIDMaterial)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método SeguimientoCerrado ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If flag = "Sí" Then
        '-----------------------------------------------
        '   -Devuelve:
        '       UltimaFechaFinSeguimiento =strFechaUltima
        '       UltimaFechaFinSeguimiento ="#ERR" & "|" & strTextoError
        '-----------------------------------------------
        flag = UltimaFechaFinSeguimiento(strIDMaterial)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método UltimaFechaFinSeguimiento ha devuelto un error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        strUltimaFechaFinSeguimiento = flag
         If IsDate(strUltimaFechaFinSeguimiento) Then
            m_SQL = "UPDATE TbMaterial SET TbMaterial.FechaEntrega = #" & Format(strUltimaFechaFinSeguimiento, "mm/dd/yyyy") & "# " & _
                    "WHERE (((TbMaterial.IDMaterial)='" & strIDMaterial & "'));"
            DoCmd.SetWarnings False
            DoCmd.RunSQL m_SQL
            DoCmd.SetWarnings True
         End If
    End If
    MaterialIntervencionAlta = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método MaterialIntervencionAlta ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    MaterialIntervencionAlta = "#ERR" & "|" & strTextoError
End Function
Public Function MaterialSeguimientoConItervencionAbierta( _
                                                            strIDMaterial As String _
                                                            ) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 12/09/14
    '   -Modificaciones:
    '
    '   -Funcionamiento:
    '       -Para el strIDMaterial consulta la tabla TbMaterialSeguimiento  si algún registro tiene UltimaIntervencion="Sí"
    '   -Llamada desde
    
    '   -Devuelve:
    '       MaterialSeguimientoConItervencionAbierta ="Sí o No"
    '       MaterialSeguimientoConItervencionAbierta ="#ERR" & "|" & strTextoError
    '------------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, strTextoError As String
    On Error GoTo errores
    If strIDMaterial = "" Then
        strTextoError = "Se ha de indicar el IDMaterial"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbMaterialSeguimiento.FechaFinIntervencion " & _
            "FROM TbMaterialSeguimiento " & _
            "WHERE (((TbMaterialSeguimiento.IDMaterial)='" & strIDMaterial & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            MaterialSeguimientoConItervencionAbierta = "No"
            Exit Function
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    m_SQL = "SELECT TbMaterialSeguimiento.FechaFinIntervencion " & _
            "FROM TbMaterialSeguimiento " & _
            "WHERE (((TbMaterialSeguimiento.IDMaterial)='" & strIDMaterial & _
            "') AND ((TbMaterialSeguimiento.FechaFinIntervencion) Is Null));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            MaterialSeguimientoConItervencionAbierta = "No"
        Else
            MaterialSeguimientoConItervencionAbierta = "Sí"
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método MaterialSeguimientoConItervencionAbierta ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    MaterialSeguimientoConItervencionAbierta = "#ERR" & "|" & strTextoError
End Function
Public Function DameIDIntervencion( _
                                    strIDMaterial As String _
                                    ) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 15/09/14
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       -strID=strIDMaterial_XX
    
    '       -Siendo XX el siguiente al ordinal máximo que haya
    '   -Llamada por:
    
    '   -Devuelve:
    '       DameIDIntervencion = strParteFija & Format(cstr(lngIDMaximo+1), "00")
    '       DameIDIntervencion = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, strValorCampo As String, strOrdinal As String, strIDSeguimiento As String, _
        lngID As Long, lngIDMaximo As Long
    Dim strParteFija As String, strTextoError As String
    On Error GoTo errores
    If strIDMaterial = "" Then
        strTextoError = "No se introducido la IDActividad"
        Err.Raise 1000
    End If
    strParteFija = strIDMaterial & "_"
    m_SQL = "SELECT TbMaterialSeguimiento.IDSeguimiento " & _
            "FROM TbMaterialSeguimiento " & _
            "WHERE (((TbMaterialSeguimiento.IDMaterial)='" & strIDMaterial & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                strIDSeguimiento = Nz(.Fields("IDSeguimiento"), "")
                If Len(strIDSeguimiento) > 2 Then
                    If IsNumeric(Right(strIDSeguimiento, 2)) Then
                        lngID = CLng(Right(strIDSeguimiento, 2))
                        If lngID > lngIDMaximo Then
                            lngIDMaximo = lngID
                        End If
                    End If
                End If
                .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    DameIDIntervencion = strParteFija & Format(CStr(lngIDMaximo + 1), "00")
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método DameIDIntervencion ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    DameIDIntervencion = "#ERR" & "|" & strTextoError
End Function
Public Function EdicionMaterialEnActividad( _
                                            strIDMaterial As String, _
                                            Optional strMaterial As String, _
                                            Optional strCOSTE As String, _
                                            Optional strTipoAccion As String, _
                                            Optional strPN As String, _
                                            Optional strNS As String, _
                                            Optional strReparadoPor As String, _
                                            Optional strDESCRIPCION As String, _
                                            Optional strFechaInicial As String, _
                                            Optional strGarantia As String, _
                                            Optional strEsReparacion As String, _
                                            Optional strFechaEntrega As String _
                                            ) As String

     '-------------------------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 25/01/13
    '   -Modificaciones:
    '       9/7/2013: es opcional strCoste y strFechaEntrega
    '       -11/09/2014: Se incluye la FechaInicial y la Garantía
    '       -06/10/2014: Se incluye un campo del tbMaterial -->EsReparacion.
    '   -Funcionamiento:
    '       -va a editar un registro en la tabla TbMaterial
    '       -la IDActividad no se puede cambiar. Habría que eliminarlo y crearlo de nuevo para la nueva actividad
    '       -strMaterial<>"" y strMaterial<>"#"
    
    
    '   -Llamada desde
    '       -Form_FormMaterialReparacionEdicion.Editar
    '       -Form_FormMaterialRepuestoEdicion.Editar
    '   -Devuelve:
    '       EdicionMaterialEnActividad = Descriptivo
    '       EdicionMaterialEnActividad = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    Dim rcdDatos As DAO.Recordset, strTextoError As String
    On Error GoTo errores
    If strIDMaterial = "" Or strIDMaterial = "#" Then
        strTextoError = "Se ha de indicar el material que se pretende editar "
        Err.Raise 1000
    End If
    If strMaterial = "" Or strMaterial = "#" Then
        strTextoError = "Se ha de indicar el material de que se trata "
        Err.Raise 1000
    End If
    If strCOSTE <> "" And strCOSTE <> "#" Then
        If Not IsNumeric(strCOSTE) Then
            strTextoError = "Si introduce un coste ha de ser numérico "
            Err.Raise 1000
        End If
    End If
    
    If strFechaEntrega <> "" And strFechaEntrega <> "#" Then
        If Not IsDate(strFechaEntrega) Then
            strTextoError = "Se ha de indicar la fecha de entrega o de reparación del material "
            Err.Raise 1000
        End If
    End If
    If Not IsDate(strFechaInicial) Then
        strTextoError = "Es obligatoria la fecha inicial"
        Err.Raise 1000
    End If
    If strGarantia <> "Sí" And strGarantia <> "No" Then
        strTextoError = "Garantía sólo puede ser Sí o No"
        Err.Raise 1000
    End If
    If strEsReparacion <> "Sí" And strEsReparacion <> "No" Then
        strTextoError = "Se ha de saber si es una reparación o no"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbMaterial.* " & _
            "FROM TbMaterial " & _
            "WHERE (((TbMaterial.IDMaterial)='" & strIDMaterial & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strTextoError = "No existe el material con código interno: " & strIDMaterial
            Err.Raise 1000
        End If
        .Edit
            .Fields("Material") = strMaterial
            .Fields("FechaIncio") = strFechaInicial
            .Fields("Garantia") = strGarantia
            If strPN <> "" And strPN <> "#" Then
                .Fields("PN") = strPN
            Else
                If strPN = "#" Then
                    .Fields("PN") = Null
                End If
            End If
            If strNS <> "" And strNS <> "#" Then
                .Fields("NS") = strNS
            Else
                .Fields("NS") = Null
            End If
            If IsNumeric(strCOSTE) Then
                .Fields("Coste") = strCOSTE
            Else
                If strCOSTE = "#" Then
                    .Fields("Coste") = Null
                End If
            End If
            If IsDate(strFechaEntrega) Then
                .Fields("FechaEntrega") = strFechaEntrega
            Else
                If strFechaEntrega = "#" Then
                    .Fields("FechaEntrega") = Null
                End If
            End If
            If strReparadoPor <> "" And strReparadoPor <> "#" Then
                .Fields("ReparadoPor") = strReparadoPor
            Else
                If strReparadoPor = "#" Then
                    .Fields("ReparadoPor") = Null
                End If
            End If
            If strTipoAccion <> "" And strTipoAccion <> "#" Then
                .Fields("TipoAccion") = strTipoAccion
            Else
                If strTipoAccion = "#" Then
                    .Fields("TipoAccion") = Null
                End If
            End If
            If strDESCRIPCION <> "" And strDESCRIPCION <> "#" Then
                .Fields("Descripcion") = strDESCRIPCION
            Else
                If strDESCRIPCION = "#" Then
                    .Fields("Descripcion") = Null
                End If
            End If
            .Fields("EsReparacion") = strEsReparacion
            If strEsReparacion = "No" Then
                If IsDate(strFechaEntrega) Then
                    .Fields("FechaEntrega") = strFechaEntrega
                Else
                    If strFechaEntrega = "#" Then
                        .Fields("FechaEntrega") = Null
                    End If
                End If
            End If
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    EdicionMaterialEnActividad = "Edición Correcta"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Evento.EdicionMaterialEnActividad ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    EdicionMaterialEnActividad = "#ERR" & "|" & strTextoError
End Function
Public Function IntervencionDatosDefechasIntermedias(strIDIntervencion As String, _
                                                    Optional strIDMaterial As String, _
                                                    Optional strNumeroIntervencionesEnSeguimiento As String _
                                                    ) As String
     '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 12/09/14
    '   -Modificaciones:
    '
    '   -Funcionamiento:
    '       -El codigo que devuelve es:
    '           strCodigoFinal =strNumeroIntervencionesEnSeguimiento;strOrdenIntervencionActual;strFechaInicioIntervencionAnterior; _
                strFechaFinIntervencionAnterior;strFechaInicionIntervencionPosterior
    '           strOrdenIntervencionActual="Primera o Intermedia o Última"
    
    '       -Obtenemos la posición de la intervención dentro del seguimiento.
    '           A) Es la única intervención
    '               -Cotejamos strFechaInicioIntervencion>=TbMaterial.FechaInicio
    '               -Cotejamos strFechaFinIntervencion>strFechaInicioIntervencion
    '           B) Más de una Intervención
    '               -Ocupa el primer Lugar
    '                   -Cotejamos strFechaInicioIntervencion>=TbMaterial.FechaInicio
    '                   -Cotejamos strFechaFinIntervencion>strFechaInicioIntervencion
    '               -Ocupa un lugar intermedio
    '                   -Cotejamos strFechaInicioIntervencion>=strFechaFinIntervencionAnterior
    '                   -strFechaFinIntervencion Obligada
    '                   -strFechaFinIntervencion>=strFechaInicioIntervencion AND strFechaFinIntervencion<=strFechaInicioIntervencionSiguiente
    '                   -strUltimaIntervencion No puede ser Sí
    '               -Ocupa el último lugar
    '                   -Cotejamos strFechaInicioIntervencion>=strFechaInicioIntervencionAnterior
    '                   -Si strFechaFinIntervencion es fecha-->strFechaFinIntervencion>=strFechaInicioIntervencion
    '       -Grabar en la tabla TbMaterialSeguimiento
    '   -Llamada desde
    '       -me.SeguimientoMaterialEliminar
    '   -Devuelve:
    '           strCodigoFinal =strNumeroIntervencionesEnSeguimiento;strOrdenIntervencionActual; _
                    strFechaFinIntervencionAnterior;strFechaInicionIntervencionPosterior
    '           strOrdenIntervencionActual="Primera o Intermedia o Última"
    '       IntervencionDatosDefechasIntermedias = strCodigoFinal
    '       IntervencionDatosDefechasIntermedias = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    Dim rcdDatos As DAO.Recordset, strOrdenIntervencionActual As String, _
        strFechaFinIntervencionAnterior As String, strFechaInicionIntervencionPosterior As String
    Dim strCodigoFinal As String, strUltimoID As String, strPrimerID As String
    Dim strIDIntervencionAnterior As String, strIDIntervencionPosterior As String, strIDIntervencionActual As String, strTextoError As String
    On Error GoTo errores
    If strIDIntervencion = "" Then
        strTextoError = "Es obligatorio el IDIntervención"
        Err.Raise 1000
    End If
    If strIDMaterial = "" Then
        '------------------------------------------
        '   -Devuelve:
        '       Dame = strValorObtenido
        '       Dame = "#ERR" & "|" & strTextoError
        '--------------------------------------------
        flag = Dame("TbMaterialSeguimiento", "IDMaterial", "IDSeguimiento", strIDIntervencion)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método Dame ha devuelto el error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        strIDMaterial = flag
    End If
    If strIDMaterial = "" Then
        strTextoError = "Intervención no registrada"
        Err.Raise 1000
    End If
    If Not IsNumeric(strNumeroIntervencionesEnSeguimiento) Then
        m_SQL = "SELECT TbMaterialSeguimiento.IDMaterial " & _
                "FROM TbMaterialSeguimiento " & _
                "WHERE (((TbMaterialSeguimiento.IDMaterial)='" & strIDMaterial & "'));"
        '-------------------------------------------------------------
        '   -Devuelve:
        '       DameNumeroRegistrosPorSQL = CStr(lngRegistros)
        '       DameNumeroRegistrosPorSQL = "#ERR" & "|" & strTextoError
        '-------------------------------------------------------------------
        flag = DameNumeroRegistrosPorSQL(m_SQL)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método DameNumeroRegistrosPorSQL ha devuelto el error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        strNumeroIntervencionesEnSeguimiento = flag
    End If
    If Not IsNumeric(strNumeroIntervencionesEnSeguimiento) Then
        strTextoError = "Intervención no registrada"
        Err.Raise 1000
    Else
        If strNumeroIntervencionesEnSeguimiento = "0" Then
            strTextoError = "Intervención no registrada"
            Err.Raise 1000
        End If
    End If
    If strNumeroIntervencionesEnSeguimiento = "1" Then
        strOrdenIntervencionActual = "Primera"
        '------------------------------------------
        '   -Devuelve:
        '       Dame = strValorObtenido
        '       Dame = "#ERR" & "|" & strTextoError
        '--------------------------------------------
        flag = Dame("TbMaterial", "FechaIncio", "IDMaterial", strIDMaterial)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método Dame ha devuelto el error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        strFechaFinIntervencionAnterior = flag
        If Not IsDate(strFechaFinIntervencionAnterior) Then
            strTextoError = "Es una intervención única y la fecha de inicio del seguimiento no está rellena"
            Err.Raise 1000
        End If
        strFechaInicionIntervencionPosterior = ""
        strCodigoFinal = strNumeroIntervencionesEnSeguimiento & ";" & strOrdenIntervencionActual & ";" & _
                strFechaFinIntervencionAnterior & ";" & strFechaInicionIntervencionPosterior
        IntervencionDatosDefechasIntermedias = strCodigoFinal
        Exit Function
    End If
    '--------------
    ' OBTENCIÓN DE strOrdenIntervencionActual
    '--------------
        '--------------
        ' Primero TbMaterialSeguimiento.IDSeguimiento
        '--------------
        m_SQL = "SELECT First(Right([IDSeguimiento],2)) AS PrimeroDeIDSeguimiento " & _
                "FROM TbMaterialSeguimiento " & _
                "WHERE (((TbMaterialSeguimiento.IDMaterial)='" & strIDMaterial & "'));"
        Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
        With rcdDatos
            If .EOF Then
                strTextoError = "Intervención no registrada"
                Err.Raise 1000
            End If
            strPrimerID = Nz(.Fields("PrimeroDeIDSeguimiento"), "")
            If Not IsNumeric(strPrimerID) Then
                strTextoError = "Intervención no registrada"
                Err.Raise 1000
            End If
                
        End With
        rcdDatos.Close
        Set rcdDatos = Nothing
        '--------------
        ' último TbMaterialSeguimiento.IDSeguimiento
        '--------------
            m_SQL = "SELECT Last(Right([IDSeguimiento],2)) AS ÚltimoDeIDSeguimiento " & _
                "FROM TbMaterialSeguimiento " & _
                "WHERE (((TbMaterialSeguimiento.IDMaterial)='" & strIDMaterial & "'));"
        Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
        With rcdDatos
            If .EOF Then
                strTextoError = "Intervención no registrada"
                Err.Raise 1000
            End If
            strUltimoID = Nz(.Fields("ÚltimoDeIDSeguimiento"), "")
            If Not IsNumeric(strUltimoID) Then
                strTextoError = "Intervención no registrada"
                Err.Raise 1000
            End If
        End With
        rcdDatos.Close
        Set rcdDatos = Nothing
        If CLng(Right(strIDIntervencion, 2)) = CLng(strPrimerID) Then
            strOrdenIntervencionActual = "Primera"
        ElseIf CLng(Right(strIDIntervencion, 2)) = CLng(strUltimoID) Then
            strOrdenIntervencionActual = "Última"
        Else
            strOrdenIntervencionActual = "Intermedia"
        End If
    '--------------
    ' OBTENCIÓN DE TODO LO DEMÁS
    '--------------
        If strOrdenIntervencionActual = "Primera" Then
            '------------------------------------------
            '   -Devuelve:
            '       Dame = strValorObtenido
            '       Dame = "#ERR" & "|" & strTextoError
            '--------------------------------------------
            flag = Dame("TbMaterial", "FechaIncio", "IDMaterial", strIDMaterial)
            If InStr(1, flag, "|") <> 0 Then
                dato = Split(flag, "|")
                strTextoError = "El método Dame ha devuelto el error: " & vbNewLine & dato(1)
                Err.Raise 1000
            End If
            strFechaFinIntervencionAnterior = flag
            If Not IsDate(strFechaFinIntervencionAnterior) Then
                strTextoError = "Es una intervención única y la fecha de inicio del seguimiento no está rellena"
                Err.Raise 1000
            End If
             strFechaInicionIntervencionPosterior = ""
        ElseIf strOrdenIntervencionActual = "Última" Then
            m_SQL = "SELECT TbMaterialSeguimiento.FechaFinIntervencion " & _
                    "FROM TbMaterialSeguimiento " & _
                    "WHERE ((Not (TbMaterialSeguimiento.IDMaterial) = '" & strIDMaterial & "')) " & _
                    "ORDER BY TbMaterialSeguimiento.IDSeguimiento DESC;"
            Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
            With rcdDatos
                If .EOF Then
                    strTextoError = "Intervención no registrada"
                    Err.Raise 1000
                End If
                .MoveNext
                strFechaInicionIntervencionPosterior = ""
                strFechaFinIntervencionAnterior = Nz(.Fields("FechaFinIntervencion"), "")
                If Not IsDate(strFechaFinIntervencionAnterior) Then
                    strTextoError = "Intervención anterior no está bien registrada al faltarle la fecha de fin de la misma"
                    Err.Raise 1000
                End If
            End With
            rcdDatos.Close
            Set rcdDatos = Nothing
        Else
            m_SQL = "SELECT TbMaterialSeguimiento.IDSeguimiento " & _
                    "FROM TbMaterialSeguimiento " & _
                    "WHERE (((Right([IDSeguimiento],2)) > " & Right(strIDIntervencion, 2) & ")) " & _
                    "ORDER BY TbMaterialSeguimiento.IDSeguimiento DESC;"
            Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
            With rcdDatos
                If .EOF Then
                    strTextoError = "Intervención no registrada"
                    Err.Raise 1000
                End If
                strIDIntervencionPosterior = .Fields("IDSeguimiento")
            End With
            rcdDatos.Close
            Set rcdDatos = Nothing
            m_SQL = "SELECT TbMaterialSeguimiento.IDSeguimiento " & _
                    "FROM TbMaterialSeguimiento " & _
                    "WHERE (((Right([IDSeguimiento],2)) < " & Right(strIDIntervencion, 2) & ")) " & _
                    "ORDER BY TbMaterialSeguimiento.IDSeguimiento DESC;"
            Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
            With rcdDatos
                If .EOF Then
                    strTextoError = "Intervención no registrada"
                    Err.Raise 1000
                End If
                strIDIntervencionAnterior = .Fields("IDSeguimiento")
            End With
            rcdDatos.Close
            Set rcdDatos = Nothing
            m_SQL = "SELECT TbMaterialSeguimiento.FechaInicioIntervencion, TbMaterialSeguimiento.FechaFinIntervencion " & _
                    "FROM TbMaterialSeguimiento " & _
                    "WHERE (((TbMaterialSeguimiento.IDSeguimiento)='" & strIDIntervencionAnterior & "'));"
            Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
            With rcdDatos
                If .EOF Then
                    strTextoError = "Intervención no registrada"
                    Err.Raise 1000
                End If
                strFechaFinIntervencionAnterior = Nz(.Fields("FechaFinIntervencion"), "")
            End With
            rcdDatos.Close
            Set rcdDatos = Nothing
            m_SQL = "SELECT TbMaterialSeguimiento.FechaInicioIntervencion, TbMaterialSeguimiento.FechaFinIntervencion " & _
                    "FROM TbMaterialSeguimiento " & _
                    "WHERE (((TbMaterialSeguimiento.IDSeguimiento)='" & strIDIntervencionPosterior & "'));"
            Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
            With rcdDatos
                If .EOF Then
                    strTextoError = "Intervención no registrada"
                    Err.Raise 1000
                End If
                strFechaInicionIntervencionPosterior = Nz(.Fields("FechaInicioIntervencion"), "")
            End With
            rcdDatos.Close
            Set rcdDatos = Nothing
        End If
    strCodigoFinal = strNumeroIntervencionesEnSeguimiento & ";" & strOrdenIntervencionActual & ";" & _
                strFechaFinIntervencionAnterior & ";" & strFechaInicionIntervencionPosterior
    IntervencionDatosDefechasIntermedias = strCodigoFinal
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método IntervencionDatosDefechasIntermedias ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    IntervencionDatosDefechasIntermedias = "#ERR" & "|" & strTextoError
End Function
Public Function EliminarMaterialEnActividad( _
                                                strIDMaterial As String, _
                                                Optional strEliminando As String, _
                                                Optional strRespetandoElFranqueo As String, _
                                                Optional strEventoFranqueado As String, _
                                                Optional strEstaFacturado As String _
                                                ) As String
     '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 25/01/13
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       -Sirve para eliminar un registro de la tabla TbMaterial
    '       -strIDMaterial ha de ser existente
    '       -No ha de estar franqueado el evento del que depende
    '       -también se ha de borrar el anexo que tuviera
    '   -Llamada desde
    '       -Form_FormEventoGestion.ComandoEliminarMaterial_Click
    '   -Devuelve:
    '       EliminarMaterialEnActividad = Descriptivo
    '       EliminarMaterialEnActividad = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    Dim rcdDatos As DAO.Recordset, strNombreArchivo As String, strURLAnexo As String, strIDActividadEnMaterial As String, _
        colAnexosAEliminar As New Collection, VarItem As Variant, strTextoError As String
    On Error GoTo errores
    If strEliminando <> "Sí" And strEliminando <> "No" Then
        strEliminando = "Sí"
    End If
    If strIDMaterial = "" Then
        strTextoError = "Se ha de indicar el IDMaterial"
        Err.Raise 1000
    End If
    If strRespetandoElFranqueo <> "Sí" And strRespetandoElFranqueo <> "No" Then
        strRespetandoElFranqueo = "No"
    End If
     If strRespetandoElFranqueo = "Sí" Then
         If strEventoFranqueado <> "Sí" And strEventoFranqueado <> "No" Then
             '-----------------------------------------------
             '   -Devuelve:
             '       EstaFranqueado = strValorFranqueado
             '       EstaFranqueado = "#ERR" & "|" & strTextoError
             '-----------------------------------------------
             flag = EstaFranqueado("MA", strIDMaterial)
             If InStr(1, flag, "|") <> 0 Then
                 dato = Split(flag, "|")
                 strTextoError = "El método EstaFranqueado ha devuelto el error: " & vbNewLine & dato(1)
                 Err.Raise 1000
             End If
             strEventoFranqueado = flag
         End If
         If strEventoFranqueado = "Sí" Then
             strTextoError = "El evento del que depende la actividad de este material ya está franqueado"
             Err.Raise 1000
         End If
    End If
    If strEstaFacturado <> "Sí" And strEstaFacturado <> "No" Then
        '-----------------------------------------------
        '   -Devuelve:
        '       EstaFacturado = "Sí" "No"
        '       EstaFacturado = "#ERR" & "|" & strTextoError
        '-----------------------------------------------
        flag = EstaFacturado("MA", strIDMaterial)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método EstaFacturado ha devuelto un error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        strEstaFacturado = flag
        If strEstaFacturado <> "Sí" And strEstaFacturado <> "No" Then
            strTextoError = "El método MiFactura.EstaFacturado ha devuelto resultado con formato desconocido"
            Err.Raise 1000
        End If
    End If
    If strEstaFacturado = "Sí" Then
        strTextoError = "EL material ya está facturado"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbAnexos.NombreArchivo " & _
            "FROM TbMaterialSeguimiento INNER JOIN TbAnexos ON " & _
            "TbMaterialSeguimiento.IDSeguimiento = TbAnexos.IDMaterialSeguimiento " & _
            "WHERE (((TbMaterialSeguimiento.IDMaterial)='" & strIDMaterial & "'));"
    
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            
            
            .MoveFirst
            Do While Not .EOF
                strNombreArchivo = Nz(.Fields("NombreArchivo"), "")
                strURLAnexo = m_ObjEntorno.URLDirectorioAnexos & strNombreArchivo
                If fso.FileExists(strURLAnexo) Then
                    '------------------------------------------
                    '   -Devuelve
                    '       FormularioAbierto = true or false
                    '       FormularioAbierto =False-->Error
                    '------------------------------------------
                    If FicheroAbierto(strURLAnexo) Then
                        strTextoError = "Al menos hay un anexo abierto: " & vbNewLine & strNombreArchivo
                        Err.Raise 1000
                    End If
                    colAnexosAEliminar.Add strURLAnexo
                End If
                
                .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If strEliminando = "Sí" Then
        m_SQL = "DELETE TbMaterialSeguimiento.IDMaterial " & _
                "FROM TbMaterialSeguimiento " & _
                "WHERE (((TbMaterialSeguimiento.IDMaterial)='" & strIDMaterial & "'));"
        DoCmd.SetWarnings False
        DoCmd.RunSQL m_SQL
        DoCmd.SetWarnings True
    End If
   m_SQL = "SELECT TbMaterial.* " & _
            "FROM TbMaterial " & _
            "WHERE (((TbMaterial.IDMaterial)='" & strIDMaterial & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strTextoError = "No se ha encontrado el material a eliminar"
            Err.Raise 1000
        End If
        If strEliminando = "Sí" Then
            .Delete
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If strEliminando = "Sí" Then
        If colAnexosAEliminar.count > 0 Then
            For Each VarItem In colAnexosAEliminar
                strURLAnexo = CStr(VarItem)
                If fso.FileExists(strURLAnexo) Then
                    fso.DeleteFile strURLAnexo, True
                End If
            Next
        End If
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método EliminarMaterialEnActividad ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    
    EliminarMaterialEnActividad = "#ERR" & "|" & strTextoError
End Function
Public Function EliminarMaterialSinActividad( _
                                                strIDMaterial As String, _
                                                Optional strEliminando As String _
                                                ) As String
     '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 25/01/13
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       -Sirve para eliminar un registro de la tabla TbMaterial
    '       -strIDMaterial ha de ser existente
    '       -No debe provenir de una actividad ---->Me.MaterialProvieneDeActividad
   
    '   -Llamada desde
    '       -Form_FormMaterialesGestion.ComandoEliminarMatSinActividad_Click
    '   -Devuelve:
    '       EliminarMaterialSinActividad =Descriptivo
    '       EliminarMaterialSinActividad ="#ERR" & "|" & strTextoError
    '-----------------------------------------------
    Dim rcdDatos As DAO.Recordset, strNombreArchivo As String, strURLAnexo As String, strIDActividadEnMaterial As String, _
        colAnexosAEliminar As New Collection, VarItem As Variant, strTextoError As String
    On Error GoTo errores
    If strIDMaterial = "" Then
        strTextoError = "Se ha de indicar el IDMaterial"
        Err.Raise 1000
    End If
     If strEliminando <> "Sí" And strEliminando <> "No" Then
        strEliminando = "Sí"
    End If
    '-----------------------------------------------
    '   -Devuelve:
    '       MaterialProvieneDeActividad = "Sí" "No"
    '       MaterialProvieneDeActividad = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = MaterialProvieneDeActividad(strIDMaterial)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método MaterialProvieneDeActividad ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If flag <> "No" Then
        strTextoError = "No se puede eliminar mediante este procedimiento un material que proviene de una actividad"
        Err.Raise 1000
    End If
    If flag = "Sí" Then
        '-----------------------------------------------
        '   -Devuelve:
        '       EstaFranqueado = "Sí" "No"
        '       EstaFranqueado = "#ERR" & "|" & strTextoError
        '-----------------------------------------------
        flag = EstaFranqueado("MA", strIDMaterial)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método EstaFranqueado ha devuelto un error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        If flag <> "Sí" And flag <> "No" Then
            strTextoError = "El método Evento.EstaFranqueado ha devuelto resultado con formato desconocido"
            Err.Raise 1000
        End If
        If flag = "Sí" Then
            strTextoError = "EL material depende de un evento que está franqueado. Ha de quitar el franqueo y posteriormente eliminar su material"
            Err.Raise 1000
        End If
    End If
    '-----------------------------------------------
    '   -Devuelve:
    '       EstaFacturado = "Sí" "No"
    '       EstaFacturado = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    flag = EstaFacturado("MA", strIDMaterial)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método EstaFacturado ha devuelto un error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    If flag <> "Sí" And flag <> "No" Then
        strTextoError = "El método EstaFacturado ha devuelto resultado con formato desconocido"
        Err.Raise 1000
    End If
    If flag = "Sí" Then
        strTextoError = "EL material ya está facturado"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbAnexos.NombreArchivo " & _
            "FROM TbMaterialSeguimiento INNER JOIN TbAnexos ON " & _
            "TbMaterialSeguimiento.IDSeguimiento = TbAnexos.IDMaterialSeguimiento " & _
            "WHERE (((TbMaterialSeguimiento.IDMaterial)='" & strIDMaterial & "'));"
    
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            
            
            .MoveFirst
            Do While Not .EOF
                strURLAnexo = Nz(.Fields("NombreArchivo"), "")
                If strURLAnexo <> "" Then
                    If InStr(1, strURLAnexo, ":\") <> 0 Or InStr(1, strURLAnexo, "\\") <> 0 Then
                        strNombreArchivo = fso.GetFileName(strURLAnexo)
                    Else
                        strNombreArchivo = strURLAnexo
                    End If
                    strURLAnexo = m_ObjEntorno.URLDirectorioAnexos & strNombreArchivo
                    If fso.FileExists(strURLAnexo) Then
                        '------------------------------------------
                        '   -Devuelve
                        '       FormularioAbierto = true or false
                        '       FormularioAbierto =False-->Error
                        '------------------------------------------
                        If FicheroAbierto(strURLAnexo) Then
                            strTextoError = "Al menos hay un anexo abierto: " & vbNewLine & strNombreArchivo
                            Err.Raise 1000
                        End If
                        colAnexosAEliminar.Add strURLAnexo
                    End If
                End If
                .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If strEliminando = "Sí" Then
        m_SQL = "DELETE TbMaterialSeguimiento.IDMaterial " & _
                "FROM TbMaterialSeguimiento " & _
                "WHERE (((TbMaterialSeguimiento.IDMaterial)='" & strIDMaterial & "'));"
        DoCmd.SetWarnings False
        DoCmd.RunSQL m_SQL
        DoCmd.SetWarnings True
    End If
    m_SQL = "SELECT TbMaterial.* " & _
            "FROM TbMaterial " & _
            "WHERE (((TbMaterial.IDMaterial)='" & strIDMaterial & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strTextoError = "No se ha encontrado el material a eliminar"
            Err.Raise 1000
        End If
        If strEliminando = "Sí" Then
            .Delete
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If strEliminando = "Sí" Then
        If colAnexosAEliminar.count > 0 Then
            For Each VarItem In colAnexosAEliminar
                strURLAnexo = CStr(VarItem)
                If fso.FileExists(strURLAnexo) Then
                    fso.DeleteFile strURLAnexo, True
                End If
            Next
        End If
    End If
    
    EliminarMaterialSinActividad = "Material eliminado correctamente"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método EliminarMaterialSinActividad ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    
    EliminarMaterialSinActividad = "#ERR" & "|" & strTextoError
End Function
Public Function MaterialDeActividadAlta( _
                                            strIDActividad As String, _
                                            strMaterial As String, _
                                            Optional strCOSTE As String, _
                                            Optional strTipoAccion As String, _
                                            Optional strPN As String, _
                                            Optional strNS As String, _
                                            Optional strReparadoPor As String, _
                                            Optional strDESCRIPCION As String, _
                                            Optional strFechaInicial As String, _
                                            Optional strGarantia As String, _
                                            Optional strEsReparacion As String, _
                                            Optional strFechaEntrega As String _
                                            ) As String

     '-------------------------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 25/01/13
    '   -Modificaciones:
    '       -09/07/13: es opcional strCoste y strFechaEntrega
    '       -11/09/2014: Se incluye la FechaInicial y la Garantía
    '       -17/09/2014: No se permite tocar la fecha de entrega. Será consecuencia de los seguimientos
    '       -06/10/2014: Se incluye un campo del tbMaterial -->EsReparacion.
    '   -Funcionamiento:
    '       -va a dar de alta un registro en la tabla TbMaterial
    '       -Siempre va a depender de una IDActividad
    '       -Ésta ha de depender siempre de un IDEvento y éste que no esté franqueado
    '       -Si no hay coste ha de ser 0
    '       -Se considera acabado cuando está relleno el Reparadopor
   
    '   -Llamada desde
    '       -Form_FormMaterialRepuestoAlta.ComandoAlta_Click
    '       -Form_FormMaterialAltaReparacion.ComandoAlta_Click
    '   -Devuelve:
    '       MaterialDeActividadAlta = strIDMaterial
    '       MaterialDeActividadAlta = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    Dim rcdDatos As DAO.Recordset, strIDMaterial As String, strWhereNS As String, strWhereMaterial As String, _
        strWhereIDActividad As String, strWherePN As String, strWhereEsReparacion As String, strParteWhere As String, _
        strSQLInicial As String, strTextoError As String
    On Error GoTo errores
    
    If strIDActividad = "" Then
        strTextoError = "Se ha de indicar la actividad de la que depende el material"
        Err.Raise 1000
    End If
    If strMaterial = "" Then
        strTextoError = "Se ha de indicar el material"
        Err.Raise 1000
    End If
    If strCOSTE <> "" Then
       If Not IsNumeric(strCOSTE) Then
            strTextoError = "Se ha de indicar el coste con un valor numérico válido"
            Err.Raise 1000
        End If
    End If
    If Not IsDate(strFechaInicial) Then
        strFechaInicial = Now()
    End If
    If strGarantia <> "Sí" And strGarantia <> "No" Then
        strGarantia = "No"
    End If
    If strEsReparacion <> "Sí" And strEsReparacion <> "No" Then
        strEsReparacion = "No"
    End If
    If strEsReparacion = "No" Then
        If strFechaEntrega <> "" Then
            If Not IsDate(strFechaEntrega) Then
                strTextoError = "El campo fecha de entrega no tiene un formato correcto"
                Err.Raise 1000
            End If
        End If
    End If
    strWhereIDActividad = "((TbMaterial.IDActividad)='" & strIDActividad & "')"
    strWhereMaterial = "((TbMaterial.Material)='" & strMaterial & "')"
    strWhereEsReparacion = "((TbMaterial.EsReparacion)='" & strEsReparacion & "')"
    If strPN = "" Then
        strWherePN = "((TbMaterial.PN) Is Null)"
    Else
        strWherePN = "((TbMaterial.PN)='" & strPN & "')"
    End If
    If strNS = "" Then
        strWhereNS = "((TbMaterial.NS) Is Null)"
    Else
        strWhereNS = "((TbMaterial.NS)='" & strNS & "')"
    End If
    strSQLInicial = "SELECT TbMaterial.Material, TbMaterial.PN, TbMaterial.NS " & _
                    "FROM TbMaterial INNER JOIN TbActividades ON TbMaterial.IDActividad = TbActividades.IDActividad "
    strParteWhere = "WHERE (" & strWhereIDActividad & " AND " & _
                            strWhereMaterial & " AND " & _
                            strWhereEsReparacion & " AND " & _
                            strWhereNS & " AND " & _
                            strWherePN & ");"
    m_SQL = strSQLInicial & strParteWhere
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            strTextoError = "Ya se ha dado de alta este material para esta actividad"
            Err.Raise 1000
        End If
    End With
    rcdDatos.Close
    '------------------------------------------
    '   -Devuelve:
    '       DameValorSiguienteAlMaximoDeUnCampo = strParteFijaInicial & Format(CStr(lngValorMaximo + 1), String(intNumeroCaracteresDesdeLaDerecha, "0"))
    '       DameValorSiguienteAlMaximoDeUnCampo = "#ERR" & "|" & strTextoError
    '--------------------------------------------
    flag = DameValorSiguienteAlMaximoDeUnCampo("TbMaterial", "IDMaterial", strIDActividad & "_", 2)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método DameValorSiguienteAlMaximoDeUnCampo ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strIDMaterial = flag
    m_SQL = "SELECT TbMaterial.* " & _
            "FROM TbMaterial " & _
            "WHERE (((TbMaterial.IDMaterial)='" & strIDMaterial & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            strTextoError = "La función DameValorSiguienteAlMaximoDeUnCampo ha devuelto un valor de IDMaterial ya existente"
            Err.Raise 1000
        End If
        .AddNew
            .Fields("IDMaterial") = strIDMaterial
            .Fields("IDActividad") = strIDActividad
            .Fields("Material") = strMaterial
            .Fields("FechaIncio") = strFechaInicial
            .Fields("Garantia") = strGarantia
            If strPN <> "" Then
                .Fields("PN") = strPN
            End If
            If strNS <> "" Then
                .Fields("NS") = strNS
            End If
            If IsNumeric(strCOSTE) Then
                .Fields("Coste") = strCOSTE
            End If
            If strTipoAccion <> "" Then
                .Fields("TipoAccion") = strTipoAccion
            End If
            If strReparadoPor <> "" Then
                .Fields("ReparadoPor") = strReparadoPor
            End If
            If strDESCRIPCION <> "" Then
                .Fields("Descripcion") = strDESCRIPCION
            End If
            .Fields("EsReparacion") = strEsReparacion
            If strEsReparacion = "No" Then
                If IsDate(strFechaEntrega) Then
                    .Fields("FechaEntrega") = strFechaEntrega
                End If
            End If
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    MaterialDeActividadAlta = strIDMaterial
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método MaterialDeActividadAlta ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    MaterialDeActividadAlta = "#ERR" & "|" & strTextoError
End Function
Public Function DameIDEquipoEnFuncionDeNodoBuiSubSistemaNombre( _
                                                                strNodo As String, _
                                                                strBUI As String, _
                                                                strSUBSISTEMA As String, _
                                                                strEquipo As String _
                                                                ) As String
    
     '-------------------------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 11/03/13
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       -va a buscar en la tabla TbEquipos el IDEquipo
    '   -Llamada desde
    '       -Form_FormEventoAlta.Registrar
    '   -Devuelve:
    '       DameIDEquipoEnFuncionDeNodoBuiSubSistemaNombre = strIDEquipo
    '       DameIDEquipoEnFuncionDeNodoBuiSubSistemaNombre = "#ERR" & "|" & strTextoError
    '------------------------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, strTextoError As String
    On Error GoTo errores
    If strNodo = "" Then
        strTextoError = "No se puede dejar vacío el Nodo"
        Err.Raise 1000
    End If
    If strBUI = "" Then
        strTextoError = "No se puede dejar vacío el BUI"
        Err.Raise 1000
    End If
    If strSUBSISTEMA = "" Then
        strTextoError = "No se puede dejar vacío el Subsistema"
        Err.Raise 1000
    End If
    If strEquipo = "" Then
        strTextoError = "No se puede dejar vacío el Equipo"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbEquipos.IDEquipo " & _
            "FROM TbEquipos " & _
            "WHERE (((TbEquipos.NODO)='" & strNodo & _
            "') AND ((TbEquipos.BUI)='" & strBUI & _
            "') AND ((TbEquipos.SUBSISTEMA)='" & strSUBSISTEMA & _
            "') AND ((TbEquipos.Equipo)='" & strEquipo & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strTextoError = "No se puede ha encontrado el IDEquipo"
            Err.Raise 1000
        End If
        DameIDEquipoEnFuncionDeNodoBuiSubSistemaNombre = Nz(.Fields("IDEquipo"), "")
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método DameIDEquipoEnFuncionDeNodoBuiSubSistemaNombre ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    DameIDEquipoEnFuncionDeNodoBuiSubSistemaNombre = "#ERR" & "|" & strTextoError
End Function
Public Function RegenerarTablaFacturacion(strURLBBDatosAntiguaConDatosBuenos As String) As String
    Dim dbOrigen As DAO.Database
    Dim strTextoError As String
    
    
    RegenerarTablaFacturacion = "OK"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método RegenerarTablaFacturacion ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
    RegenerarTablaFacturacion = "#ERR" & "|" & strTextoError
End Function
Public Function ActividadAlta( _
                        strAliasTecnico As String, _
                        strDESCRIPCION As String, _
                        strFechaAlta As String, _
                        strUbicacion As String, _
                        Optional strHorasLaborables As String, _
                        Optional strHORASEXTRAS As String, _
                        Optional strTipoActividad As String, _
                        Optional strIDEvento As String, _
                        Optional strIDEventoGenerado As String, _
                        Optional strAvisoMaximasHoras As String, _
                        Optional strEventoFacturado As String _
                        ) As String

     '-------------------------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 10/01/13
    '   -Modificaciones:
    '       -30/05/2013: se ha de poder poner 0 horas laborables y 0 horas no laborables para poder meter un material que sólo dependa de evento
    '       -04/06/2013: No se puede dar de ActividadAlta con horas lab superiores a 7,5--->DameHorasTecnicoDia
    '   -Funcionamiento:
    '       -va a tomar todos los datos que entran y los va a meter TbActividades--->
    '       -una actividad, puede depender de un evento o no
    '       -una actividad sólo puede ser en una jornada, no hay fecha final, si se acaba en el día se empieza otra en el día siguiente
    '       -1) Verificación de datos
    '           -strIDEvento Existente si está relleno
    '           -strIDEventoGenerado Existente si está relleno
    '           -strDescripcion obligatorio
    '           -strFechaAlta Fecha, si depende de un strIDEvento ha de ser igual o mayor que ésta
    '           -strHorasLaborables número distinto de 0 si strHorasExtras="" o strHorasExtras="#" or strHorasExtras="0"
    '           -strHorasExtras si relleno número
    '           -strTipoActividad obligatoria <>"#"
    '           -strHoraFin>strHoraAlta
    '   -Llamada desde
    '       -Form_FormActividadAlta.ComandoAlta_Click
    '   -Devuelve:
    '       ActividadAlta = strIDActividad
    '       ActividadAlta = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    Dim rcdDatos As DAO.Recordset, strEventoFranqueado As String, strIDActividad As String, strFECHAALTAEVENTO As String, strFranqueoEvento As String
    Dim strAño As String, strMes As String, strParteFija As String, strHoraMinuto As String, strDia As String, strHora As String, strMinuto As String
    Dim strFechaAhora As String, strHorasYaParaElDia As String, dblHorasYa As Double, dblHorasTotales As Double, strTextoError As String
    On Error GoTo errores
    
    If strAvisoMaximasHoras <> "Sí" And strAvisoMaximasHoras <> "No" Then
        strAvisoMaximasHoras = "No"
    End If
    If strIDEvento <> "" And strIDEventoGenerado <> "" Then
        If strIDEventoGenerado = strIDEvento Then
            strTextoError = "No puede ser igual el IDEvento que el IDEvento generado"
            Err.Raise 1000
        End If
    End If
    If Not IsDate(strFechaAlta) Then
        strTextoError = "La fecha de ActividadAlta ha de ser una fecha válida "
        Err.Raise 1000
    End If
    If strIDEvento <> "" Then
        '------------------------------------------
        '   -Devuelve:
        '       Dame = strValorObtenido
        '       Dame = "#ERR" & "|" & strTextoError
        '--------------------------------------------
        flag = Dame("TbEventos", "FECHAALTAEVENTO", "IDEvento", strIDEvento)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método Dame ha devuelto el error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        strFECHAALTAEVENTO = flag
        If Not IsDate(strFECHAALTAEVENTO) Then
            strTextoError = "El ID del evento ha de existir previamente "
            Err.Raise 1000
        End If
        If CDate(strFechaAlta) < CDate(strFECHAALTAEVENTO) Then
            strTextoError = "La fecha de ActividadAlta de la actividad no puede ser inferior de la de apertura del evento a la que pertenece "
            Err.Raise 1000
        End If
        '-----------------------------------------------
        '   -Devuelve:
        '       EstaFranqueado = strEventoFranqueado
        '       EstaFranqueado = "#ERR" & "|" & strTextoError
        '-----------------------------------------------
        flag = EstaFranqueado("EV", strIDEvento)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método EstaFranqueado ha devuelto el error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        strEventoFranqueado = flag
        If strEventoFranqueado = "Sí" Then
            strTextoError = "El Evento del que depende esta actividad ya está franqueado. Ha de quitar el franqueo para dar de ActividadAlta una nueva actividad"
            Err.Raise 1000
        End If
        If strEventoFacturado <> "Sí" And strEventoFacturado <> "No" Then
            '-----------------------------------------------
            '   -Devuelve:
            '       EstaFacturado = Sí/No
            '       EstaFacturado = "#ERR" & "|" & strTextoError
            '-----------------------------------------------
            flag = EstaFacturado("EV", strIDEvento)
            If InStr(1, flag, "|") <> 0 Then
                dato = Split(flag, "|")
                strTextoError = "El método EstaFacturado ha devuelto un error: " & vbNewLine & dato(1)
                Err.Raise 1000
            End If
            strEventoFacturado = flag
        End If
        If strEventoFacturado = "Sí" Then
            strTextoError = "No se puede dar de alta una actividad de un evento ya facturado"
            Err.Raise 1000
        End If
    End If
    If Not IsNumeric(strHorasLaborables) Then
        strHorasLaborables = "0"
    End If
    If Not IsNumeric(strHORASEXTRAS) Then
        strHORASEXTRAS = "0"
    End If
    If strAvisoMaximasHoras = "Sí" Then
        If CDbl(strHorasLaborables) > 7.5 Then
            strTextoError = "Las horas laborables no pueden exceder de 7,5 h"
            Err.Raise 1000
        End If
         '-------------------------------------------------------------------
        '   -Devuelve:
        '       DameHorasTecnico = cstr(dblHoras)
        '       DameHorasTecnico = "#ERR" & "|" & strTextoError
        '-------------------------------------------------------------------
        flag = DameHorasTecnico(strAliasTecnico, strFechaAlta)
        If InStr(1, flag, "|") <> 0 Then
            flag = ""
        End If
        strHorasYaParaElDia = flag
        If IsNumeric(strHorasYaParaElDia) Then
            dblHorasYa = CDbl(strHorasYaParaElDia)
            dblHorasTotales = dblHorasYa + CDbl(strHorasLaborables)
            If dblHorasTotales > 7.5 Then
                strTextoError = "Para el " & strFechaAlta & " " & strAliasTecnico & " ya llevaba " & CStr(dblHorasYa) & " y con " & strHorasLaborables & " horas sumaría " & CStr(dblHorasTotales) & " que supera el tope de 7,5"
                Err.Raise 1000
            End If
        End If
    End If
    If strTipoActividad = "" Then
        strTextoError = "Se ha de indicar el tipo de actividad "
        Err.Raise 1000
    End If
    strFechaAhora = Now()
    strAño = Format(CDate(strFechaAlta), "yy")
    strMes = Format(CDate(strFechaAlta), "mm")
    strDia = Format(CDate(strFechaAlta), "dd")
    strHoraMinuto = Format(CDate(strFechaAhora), "hh:mm")
    dato = Split(strHoraMinuto, ":")
    strHora = Nz(dato(0), "")
    strMinuto = Nz(dato(1), "")
    strParteFija = strAliasTecnico & "_" & strAño & strMes & strDia & strHora & strMinuto
    '------------------------------------------
    '   -Devuelve:
    '       DameValorSiguienteAlMaximoDeUnCampo = strParteFijaInicial & Format(CStr(lngValorMaximo + 1), String(intNumeroCaracteresDesdeLaDerecha, "0"))
    '       DameValorSiguienteAlMaximoDeUnCampo = "#ERR" & "|" & strTextoError
    '--------------------------------------------
    flag = DameValorSiguienteAlMaximoDeUnCampo("TbActividades", "IDActividad", strParteFija, 2)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método DameValorSiguienteAlMaximoDeUnCampo ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strIDActividad = flag
    If strIDActividad = "#ERR" Then
        strTextoError = "La función DameIDActividad ha devuelto un valor con un formato desconocido"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbActividades.* " & _
            "FROM TbActividades " & _
            "WHERE (((TbActividades.IDActividad)='" & strIDActividad & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            strTextoError = "La función DameIDActividad ha devuelto un valor de IDActividad ya existente"
            Err.Raise 1000
        End If
        .AddNew
            .Fields("IDActividad") = strIDActividad
            If strIDEvento <> "" Then
                .Fields("IDEvento") = strIDEvento
            End If
            If strAliasTecnico <> "" Then
                .Fields("ALIASTecnico") = strAliasTecnico
            End If
            If IsDate(strFechaAlta) Then
                .Fields("FechaAlta") = strFechaAlta
            End If
            .Fields("HorasLaborables") = CDbl(strHorasLaborables)
            .Fields("HorasExtras") = CDbl(strHORASEXTRAS)
            
            If strTipoActividad <> "" Then
                .Fields("Actividad") = strTipoActividad
            End If
            If strUbicacion <> "" Then
                .Fields("Ubicacion") = strUbicacion
            End If
            If strDESCRIPCION <> "" Then
                .Fields("DESCRIPCION") = strDESCRIPCION
            End If
            
            If strIDEventoGenerado <> "" Then
                .Fields("IDEVENTOGENERADO") = strIDEventoGenerado
            End If
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    ActividadAlta = strIDActividad
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método ActividadAlta ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    ActividadAlta = "#ERR" & "|" & strTextoError
End Function
Public Function ActividadEliminar( _
                                    strIDActividad As String, _
                                    Optional strEliminando As String, _
                                    Optional strRespetandoElFranqueo As String, _
                                    Optional strEstaFacturado As String _
                                    ) As String
     '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 14/01/13
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       -Sirve para ActividadEliminar un registro de la tabla TbActividades y también los que haya en TbMateriales vinculados a la actividad
    '       -1) Verificación de datos
    '           -strIDActividad Existente
    '           -si tiene anexos se han de ActividadEliminar también
    '           -si pertenecen a un evento que está franqueado no se puede
    
    '   -Llamada desde
    '       -Form_FormEventoGestion.ComandoEliminarActividad_Click
    '   -Devuelve:
    '       ActividadEliminar = cstr(intNumeroMaterialesEliminados) & "[" & strCadenaMaterialesNoEliminados
    '       ActividadEliminar = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    Dim rcdDatos As DAO.Recordset, strIDEventoDeActividad As String
    Dim intNumeroMaterialesEliminados As Integer, strCadenaMaterialesNoEliminados As String, strEventoFranqueado As String
    Dim strCadenaAnexosAEliminar As String, strCadenaAnexosEliminados As String, strIDMaterial As String, strTextoError As String
    On Error GoTo errores
    
    If strIDActividad = "" Then
        strTextoError = "Se ha de indicar el IDActividad a ActividadEliminar"
        Err.Raise 1000
    End If
     If strEliminando <> "Sí" And strEliminando <> "No" Then
        strEliminando = "Sí"
    End If
    If strRespetandoElFranqueo <> "Sí" And strRespetandoElFranqueo <> "No" Then
        strRespetandoElFranqueo = "No"
    End If
    If strEstaFacturado <> "Sí" And strEstaFacturado <> "No" Then
        '-------------------------------------------------------------------
        '   -Devuelve:
        '       EstaFacturado = strEstaFacturado
        '       EstaFacturado = "#ERR" & "|" & strTextoError
        '-------------------------------------------------------------------
        flag = EstaFacturado("AC", strIDActividad)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método EstaFacturado ha devuelto el error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        strEstaFacturado = flag
        If strEstaFacturado <> "Sí" And strEstaFacturado <> "No" Then
            strTextoError = "No se puede determinar si la actividad está o no facturada"
            Err.Raise 1000
        End If
    End If
    If strEstaFacturado = "Sí" Then
        strTextoError = "Esta actividad ha formado parte de una factura, hasta que no la elimine, no se puede borrar la misma"
        Err.Raise 1000
    End If
    If strRespetandoElFranqueo = "Sí" Then
         '------------------------------------------
        '   -Devuelve:
        '       Dame = strValorObtenido
        '       Dame = "#ERR" & "|" & strTextoError
        '--------------------------------------------
        flag = Dame("TbActividades", "IDEvento", "IDActividad", strIDActividad)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método Dame ha devuelto el error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        strIDEventoDeActividad = flag
        If strIDEventoDeActividad <> "" Then
            '-----------------------------------------------
            '   -Devuelve:
            '       EstaFranqueado = strValorFranqueado
            '       EstaFranqueado = "#ERR" & "|" & strTextoError
            '-----------------------------------------------
            flag = EstaFranqueado("EV", strIDEventoDeActividad)
            If InStr(1, flag, "|") <> 0 Then
                dato = Split(flag, "|")
                strTextoError = "El método EstaFranqueado ha devuelto el error: " & vbNewLine & dato(1)
                Err.Raise 1000
            End If
            strEventoFranqueado = flag
            If strEventoFranqueado = "Sí" Then
                strTextoError = "La actuvidad : " & strIDActividad & " depende del evento: " & strIDEventoDeActividad & " que está franqueado"
                Err.Raise 1000
            End If
        End If
    End If
    Dim m_Actividad As Actividad
    Set m_Actividad = Constructor.getActividad(strIDActividad, strTextoError)
    If strTextoError <> "" Then
        Err.Raise 1000
    End If
    If strEliminando <> "Sí" Then
        If Not m_Actividad.ColAnexos Is Nothing Then
            If m_Actividad.AnexosCerrados <> EnumSino.Sí Then
                strTextoError = "Al menos hay un anexo abierto"
                Err.Raise 1000
            End If
        End If
    Else
        m_Actividad.EliminarAnexos strTextoError
        If strTextoError <> "" Then
            Err.Raise 1000
        End If
    End If
   
    '--------------
    ' borramos todos los materiales que pudiera tener asociados la actividad
    '-----------------
    m_SQL = "SELECT TbMaterial.IDMaterial " & _
                "FROM TbMaterial " & _
                "WHERE (((TbMaterial.IDActividad)='" & strIDActividad & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                strIDMaterial = Nz(.Fields("IDMaterial"), "")
                '-------------------------------------------------------------------
                '   -Devuelve:
                '       EliminarMaterialEnActividad = strValorObtenido
                '       EliminarMaterialEnActividad = "#ERR" & "|" & strTextoError
                '-------------------------------------------------------------------
                flag = EliminarMaterialEnActividad(strIDMaterial, strEliminando, strRespetandoElFranqueo, strEstaFacturado)
                If InStr(1, flag, "|") <> 0 Then
                    dato = Split(flag, "|")
                    strTextoError = "El método EliminarMaterialEnActividad ha devuelto un error: " & vbNewLine & dato(1)
                    Err.Raise 1000
                End If
                If InStr(1, flag, "|") <> 0 Then
                    If strCadenaMaterialesNoEliminados = "" Then
                        strCadenaMaterialesNoEliminados = strIDMaterial
                    Else
                        strCadenaMaterialesNoEliminados = strCadenaMaterialesNoEliminados & ";" & strIDMaterial
                    End If
                Else
                    intNumeroMaterialesEliminados = intNumeroMaterialesEliminados + 1
                End If
                .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    m_SQL = "SELECT TbActividades.* " & _
            "FROM TbActividades " & _
            "WHERE (((TbActividades.IDActividad)='" & strIDActividad & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strTextoError = "No se ha encontrado la actividad a copiar"
            Err.Raise 1000
        End If
        If strEliminando = "Sí" Then
            .Delete
        End If
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    ActividadEliminar = CStr(intNumeroMaterialesEliminados) & "[" & strCadenaMaterialesNoEliminados
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Actividad.ActividadEliminar ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    ActividadEliminar = "#ERR" & "|" & strTextoError
End Function

Public Function EstaFacturado( _
                                strPrefijo As String, _
                                strID As String _
                                ) As String
                               
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 07/05/13
    '   -Modificaciones:
    '       -05/9/14:Incluye los gastos
    '   -Funcionamiento:
    '       -se va a mirar en la tabla TbFacturaEventosInvolucrados ó TbFacturaActividadesInvolucradas ó TbFacturaMaterialesInvolucrados
    '          en función de qué variables se rellenen
    '   -Llamada por:
    '       -me.AñadirRegistroEnTbAux
    '       -Form_FormEventoGestion.ComandoEliminarEvento_Click
    '       -Form_FormActividadEdicion.EstablecerDatos
    '       -Actividad.Eliminar
    '       -Actividad.Edicion
    '       -Form_FormMaterialReparacionEdicion.EstablecerDatos
    '   -Devuelve:
    '       strPrefijo=EV,AC,MA,SUB,GA
    '       EstaFacturado = strEstaFacturado
    '       EstaFacturado = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, strEstaFacturado As String, strNombreTabla As String, strNombreCampo As String, strEsNumericoCampoClave As String, _
        strParteWhere As String, stSQLInicial As String, strTextoError As String
    On Error GoTo errores
    If strPrefijo = "EV" Then
        strNombreTabla = "TbFacturaEventosInvolucrados"
        strNombreCampo = "IDEvento"
        strEsNumericoCampoClave = "NO"
    ElseIf strPrefijo = "AC" Then
        strNombreTabla = "TbFacturaActividadesInvolucradas"
        strNombreCampo = "IDActividad"
        strEsNumericoCampoClave = "NO"
    ElseIf strPrefijo = "MA" Then
        strNombreTabla = "TbFacturaMaterialesInvolucrados"
        strNombreCampo = "IDMaterial"
        strEsNumericoCampoClave = "NO"
    ElseIf strPrefijo = "SUB" Then
        strNombreTabla = "TbFacturaSubcontratacionesInvolucradas"
        strNombreCampo = "IDSubContratacion"
        strEsNumericoCampoClave = "Sí"
    ElseIf strPrefijo = "GA" Then
        strNombreTabla = "TbFacturaGastosInvolucrados"
        strNombreCampo = "IDGasto"
        strEsNumericoCampoClave = "Sí"
    Else
        strTextoError = "El tipo de Anexo ha de ser EV,AC,MA,AN o PAR"
        Err.Raise 1000
    End If
    stSQLInicial = "SELECT " & strNombreTabla & "." & strNombreCampo & " as IDDato " & _
                "FROM " & strNombreTabla & " "

    If strEsNumericoCampoClave = "No" Then
        strParteWhere = "WHERE (((" & strNombreTabla & "." & strNombreCampo & ")='" & strID & "'));"
    Else
        strParteWhere = "WHERE (((" & strNombreTabla & "." & strNombreCampo & ")=" & strID & "));"
    End If
    m_SQL = stSQLInicial & strParteWhere
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    If rcdDatos.EOF Then
        strEstaFacturado = "No"
    Else
        strEstaFacturado = "Sí"
    End If
    rcdDatos.Close
    Set rcdDatos = Nothing
    EstaFacturado = strEstaFacturado
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Facturacion.EstaFacturado ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    EstaFacturado = "#ERR" & "|" & strTextoError
End Function













Public Function DameCadenaMaterialesSinEntregaPorEvento( _
                                                            strIDEvento As String _
                                                        ) As String
     '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 11/01/13
    '   -Modificaciones:
    '
    '   -Funcionamiento:
    '       -va a partir de una consulta de materiales por evento y si a alguno le falta la fecha de entrega se añade a la cadena strCadenaReparaciones
    '       strCadenaMaterialesSinEntregaPorEvento= IDMaterial1;IDMaterial2;....;IDMaterialn
    '       -1) Verificación de datos
    '           -strIDEvento  existente
        
    '   -Llamada desde
    '       -Form_FormEventoGestion.ComandoEliminarEvento_Click
    '       -Evento.Franquear
    '   -Devuelve:
        
    '       DameCadenaMaterialesSinEntregaPorEvento =strCadenaMaterialesSinEntregaPorEvento
    '       DameCadenaMaterialesSinEntregaPorEvento ="#ERR" & "|" & strTextoError
    '-----------------------------------------------
    Dim rcdDatos As DAO.Recordset, strCadenaMaterialesSinEntregaPorEvento As String, strFechaEntrega As String, strIDMaterial As String, strTextoError As String
    On Error GoTo errores
    If strIDEvento = "" Then
        strTextoError = "No se puede dejar en blanco el IDEvento"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbMaterial.IDMaterial, TbMaterial.FechaEntrega " & _
                "FROM (TbEventos INNER JOIN TbActividades ON TbEventos.IDEvento = TbActividades.IDEvento) " & _
                "INNER JOIN TbMaterial ON TbActividades.IDActividad = TbMaterial.IDActividad " & _
                "WHERE (((TbEventos.IDEvento)='" & strIDEvento & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                strIDMaterial = Nz(.Fields("IDMaterial"), "")
                strFechaEntrega = Nz(.Fields("FechaEntrega"), "")
                If Not IsDate(strFechaEntrega) Then
                    If strCadenaMaterialesSinEntregaPorEvento = "" Then
                        strCadenaMaterialesSinEntregaPorEvento = strIDMaterial
                    Else
                        strCadenaMaterialesSinEntregaPorEvento = strCadenaMaterialesSinEntregaPorEvento & ";" & strIDMaterial
                    End If
                End If
                .MoveNext
            Loop
        End If
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    DameCadenaMaterialesSinEntregaPorEvento = strCadenaMaterialesSinEntregaPorEvento
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Evento.DameCadenaMaterialesSinEntregaPorEvento ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    DameCadenaMaterialesSinEntregaPorEvento = "#ERR" & "|" & strTextoError
End Function
Public Function DameCadenaActividadesPorEvento(strIDEvento As String) As String
     '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 11/01/13
    '   -Modificaciones:
    '
    '   -Funcionamiento:
    '       -va a ir una a una las actividades del evento dado y devuelve para cada una
    '           -IDActividad & ";" & "Sí/No"
    '       strCadenaActividades= IDActividad1 & ";" & "Sí/No" & "#" & _
    '                                       IDActividad2 & ";" & "Sí/No" & "#" & _
    '                                       IDActividad3 & ";" & "Sí/No" & "#" & _
    '                                       IDActividadn & ";" & "Sí/No" & "#" & _
    '       -1) Verificación de datos
    '           -strIDEvento  existente
        
    '   -Llamada desde
    '       -Evento.Franquear
    '   -Devuelve:
        
    '       DameCadenaActividadesPorEvento = strCadenaActividades
    '       DameCadenaActividadesPorEvento = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    Dim rcdDatos As DAO.Recordset, strIDEventoEnBD As String, strCadenaActividades As String, strIDExportacion As String
    Dim strIDActividad As String, strValorExportacion As String, strTextoError As String
    On Error GoTo errores
    If strIDEvento = "" Then
        strTextoError = "No se puede dejar en blanco el IDEvento"
        Err.Raise 1000
    End If
     m_SQL = "SELECT TbActividades.* " & _
            "FROM TbActividades " & _
            "WHERE (((TbActividades.IDEvento)='" & strIDEvento & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                strIDActividad = Nz(.Fields("IDACTIVIDAD"), "")
                If Not strIDActividad = "" Then
                    strIDExportacion = Nz(.Fields("IDExportacion"), "")
                    If IsDate(strIDExportacion) Then
                        strValorExportacion = "Sí"
                    Else
                        strValorExportacion = "No"
                    End If
                    If strCadenaActividades = "" Then
                        strCadenaActividades = strIDActividad & ";" & strValorExportacion
                    Else
                        strCadenaActividades = strCadenaActividades & "#" & strIDActividad & ";" & strValorExportacion
                    End If
                End If
                .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    DameCadenaActividadesPorEvento = strCadenaActividades
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Evento.DameCadenaActividadesPorEvento ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    DameCadenaActividadesPorEvento = "#ERR" & "|" & strTextoError
End Function

Public Function DameCadenaMaterialesPorEvento(strIDEvento As String) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 11/01/13
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       -va a mirar en la tabla Materiales con la visibilidad que se le pide la cadena de Materiales de ese evento
    '   -Llamada desde
    '   -Devuelve:
    '       DameCadenaMaterialesPorEvento = strCadenaMateriales
    '       DameCadenaMaterialesPorEvento = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    Dim rcdDatos As DAO.Recordset, strIDMat As String, strCadenaMateriales As String, strTextoError As String
    On Error GoTo errores
    m_SQL = "SELECT TbMaterial.IDMaterial " & _
                "FROM (TbEventos INNER JOIN TbActividades ON TbEventos.IDEvento = TbActividades.IDEvento) " & _
                "INNER JOIN TbMaterial ON TbActividades.IDActividad = TbMaterial.IDActividad " & _
                "WHERE (((TbEventos.IDEvento)='" & strIDEvento & "'));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
                strIDMat = Nz(.Fields("IDMaterial"), "")
                If strCadenaMateriales = "" Then
                    strCadenaMateriales = strIDMat
                Else
                    strCadenaMateriales = strCadenaMateriales & ";" & strIDMat
                End If
                .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    DameCadenaMaterialesPorEvento = strCadenaMateriales
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método DameCadenaMaterialesPorEvento ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    DameCadenaMaterialesPorEvento = "#ERR" & "|" & strTextoError
End Function

Public Function EstaFranqueado( _
                                strPrefijo As String, _
                                strID As String _
                                ) As String
     '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 15/01/13
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       -Si introducen strIDEvento ó strIDActividad ó strIDMaterial, consiste en obtener el strIDEventoObtenido que es único y ver si
    '           TbEventos.Franqueado es "Verdadero"
    '       -Si introducen strIDEquipo, he de recorrer la tabla TbEventos y ver si algún campo TbEventos.Franqueado es "Verdadero"
   
    '   -Llamada desde
    '       -Form_FormEventoGestion.ComandoCopiarActividad_Click
    '       -Actividad.EsActividadDeEventoFranquead
    '       -Material.MaterialDeActividadAlta
    '       -Actividad.Alta
    '       -Material.EdicionMaterialEnActividad
    '       -Form_FormEventoGestion.ComandoAltaActividad_Click
    '       -Form_FormEventoGestion.ComandoAltaMaterial_Click
    '       -Material.EliminarMaterialEnActividad
    '       -Material.EditarIDEquipo
    '       -Form_FormEventoGestion.ComandoEditarMaterial_Click
    '       -Form_FormEventoGestion.ComandoEliminarMaterial_Click
    '       -Form_FormEventoGestion.ComandoFranquear_Click
    '       -Form_FormEventoGestion.ComandoEliminarEvento_Click
    '       -Actividad.Eliminar
    '       -Form_FormEventoGestion.ComandoEditarEvento_Click
    '       -Form_FormActividadGestion.ComandoEditarActividad_Click
    '       -Parte.CadenaIDEventoPreparadoParaParte
    '       -Facturacion.GrabarRegistroEnTbAux
    '       -Form_FormMaterialReparacionEdicion.EstablecerDatos
    '       -Form_FormEquipoEdicion.EditarIDEquipo
    '   -Devuelve:
    '       strPrefijo=EV,AC,MA,EQ,
    '       EstaFranqueado = strValorFranqueado
    '       EstaFranqueado = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    Dim rcdDatos As DAO.Recordset, strValorFranqueado As String, strTextoError As String
    On Error GoTo errores
    If strPrefijo = "EV" Then
        '-----------------------------------------------
        '   -Devuelve:
        '       Dame = strValor
        '       Dame = "#ERR" & "|" & strTextoError
        '-----------------------------------------------
        flag = Dame("TbEventos", "Franqueado", "IDEvento", strID)
        If InStr(1, flag, "|") <> 0 Then
            dato = Split(flag, "|")
            strTextoError = "El método Dame ha devuelto un error: " & vbNewLine & dato(1)
            Err.Raise 1000
        End If
        If flag = "Verdadero" Or flag = "True" Then
            EstaFranqueado = "Sí"
        Else
            EstaFranqueado = "No"
        End If
        Exit Function
    ElseIf strPrefijo = "AC" Then
        m_SQL = "SELECT TbEventos.Franqueado " & _
                "FROM TbEventos INNER JOIN TbActividades ON TbEventos.IDEvento = TbActividades.IDEvento " & _
                "WHERE (((TbActividades.IDActividad)='" & strID & "'));"
        Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
        With rcdDatos
            If .EOF Then
                EstaFranqueado = "No"
            Else
                If .Fields("Franqueado") Then
                    EstaFranqueado = "Sí"
                Else
                    EstaFranqueado = "No"
                End If
            End If
        End With
        rcdDatos.Close
        Set rcdDatos = Nothing
        Exit Function
    ElseIf strPrefijo = "MA" Then
        m_SQL = "SELECT TbEventos.Franqueado " & _
                    "FROM (TbEventos INNER JOIN TbActividades ON TbEventos.IDEvento = TbActividades.IDEvento) INNER JOIN TbMaterial ON " & _
                    "TbActividades.IDActividad = TbMaterial.IDActividad " & _
                    "WHERE (((TbMaterial.IDMaterial)='" & strID & "'))"
        Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
        With rcdDatos
            If .EOF Then
                EstaFranqueado = "No"
            Else
                If .Fields("Franqueado") Then
                    EstaFranqueado = "Sí"
                Else
                    EstaFranqueado = "No"
                End If
            End If
        End With
        rcdDatos.Close
        Set rcdDatos = Nothing
        Exit Function
    ElseIf strPrefijo = "EQ" Then
        m_SQL = "SELECT TbEventos.Franqueado " & _
                    "FROM TbEventos " & _
                    "WHERE (((TbEventos.IDEquipo)=" & strID & "));"
        Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
        With rcdDatos
            If Not .EOF Then
                .MoveFirst
                Do While Not .EOF
                    strValorFranqueado = Nz(.Fields("Franqueado"), "")
                    If strValorFranqueado = "Verdadero" Then
                        rcdDatos.Close
                        Set rcdDatos = Nothing
                        EstaFranqueado = "Sí"
                        Exit Function
                    End If
                    .MoveNext
                Loop
            End If
        End With
        rcdDatos.Close
        Set rcdDatos = Nothing
        EstaFranqueado = "No"
        Exit Function
    Else
        strTextoError = "El tipo de ID ha de ser EV,AC,MA,EQ"
        Err.Raise 1000
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método EstaFranqueado ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    EstaFranqueado = "#ERR" & "|" & strTextoError
End Function
Public Function DameIDEvento(strBUI As String, strFECHAALTAEVENTO As String) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 18/01/13
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       -En general el código es strInicialesBUI & aamm & xxx
    '       -Donde strInicialesBUI en principio es la que venga en la tabla TbBuiIDEvento, de no venir, se ponen las tres primeras iniciales
    '       -Donde aamm es las dos ultimas cifras del año de la strFechaAltaEvento y mm es el mes de la misma
    '       - xxx es el ordinal expresado en tres cifras siguiente al que haya
    '   -Llamada desde
    '       -Evento.Edicion
    '       -Evento.Alta
    '   -Devuelve:
        
    '       DameIDEvento = strIDEvento
    '       DameIDEvento = "#ERR" & "|" & strTextoError
    '-----------------------------------------------
    Dim rcdDatos As DAO.Recordset, strIDEvento As String, strOrdinal As String, strInicialesBUI As String, strMM As String, strAA As String
    Dim strParticulaComparativa As String, strTextoError As String
    On Error GoTo errores
    If strBUI = "" Then
        strTextoError = "No se puede dejar en blanco el BUI"
        Err.Raise 1000
    End If
    If Not IsDate(strFECHAALTAEVENTO) Then
        strTextoError = "No se puede dejar en blanco la fecha de alta del evento"
        Err.Raise 1000
    End If
    '------------------------------------------
    '   -Devuelve:
    '       Dame = strValorObtenido
    '       Dame = "#ERR" & "|" & strTextoError
    '--------------------------------------------
    flag = Dame("TbBuiIDEvento", "CODBUIIDVENTO", "BUI", strBUI)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método Dame ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strInicialesBUI = flag
    If strInicialesBUI = "" Then
        If Len(strBUI) < 3 Then
            strTextoError = "El BUI no viene en la tabla de BUI Código Evento y además no tiene una longitud mínima de 3 caracteres"
            Err.Raise 1000
        End If
        strInicialesBUI = Left(strBUI, 3)
    End If
    strMM = Format(CDate(strFECHAALTAEVENTO), "mm")
    strAA = Format(CDate(strFECHAALTAEVENTO), "yy")
    strParticulaComparativa = strInicialesBUI & strAA & strMM
    '--------------
    ' ordinal
    '-------------
    m_SQL = "SELECT Right([IDEvento],3) AS Ordinal " & _
            "FROM TbEventos " & _
            "WHERE (((TbEventos.IDEvento) Like '" & strParticulaComparativa & "*')) " & _
            "ORDER BY Right([IDEvento],3) DESC;"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strOrdinal = "0"
        Else
            strOrdinal = Nz(.Fields("Ordinal"), "0")
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If Not IsNumeric(strOrdinal) Then
        strOrdinal = "0"
    End If
    strIDEvento = strParticulaComparativa & Format(CLng(strOrdinal) + 1, "000")
    DameIDEvento = strIDEvento
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Evento.DameIDEvento ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    DameIDEvento = "#ERR" & "|" & strTextoError
End Function

Public Function EdicionMaterialSinActividad( _
                                            strIDMaterial As String, _
                                            Optional strMaterial As String, _
                                            Optional strCOSTE As String, _
                                            Optional strTipoAccion As String, _
                                            Optional strPN As String, _
                                            Optional strNS As String, _
                                            Optional strReparadoPor As String, _
                                            Optional strDESCRIPCION As String, _
                                            Optional strFechaInicial As String, _
                                            Optional strGarantia As String, _
                                            Optional strEsReparacion As String, _
                                            Optional strFechaEntrega As String _
                                            ) As String
    '-------------------------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 10/04/13
    '   -Modificaciones:
    '       -11/09/2014: Se incluye la FechaInicial y la Garantía
    '       -06/10/2014: Se incluye un campo del tbMaterial -->EsReparacion.
    '   -Funcionamiento:
    '       -va a editar un registro en la tabla TbMaterial
    
    '       -strMaterial<>"" y strMaterial<>"#"
    '       -Si no hay coste ha de ser 0
    
    '   -Llamada desde
    
    '   -Devuelve:
    '       EdicionMaterialSinActividad = Descriptivo
    '       EdicionMaterialSinActividad = "#ERR" & "|" & strTextoError
    '----------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, strFacturado As String, strTextoError As String
    On Error GoTo errores
    
    If strIDMaterial = "" Or strIDMaterial = "#" Then
        strTextoError = "Se ha de indicar el material que se pretende editar "
        Err.Raise 1000
    End If
    If strMaterial = "" Or strMaterial = "#" Then
        strTextoError = "Se ha de indicar el material de que se trata "
        Err.Raise 1000
    End If
    If strCOSTE <> "" And strCOSTE <> "#" Then
        If Not IsNumeric(strCOSTE) Then
            strTextoError = "Se introducir un coste distinto de nulo ha de ser un número "
            Err.Raise 1000
        End If
    End If
     If Not IsDate(strFechaInicial) Then
        strTextoError = "Es obligatoria la fecha inicial"
        Err.Raise 1000
    End If
    If strGarantia <> "Sí" And strGarantia <> "No" Then
        strTextoError = "Garantía sólo puede ser Sí o No"
        Err.Raise 1000
    End If
     If strEsReparacion <> "Sí" And strEsReparacion <> "No" Then
        strTextoError = "Se ha de saber si es una reparación o no"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbFacturaMaterialesInvolucrados.IDFactura " & _
            "FROM TbFacturaMaterialesInvolucrados " & _
            "WHERE IDMaterial='" & strIDMaterial & "';"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            strFacturado = Nz(.Fields("IDFactura"), "")
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    If strFacturado <> "" Then
        strTextoError = "Material ya facturado"
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbMaterial.* " & _
            "FROM TbMaterial " & _
            "WHERE IDMaterial='" & strIDMaterial & "';"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            strTextoError = "No existe el material con código interno: " & strIDMaterial
            Err.Raise 1000
        End If
        .Edit
            .Fields("Material") = strMaterial
            .Fields("FechaIncio") = strFechaInicial
            .Fields("Garantia") = strGarantia
            If strPN <> "" And strPN <> "#" Then
                .Fields("PN") = strPN
            Else
                If strPN = "#" Then
                    .Fields("PN") = Null
                End If
            End If
            If strNS <> "" And strNS <> "#" Then
                .Fields("NS") = strNS
            Else
                .Fields("NS") = Null
            End If
            If IsNumeric(strCOSTE) Then
                .Fields("Coste") = strCOSTE
            Else
                If strCOSTE = "#" Then
                    
                    .Fields("Coste") = Null
                End If
            End If
            If IsDate(strFechaEntrega) Then
                .Fields("FechaEntrega") = strFechaEntrega
            End If
            If strReparadoPor <> "" And strReparadoPor <> "#" Then
                .Fields("ReparadoPor") = strReparadoPor
            Else
                If strReparadoPor = "#" Then
                    .Fields("ReparadoPor") = Null
                End If
            End If
            If strTipoAccion <> "" And strTipoAccion <> "#" Then
                .Fields("TipoAccion") = strTipoAccion
            Else
                If strTipoAccion = "#" Then
                    .Fields("TipoAccion") = Null
                End If
            End If
            If strDESCRIPCION <> "" And strDESCRIPCION <> "#" Then
                .Fields("Descripcion") = strDESCRIPCION
            Else
                If strDESCRIPCION = "#" Then
                    .Fields("Descripcion") = Null
                End If
            End If
             .Fields("EsReparacion") = strEsReparacion
            If strEsReparacion = "No" Then
                If IsDate(strFechaEntrega) Then
                    .Fields("FechaEntrega") = strFechaEntrega
                Else
                    If strFechaEntrega = "#" Then
                        .Fields("FechaEntrega") = Null
                    End If
                End If
            End If
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    EdicionMaterialSinActividad = "Edición Correcta"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método Material.EdicionMaterialSinActividad ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not rcdDatos Is Nothing Then
        rcdDatos.Close
        Set rcdDatos = Nothing
    End If
    EdicionMaterialSinActividad = "#ERR" & "|" & strTextoError
End Function

Public Function LimpiarDirectorioTrabajo() As String
    Dim fld As Folder, archivo As File, ColArchivosNoTocar As New Collection, strTextoError As String, intNumeroArchivosEliminados As Integer
    On Error GoTo errores
    
    If Not fso.FolderExists(m_ObjEntorno.URLDirectorioLocal) Then
        
        LimpiarDirectorioTrabajo = CStr(intNumeroArchivosEliminados)
        Exit Function
    End If
    ColArchivosNoTocar.Add "IconoAplicacion.ico"
    ColArchivosNoTocar.Add fso.GetBaseName(CurrentDb().Name) & ".ini"
    ColArchivosNoTocar.Add fso.GetBaseName(CurrentDb().Name) & ".accdb"
    ColArchivosNoTocar.Add fso.GetBaseName(CurrentDb().Name) & ".accde"
    ColArchivosNoTocar.Add fso.GetBaseName(CurrentDb().Name) & ".laccdb"
    For Each fld In fso.GetFolder(m_ObjEntorno.URLDirectorioLocal).SubFolders
        intNumeroArchivosEliminados = CInt(fld.Files.count)
        flag = Eliminar_Directorio(fld.Path)
    Next
    For Each archivo In fso.GetFolder(m_ObjEntorno.URLDirectorioLocal).Files
        '-------------------------------------------------------------------
        '   -Devuelve:
        '       EstaEnColeccion = Sí/No
        '       EstaEnColeccion = "ERR" & "|" & strTextoError
        '-------------------------------------------------------------------
        If EstaEnColeccion(ColArchivosNoTocar, archivo.Name) = "No" Then
            If Not FicheroAbierto(archivo.Path) Then
                fso.DeleteFile archivo.Path, True
                intNumeroArchivosEliminados = intNumeroArchivosEliminados + 1
            
            End If
        End If
    Next
    
    LimpiarDirectorioTrabajo = CStr(intNumeroArchivosEliminados)
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método LimpiarDirectorioTrabajo ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
    LimpiarDirectorioTrabajo = "#ERR" & "|" & strTextoError
End Function
Function Eliminar_Directorio(Path As String) As Boolean
  
On Error GoTo Error_Sub
  
   
  
    'Le pasamos a DeleTeFolder el Path a eliminar
    fso.DeleteFolder Path, True
  
    If Err.Number = 0 Then
       ' Ok
       Eliminar_Directorio = True
       
    End If
      

Exit Function
Error_Sub:
  
MsgBox Err.Description, vbCritical
  
End Function
Public Function EstaEnColeccion(ByRef col As Collection, strElemento As String, Optional strSeparador As String, Optional intOrdinal As Integer) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día X
    '   -Modificaciones:
    
    '   -Funcionamiento:
    '       -Va a rellenar una tabla auxiliar TbAuxSumSinEmail con los que no tengan correo
    '           Public intNFacturasPVT
    '           Public intNumeroDPDsAFaltaDeAlgunAnexo As Integer
    '           Public dteHoraUltimoContador As date
    '   -Llamada desde
   
    '   -Devuelve:
    '       EstaEnColeccion = Sí/No
    '       EstaEnColeccion = "ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim VarItem As Variant, strTextoError As String, dato As Variant, strElementoDeCol As String
    On Error GoTo errores
    If strElemento = "" Then
        strTextoError = "No se ha indicado el elemento a buscar"
        Err.Raise 1000
    End If
    
    If col.count = 0 Then
        EstaEnColeccion = "No"
        Exit Function
    End If
    For Each VarItem In col
        strElementoDeCol = CStr(VarItem)
        If strSeparador <> "" Then
            If InStr(1, strElementoDeCol, strSeparador) <> 0 Then
                dato = Split(strElementoDeCol, strSeparador)
                If UBound(dato) >= intOrdinal Then
                    If dato(intOrdinal) = strElemento Then
                        EstaEnColeccion = "Sí"
                        Exit Function
                    End If
                End If
            End If
        Else
            If strElementoDeCol = strElemento Then
                EstaEnColeccion = "Sí"
                Exit Function
            End If
        End If
        
    Next
    EstaEnColeccion = "No"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método EstaEnColeccion ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    EstaEnColeccion = "ERR" & "|" & strTextoError
End Function
Public Function DameIDActividad(strAliasTecnico As String, strFechaAlta As String) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 15/09/2020
    '   -Modificaciones:
    
    '   -Funcionamiento:
    
    '   -Llamada desde
   
    '   -Devuelve:
    '       DameIDActividad = strIDActividad
    '       DameIDActividad = "ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim strFechaAhora As String, strAño As String, strMes As String, strDia As String, strHoraMinuto As String, strHora As String, strMinuto As String, strParteFija As String, _
        strIDActividad As String, strTextoError As String
    On Error GoTo errores
    If strAliasTecnico = "" Then
        strTextoError = "Se ha de indicar el Alias del técnico"
        Err.Raise 1000
    End If
    If Not IsDate(strFechaAlta) Then
        strTextoError = "Se ha de indicar la fecha del alta"
        Err.Raise 1000
    End If
    
    strAño = Format(CDate(strFechaAlta), "yy")
    strMes = Format(CDate(strFechaAlta), "mm")
    strDia = Format(CDate(strFechaAlta), "dd")
    strHoraMinuto = Format(CDate(Now()), "hh:mm")
    dato = Split(strHoraMinuto, ":")
    strHora = Nz(dato(0), "")
    strMinuto = Nz(dato(1), "")
    strParteFija = strAliasTecnico & "_" & strAño & strMes & strDia & strHora & strMinuto
    '------------------------------------------
    '   -Devuelve:
    '       DameValorSiguienteAlMaximoDeUnCampo = strParteFijaInicial & Format(CStr(lngValorMaximo + 1), String(intNumeroCaracteresDesdeLaDerecha, "0"))
    '       DameValorSiguienteAlMaximoDeUnCampo = "#ERR" & "|" & strTextoError
    '--------------------------------------------
    flag = DameValorSiguienteAlMaximoDeUnCampo("TbActividades", "IDActividad", strParteFija, 2)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método DameValorSiguienteAlMaximoDeUnCampo ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strIDActividad = flag
    DameIDActividad = strIDActividad
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método DameIDActividad ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    DameIDActividad = "ERR" & "|" & strTextoError
End Function
Public Function DameIDMaterial(strIDActividad As String) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 15/09/2020
    '   -Modificaciones:
    
    '   -Funcionamiento:
    
    '   -Llamada desde
   
    '   -Devuelve:
    '       DameIDMaterial = strIDActividad
    '       DameIDMaterial = "ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim strIDMaterial As String, strTextoError As String
    On Error GoTo errores
    If strIDActividad = "" Then
        strTextoError = "Se ha de indicar el IDActividad"
        Err.Raise 1000
    End If
    '------------------------------------------
    '   -Devuelve:
    '       DameValorSiguienteAlMaximoDeUnCampo = strParteFijaInicial & Format(CStr(lngValorMaximo + 1), String(intNumeroCaracteresDesdeLaDerecha, "0"))
    '       DameValorSiguienteAlMaximoDeUnCampo = "#ERR" & "|" & strTextoError
    '--------------------------------------------
    flag = DameValorSiguienteAlMaximoDeUnCampo("TbMaterial", "IDMaterial", strIDActividad & "_", 2)
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método DameValorSiguienteAlMaximoDeUnCampo ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strIDMaterial = flag
    DameIDMaterial = strIDMaterial
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método DameIDMaterial ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    DameIDMaterial = "ERR" & "|" & strTextoError
End Function

Public Function RellenaMaterialesCosteCero(strIDEvento As String, ColMateriales As Collection) As String
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día 17/09/2020
    '   -Modificaciones:
   
    '   -Funcionamiento:
   
    '   -llamada por:
    
    '   -Devuelve:
    '       RellenaMaterialesCosteCero = CStr(colMateriales.Count)
    '       RellenaMaterialesCosteCero = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim rcdDatos As DAO.Recordset, strIDMaterial As String, strMaterial As String, strTextoError As String
    On Error GoTo errores
    Set ColMateriales = New Collection
    m_SQL = "SELECT  TbMaterial.IDMaterial " & _
            "FROM (TbEventos INNER JOIN TbActividades ON TbEventos.IDEvento = TbActividades.IDEvento) INNER JOIN TbMaterial ON TbActividades.IDActividad = TbMaterial.IDActividad " & _
            "WHERE (((TbEventos.IDEvento)='" & strIDEvento & "') AND ((TbMaterial.COSTE)=0) AND ((TbMaterial.TipoAccion)<>'IDENTIFICADO_COMO_IRREPARABLE_OBSOLETO' Or (TbMaterial.TipoAccion) Is Null));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
             .MoveFirst
            Do While Not .EOF
                strIDMaterial = Nz(.Fields("IDMaterial"), "")
                ColMateriales.Add strIDMaterial
                .MoveNext
            Loop
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    RellenaMaterialesCosteCero = CStr(ColMateriales.count)
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método DameIDMaterial ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    RellenaMaterialesCosteCero = "ERR" & "|" & strTextoError
End Function

Public Function EstablecerComboContacto(Optional strBUI As String) As String
   
    Dim strTextoError As String, m_SQL As String
    On Error GoTo errores
    If strBUI = "" Then strBUI = "-1"
    m_SQL = "SELECT distinct TbEventos.CONTACTO " & _
            "FROM TbEventos " & _
            "WHERE ((Not (TbEventos.CONTACTO) Is Null) " & _
            "AND ((TbEventos.BUI)='" & strBUI & "')) " & _
            "ORDER BY TbEventos.CONTACTO;"
    EstablecerComboContacto = m_SQL
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método EstablecerComboContacto ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    EstablecerComboContacto = "ERR" & "|" & strTextoError
End Function

Public Function DiaLaborableParaTecnico( _
                                            m_Alias As String, _
                                            m_Fecha As String, _
                                            Optional ByRef p_Error As String _
                                            ) As EnumSino
    Dim rcdDatos As DAO.Recordset
    Dim Dia As Long
    
   
    
    On Error GoTo errores
    If m_Alias = "" Then
        p_Error = "Falta el alias"
        Err.Raise 1000
    End If
    If Not IsDate(m_Fecha) Then
        p_Error = "Falta la fecha"
        Err.Raise 1000
    End If
    Dia = VBA.Weekday(m_Fecha, vbMonday)
    If Dia = 6 Or Dia = 7 Then
        DiaLaborableParaTecnico = EnumSino.No
        Exit Function
    End If
    m_SQL = "SELECT TbTecnicosAusencias.FechaLibranza " & _
                "FROM TbTecnicosAusencias " & _
                "WHERE (((TbTecnicosAusencias.Alias)='" & m_Alias & "') " & _
                "AND ((TbTecnicosAusencias.FechaLibranza)=#" & Format(m_Fecha, "mm/dd/yyyy") & "#));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            DiaLaborableParaTecnico = EnumSino.No
            Exit Function
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    m_SQL = "SELECT TbTecnicosFiestas.FechaFiesta " & _
                "FROM TbTecnicosFiestas " & _
                "WHERE (((TbTecnicosFiestas.FechaFiesta)=#" & Format(m_Fecha, "mm/dd/yyyy") & "#));"
    Set rcdDatos = CurrentDb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            DiaLaborableParaTecnico = EnumSino.No
            Exit Function
        End If
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    DiaLaborableParaTecnico = EnumSino.Sí
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método DiaLaborableParaTecnico ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function

Public Function getValoresDistintos( _
                                    p_Control As ComboBox, _
                                    p_NombreTabla As String, _
                                    p_NombreCampo As String, _
                                    Optional ByRef p_Error As String _
                                    ) As String
    Dim rcdDatos As DAO.Recordset
    
    On Error GoTo errores
    p_Control.RowSource = ""
    
    m_SQL = "SELECT DISTINCT " & p_NombreCampo & " " & _
            "FROM " & p_NombreTabla & " " & _
            "WHERE NOT " & p_NombreCampo & " Is Null;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            p_Control.AddItem .Fields(0).Value
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getValoresDistintos ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function

Public Function getFechaInicialYFinalDeColeEventos(p_ColEventos As Scripting.Dictionary, Optional ByRef p_Error As String) As String
    
    Dim m_FechaInicialMinima As String
    Dim m_FechaFinalMaxima As String
    Dim m_FechaFinalCalculada As String
    Dim m_ID As Variant
    Dim m_ObjEvento As Evento
    
    On Error GoTo errores
    If p_ColEventos Is Nothing Then
        Exit Function
    End If
    For Each m_ID In p_ColEventos
        Set m_ObjEvento = p_ColEventos(m_ID)
        m_FechaFinalCalculada = m_ObjEvento.FechaFinal
        If IsDate(m_FechaFinalCalculada) Then
            If Not IsDate(m_FechaInicialMinima) Then
                m_FechaInicialMinima = m_FechaFinalCalculada
            Else
                If CDate(m_FechaFinalCalculada) < CDate(m_FechaInicialMinima) Then
                    m_FechaInicialMinima = m_FechaFinalCalculada
                End If
            End If
        End If
        If IsDate(m_FechaFinalCalculada) Then
            If Not IsDate(m_FechaFinalMaxima) Then
                m_FechaFinalMaxima = m_FechaFinalCalculada
            Else
                If CDate(m_FechaFinalCalculada) > CDate(m_FechaFinalMaxima) Then
                    m_FechaFinalMaxima = m_FechaFinalCalculada
                End If
            End If
        End If
        Set m_ObjEvento = Nothing
    Next
    If IsDate(m_FechaInicialMinima) And IsDate(m_FechaFinalMaxima) Then
        getFechaInicialYFinalDeColeEventos = m_FechaInicialMinima & "|" & m_FechaFinalMaxima
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getFechaInicialYFinalDeColeEventos ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function
Public Function RellenarInformeEnWord( _
                                        ByRef appWord As Word.Application, _
                                        ByRef Plantilla As Word.Document, _
                                        ByRef p_ObjColEventos As Scripting.Dictionary, _
                                        ByRef p_BUI As String, _
                                        ByRef p_URLCarpeta As String, _
                                        p_FechaMinima As String, _
                                        p_FechaMaxima As String, _
                                        Optional ByRef p_Error As String _
                                        ) As String
    
    
    
    Dim Informe As Word.Document
    Dim Tabla As Word.Table
    Dim VarItem As Variant
    Dim m_ObjEvento As Evento
    Dim m_IDEvento As Variant
    Dim m_URLInformeDoc As String
    Dim m_NombreOrigen As String
    Dim m_NombreDestino As String
    
    
    On Error GoTo errores
    
    If p_ObjColEventos Is Nothing Then
        p_Error = "No hay eventos"
        Err.Raise 1000
    End If
    If p_BUI = "" Then
        p_Error = "Se ha de indicar el centro para el informe"
        Err.Raise 1000
    End If
    If Not fso.FolderExists(p_URLCarpeta) Then
        p_Error = "La carpeta de destino"
        Err.Raise 1000
    End If
    
    m_URLInformeDoc = p_URLCarpeta & p_BUI & ".docx"
    
    If fso.FileExists(m_URLInformeDoc) Then
        If FicheroAbierto(m_URLInformeDoc) Then
            p_Error = "Debe cerrar el word que está abierto"
            Err.Raise 1000
        End If
    End If
    
     
    
    Set Informe = appWord.Documents.Add()
    
    m_NombreOrigen = appWord.Documents(Plantilla).Name
    m_NombreDestino = appWord.Documents(Informe).Name
    With appWord
        .Documents(Informe.Name).PageSetup.Orientation = wdOrientLandscape
        .Windows(.Documents(m_NombreOrigen)).Activate
        .Documents(m_NombreOrigen).tables(1).Select
        .Selection.Copy
        .Windows(.Documents(m_NombreDestino)).Activate
        .Selection.EndKey Unit:=wdStory
        .Selection.TypeParagraph
        .Selection.PasteAndFormat (wdFormatOriginalFormatting)
    End With
    With appWord.Documents(m_NombreDestino).tables(appWord.Documents(m_NombreDestino).tables.count)
        'FECHA INICIAL
        .Cell(2, 2).Range.Text = Format(p_FechaMinima, "dd/mm/yyyy")
        'FECHA FINAL
        .Cell(2, 4).Range.Text = Format(p_FechaMaxima, "dd/mm/yyyy")
    End With
    
    For Each m_IDEvento In p_ObjColEventos.Keys
        If Nz(m_IDEvento, "") = "" Then
            GoTo siguienteEvento
        End If
        Set m_ObjEvento = p_ObjColEventos(m_IDEvento)
        m_Linea = "Evento " & m_ObjEvento.IDEVENTO
        Avance m_Linea
        If m_ObjEvento.BUI <> p_BUI Then
             Set m_ObjEvento = Nothing
            GoTo siguienteEvento
        End If
        With appWord
            .Documents(Informe.Name).PageSetup.Orientation = wdOrientLandscape
            .Windows(.Documents(m_NombreOrigen)).Activate
            .Documents(m_NombreOrigen).tables(2).Select
            .Selection.Copy
            .Windows(.Documents(m_NombreDestino)).Activate
            .Selection.EndKey Unit:=wdStory
            .Selection.TypeParagraph
            .Selection.PasteAndFormat (wdFormatOriginalFormatting)
        End With
        With appWord.Documents(m_NombreDestino).tables(appWord.Documents(m_NombreDestino).tables.count)
        
            'Código PT(2,1)
            .Cell(2, 1).Range.Text = m_IDEvento
            'PM Cliente(2,2)
            .Cell(2, 2).Range.Text = m_ObjEvento.PMPR
            'Creador(2,3)
            .Cell(2, 3).Range.Text = m_ObjEvento.Tecnico.Nombre
            'Militar(2,4)
            .Cell(2, 4).Range.Text = "USUARIO/POC"
            'Apertura(2,5)
            .Cell(2, 5).Range.Text = m_ObjEvento.FECHAALTAEVENTO & " " & m_ObjEvento.HORAINICIALEVENTO
            'Cierre(2,6)
            .Cell(2, 6).Range.Text = m_ObjEvento.FechaFinal
            'T.Resp.(2,7)
            .Cell(2, 7).Range.Text = m_ObjEvento.TIEMPORESPUESTAEVENTO
            'SUBSIST(2,8)
            .Cell(2, 8).Range.Text = m_ObjEvento.SubSistema & " " & m_ObjEvento.Equipo.Equipo
            '
            'Causas(4,1)
            .Cell(4, 1).Range.Text = m_ObjEvento.Descripcion
            'Causas(4,2)
            .Cell(4, 2).Range.Text = m_ObjEvento.CausaFin
            VBA.DoEvents
            'Debug.Print m_IDEvento
            VBA.DoEvents
        End With
        Set m_ObjEvento = Nothing
siguienteEvento:
    Next
    With appWord
        .Documents(Informe.Name).PageSetup.Orientation = wdOrientLandscape
        .Windows(.Documents(m_NombreOrigen)).Activate
        .Documents(m_NombreOrigen).tables(3).Select
        .Selection.Copy
        .Windows(.Documents(m_NombreDestino)).Activate
        '.Selection.EndKey Unit:=wdStory
        .Selection.InsertBreak Type:=7
        .Selection.TypeParagraph
        .Selection.PasteAndFormat (wdFormatOriginalFormatting)
    End With
    With appWord
        If .ActiveWindow.View.SplitSpecial <> wdPaneNone Then
            .ActiveWindow.Panes(2).Close
        End If
        If .ActiveWindow.ActivePane.View.Type = wdNormalView Or .ActiveWindow.ActivePane.View.Type = wdOutlineView Then
            .ActiveWindow.ActivePane.View.Type = wdPrintView
        End If
        .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
        .Selection.TypeText Text:="Página "
        .Selection.Fields.Add Range:=.Selection.Range, Type:=wdFieldEmpty, Text:="PAGE  ", PreserveFormatting:=True
        .Selection.TypeText Text:=" de "
        .Selection.Fields.Add Range:=.Selection.Range, Type:=wdFieldEmpty, Text:="NUMPAGES  ", PreserveFormatting:=True
        .ActiveWindow.ActivePane.View.Type = wdNormalView
        .Selection.PageSetup.TopMargin = .CentimetersToPoints(2.75)
        .Selection.PageSetup.BottomMargin = .CentimetersToPoints(1.5)
    End With
    If fso.FileExists(m_URLInformeDoc) Then
        fso.DeleteFile m_URLInformeDoc, True
    End If
    m_Linea = "Grabando el documento "
    Avance m_Linea
    
    
    appWord.Documents(m_NombreDestino).SaveAs2 m_URLInformeDoc
    appWord.Documents(m_URLInformeDoc).Close False
    
    
    Set Informe = Nothing
    
    RellenarInformeEnWord = m_URLInformeDoc
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarInformeEnWord ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    If Not appWord.Documents(m_NombreOrigen) Is Nothing Then
        appWord.Documents(m_NombreOrigen).Close False
        Set Plantilla = Nothing
    End If
    
    
    
End Function



Public Function getListaEventosEntreFechas( _
                                            Optional p_ObjColEventos As Scripting.Dictionary, _
                                            Optional p_FechaInicial As String, _
                                            Optional p_FechaFinal As String, _
                                            Optional p_SoloFranqueados As EnumSino = EnumSino.No, _
                                            Optional p_BUI As String, _
                                            Optional ByRef p_Error As String _
                                            ) As Scripting.Dictionary
    Dim rcdDatos As DAO.Recordset
    Dim m_Campo As Variant
    Dim m_SQL As String
    Dim m_ObjEvento As Evento
    Dim m_IDEvento As Variant
    Dim m_FechaInicial As String
    Dim m_FechaFinal As String
    Dim m_Franqueado As String
    Dim m_DesdeCol As Boolean
    
    Dim m_SQLLImitante As String
    
    On Error GoTo errores
    If p_ObjColEventos Is Nothing Then
        m_DesdeCol = False
    Else
        If p_ObjColEventos.count = 0 Then
            m_DesdeCol = False
        Else
            m_DesdeCol = True
        End If
    End If
    If Not IsDate(p_FechaInicial) Then
        p_FechaInicial = "01/01/1900"
    End If
    If Not IsDate(p_FechaFinal) Then
        p_FechaFinal = Date
    End If
    
    If m_DesdeCol = True Then
        For Each m_IDEvento In p_ObjColEventos.Keys
            'Debug.Print m_IDEvento
            'If m_IDEvento = "BER2104001" Then Stop
            
            Set m_ObjEvento = Constructor.getEvento(CStr(m_IDEvento), p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            
            If p_BUI <> "" Then
                If p_BUI <> m_ObjEvento.BUI Then
                    GoTo siguiente1
                End If
            End If
            m_FechaInicial = m_ObjEvento.FECHAALTAEVENTO
            m_FechaFinal = m_ObjEvento.FechaFinal
            p_Error = m_ObjEvento.Error
            If p_Error <> "" Then
                
            End If
            m_Franqueado = m_ObjEvento.Franqueado
            If p_SoloFranqueados = EnumSino.No Then
                
                If IsDate(m_FechaFinal) Then
                    If CDate(m_FechaInicial) < CDate(p_FechaInicial) And CDate(m_FechaFinal) < CDate(p_FechaInicial) Then
                        GoTo siguiente1
                    End If
                    If CDate(m_FechaInicial) > CDate(p_FechaFinal) And CDate(m_FechaFinal) > CDate(p_FechaFinal) Then
                        GoTo siguiente1
                    End If
                End If
                
            Else
                If Not m_Franqueado Then
                    GoTo siguiente1
                End If
                If IsDate(m_FechaFinal) Then
                    If Not (CDate(m_FechaFinal) >= CDate(p_FechaInicial) And CDate(m_FechaFinal) <= CDate(p_FechaFinal)) Then
                        GoTo siguiente1
                    End If
                End If
            End If
            If getListaEventosEntreFechas Is Nothing Then
                Set getListaEventosEntreFechas = New Scripting.Dictionary
                getListaEventosEntreFechas.CompareMode = TextCompare
            End If
            If Not getListaEventosEntreFechas.Exists(m_ObjEvento.IDEVENTO) Then
                getListaEventosEntreFechas.Add m_ObjEvento.IDEVENTO, m_ObjEvento
            End If
            
            
            Set m_ObjEvento = Nothing
siguiente1:
        Next
        
        Exit Function
    End If
    If p_SoloFranqueados = EnumSino.Sí Then
        m_SQLLImitante = "SELECT TbEventos.IDEvento " & _
                        "FROM TbEventos " & _
                        "WHERE FechaFinal Between #" & _
                        Format(p_FechaInicial, "mm/dd/yyyy") & "# And #" & Format(p_FechaFinal, "mm/dd/yyyy") & _
                        "# AND Franqueado=True;"
    
    ElseIf p_SoloFranqueados = EnumSino.No Then
        m_SQLLImitante = "SELECT TbEventos.IDEvento " & _
                        "FROM TbEventos " & _
                        "WHERE Franqueado=false AND " & _
                        "FECHAALTAEVENTO Between #" & Format(p_FechaInicial, "mm/dd/yyyy") & _
                            "# And #" & Format(p_FechaFinal, "mm/dd/yyyy") & "#;"
        
    Else
        m_SQLLImitante = "SELECT TbEventos.IDEvento " & _
                        "FROM TbEventos " & _
                        "WHERE FECHAALTAEVENTO Between #" & Format(p_FechaInicial, "mm/dd/yyyy") & _
                            "# And #" & Format(p_FechaFinal, "mm/dd/yyyy") & "#;"
                            
      
    End If
    
    If p_BUI = "" Then
        m_SQL = "SELECT TbEventos.* " & _
                "FROM TbEventos " & _
                "WHERE (((TbEventos.IDEvento) In (" & m_SQLLImitante & ")));"
                
        
    Else
        m_SQL = "SELECT TbEventos.* " & _
                "FROM TbEventos " & _
                "WHERE (((TbEventos.IDEvento) In (" & m_SQLLImitante & ")) AND ((TbEventos.BUI)='" & p_BUI & "'));"
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            rcdDatos.Close
            Set rcdDatos = Nothing
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Set m_ObjEvento = New Evento
            For Each m_Campo In m_ObjEvento.ColCampos
                m_ObjEvento.SetPropiedad m_Campo, Nz(.Fields(m_Campo).Value, ""), p_Error
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            Next
            If getListaEventosEntreFechas Is Nothing Then
                Set getListaEventosEntreFechas = New Scripting.Dictionary
                getListaEventosEntreFechas.CompareMode = TextCompare
            End If
            If Not getListaEventosEntreFechas.Exists(CStr(m_ObjEvento.IDEVENTO)) Then
                getListaEventosEntreFechas.Add CStr(m_ObjEvento.IDEVENTO), m_ObjEvento
            End If
            Set m_ObjEvento = Nothing
            .MoveNext
        Loop
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getListaEventosEntreFechas ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function ValorExistente( _
                                p_NombreTabla As String, _
                                p_NombreCampo As String, _
                                p_ValorCampo As String, _
                                Optional ValorEsNumerico As EnumSino = EnumSino.No, _
                                Optional p_db As DAO.Database, _
                                Optional ByRef p_Error As String _
                                ) As EnumSino
    
    Dim rcdDatos As DAO.Recordset
    
    
    On Error GoTo errores
    
    If p_db Is Nothing Then
        Set p_db = getdb()
    End If
    If p_NombreTabla = "" Or p_NombreCampo = "" Then
        p_Error = "falta p_NombreTabla o p_NombreCampo"
        Err.Raise 1000
    End If
    If ValorEsNumerico = EnumSino.No Then
        m_SQL = "SELECT " & p_NombreTabla & "." & p_NombreCampo & " " & _
                "FROM " & p_NombreTabla & " " & _
                "WHERE " & p_NombreCampo & "='" & p_ValorCampo & "';"
    Else
        m_SQL = "SELECT " & p_NombreTabla & "." & p_NombreCampo & " " & _
                "FROM " & p_NombreTabla & " " & _
                "WHERE " & p_NombreCampo & "=" & p_ValorCampo & ";"
    End If
    
    Set rcdDatos = p_db.OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
           ValorExistente = EnumSino.No
        Else
            ValorExistente = EnumSino.Sí
        End If
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "EL método ValorExistente ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function RellenarComboTecnicos( _
                                        p_Cmb As ComboBox, _
                                        Optional p_EnTodos As EnumSino, _
                                        Optional p_SoloEnActivo As EnumSino = EnumSino.Sí, _
                                        Optional ByRef p_Error As String _
                                        ) As String

    
    Dim m_ObjTecnico As Tecnico
    Dim m_ID As Variant
    Dim m_Col As Scripting.Dictionary
    On Error GoTo errores
    
    
    p_Cmb.RowSource = "ALIAS;NOMBRE;TIPO"
    Set m_Col = Constructor.getTecnicos(p_EnTodos, p_SoloEnActivo, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Exit Function
    End If
    For Each m_ID In m_Col
        Set m_ObjTecnico = m_Col(m_ID)
        With m_ObjTecnico
            p_Cmb.AddItem .Alias & ";" & .Nombre & ";" & .Tipo
        End With
        Set m_ObjTecnico = Nothing
    Next
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarComboTecnicos ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function


Public Function Avance( _
                        p_Linea As String, _
                        Optional ByRef p_Error As String _
                        ) As String
    Dim lbl As Label
    Dim m_FormActivo As Form
    On Error GoTo errores
    Set m_FormActivo = Screen.ActiveForm
    If m_FormActivo Is Nothing Then
        Exit Function
    End If
    On Error Resume Next
    Set lbl = m_FormActivo.Controls("lblEstado")
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    
    lbl.Visible = True
    VBA.DoEvents
    lbl.Caption = p_Linea
    VBA.DoEvents
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Avance ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
End Function


Public Function RellenarListaEquiposMedida( _
                                            p_Form As Form, _
                                            Optional ByRef p_Error As String _
                                            ) As String
    
    Dim m_Col As Scripting.Dictionary
    Dim m_ID As Variant
    Dim m_EquipoMedida As EquipoMedida
    Dim m_Cmb As ComboBox
    
    On Error GoTo errores
    Set m_Cmb = p_Form.ComboEquipoDeMedida
    m_Cmb.RowSource = "ID;Equipo Medida;Calibración"
    Set m_Col = m_ObjEntorno.ColEquiposMedida
    p_Error = m_ObjEntorno.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Col Is Nothing Then
        Exit Function
    End If
    For Each m_ID In m_Col
        Set m_EquipoMedida = m_Col(m_ID)
        With m_EquipoMedida
            m_Cmb.AddItem .IDEquipoMedida & ";" & .NombreParaCombo & ";" & .EstadoCalibracion
        End With
        Set m_EquipoMedida = Nothing
    Next
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RellenarListaEquiposMedida ha devuelto el error: " & Err.Description
    End If
    
End Function




Public Function FechaEnIntervalo( _
                                    m_FechaDada As String, _
                                    m_FechaInicial As String, _
                                    m_FechaFinal As String, _
                                    Optional ByRef p_Error As String _
                                    ) As EnumSino
    On Error GoTo errores
    If Not IsDate(m_FechaDada) Or Not IsDate(m_FechaInicial) Or Not IsDate(m_FechaFinal) Then
        Exit Function
    End If
    If CDate(m_FechaInicial) > CDate(m_FechaFinal) Then
        p_Error = "La fecha inicial no puede ser posterior a la final"
        Err.Raise 1000
    End If
    If CDate(m_FechaDada) < CDate(m_FechaInicial) Then
        FechaEnIntervalo = EnumSino.No
        Exit Function
    End If
    If CDate(m_FechaDada) > CDate(m_FechaFinal) Then
        FechaEnIntervalo = EnumSino.No
        Exit Function
    End If
    FechaEnIntervalo = EnumSino.Sí
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método FechaEnIntervalo ha devuelto el error: " & Err.Description
    End If
    Debug.Print p_Error
End Function

Public Function EjecutarShell( _
                                strComando As String, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    
    Dim ManejadorProceso As Long
    Dim IDProceso As Long
    Dim lpExitCode As Long
    On Error GoTo errores
    
    If strComando = "" Then
        p_Error = "No se ha indicado el comando"
        Err.Raise 1000
    End If
    IDProceso = Shell(strComando, vbHide)
    ManejadorProceso = OpenProcess(PROCESS_QUERY_INFORMATION, False, IDProceso)
    ' Mientras lp_ExitCode = STATUS_PENDING, se ejecuta el do
    Do
        Call GetExitCodeProcess(ManejadorProceso, lpExitCode)
        DoEvents
    Loop While lpExitCode = STATUS_PENDING
    Call CloseHandle(ManejadorProceso)

    
    EjecutarShell = "OK"
    Exit Function
errores:
    
    If Err.Number <> 1000 Then
        p_Error = "El método EjecutarShell ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function
Public Function EnOficina(Optional ByRef p_Error As String) As EnumSino
    
    Dim m_EnOficinaPorArchivo As EnumSino
    On Error GoTo errores
    
    
    m_EnOficinaPorArchivo = EnOficinaPorArchivo
    If m_EnOficinaPorArchivo <> Empty Then
        EnOficina = m_EnOficinaPorArchivo
        Exit Function
    End If
    EnOficina = EnOficinaPorComando
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EnOficina ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function
Private Function EnOficinaPorComando(Optional ByRef p_Error As String) As EnumSino
    
    Dim fichero As Scripting.TextStream
    Dim m_Linea As String
    Dim m_URLTemp As String
    Dim m_TextoEsperado As String
    Dim m_Comando As String
    Dim m_IP As String
    
    On Error GoTo errores
    
   
    m_URLTemp = Environ("APPDATA") & "\" & "Aplicaciones DYSN\" & fso.GetTempName() & ".txt"
    m_IP = "10.14.7.44"
    m_TextoEsperado = m_IP & ": bytes=32"
    m_Comando = "cmd /c ping " & m_IP & " >" & Chr(34) & m_URLTemp & Chr(34)
    EjecutarShell m_Comando, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set fichero = fso.OpenTextFile(m_URLTemp)
    Do Until fichero.AtEndOfStream
        m_Linea = fichero.ReadLine
        If InStr(1, m_Linea, m_TextoEsperado) <> 0 Then
            EnOficinaPorComando = EnumSino.Sí
             fichero.Close
            Set fichero = Nothing
            fso.DeleteFile m_URLTemp, True
            Exit Function
        End If
    Loop
    fichero.Close
    Set fichero = Nothing
    fso.DeleteFile m_URLTemp, True
    EnOficinaPorComando = EnumSino.No
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EnOficinaPorComando ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function
Private Function EnOficinaPorArchivo(Optional ByRef p_Error As String) As EnumSino
    
    Dim m_Texto As String
    Dim m_EnOficinaPorArchivo As String
    
    
    On Error GoTo errores
    
    m_Texto = DAMESSIDPorArchivo(p_Error)
    If p_Error <> "" Then
        Exit Function
    End If
    If InStr(1, m_Texto, "|") = 0 Then
        Exit Function
    End If
    dato = Split(m_Texto, "|")
    If UBound(dato) = 2 Then
        m_EnOficinaPorArchivo = dato(2)
    End If
    If m_EnOficinaPorArchivo = "En Oficina" Or m_EnOficinaPorArchivo = "Fuera de Oficina" Then
        If m_EnOficinaPorArchivo = "En Oficina" Then
            EnOficinaPorArchivo = EnumSino.Sí
        Else
            EnOficinaPorArchivo = EnumSino.No
        End If
        Exit Function
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método EnOficinaPorArchivo ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function

Public Function DAMESSIDPorArchivo(Optional ByRef p_Error As String) As String
    
    
    Dim fichero As Scripting.TextStream
    
    Dim m_ArchivoSSID As String
    
    
    On Error GoTo errores
    If m_ObjEntorno Is Nothing Then
        Exit Function
    End If
    m_ArchivoSSID = m_ObjEntorno.URLArchivoSSID
    p_Error = m_ObjEntorno.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not fso.FileExists(m_ArchivoSSID) Then
        Exit Function
    End If
    Set fichero = fso.OpenTextFile(m_ArchivoSSID)
    DAMESSIDPorArchivo = Trim(fichero.ReadLine)
    fichero.Close
    

    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método DAMESSIDPorArchivo ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
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

Public Function RegistroApertura( _
                                Optional ByRef p_Error As String _
                                ) As String

    Dim m_SQL As String
    Dim m_ID As String
    Dim rcdDatos As DAO.Recordset
    Dim m_TextoEnOficina As String
    Dim m_NombreUsuario As String
   
    Dim m_NombreAplicacion As String
    Dim m_VersionAplicacion As String
    On Error GoTo errores
    
    
    If Not m_ObjUsuarioConectado Is Nothing Then
        m_NombreUsuario = m_ObjUsuarioConectado.Nombre
    Else
        m_NombreUsuario = "Desconocido"
    End If
    
    If m_EnOficina = Empty Then
        m_TextoEnOficina = "NA"
    Else
        If m_EnOficina = EnumSino.Sí Then
            m_TextoEnOficina = "Sí"
        Else
            m_TextoEnOficina = "No"
        End If
    End If
    m_VersionAplicacion = m_ObjEntorno.VersionAplicacion
    If m_VersionAplicacion = "" Then
        m_VersionAplicacion = "Desconocida"
    End If
    If Application.TempVars("EnPruebas") = "Sí" Then
        m_NombreAplicacion = "BRASS PRUEBAS"
    Else
        m_NombreAplicacion = "BRASS"
    End If
    m_ID = DameID1("TbAplicacionesAperturas", "IDApertura", getdbLanzadera(), p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_SQL = "SELECT * FROM TbAplicacionesAperturas;"
    Set rcdDatos = getdbLanzadera().OpenRecordset(m_SQL)
    With rcdDatos
        .AddNew
            .Fields("IDApertura") = m_ID
            .Fields("IDAplicacion") = IDAplicacion
            .Fields("NombreUsuario") = m_NombreUsuario
            .Fields("FechaApertura") = Date
            .Fields("HoraApertura") = getHora()
            .Fields("NombreAplicacion") = m_NombreAplicacion
            .Fields("EnOficina") = m_TextoEnOficina
            .Fields("UsuarioMaquina") = getUsuariodeMaquina()
            .Fields("NombreMaquina") = getMaquina()
            .Fields("VersionAplicacion") = m_VersionAplicacion

        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RegistroApertura ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Debug.Print p_Error
End Function

Public Function RegistroCierre( _
                                Optional ByRef p_Error As String _
                                ) As String

    Dim m_SQL As String
    Dim m_NombreUsuario As String
    On Error GoTo errores
    
    If Not m_ObjUsuarioConectado Is Nothing Then
        m_NombreUsuario = m_ObjUsuarioConectado.Nombre
    Else
        m_NombreUsuario = "Desconocido"
    End If
    
    m_SQL = "UPDATE TbAplicacionesAperturas SET FechaCierre = Date(), HoraCierre = #" & getHora() & "# " & _
            "WHERE NombreUsuario='" & m_NombreUsuario & "' " & _
            "AND HoraCierre Is Null " & _
            "AND IDAplicacion=" & IDAplicacion & ";"
    getdbLanzadera().Execute m_SQL
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RegistroCierre ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Debug.Print p_Error
End Function
Public Function getHora() As String
    Dim xmlHttp As Object
    Dim url As String
    Dim response As String
    Dim json As Object
    Dim hora As String
    Dim dato As Variant
    Dim flag As String
    On Error GoTo errores
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
    Exit Function
errores:
    getHora = Format(Now(), "hh:mm:ss")
End Function


Public Function getCalibracionPorFecha( _
                                        p_IDEquipoMedida As String, _
                                        Optional p_IDEvento As String, _
                                        Optional p_FechaAltaEvento As String, _
                                        Optional ByRef p_Error As String _
                                        ) As EquipoMedidaCalibracion
    
    Dim m_Evento As Evento
    Dim m_ID As Variant
    Dim m_Equipo As EquipoMedida
    Dim m_ColCalibraciones As Scripting.Dictionary
    
    Dim m_Calibracion As EquipoMedidaCalibracion
    Dim m_FechaEnIntervalo As EnumSino
    
    On Error GoTo errores
    If p_IDEquipoMedida = "" Then
        Exit Function
    End If
    If p_IDEvento = "" And Not IsDate(p_FechaAltaEvento) Then
        Exit Function
    End If
    If Not IsDate(p_FechaAltaEvento) Then
        Set m_Evento = Constructor.getEvento(p_IDEvento, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If m_Evento Is Nothing Then
            Exit Function
        End If
        p_FechaAltaEvento = m_Evento.FECHAALTAEVENTO
        
    End If
    If Not IsDate(p_FechaAltaEvento) Then
        Exit Function
    End If
    Set m_Equipo = Constructor.getEquipoMedida(p_IDEquipoMedida, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_Equipo Is Nothing Then
        Exit Function
    End If
    Set m_ColCalibraciones = m_Equipo.Calibraciones
    p_Error = m_Equipo.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ColCalibraciones Is Nothing Then
        Exit Function
    End If
    For Each m_ID In m_ColCalibraciones
        Set m_Calibracion = m_ColCalibraciones(m_ID)
        m_FechaEnIntervalo = FechaEnIntervalo(p_FechaAltaEvento, m_Calibracion.FechaCalibracion, m_Calibracion.FechaFinCalibracion, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If m_FechaEnIntervalo = EnumSino.Sí Then
            Set getCalibracionPorFecha = m_Calibracion
            Exit Function
        End If
        Set m_Calibracion = Nothing
    Next
    Exit Function
errores:
    If Err.p_Error <> 1000 Then
        p_Error = "El método EventoEquipoMedida.getCalibracionPorFecha ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function

Private Function getDirectorioOneDrive(Optional ByRef p_Error As String) As String
    Dim fso As Object
    Dim carpetaRaiz As Object
    Dim subCarpeta As Object
    Dim rutaEncontrada As String
    Dim encontrado As Boolean
    
    On Error GoTo errores
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
            rutaEncontrada = subCarpeta.Path
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
