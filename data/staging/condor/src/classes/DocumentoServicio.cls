VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DocumentoServicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' ==========================================================================
' EN: DocumentoServicio.cls
' ACCIÓN: Añadir método específico para PDF de CDCASUB para evitar Type Mismatch
' ==========================================================================

' --- CONSTANTES DE EXPORTACIÓN A WORD ---
' Límite inmutable de caracteres por marcador en plantillas Word
' GAP-P3-02: Word trunca a 255 caracteres por campo, no 256
Private Const MAX_CHARS_POR_MARCADOR As Long = 255



Public Function GenerarDocumentoPDF_CDCASUB(ByRef objData As DatosCDCASUB, ByRef objSolicitud As Solicitud) As String
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim rutaWordTemporal As String
    Dim rutaPDFDestino As String
    Const wdFormatPDF As Long = 17

    On Error GoTo Errores

    ' 1. Generar el documento Word poblado (Paso intermedio)
    ' Usamos el generador específico de CDCASUB que ya creamos
    rutaWordTemporal = Me.GenerarDocumentoCDCASUB(objData, objSolicitud)

    If Not fso.FileExists(rutaWordTemporal) Then
        Err.Raise 514, "GenerarDocumentoPDF_CDCASUB", "Falló la generación del archivo Word intermedio."
    End If

    ' 2. Instanciar Word y abrir el documento
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    Set wordDoc = wordApp.Documents.Open(rutaWordTemporal)

    ' 3. Calcular ruta destino y Exportar a PDF
    rutaPDFDestino = Replace(rutaWordTemporal, ".docx", ".pdf", , , vbTextCompare)
    wordDoc.ExportAsFixedFormat OutputFileName:=rutaPDFDestino, ExportFormat:=wdFormatPDF

    ' 4. Limpieza (Camino Exitoso)
    wordDoc.Close SaveChanges:=False
    wordApp.Quit

    ' Borrar el temporal DOCX
    On Error Resume Next
    fso.DeleteFile rutaWordTemporal, True
    On Error GoTo Errores

    ' 5. Devolver ruta del PDF
    GenerarDocumentoPDF_CDCASUB = rutaPDFDestino

LimpiarYSalir:
    Set wordDoc = Nothing
    Set wordApp = Nothing
    Exit Function

Errores:
    Dim savedErr As Long, savedDesc As String
    savedErr = Err.Number: savedDesc = Err.description

    On Error Resume Next
    If Not wordDoc Is Nothing Then wordDoc.Close SaveChanges:=False
    If Not wordApp Is Nothing Then wordApp.Quit
    If fso.FileExists(rutaWordTemporal) Then fso.DeleteFile rutaWordTemporal, True
    If fso.FileExists(rutaPDFDestino) Then fso.DeleteFile rutaPDFDestino, True

    Dim errObj As New CondorError
    errObj.Create savedErr, savedDesc, "DocumentoServicio.GenerarDocumentoPDF_CDCASUB"
    errObj.Raise
End Function

Public Function GenerarDocumentoPDF_PCSUB(ByRef objPCSUB As Object, ByRef objSolicitud As Solicitud) As String
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim rutaWordTemporal As String
    Dim rutaPDFDestino As String
    Const wdFormatPDF As Long = 17

    On Error GoTo Errores

    ' 1. Generar el documento Word poblado (Paso intermedio)
    rutaWordTemporal = Me.GenerarDocumentoPCSUB(objPCSUB, objSolicitud)

    If Not fso.FileExists(rutaWordTemporal) Then
        Err.Raise 514, "GenerarDocumentoPDF_PCSUB", "Falló la generación del archivo Word intermedio."
    End If

    ' 2. Instanciar Word y abrir el documento
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    Set wordDoc = wordApp.Documents.Open(rutaWordTemporal)

    ' 3. Calcular ruta destino y Exportar a PDF
    rutaPDFDestino = Replace(rutaWordTemporal, ".docx", ".pdf", , , vbTextCompare)
    wordDoc.ExportAsFixedFormat OutputFileName:=rutaPDFDestino, ExportFormat:=wdFormatPDF

    ' 4. Limpieza (Camino Exitoso)
    wordDoc.Close SaveChanges:=False
    wordApp.Quit

    ' Borrar el temporal DOCX
    On Error Resume Next
    fso.DeleteFile rutaWordTemporal, True
    On Error GoTo Errores

    ' 5. Devolver ruta del PDF
    GenerarDocumentoPDF_PCSUB = rutaPDFDestino

LimpiarYSalir:
    Set wordDoc = Nothing
    Set wordApp = Nothing
    Exit Function

Errores:
    Dim savedErr As Long, savedDesc As String
    savedErr = Err.Number: savedDesc = Err.description

    On Error Resume Next
    If Not wordDoc Is Nothing Then wordDoc.Close SaveChanges:=False
    If Not wordApp Is Nothing Then wordApp.Quit
    If fso.FileExists(rutaWordTemporal) Then fso.DeleteFile rutaWordTemporal, True
    If fso.FileExists(rutaPDFDestino) Then fso.DeleteFile rutaPDFDestino, True

    Dim errObj As New CondorError
    errObj.Create savedErr, savedDesc, "DocumentoServicio.GenerarDocumentoPDF_PCSUB"
    errObj.Raise
End Function

Public Function GenerarDocumentoPDFparaSolicitud(ByVal idSolicitud As Long) As String
    Dim solServ As New SolicitudServicio
    Dim sol As Solicitud
    
    On Error GoTo Errores
    
    Set sol = solServ.getSolicitudPorID(idSolicitud)
    If sol Is Nothing Then Err.Raise 513, "GenerarPDF", "No se encontró la solicitud."
    
    Select Case sol.tipoSolicitud
        Case "PC"
            Dim pcServ As New DatosPCServicio, pcData As DatosPC
            Set pcData = pcServ.getDatosPC(idSolicitud)
            If pcData Is Nothing Then Err.Raise 513, "GenerarPDF", "Faltan los datos técnicos (PC)."
            
            ' Validaciones de Workflow (Firmas y RAC)
            If Not pcServ.EsAprobacionSuministradorCompleta(idSolicitud) Then Err.Raise 513, "GenerarPDF", "Falta el Cuadro de Firmas."
            If Not pcServ.EsDictamenRACCompleto(idSolicitud) Then Err.Raise 513, "GenerarPDF", "Faltan los datos del RAC."
            
            GenerarDocumentoPDFparaSolicitud = Me.GenerarDocumentoPDF_PC(pcData, sol)
            
        Case "CD_CA"
            Dim cdcaServ As New DatosCDCAServicio, cdcaData As DatosCDCA
            Set cdcaData = cdcaServ.GetDatosCDCA(idSolicitud)
            If cdcaData Is Nothing Then Err.Raise 513, "GenerarPDF", "Faltan los datos técnicos (CD/CA)."
            
            If Not cdcaServ.EsAprobacionSuministradorCompleta(idSolicitud) Then Err.Raise 513, "GenerarPDF", "Falta el Cuadro de Firmas."
            If Not cdcaServ.EsDictamenRACCompleto(idSolicitud) Then Err.Raise 513, "GenerarPDF", "Faltan los datos del RAC."
            
            GenerarDocumentoPDFparaSolicitud = Me.GenerarDocumentoPDF_CDCA(cdcaData, sol)

        Case "CD_CA_SUB"
            ' --- NUEVO CASO IMPLEMENTADO ---
            Dim cdcaSubServ As New DatosCDCASUBServicio, subData As DatosCDCASUB
            Set subData = cdcaSubServ.getDatosCDCASUB(idSolicitud)
            If subData Is Nothing Then Err.Raise 513, "GenerarPDF", "Faltan los datos técnicos (CD/CA SUB)."
            
            If Not cdcaSubServ.EsAprobacionSuministradorCompleta(idSolicitud) Then Err.Raise 513, "GenerarPDF", "Falta el Cuadro de Firmas."
            If Not cdcaSubServ.EsDictamenRACCompleto(idSolicitud) Then Err.Raise 513, "GenerarPDF", "Faltan los datos del RAC."
            
            GenerarDocumentoPDFparaSolicitud = Me.GenerarDocumentoPDF_CDCASUB(subData, sol)
        
        Case "PC_SUB"
            ' --- PC_SUB ---
            Dim pcsubServPDF As New DatosPCSUBServicio, pcsubDataPDF As Object
            Set pcsubDataPDF = pcsubServPDF.getDatosPCSUB(idSolicitud)
            If pcsubDataPDF Is Nothing Then Err.Raise 513, "GenerarPDF", "Faltan los datos técnicos (PC SUB)."
            
            If Not pcsubServPDF.EsAprobacionSuministradorCompleta(idSolicitud) Then Err.Raise 513, "GenerarPDF", "Falta el Cuadro de Firmas."
            If Not pcsubServPDF.EsDictamenRACCompleto(idSolicitud) Then Err.Raise 513, "GenerarPDF", "Faltan los datos del RAC."
            
            GenerarDocumentoPDFparaSolicitud = Me.GenerarDocumentoPDF_PCSUB(pcsubDataPDF, sol)

        Case Else
            Err.Raise 513, "GenerarPDF", "Tipo de solicitud no reconocido."
    End Select
    
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DocumentoServicio.GenerarDocumentoPDFparaSolicitud"
    errObj.Raise
End Function
' ==========================================================================
' MÓDULO: DocumentoServicio.cls
' ACCIÓN: Reemplazar ImportarDatosRACDesdeWord para que lea campos extendidos.
' ==========================================================================
Public Function ImportarDatosRACDesdeWord(ByVal rutaDocumento As String) As DatosPC
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim mapeoServ As New MapeoServicio
    Dim colMapeos As Scripting.Dictionary
    Dim unMapeo As MapeoCampos
    Dim key As Variant
    Dim objPC As New DatosPC
    
    On Error GoTo Errores
    
    If Not fso.FileExists(rutaDocumento) Then Err.Raise 513, "ImportarWord", "El archivo especificado no existe: " & rutaDocumento
    
    Set colMapeos = mapeoServ.getMapeoPC()
    
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    Set wordDoc = wordApp.Documents.Open(rutaDocumento, ReadOnly:=True)
    
    For Each key In colMapeos.Keys
        Set unMapeo = colMapeos(key)
        
        ' REGLA: Solo nos interesan los campos de la sección del RAC
        If unMapeo.nombreCampoTabla = "racCodigo" Or _
           unMapeo.nombreCampoTabla = "observacionesRAC" Or _
           unMapeo.nombreCampoTabla = "racNombre" Then
           
            Dim valorFinal As Variant
            
            ' --- INICIO DE LA NUEVA LÓGICA PARA LEER EXTENSIONES ---
            If unMapeo.numExtensiones > 0 Then
                ' Este es un campo largo, debemos concatenar sus partes.
                Dim textoConcatenado As String
                Dim i As Integer
                
                ' 1. Leer el campo base
                textoConcatenado = Nz(GetFormFieldValue(wordDoc, unMapeo.nombreCampoWord), "")
                
                ' 2. Leer y añadir las extensiones
                For i = 1 To unMapeo.numExtensiones
                    Dim valorExtension As String
                    valorExtension = Nz(GetFormFieldValue(wordDoc, unMapeo.nombreCampoWord & "_ext" & i), "")
                    
                    ' Si la extensión tiene contenido, la añadimos.
                    ' Evitamos añadir espacios en blanco si el campo está vacío.
                    If Len(valorExtension) > 0 Then
                        textoConcatenado = textoConcatenado & valorExtension
                    End If
                Next i
                
                valorFinal = textoConcatenado
            ElseIf InStr(1, unMapeo.nombreCampoWord, "_ext") > 0 Then
                ' Ignoramos explícitamente los campos de extensión, ya que son procesados por su campo base.
                GoTo SiguienteMapeo
            Else
                ' Este es un campo normal, sin extensiones.
                valorFinal = GetFormFieldValue(wordDoc, unMapeo.nombreCampoWord)
            End If
            ' --- FIN DE LA NUEVA LÓGICA ---
            
            ' Asignamos el valor final (ya sea simple o concatenado) al objeto DatosPC
            If IsDate(valorFinal) Then
                objPC.SetPropiedad unMapeo.nombreCampoTabla, CDate(valorFinal)
            ElseIf unMapeo.nombreCampoTabla = "observacionesRAC" Then
                objPC.SetPropiedad unMapeo.nombreCampoTabla, Nz(valorFinal, "")
            ElseIf Len(Trim$(Nz(valorFinal, ""))) > 0 Then
                objPC.SetPropiedad unMapeo.nombreCampoTabla, Nz(valorFinal, "")
            End If
        End If
SiguienteMapeo:
    Next key

    ' ========================================================================
    ' GAP-P4-02: Importar comentarios Word al campo observacionesRAC
    ' ========================================================================
    Dim com As Object
    Dim comentarios As String
    Dim comentariosExistentes As String

    comentariosExistentes = Nz(objPC.observacionesRAC, "")

    For Each com In wordDoc.Comments
        Dim textoComentario As String
        textoComentario = Trim$(com.Range.text)
        If Len(textoComentario) > 0 Then
            comentarios = comentarios & com.Author & ": " & textoComentario & vbCrLf
        End If
    Next com

    If Len(comentarios) > 0 Then
        If Len(comentariosExistentes) > 0 Then
            objPC.SetPropiedad "observacionesRAC", comentariosExistentes & vbCrLf & comentarios
        Else
            objPC.SetPropiedad "observacionesRAC", comentarios
        End If
    End If

    Set ImportarDatosRACDesdeWord = objPC

LimpiarYSalir:
    On Error Resume Next
    If Not wordDoc Is Nothing Then wordDoc.Close SaveChanges:=False
    If Not wordApp Is Nothing Then wordApp.Quit
    Set wordApp = Nothing
    Set wordDoc = Nothing
    Exit Function
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DocumentoServicio.ImportarDatosRACDesdeWord"
    errObj.Raise
End Function

Private Function GetFormFieldValue(ByRef doc As Word.Document, ByVal fieldName As String) As Variant
    ' RESPONSABILIDAD: Leer de forma segura el valor de un FormField.
    Dim ff As FormField
    On Error Resume Next ' Si el campo no existe, devolverá Empty
    
    Set ff = doc.FormFields(fieldName)
    
    If Not ff Is Nothing Then
        Select Case ff.Type
            Case wdFieldFormCheckBox
                GetFormFieldValue = ff.CheckBox.value
            Case Else ' wdFieldFormTextInput, etc.
                GetFormFieldValue = ff.result
        End Select
    Else
        GetFormFieldValue = Empty
    End If
End Function
' ==========================================================================
' MÓDULO: DocumentoServicio.cls
' GAP-P3-01: Refactorizado para ser gemelo arquitectónico de CDCA/CDCASUB
' ==========================================================================
Public Function GenerarDocumentoPC(ByRef objPC As DatosPC, ByRef objSolicitud As Solicitud, Optional ByVal numVersion As Long = 0) As String
    GenerarDocumentoPC = GenerarDocumentoGenerico(objPC, objSolicitud, "PC", m_ObjEntorno.URLPlantillaPC)
End Function

Public Function GenerarDocumentoCDCA(ByRef objCDCA As DatosCDCA, ByRef objSolicitud As Solicitud, Optional ByVal numVersion As Long = 0) As String
    GenerarDocumentoCDCA = GenerarDocumentoGenerico(objCDCA, objSolicitud, "CD_CA", m_ObjEntorno.URLPlantillaCDCA)
End Function

Public Function GenerarDocumentoCDCASUB(ByRef objCDCASUB As DatosCDCASUB, ByRef objSolicitud As Solicitud, Optional ByVal numVersion As Long = 0) As String
    GenerarDocumentoCDCASUB = GenerarDocumentoGenerico(objCDCASUB, objSolicitud, "CD_CA_SUB", m_ObjEntorno.URLPlantillaCDCASUB)
End Function

Public Function GenerarDocumentoPCSUB(ByRef objPCSUB As Object, ByRef objSolicitud As Solicitud, Optional ByVal numVersion As Long = 0) As String
    GenerarDocumentoPCSUB = GenerarDocumentoGenerico(objPCSUB, objSolicitud, "PC_SUB", m_ObjEntorno.URLPlantillaPCSUB)
End Function

' ==========================================================================
' MÓDULO: DocumentoServicio.cls
' ACCIÓN: REEMPLAZO de GenerarDocumentoGenerico para que sea verdaderamente polimórfico.
' ==========================================================================
Private Function GenerarDocumentoGenerico(ByRef objDatosTecnicos As Object, ByRef objSolicitud As Solicitud, _
    ByVal nombrePlantilla As String, ByVal rutaPlantilla As String) As String
    
    Dim wordApp As Object, wordDoc As Object
    Dim mapeoServ As New MapeoServicio, colMapeos As Scripting.Dictionary
    Dim unMapeo As MapeoCampos, key As Variant
    Dim rutaDestino As String, valorCampo As Variant
    Dim versionBorrador As String
    Dim wfServ As New WorkflowServicio
    Const MAX_CHUNK_SIZE As Long = 255
    On Error GoTo Errores
    
    If Not fso.FileExists(rutaPlantilla) Then Err.Raise 513, , "No se encuentra la plantilla: " & rutaPlantilla
    
    versionBorrador = wfServ.GetProximaVersionBorrador(objSolicitud.idSolicitud)
    
    If nombrePlantilla = "CD_CA_SUB" Then
        rutaDestino = m_ObjEntorno.URLDirectorioLocal & "SUB_" & Replace(objSolicitud.codigoSolicitud, "/", "_") & ".v" & Mid(versionBorrador, 2) & ".docx"
    ElseIf nombrePlantilla = "CD_CA" Then
        rutaDestino = m_ObjEntorno.URLDirectorioLocal & "CD_CA_" & Replace(objSolicitud.codigoSolicitud, "/", "_") & ".v" & Mid(versionBorrador, 2) & ".docx"
    ElseIf nombrePlantilla = "PC_SUB" Then
        rutaDestino = m_ObjEntorno.URLDirectorioLocal & "PC_SUB_" & Replace(objSolicitud.codigoSolicitud, "/", "_") & ".v" & Mid(versionBorrador, 2) & ".docx"
    Else
        rutaDestino = m_ObjEntorno.URLDirectorioLocal & nombrePlantilla & "_" & Replace(objSolicitud.codigoSolicitud, "/", "_") & " " & versionBorrador & ".docx"
    End If
    fso.CopyFile rutaPlantilla, rutaDestino
    
    ' Obtenemos el mapeo correspondiente a la plantilla
    Select Case nombrePlantilla
        Case "PC": Set colMapeos = mapeoServ.getMapeoPC()
        Case "CD_CA": Set colMapeos = mapeoServ.getMapeoCDCA()
        Case "CD_CA_SUB": Set colMapeos = mapeoServ.getMapeoCDCASUB()
        Case "PC_SUB": Set colMapeos = mapeoServ.getMapeoPCSUB()
        Case Else: Err.Raise 513, , "Nombre de plantilla desconocido: " & nombrePlantilla
    End Select
    
    If colMapeos Is Nothing Then Err.Raise 513, , "No se pudo obtener el mapeo de campos para " & nombrePlantilla
    If colMapeos.count = 0 Then Err.Raise 513, , "No se pudo obtener el mapeo de campos para " & nombrePlantilla
    
    ' === PRE-CACHEAR EL BLOQUE CONTINUATION (optimización: buscar solo 1 vez, no en cada overflow) ===
    Dim contBaseName As String
    Dim mapeoCont As MapeoCampos
    contBaseName = ""
    Set mapeoCont = Nothing
    Dim keyCont As Variant
    For Each keyCont In colMapeos.Keys
        Dim mapeoTemp As MapeoCampos
        Set mapeoTemp = colMapeos(keyCont)
        If Right(mapeoTemp.nombreCampoWord, 4) = "Cont" And mapeoTemp.numExtensiones > 0 Then
            contBaseName = mapeoTemp.nombreCampoWord
            Set mapeoCont = mapeoTemp
            Exit For
        End If
    Next keyCont
    
    ' === GUARDAR PIDs DE WINWORD ANTES ===
    Dim pidsAntes As String
    pidsAntes = GetWordPidsString()
    
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    Set wordDoc = wordApp.Documents.Open(rutaDestino)
    
    ' Bucle de relleno (ahora funciona con cualquier objeto que implemente .getPropiedad)
    For Each key In colMapeos.Keys
        Set unMapeo = colMapeos(key)
        'Call LogDebug unMapeo.nombreCampoWord, unMapeo.nombreCampoTabla
       ' If unMapeo.nombreCampoWord = "Parte0_02" Then Stop
        Dim valorPropiedad As Variant
        'If unMapeo.nombreCampoTabla = "firmaAprobacionRespDisenioNombre" Then Stop
        ' --- INICIO DE LA MEJORA: Depuración de Campos Requeridos ---
        ' Obtenemos el valor de forma polimórfica
        On Error Resume Next
        valorPropiedad = objDatosTecnicos.getPropiedad(unMapeo.nombreCampoTabla)
        
        If Err.Number <> 0 Then
            Dim campoFallo As String
            campoFallo = unMapeo.nombreCampoTabla
            On Error GoTo Errores
            Err.Raise 513, "GenerarDoc", "Error al obtener la propiedad '" & campoFallo & "' del objeto técnico. Verifique que el campo existe en la clase y en la tabla."
        End If
        On Error GoTo Errores
        ' --- FIN DE LA MEJORA ---
        
        ' === GAP-P3-01: Solo logueamos campos con overflow o _Cont ===
        
        ' === GAP-P3-01 FIX: *_Cont se verifica PRIMERO para evitar doble procesamiento ===
        ' El bloque Cont es lleno EXCLUSIVAMENTE por overflow desde el campo base.
        ' Nunca se procesa aquí porque no tiene datos propios en la BD.
        If Right(unMapeo.nombreCampoWord, 4) = "Cont" Then
            ' Skip - filled by overflow from base field only
        ElseIf unMapeo.numExtensiones > 0 Then
            ' === FASE 1: Campo base + extensiones del mapping ===
            
            Dim textoLargo As String, chunk As String, i As Integer
            textoLargo = Nz(valorPropiedad, "")
            
            ' Campo base
            chunk = Left(textoLargo, MAX_CHARS_POR_MARCADOR)
            Call SetFormFieldValue(wordDoc, unMapeo.nombreCampoWord, chunk, nombrePlantilla, unMapeo.nombreCampoTabla, rutaPlantilla, rutaDestino)
            textoLargo = Mid(textoLargo, MAX_CHARS_POR_MARCADOR + 1)
            
            ' Extensiones del mapping (_ext1, _ext2, ...)
            For i = 1 To unMapeo.numExtensiones
                If Len(textoLargo) > 0 Then
                    chunk = Left(textoLargo, MAX_CHARS_POR_MARCADOR)
                    Call SetFormFieldValue(wordDoc, unMapeo.nombreCampoWord & "_ext" & i, chunk, nombrePlantilla, unMapeo.nombreCampoTabla, rutaPlantilla, rutaDestino)
                    textoLargo = Mid(textoLargo, MAX_CHARS_POR_MARCADOR + 1)
                Else
                    Call SetFormFieldValue(wordDoc, unMapeo.nombreCampoWord & "_ext" & i, "", nombrePlantilla, unMapeo.nombreCampoTabla, rutaPlantilla, rutaDestino)
                End If
            Next i
            
            ' === FASE 2: Bloque de continuación (overflow a *_Cont) ===
            If Len(textoLargo) > 0 Then
                Call ProcesarContinuationBlock(textoLargo, wordDoc, contBaseName, mapeoCont, nombrePlantilla, rutaPlantilla, rutaDestino)
            End If
            
        Else
            ' === FASE 3: Campos normales (sin extensiones) - sin log ===
            Dim textoParaCampo As String
            If VarType(valorPropiedad) = vbBoolean Then
                If valorPropiedad = True Then Call SetFormFieldValue(wordDoc, unMapeo.nombreCampoWord, True, nombrePlantilla, unMapeo.nombreCampoTabla, rutaPlantilla, rutaDestino)
            ElseIf Nz(unMapeo.valorAsociado, "") <> "" Then
                If UCase(Trim(CStr(valorPropiedad))) = UCase(Trim(CStr(unMapeo.valorAsociado))) Then Call SetFormFieldValue(wordDoc, unMapeo.nombreCampoWord, True, nombrePlantilla, unMapeo.nombreCampoTabla, rutaPlantilla, rutaDestino)
            Else
                If IsDate(valorPropiedad) Then
                    If CDate(Nz(valorPropiedad, 0)) > 0 Then textoParaCampo = Format(valorPropiedad, "dd/mm/yyyy") Else textoParaCampo = ""
                Else
                    textoParaCampo = Nz(valorPropiedad, "")
                End If
                
                ' Lógica para añadir saltos de línea
                Select Case unMapeo.nombreCampoTabla
                    Case "suministradorNombreDir", "objetoContrato", "descripcionMaterialAfectado", "numPlanoEspecificacion", "descripcionPropuestaCambio"
                        If textoParaCampo <> "" Then textoParaCampo = textoParaCampo & vbCrLf
                End Select
                
                Call SetFormFieldValue(wordDoc, unMapeo.nombreCampoWord, textoParaCampo, nombrePlantilla, unMapeo.nombreCampoTabla, rutaPlantilla, rutaDestino)
            End If
        End If
    Next key
    
    wordDoc.Save
    
    GenerarDocumentoGenerico = rutaDestino
    
LimpiarYSalir:
    ' --- BLOQUE 3: LIMPIEZA GARANTIZADA ---
    ' Este bloque se ejecuta siempre, tanto en caso de éxito como de error.
    On Error Resume Next ' La limpieza NUNCA debe generar un nuevo error.
    If Not wordDoc Is Nothing Then wordDoc.Close SaveChanges:=False
    If Not wordApp Is Nothing Then wordApp.Quit
    Set wordDoc = Nothing
    Set wordApp = Nothing
    ' Esperar 300ms para que Word libere recursos (antes eran 2000ms - optimizado)
    Sleep 300
    ' Matar cualquier WINWORD huerfano que haya quedado
    Call KillNewWordProcesses(pidsAntes)
    Exit Function

Errores:
    ' 1. CAPTURA
    Dim savedErr As Long, savedDesc As String
    savedErr = Err.Number: savedDesc = Err.description
    
    ' 2. LIMPIEZA
    On Error Resume Next
    If Not wordDoc Is Nothing Then wordDoc.Close SaveChanges:=False
    If Not wordApp Is Nothing Then wordApp.Quit
    If fso.FileExists(rutaDestino) Then fso.DeleteFile rutaDestino, True
    ' Esperar 300ms para que Word libere recursos
    Sleep 300
    ' Matar cualquier WINWORD huerfano que haya quedado
    Call KillNewWordProcesses(pidsAntes)
    
    ' 3. PROPAGACIÓN
    Dim errObj As New CondorError
    errObj.Create savedErr, savedDesc, "DocumentoServicio.GenerarDocumentoGenerico"
    errObj.AddToCallStack "Plantilla: " & nombrePlantilla
    errObj.Raise
End Function


Private Sub SetFormFieldValue(ByRef doc As Object, ByVal fieldName As String, ByVal newValue As Variant, _
    Optional ByVal nombrePlantilla As String = "", Optional ByVal nombreCampoTabla As String = "", _
    Optional ByVal rutaPlantilla As String = "", Optional ByVal rutaDestino As String = "")
    Dim ff As Object
    
    Dim savedErr As Long, savedDesc As String
    
    On Error GoTo Errores
    
    ' Verificar si el documento está en modo de solo lectura
    If doc.ReadOnly Then
        Err.Raise 600, "SetFormFieldValue", "El documento está abierto en modo de solo lectura y no se puede modificar."
    End If
    
    ' Buscamos el campo de formulario por su nombre. Si el mapping apunta a
    ' un FormField inexistente, fallamos explícitamente en vez de saltarlo.
    Call ValidarCampoWordMapeado(doc, fieldName, nombrePlantilla, nombreCampoTabla, rutaPlantilla, rutaDestino)
    Set ff = doc.FormFields(fieldName)
    
    If Not ff Is Nothing Then
        ' La forma de asignar el valor depende del tipo de campo
        Select Case ff.Type
            Case 86 ' wdFieldFormCheckBox
                ' Para checkboxes, asignamos un booleano
                If VarType(newValue) <> vbBoolean Then
                    newValue = False
                End If
                ff.CheckBox.value = CBool(newValue)
            Case 71 ' wdFieldFormDropDown
                ' Para dropdowns, intentamos con DropDown.Value
                On Error Resume Next
                ff.Dropdown.value = newValue
                savedErr = Err.Number
                On Error GoTo Errores
                If savedErr <> 0 Then
                    ' Si falla DropDown.Value, intentamos con result
                    ff.result = CStr(newValue)
                End If
            Case 70 ' wdFieldFormTextInput
                ' Para campos de texto, asignamos el resultado
                ff.result = CStr(newValue)
            Case Else
                ' Para otros tipos de campos de formulario
                On Error Resume Next
                ff.result = CStr(newValue)
                savedErr = Err.Number
                On Error GoTo Errores
        End Select
    End If

LimpiarYSalir:
    Set ff = Nothing
    Exit Sub
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DocumentoServicio.SetFormFieldValue"
    errObj.AddToCallStack "FormField: '" & fieldName & "'"
    errObj.Raise
End Sub
' ==========================================================================
' MÓDULO: DocumentoServicio.cls
' ACCIÓN: AÑADIR esta nueva función para generar el PDF.
' ==========================================================================
Public Function GenerarDocumentoPDF_PC(ByRef objPC As DatosPC, ByRef objSolicitud As Solicitud) As String
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim rutaWordTemporal As String
    Dim rutaPDFDestino As String
    
    
    On Error GoTo ManejadorErrores
    
    ' --- PASO 1: Generar el documento Word poblado ---
    ' Reutilizamos la función que ya tenemos, que sigue el patrón de "Copiar y Poblar".
    rutaWordTemporal = Me.GenerarDocumentoPC(objPC, objSolicitud)
    
    If Not fso.FileExists(rutaWordTemporal) Then
        Err.Raise 513, "GenerarPDF", "Falló la generación del archivo Word intermedio."
    End If
    
    ' --- PASO 2: Convertir el Word a PDF ---
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    
    Set wordDoc = wordApp.Documents.Open(rutaWordTemporal)
    
    ' Definimos la ruta del PDF de salida
    rutaPDFDestino = Replace(rutaWordTemporal, ".docx", ".pdf", , , vbTextCompare)
    
    ' Exportamos a PDF
    wordDoc.ExportAsFixedFormat OutputFileName:=rutaPDFDestino, ExportFormat:=wdFormatPDF
    
    ' --- PASO 3: Limpieza ---
    wordDoc.Close SaveChanges:=False
    wordApp.Quit
    
    On Error Resume Next ' Ignoramos errores al eliminar el temporal
    fso.DeleteFile rutaWordTemporal, True
    On Error GoTo ManejadorErrores
    
    ' --- PASO 4: Devolver la ruta del PDF ---
    GenerarDocumentoPDF_PC = rutaPDFDestino

LimpiarYSalir:
    Set wordDoc = Nothing
    Set wordApp = Nothing
    Exit Function
    
ManejadorErrores:
    ' Aseguramos la limpieza incluso si hay un error
    On Error Resume Next
    If Not wordDoc Is Nothing Then wordDoc.Close SaveChanges:=False
    If Not wordApp Is Nothing Then wordApp.Quit
    ' Intentamos eliminar los archivos temporales que se hayan podido crear
    If fso.FileExists(rutaWordTemporal) Then fso.DeleteFile rutaWordTemporal, True
    If fso.FileExists(rutaPDFDestino) Then fso.DeleteFile rutaPDFDestino, True
    On Error GoTo 0
    
    ' Propagamos el error original usando nuestro sistema
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DocumentoServicio.GenerarDocumentoPDF_PC"
    errObj.Raise
End Function
Public Function GenerarDocumentoParaSolicitud(ByVal idSolicitud As Long, Optional ByVal numVersion As Long = 0, Optional ByRef db As DAO.Database) As String
    Dim solServ As New SolicitudServicio
    Dim sol As Solicitud
    
    On Error GoTo Errores
    
    Set sol = solServ.getSolicitudPorID(idSolicitud, db)
    If sol Is Nothing Then Err.Raise 513, "GenerarDoc", "No se encontró la solicitud."
    
    Select Case sol.tipoSolicitud
        Case "PC"
            ' Lógica para PC (sin cambios)
            Dim pcServ As New DatosPCServicio, pcData As DatosPC
            Set pcData = pcServ.getDatosPC(idSolicitud, db)
            GenerarDocumentoParaSolicitud = Me.GenerarDocumentoPC(pcData, sol, numVersion)
            
        Case "CD_CA"
            ' --- INICIO DE LA CORRECCIÓN ---
            ' Se activa la lógica para CD/CA, eliminando el error deliberado.
            Dim cdcaServ As New DatosCDCAServicio, cdcaData As DatosCDCA
            Set cdcaData = cdcaServ.GetDatosCDCA(idSolicitud, db)
            
            ' Validamos que los datos existan antes de intentar generar el documento.
            If cdcaData Is Nothing Then
                Err.Raise 513, "GenerarDoc", "No se encontraron los datos técnicos (CD/CA) para esta solicitud."
            End If
            
            ' Llamamos a la función helper que ya existe para generar el documento CD/CA.
            GenerarDocumentoParaSolicitud = Me.GenerarDocumentoCDCA(cdcaData, sol, numVersion)
            ' --- FIN DE LA CORRECCIÓN ---
 
        Case "CD_CA_SUB"
            ' --- LÓGICA DESBLOQUEADA ---
            Dim cdcaSubServ As New DatosCDCASUBServicio, subData As DatosCDCASUB
            Set subData = cdcaSubServ.getDatosCDCASUB(idSolicitud, db)
            
            If subData Is Nothing Then
                Err.Raise 513, "GenerarDoc", "No se encontraron los datos técnicos (CD/CA SUB)."
            End If
            
            GenerarDocumentoParaSolicitud = Me.GenerarDocumentoCDCASUB(subData, sol, numVersion)
        
        Case "PC_SUB"
            ' --- PCSUB ---
            Dim pcsubServ As New DatosPCSUBServicio, pcsubData As Object
            Set pcsubData = pcsubServ.getDatosPCSUB(idSolicitud, db)
            
            If pcsubData Is Nothing Then
                Err.Raise 513, "GenerarDoc", "No se encontraron los datos técnicos (PC SUB)."
            End If
            
            GenerarDocumentoParaSolicitud = Me.GenerarDocumentoPCSUB(pcsubData, sol, numVersion)
 
        Case Else
            Err.Raise 513, "GenerarDoc", "Tipo de solicitud no reconocido para generar documento."
    End Select
    
LimpiarYSalir:
    Exit Function
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DocumentoServicio.GenerarDocumentoParaSolicitud"
    errObj.Raise
End Function
' ==========================================================================
' MÓDULO: DocumentoServicio.cls
' ACCIÓN: AÑADIR esta nueva función para la importación de datos RAC de CD/CA.
' ==========================================================================
Public Function ImportarDatosRACDesdeWord_CDCA(ByVal rutaDocumento As String) As DatosCDCA
    Dim wordApp As Object ' Late binding para evitar dependencia de Word
    Dim wordDoc As Object
    Dim mapeoServ As New MapeoServicio
    Dim colMapeos As Scripting.Dictionary
    Dim unMapeo As MapeoCampos
    Dim key As Variant
    Dim objCDCA As New DatosCDCA
    
    On Error GoTo Errores
    
    If Not fso.FileExists(rutaDocumento) Then Err.Raise 513, "ImportarWord", "El archivo especificado no existe: " & rutaDocumento
    
    ' Usamos el mapeo específico para CD/CA
    Set colMapeos = mapeoServ.getMapeoCDCA()
    
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    Set wordDoc = wordApp.Documents.Open(rutaDocumento, ReadOnly:=True)
    
    For Each key In colMapeos.Keys
        Set unMapeo = colMapeos(key)
        
        ' REGLA: Solo nos interesan los campos de la sección del RAC
        If unMapeo.nombreCampoTabla = "racCodigo" Or _
           unMapeo.nombreCampoTabla = "observacionesRAC" Or _
           unMapeo.nombreCampoTabla = "racNombre" Then
           
            Dim valorFinal As Variant
            
            If unMapeo.numExtensiones > 0 Then
                Dim textoConcatenado As String
                Dim i As Integer
                
                textoConcatenado = Nz(GetFormFieldValue(wordDoc, unMapeo.nombreCampoWord), "")
                
                For i = 1 To unMapeo.numExtensiones
                    Dim valorExtension As String
                    valorExtension = Nz(GetFormFieldValue(wordDoc, unMapeo.nombreCampoWord & "_ext" & i), "")
                    
                    If Len(valorExtension) > 0 Then
                        textoConcatenado = textoConcatenado & valorExtension
                    End If
                Next i
                
                valorFinal = textoConcatenado
            ElseIf InStr(1, unMapeo.nombreCampoWord, "_ext") > 0 Then
                GoTo SiguienteMapeoCDCA
            Else
                valorFinal = GetFormFieldValue(wordDoc, unMapeo.nombreCampoWord)
            End If
            
            If IsDate(valorFinal) Then
                objCDCA.SetPropiedad unMapeo.nombreCampoTabla, CDate(valorFinal)
            ElseIf unMapeo.nombreCampoTabla = "observacionesRAC" Then
                objCDCA.SetPropiedad unMapeo.nombreCampoTabla, Nz(valorFinal, "")
            ElseIf Len(Trim$(Nz(valorFinal, ""))) > 0 Then
                objCDCA.SetPropiedad unMapeo.nombreCampoTabla, Nz(valorFinal, "")
            End If
        End If
SiguienteMapeoCDCA:
    Next key

    ' ========================================================================
    ' GAP-P4-02: Importar comentarios Word al campo observacionesRAC
    ' ========================================================================
    Dim comCDCA As Object
    Dim comentariosCDCA As String
    Dim comentariosExistentesCDCA As String

    comentariosExistentesCDCA = Nz(objCDCA.observacionesRAC, "")

    For Each comCDCA In wordDoc.Comments
        Dim textoComentarioCDCA As String
        textoComentarioCDCA = Trim$(comCDCA.Range.text)
        If Len(textoComentarioCDCA) > 0 Then
            comentariosCDCA = comentariosCDCA & comCDCA.Author & ": " & textoComentarioCDCA & vbCrLf
        End If
    Next comCDCA

    If Len(comentariosCDCA) > 0 Then
        If Len(comentariosExistentesCDCA) > 0 Then
            objCDCA.SetPropiedad "observacionesRAC", comentariosExistentesCDCA & vbCrLf & comentariosCDCA
        Else
            objCDCA.SetPropiedad "observacionesRAC", comentariosCDCA
        End If
    End If

    Set ImportarDatosRACDesdeWord_CDCA = objCDCA

LimpiarYSalir:
    On Error Resume Next
    If Not wordDoc Is Nothing Then wordDoc.Close SaveChanges:=False
    If Not wordApp Is Nothing Then wordApp.Quit
    Set wordApp = Nothing: Set wordDoc = Nothing
    Exit Function
    
Errores:
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DocumentoServicio.ImportarDatosRACDesdeWord_CDCA"
    errObj.Raise
End Function

Public Function ImportarDatosRACDesdeWord_PCSUB(ByVal ruta As String) As DatosPCSUB
    Dim wordApp As Object, wordDoc As Object, mapeoServ As New MapeoServicio
    Dim colMapeos As Scripting.Dictionary, unMapeo As MapeoCampos, key As Variant
    Dim obj As New DatosPCSUB
    On Error GoTo Errores

    If Not fso.FileExists(ruta) Then Err.Raise 513, , "Archivo no encontrado"
    Set colMapeos = mapeoServ.getMapeoPCSUB()
    Set wordApp = CreateObject("Word.Application"): wordApp.Visible = False
    Set wordDoc = wordApp.Documents.Open(ruta, ReadOnly:=True)

    For Each key In colMapeos.Keys
        Set unMapeo = colMapeos(key)
        ' Solo campos RAC
        If unMapeo.nombreCampoTabla Like "rac*" Or unMapeo.nombreCampoTabla Like "*RAC*" Then
            Dim valorFinal As Variant

            If unMapeo.numExtensiones > 0 Then
                Dim textoConcatenado As String
                Dim i As Integer

                textoConcatenado = Nz(GetFormFieldValue(wordDoc, unMapeo.nombreCampoWord), "")

                For i = 1 To unMapeo.numExtensiones
                    Dim valorExtension As String
                    valorExtension = Nz(GetFormFieldValue(wordDoc, unMapeo.nombreCampoWord & "_ext" & i), "")

                    If Len(valorExtension) > 0 Then
                        textoConcatenado = textoConcatenado & valorExtension
                    End If
                Next i

                valorFinal = textoConcatenado
            ElseIf InStr(1, unMapeo.nombreCampoWord, "_ext") > 0 Then
                GoTo SiguienteMapeoPCSUB
            Else
                valorFinal = GetFormFieldValue(wordDoc, unMapeo.nombreCampoWord)
            End If

            If IsDate(valorFinal) Then
                obj.SetPropiedad unMapeo.nombreCampoTabla, CDate(valorFinal)
            ElseIf unMapeo.nombreCampoTabla = "observacionesRAC" Then
                obj.SetPropiedad unMapeo.nombreCampoTabla, Nz(valorFinal, "")
            ElseIf Len(Trim$(Nz(valorFinal, ""))) > 0 Then
                obj.SetPropiedad unMapeo.nombreCampoTabla, Nz(valorFinal, "")
            End If
        End If
SiguienteMapeoPCSUB:
    Next key

    ' ========================================================================
    ' GAP-P4-02: Importar comentarios Word al campo observacionesRAC
    ' ========================================================================
    Dim comSUB As Object
    Dim comentariosSUB As String
    Dim comentariosExistentesSUB As String

    comentariosExistentesSUB = Nz(obj.observacionesRAC, "")

    For Each comSUB In wordDoc.Comments
        Dim textoComentarioSUB As String
        textoComentarioSUB = Trim$(comSUB.Range.text)
        If Len(textoComentarioSUB) > 0 Then
            comentariosSUB = comentariosSUB & comSUB.Author & ": " & textoComentarioSUB & vbCrLf
        End If
    Next comSUB

    If Len(comentariosSUB) > 0 Then
        If Len(comentariosExistentesSUB) > 0 Then
            obj.SetPropiedad "observacionesRAC", comentariosExistentesSUB & vbCrLf & comentariosSUB
        Else
            obj.SetPropiedad "observacionesRAC", comentariosSUB
        End If
    End If

    ' ========================================================================
    ' FALLBACK ROBUSTO (sin dependencia de mapeo) para PCSUB
    ' Si el mapeo no trae datos, intentamos leer bookmarks conocidos de plantilla.
    ' ========================================================================
    Dim obsFallback As String
    Dim codFallback As String
    Dim nomFallback As String

    obsFallback = LeerCampoWordConExtensiones(wordDoc, Array("Parte2_21_3", "Parte2_23_3"), 3)
    codFallback = LeerPrimerCampoWordNoVacio(wordDoc, Array("Parte2_21_2", "Parte2_23_2"))
    nomFallback = LeerPrimerCampoWordNoVacio(wordDoc, Array("Parte2_22_1", "Parte2_24_1"))

    If Len(Trim$(Nz(obj.observacionesRAC, ""))) = 0 And Len(Trim$(obsFallback)) > 0 Then
        obj.SetPropiedad "observacionesRAC", obsFallback
    End If
    If Len(Trim$(Nz(obj.racCodigo, ""))) = 0 And Len(Trim$(codFallback)) > 0 Then
        obj.SetPropiedad "racCodigo", codFallback
    End If
    If Len(Trim$(Nz(obj.racNombre, ""))) = 0 And Len(Trim$(nomFallback)) > 0 Then
        obj.SetPropiedad "racNombre", nomFallback
    End If

    wordDoc.Close False: wordApp.Quit
    Set ImportarDatosRACDesdeWord_PCSUB = obj
    Exit Function

Errores:
    On Error Resume Next
    If Not wordDoc Is Nothing Then wordDoc.Close SaveChanges:=False
    If Not wordApp Is Nothing Then wordApp.Quit
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DocumentoServicio.ImportarDatosRACDesdeWord_PCSUB"
    errObj.Raise
End Function

' EN: DocumentoServicio.cls

Public Function GenerarDocumentoPDF_CDCA(ByRef objCDCA As DatosCDCA, ByRef objSolicitud As Solicitud) As String
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim rutaWordTemporal As String
    Dim rutaPDFDestino As String
    Const wdFormatPDF As Long = 17
    
    On Error GoTo Errores
    
    ' 1. Generar el documento Word poblado (Paso intermedio)
    rutaWordTemporal = Me.GenerarDocumentoCDCA(objCDCA, objSolicitud)
    
    If Not fso.FileExists(rutaWordTemporal) Then
        ' Usamos 514 (Error Operativo) en lugar de 513 (Validación)
        Err.Raise 514, "GenerarDocumentoPDF_CDCA", "Falló la generación del archivo Word intermedio."
    End If
    
    ' 2. Instanciar Word y abrir el documento
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    Set wordDoc = wordApp.Documents.Open(rutaWordTemporal)
    
    ' 3. Calcular ruta destino y Exportar a PDF
    rutaPDFDestino = Replace(rutaWordTemporal, ".docx", ".pdf", , , vbTextCompare)
    wordDoc.ExportAsFixedFormat OutputFileName:=rutaPDFDestino, ExportFormat:=wdFormatPDF
    
    ' 4. Limpieza (Camino Exitoso)
    wordDoc.Close SaveChanges:=False
    wordApp.Quit
    
    ' Borrar el temporal DOCX
    On Error Resume Next
    fso.DeleteFile rutaWordTemporal, True
    On Error GoTo Errores
    
    ' 5. Devolver ruta del PDF
    GenerarDocumentoPDF_CDCA = rutaPDFDestino

LimpiarYSalir:
    Set wordDoc = Nothing
    Set wordApp = Nothing
    Exit Function
    
Errores:
    ' --- PATRÓN DE SEGURIDAD ANTI-ERROR 0 ---
    
    ' A. CAPTURA INMEDIATA (Antes de cualquier otra instrucción)
    Dim savedErr As Long, savedDesc As String
    savedErr = Err.Number
    savedDesc = Err.description
    
    ' B. LIMPIEZA SEGURA (On Error Resume Next impide que un fallo aquí borre el error original)
    On Error Resume Next
    If Not wordDoc Is Nothing Then wordDoc.Close SaveChanges:=False
    If Not wordApp Is Nothing Then wordApp.Quit
    
    ' Intentamos limpiar archivos temporales o corruptos
    If fso.FileExists(rutaWordTemporal) Then fso.DeleteFile rutaWordTemporal, True
    ' Si falló la exportación, el PDF podría estar corrupto o a medio escribir, mejor borrarlo
    If fso.FileExists(rutaPDFDestino) Then fso.DeleteFile rutaPDFDestino, True
    
    ' C. PROPAGACIÓN (Usando las variables capturadas)
    Dim errObj As New CondorError
    errObj.Create savedErr, savedDesc, "DocumentoServicio.GenerarDocumentoPDF_CDCA"
    errObj.Raise
End Function


Public Function ImportarDatosRACDesdeWord_CDCASUB(ByVal ruta As String) As DatosCDCASUB
    Dim wordApp As Object, wordDoc As Object, mapeoServ As New MapeoServicio
    Dim colMapeos As Scripting.Dictionary, unMapeo As MapeoCampos, key As Variant
    Dim obj As New DatosCDCASUB
    On Error GoTo Errores
    
    If Not fso.FileExists(ruta) Then Err.Raise 513, , "Archivo no encontrado"
    Set colMapeos = mapeoServ.getMapeoCDCASUB()
    Set wordApp = CreateObject("Word.Application"): wordApp.Visible = False
    Set wordDoc = wordApp.Documents.Open(ruta, ReadOnly:=True)
    
    For Each key In colMapeos.Keys
        Set unMapeo = colMapeos(key)
        ' Solo campos RAC
        If unMapeo.nombreCampoTabla Like "rac*" Or unMapeo.nombreCampoTabla Like "*RAC*" Then
            Dim valorFinal As Variant
            
            If unMapeo.numExtensiones > 0 Then
                Dim textoConcatenado As String
                Dim i As Integer
                
                textoConcatenado = Nz(GetFormFieldValue(wordDoc, unMapeo.nombreCampoWord), "")
                
                For i = 1 To unMapeo.numExtensiones
                    Dim valorExtension As String
                    valorExtension = Nz(GetFormFieldValue(wordDoc, unMapeo.nombreCampoWord & "_ext" & i), "")
                    
                    If Len(valorExtension) > 0 Then
                        textoConcatenado = textoConcatenado & valorExtension
                    End If
                Next i
                
                valorFinal = textoConcatenado
            ElseIf InStr(1, unMapeo.nombreCampoWord, "_ext") > 0 Then
                GoTo SiguienteMapeoCDCASUB
            Else
                valorFinal = GetFormFieldValue(wordDoc, unMapeo.nombreCampoWord)
            End If
            
            If IsDate(valorFinal) Then
                obj.SetPropiedad unMapeo.nombreCampoTabla, CDate(valorFinal)
            ElseIf unMapeo.nombreCampoTabla = "observacionesRAC" Then
                obj.SetPropiedad unMapeo.nombreCampoTabla, Nz(valorFinal, "")
            ElseIf Len(Trim$(Nz(valorFinal, ""))) > 0 Then
                obj.SetPropiedad unMapeo.nombreCampoTabla, Nz(valorFinal, "")
            End If
        End If
SiguienteMapeoCDCASUB:
    Next key

    ' ========================================================================
    ' GAP-P4-02: Importar comentarios Word al campo observacionesRAC
    ' ========================================================================
    Dim comSUB As Object
    Dim comentariosSUB As String
    Dim comentariosExistentesSUB As String

    comentariosExistentesSUB = Nz(obj.observacionesRAC, "")

    For Each comSUB In wordDoc.Comments
        Dim textoComentarioSUB As String
        textoComentarioSUB = Trim$(comSUB.Range.text)
        If Len(textoComentarioSUB) > 0 Then
            comentariosSUB = comentariosSUB & comSUB.Author & ": " & textoComentarioSUB & vbCrLf
        End If
    Next comSUB

    If Len(comentariosSUB) > 0 Then
        If Len(comentariosExistentesSUB) > 0 Then
            obj.SetPropiedad "observacionesRAC", comentariosExistentesSUB & vbCrLf & comentariosSUB
        Else
            obj.SetPropiedad "observacionesRAC", comentariosSUB
        End If
    End If

    ' ========================================================================
    ' FALLBACK ROBUSTO (sin dependencia de mapeo) para CDCASUB
    ' Si el mapeo no trae datos, intentamos leer bookmarks conocidos de plantilla.
    ' ========================================================================
    Dim obsFallback As String
    Dim codFallback As String
    Dim nomFallback As String
    
    obsFallback = LeerCampoWordConExtensiones(wordDoc, Array("Parte2_21_3", "Parte2_23_3"), 3)
    codFallback = LeerPrimerCampoWordNoVacio(wordDoc, Array("Parte2_21_2", "Parte2_23_2"))
    nomFallback = LeerPrimerCampoWordNoVacio(wordDoc, Array("Parte2_22_1", "Parte2_24_1"))
    
    If Len(Trim$(Nz(obj.observacionesRAC, ""))) = 0 And Len(Trim$(obsFallback)) > 0 Then
        obj.SetPropiedad "observacionesRAC", obsFallback
    End If
    If Len(Trim$(Nz(obj.racCodigo, ""))) = 0 And Len(Trim$(codFallback)) > 0 Then
        obj.SetPropiedad "racCodigo", codFallback
    End If
    If Len(Trim$(Nz(obj.racNombre, ""))) = 0 And Len(Trim$(nomFallback)) > 0 Then
        obj.SetPropiedad "racNombre", nomFallback
    End If

    wordDoc.Close False: wordApp.Quit
    Set ImportarDatosRACDesdeWord_CDCASUB = obj
    Exit Function

Errores:
    On Error Resume Next
    If Not wordDoc Is Nothing Then wordDoc.Close SaveChanges:=False
    If Not wordApp Is Nothing Then wordApp.Quit
    Dim errObj As New CondorError
    errObj.Create Err.Number, Err.description, "DocumentoServicio.ImportarDatosRACDesdeWord_CDCASUB"
    errObj.Raise
End Function

Private Function LeerPrimerCampoWordNoVacio(ByVal wordDoc As Object, ByVal nombres As Variant) As String
    Dim i As Long
    Dim val As String
    For i = LBound(nombres) To UBound(nombres)
        val = Trim$(Nz(GetFormFieldValue(wordDoc, CStr(nombres(i))), ""))
        If Len(val) > 0 Then
            LeerPrimerCampoWordNoVacio = val
            Exit Function
        End If
    Next i
    LeerPrimerCampoWordNoVacio = ""
End Function

Private Function LeerCampoWordConExtensiones(ByVal wordDoc As Object, ByVal bases As Variant, ByVal maxExt As Long) As String
    Dim i As Long, j As Long
    Dim base As String
    Dim txt As String, ext As String
    
    For i = LBound(bases) To UBound(bases)
        base = CStr(bases(i))
        txt = Trim$(Nz(GetFormFieldValue(wordDoc, base), ""))
        
        For j = 1 To maxExt
            ext = Nz(GetFormFieldValue(wordDoc, base & "_ext" & j), "")
            If Len(ext) > 0 Then txt = txt & ext
        Next j
        
        If Len(Trim$(txt)) > 0 Then
            LeerCampoWordConExtensiones = txt
            Exit Function
        End If
    Next i
    
    LeerCampoWordConExtensiones = ""
End Function

' =============================================================================
' GAP-P3-01: Helper para overflow de texto a bloque Continuation
' OPTIMIZADO: ahora recibe contBaseName y mapeoCont pre-cachados (1 sola búsqueda)
' =============================================================================
Private Sub ProcesarContinuationBlock(ByRef textoRestante As String, ByVal wordDoc As Object, _
    ByVal contBaseName As String, ByRef mapeoCont As MapeoCampos, _
    ByVal nombrePlantilla As String, ByVal rutaPlantilla As String, ByVal rutaDestino As String)
    
    Dim chunk As String
    Dim contIndex As Long
    
    If contBaseName = "" Or mapeoCont Is Nothing Then
        Exit Sub
    End If
    
    ' Procesar el campo base del bloque (ej: Parte1_20Cont)
    If Len(textoRestante) > 0 Then
        chunk = Left(textoRestante, MAX_CHARS_POR_MARCADOR)
        Call SetFormFieldValue(wordDoc, contBaseName, chunk, nombrePlantilla, mapeoCont.nombreCampoTabla, rutaPlantilla, rutaDestino)
        textoRestante = Mid(textoRestante, MAX_CHARS_POR_MARCADOR + 1)
    Else
        Call SetFormFieldValue(wordDoc, contBaseName, "", nombrePlantilla, mapeoCont.nombreCampoTabla, rutaPlantilla, rutaDestino)
    End If
    
    ' Procesar las extensiones del bloque de continuación (ej: _ext1, _ext2, ...)
    For contIndex = 1 To mapeoCont.numExtensiones
        If Len(textoRestante) > 0 Then
            chunk = Left(textoRestante, MAX_CHARS_POR_MARCADOR)
            Call SetFormFieldValue(wordDoc, contBaseName & "_ext" & contIndex, chunk, nombrePlantilla, mapeoCont.nombreCampoTabla, rutaPlantilla, rutaDestino)
            textoRestante = Mid(textoRestante, MAX_CHARS_POR_MARCADOR + 1)
        Else
            Call SetFormFieldValue(wordDoc, contBaseName & "_ext" & contIndex, "", nombrePlantilla, mapeoCont.nombreCampoTabla, rutaPlantilla, rutaDestino)
        End If
    Next
End Sub

' =============================================================================
' Helper para verificar existencia de campo Word antes de escribir
' =============================================================================
Public Sub ValidarCampoWordMapeado(ByVal wordDoc As Object, ByVal fieldName As String, _
    Optional ByVal nombrePlantilla As String = "", Optional ByVal nombreCampoTabla As String = "", _
    Optional ByVal rutaPlantilla As String = "", Optional ByVal rutaDestino As String = "")

    If Len(Trim$(fieldName)) = 0 Then
        Err.Raise 513, "DocumentoServicio.ValidarCampoWordMapeado", _
            "El mapeo de documento contiene un nombreCampoWord vacío. " & _
            "Plantilla='" & nombrePlantilla & "'; CampoTabla='" & nombreCampoTabla & _
            "'; RutaPlantilla='" & rutaPlantilla & "'; RutaDestino='" & rutaDestino & "'"
    End If

    If wordDoc Is Nothing Then
        Err.Raise 513, "DocumentoServicio.ValidarCampoWordMapeado", _
            "No hay documento Word abierto para validar el campo mapeado. " & _
            "Plantilla='" & nombrePlantilla & "'; CampoTabla='" & nombreCampoTabla & _
            "'; CampoWord='" & fieldName & "'; RutaPlantilla='" & rutaPlantilla & _
            "'; RutaDestino='" & rutaDestino & "'"
    End If

    If Not ExisteCampoWord(wordDoc, fieldName) Then
        Err.Raise 513, "DocumentoServicio.ValidarCampoWordMapeado", _
            "El campo Word mapeado no existe en la plantilla. " & _
            "Plantilla='" & nombrePlantilla & "'; CampoTabla='" & nombreCampoTabla & _
            "'; CampoWord='" & fieldName & "'; RutaPlantilla='" & rutaPlantilla & _
            "'; RutaDestino='" & rutaDestino & "'"
    End If
End Sub

Private Function ExisteCampoWord(ByVal wordDoc As Object, ByVal fieldName As String) As Boolean
    Dim ff As Object
    On Error Resume Next
    Err.Clear
    Set ff = wordDoc.FormFields(fieldName)
    ExisteCampoWord = (Err.Number = 0 And Not ff Is Nothing)
    Set ff = Nothing
    On Error GoTo 0
End Function

' =============================================================================
' Helper para obtener PIDs de WINWORD como string separado por comas
' =============================================================================
Private Function GetWordPidsString() As String
    Dim pids As String
    pids = ""
    Dim oSvc As Object
    Dim oProc As Object
    On Error Resume Next
    Set oSvc = GetObject("winmgmts:\\.\root\cimv2").ExecQuery("SELECT ProcessId FROM Win32_Process WHERE Name = 'WINWORD.EXE'")
    If oSvc Is Nothing Then
        GetWordPidsString = ""
        Exit Function
    End If
    For Each oProc In oSvc
        If pids <> "" Then pids = pids & ","
        pids = pids & oProc.ProcessId
    Next
    GetWordPidsString = pids
    Set oSvc = Nothing
    On Error GoTo 0
End Function

' =============================================================================
' Helper para matar procesos WINWORD que no existían antes (huerfanos)
' =============================================================================
Private Sub KillNewWordProcesses(pidsAntes As String)
    Dim pidsDepois As String
    pidsDepois = GetWordPidsString()
    
    If pidsAntes = "" Then
        Exit Sub
    End If
    
    Dim antesArr() As String
    Dim i As Long
    
    antesArr = Split(pidsAntes, ",")
    
    Dim oSvc As Object
    Dim oProc As Object
    On Error Resume Next
    Set oSvc = GetObject("winmgmts:\\.\root\cimv2").ExecQuery("SELECT ProcessId, Name FROM Win32_Process WHERE Name = 'WINWORD.EXE'")
    
    If oSvc Is Nothing Then Exit Sub
    
    For Each oProc In oSvc
        Dim esNuevo As Boolean
        esNuevo = True
        For i = 0 To UBound(antesArr)
            Dim aux As Long
            aux = CLng(antesArr(i))
            If aux = CLng(oProc.ProcessId) Then
                esNuevo = False
                Exit For
            End If
        Next
        If esNuevo Then
            oProc.Terminate
        End If
    Next
    
    Set oSvc = Nothing
    On Error GoTo 0
End Sub












