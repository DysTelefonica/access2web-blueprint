Attribute VB_Name = "HTML"
Option Compare Database
Option Explicit
Private Function BorraHTMLs( _
                            Optional ByRef p_Error As String) As String
    
    Dim fichero As File
    Dim fso As New Scripting.FileSystemObject
    On Error GoTo errores
    
    For Each fichero In fso.GetFolder(m_ObjEntorno.URLDirectorioLocal).Files
        If fso.GetExtensionName(fichero.Path) = "html" Or fso.GetExtensionName(fichero.Path) = "htm" Then
            If Not FicheroAbierto(fichero.Path) Then
                fso.DeleteFile fichero.Path
            End If
        End If
    Next
    Set fso = Nothing
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método BorraHTMLs ha producido el error nº: " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
End Function
Function DameUntxtYHtml()
    '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día x
    '   -Modificaciones:
  
    '   -Funcionamiento:
    '       -Mira en m_ObjEntorno.URLDirectorioLocal si no hay HTML1.txt si está pasa un bucle de 50 hasta que dé con uno que no esté
    
    '   -Llamada desde
   
    '   -Devuelve:
    '       DameUntxtYHtml =  strURLTXT & ";" & strURLHTML
    '       DameUntxtYHtml = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim fso As New FileSystemObject, i As Integer, strURLTXT As String, strURLHTML As String, strNombreHTML As String, strNombretxt As String, strResultado As String
    Dim strTextoError As String
    On Error GoTo errores
    
    BorraHTMLs
    For i = 1 To 50
        strNombretxt = "HTML" & i & ".txt"
        strNombreHTML = "HTML" & i & ".html"
        strURLTXT = m_ObjEntorno.URLDirectorioLocal & strNombretxt
        strURLHTML = m_ObjEntorno.URLDirectorioLocal & strNombreHTML
        If Not fso.FileExists(strURLTXT) And Not fso.FileExists(strURLHTML) Then
            strResultado = strURLTXT & ";" & strURLHTML
            DameUntxtYHtml = strResultado
            If Not fso Is Nothing Then
                Set fso = Nothing
            End If
            Exit Function
        End If
        
    Next
    DameUntxtYHtml = ""
    If Not fso Is Nothing Then
        Set fso = Nothing
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método DameUntxtYHtml ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    If Not fso Is Nothing Then
        Set fso = Nothing
    End If
    DameUntxtYHtml = "#ERR" & "|" & strTextoError
End Function
Function HTMLENTXT(strMensaje As String) As String
    
    Dim fso As New FileSystemObject, F1 As Object, strURLHTML As String, strNombreHTML As String, _
        strURLTXT As String, strURLCompletaArchivo As String, strNombretxt As String
    Dim strResultado As String, strTextoError As String
    On Error GoTo errores
    If strMensaje = "" Then
        strTextoError = "No se ha incluido el texto del HTML"
        Err.Raise 1000
    End If
   
    flag = DameUntxtYHtml()
    If InStr(1, flag, "|") <> 0 Then
        dato = Split(flag, "|")
        strTextoError = "El método DameUntxtYHtml ha devuelto el error: " & vbNewLine & dato(1)
        Err.Raise 1000
    End If
    strResultado = flag
    If InStr(1, strResultado, ";") = 0 Then
        strTextoError = "El método DameUntxtYHtml ha devuelto un resultado con formato desconocido"
        Err.Raise 1000
    End If
    dato = Split(strResultado, ";")
    strURLTXT = dato(0)
    strURLHTML = dato(1)
    Set F1 = fso.CreateTextFile(strURLTXT, True)
    F1.WriteLine strMensaje
    F1.Close
    fso.GetFile(strURLTXT).Name = fso.GetBaseName(strURLTXT) & ".html"
    Set fso = Nothing
    Ejecutar 1, "open", strURLHTML, "", "", 1
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método HTMLENTXT ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
    If Not fso Is Nothing Then
        Set fso = Nothing
    End If
    HTMLENTXT = "#ERR" & "|" & strTextoError
End Function
Public Function DameHTML( _
                            col As Collection, _
                            Optional strEsEquipo As String, _
                            Optional strParteTitulo As String _
                            ) As String
     '--------------------------------------------------------
    ' Función creada por Andrés Román del Peral el día x
    '   -Modificaciones:
  
    '   -Funcionamiento:
    '       -Copia el html en un archivo de texto
    '       -Lo renombra a html
    '       -abre el explorador windows de microsoft y muestra el archivo
    '   -Llamada desde
    '       -Form_FormPlanificacion.PreventivoEquipos_Click
    '       -Form_FormPlanificacion.PreventivosEquiposNoProgramados_Click
    '       -Form_FormPlanificacion.PreventivosEquiposProgramados_Click
    '       -Form_FormPlanificacion.PreventivosNoRealizados_Click
    '       -Form_FormPlanificacion.PreventivosNoRealizadosProximos_Click
    '       -Form_FormPlanificacion.PreventivosNoRealizadosVencidos_Click
    '       -Form_FormPlanificacion.PreventivosProgramados_Click
    '       -Form_FormPlanificacion.PreventivosRealizados_Click
    '       -Form_FormPlanificacion.PreventivosRealizadosEnPlazo_Click
    '       -Form_FormPlanificacion.PreventivosRealizadosFueraDePlazo_Click
    '   -Devuelve:
    '       DameHTML = strMensaje
    '       DameHTML = "#ERR" & "|" & strTextoError
    '-------------------------------------------------------------------
    Dim strMensajeCSS As String, VarItem, i As Integer, strMensaje As String, strTextoError As String
    Dim strIDEquipo As String, strFechabajaParaPlanificacion As String, strBUI As String, strSUBSISTEMA As String, _
        strPeriodicidadEnMesesRecomendada As String
    Dim strMotivoCierre As String, strIDEvento As String, strIDNOEvento As String, strFechaPrevistaCierre As String, strFechaCierre As String, _
        strMotivacionReprogramacion As String, strIDNuevaPlanificacion As String, strIDPlanificacion As String
        '----------------------------------------------------------------
        ' CABECERA DE HTML
        '----------------------------------------------------------------
            strMensaje = "<html>" & vbCrLf
            strMensaje = strMensaje & "<html lang=""es"">" & vbCrLf
            strMensaje = strMensaje & "<head>" & vbCrLf
                strMensaje = strMensaje & "<meta charset=""ISO-8859-1"" />" & vbCrLf
                strMensaje = strMensaje & "<title>MANTENIMIENTOS PREVENTIVOS " & strParteTitulo & "</title>" & vbCrLf
                strMensaje = strMensaje & "<style type=""text/css"">" & vbCrLf
                    strMensaje = strMensaje & m_ObjEntorno.CSS & vbCrLf
                strMensaje = strMensaje & "</style>" & vbCrLf
            strMensaje = strMensaje & "</head>" & vbCrLf
            strMensaje = strMensaje & "<body>" & vbCrLf
            
            '----------------------------------------------------------------
            ' COMIENZO DE TABLA
            '----------------------------------------------------------------
                strMensaje = strMensaje & "<table>" & vbCrLf
                strMensaje = strMensaje & "<tbody>" & vbCrLf
                strMensaje = strMensaje & "<tr>" & vbCrLf
                    If strEsEquipo = "No" Then
                        strMensaje = strMensaje & "<td colspan=""12"" class=""ColespanArriba""> <div> MANTENIMIENTOS PREVENTIVOS " & strParteTitulo & "</div></td>" & vbCrLf
                    Else
                        strMensaje = strMensaje & "<td colspan=""4"" class=""ColespanArriba""> <div> MANTENIMIENTOS PREVENTIVOS " & strParteTitulo & "</div></td>" & vbCrLf
                    End If
                strMensaje = strMensaje & "</tr>" & vbCrLf
                strMensaje = strMensaje & "<tr class =""Cabecera"">" & vbCrLf
                    strMensaje = strMensaje & "<th> BUI</th>" & vbCrLf
                    strMensaje = strMensaje & "<th> SUBSISTEMA</th>" & vbCrLf
                    strMensaje = strMensaje & "<th> EQUIPO</th>" & vbCrLf
                    strMensaje = strMensaje & "<th> PERIODICIDAD</th>" & vbCrLf
                    If strEsEquipo = "No" Then
                        strMensaje = strMensaje & "<th>ID.PLANIF</th>" & vbCrLf
                        strMensaje = strMensaje & "<th> F.PREVISTA</th>" & vbCrLf
                        strMensaje = strMensaje & "<th> F.CIERRE</th>" & vbCrLf
                        strMensaje = strMensaje & "<th> MOTIVO CIERRE</th>" & vbCrLf
                        strMensaje = strMensaje & "<th> EVENTO</th>" & vbCrLf
                        strMensaje = strMensaje & "<th> NO EVENTO</th>" & vbCrLf
                        strMensaje = strMensaje & "<th> MOTIVO REPROGRAMA.</th>" & vbCrLf
                        strMensaje = strMensaje & "<th> NUEVA REPROG.</th>" & vbCrLf
                    End If
                strMensaje = strMensaje & "</tr>" & vbCrLf
            
        
            For Each VarItem In col
                'strResultado = strIDEquipo & "|" & strBUI & "|" & strSUBSISTEMA & "|" & strIDEquipo & "|" & strPeriodicidadEnMesesRecomendada & "|" &
                                'strIDPlanificacion & "|" & strFechaPrevistaCierre & "|" & strFechaCierre & "|" & _
                                'strMotivoCierre & "|" & strIDEvento & "|" & strIDNOEvento & "|" & strMotivacionReprogramacion & "|" & strIDNuevaPlanificacion
                
                dato = Split(VarItem, "|")
                strBUI = CStr(dato(1))
                strSUBSISTEMA = CStr(dato(2))
                strIDEquipo = CStr(dato(3))
                strPeriodicidadEnMesesRecomendada = CStr(dato(4))
                If strEsEquipo = "No" Then
                    strIDPlanificacion = CStr(dato(5))
                    strFechaPrevistaCierre = CStr(dato(6))
                    strFechaCierre = CStr(dato(7))
                    strMotivoCierre = CStr(dato(8))
                    strIDEvento = CStr(dato(9))
                    strIDNOEvento = CStr(dato(10))
                    strMotivacionReprogramacion = CStr(dato(11))
                    strIDNuevaPlanificacion = CStr(dato(12))
                End If
                
              
                strMensaje = strMensaje & "<tr>" & vbCrLf
                
                    strMensaje = strMensaje & "<td>" & strBUI & "</td>" & vbCrLf
                    strMensaje = strMensaje & "<td>" & strSUBSISTEMA & "</td>" & vbCrLf
                    strMensaje = strMensaje & "<td>" & strIDEquipo & "</td>" & vbCrLf
                    strMensaje = strMensaje & "<td>" & strPeriodicidadEnMesesRecomendada & "</td>" & vbCrLf
                    If strEsEquipo = "No" Then
                        strMensaje = strMensaje & "<td class=""centrado"">" & strIDPlanificacion & "</td>" & vbCrLf
                        strMensaje = strMensaje & "<td class=""centrado"">" & strFechaPrevistaCierre & "</td>" & vbCrLf
                        strMensaje = strMensaje & "<td class=""centrado"">" & strFechaCierre & "</td>" & vbCrLf
                        strMensaje = strMensaje & "<td>" & strMotivoCierre & "</td>" & vbCrLf
                         strMensaje = strMensaje & "<td class=""centrado"">" & strIDEvento & "</td>" & vbCrLf
                        strMensaje = strMensaje & "<td class=""centrado"">" & strIDNOEvento & "</td>" & vbCrLf
                        strMensaje = strMensaje & "<td>" & strMotivacionReprogramacion & "</td>" & vbCrLf
                        strMensaje = strMensaje & "<td class=""centrado"">" & strIDNuevaPlanificacion & "</td>" & vbCrLf
                    End If
                strMensaje = strMensaje & "</tr>" & vbCrLf
            Next
            '----------------------------------------------------------------
            ' FIN DE TABLA DE SUGERENCIAS
            '----------------------------------------------------------------
            strMensaje = strMensaje & "</table>" & vbCrLf
    '----------------------------------------------------------------
    ' FIN DE HTML
    '----------------------------------------------------------------
        strMensaje = strMensaje & "</tbody>" & vbCrLf
        strMensaje = strMensaje & "</body>" & vbCrLf
        strMensaje = strMensaje & "</html>" & vbCrLf
        
    DameHTML = strMensaje
    Exit Function
errores:
    If Err.Number <> 1000 Then
        strTextoError = "El método DameHTML ha devuelto el error: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    DameHTML = "#ERR" & "|" & strTextoError
End Function







