VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HTML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public Mensaje As String
Public Titulo As String
Public URLArchivo As String
Private m_sCabecera As String
Private m_sError As String
Public Error As String
Private m_Error As String


Public Property Get Cabecera() As String
    
    Dim m_Cabecera As String
    On Error GoTo errores
    Me.Error = ""
    
    If m_sCabecera <> "" Then
        Cabecera = m_sCabecera
        Exit Property
    End If
    If Me.Titulo = "" Then
        m_Error = "Falta el título"
        Err.Raise 1000
    End If
    m_Cabecera = m_ObjEntorno.CabeceraHTML
    m_Error = m_ObjEntorno.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If InStr(1, m_Cabecera, "#titulo") <> 0 Then
        m_sCabecera = Replace(m_Cabecera, "#titulo", Me.Titulo)
    End If
    If m_sCabecera <> "" Then
        Cabecera = m_sCabecera
        Exit Property
    End If
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método HTML.Cabecera ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Function MostrarEdicionPublicacion( _
                                            p_ObjDatos As Object, _
                                            Optional ByRef p_Error As String _
                                            ) As String
    
    Dim m_mensaje As ADODB.stream
    
    On Error GoTo errores
    
    Me.Error = ""
    m_mensaje = HTMLEdicionPublicacion(p_ObjDatos, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    MostrarHTML p_Mensaje:=m_mensaje, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HTML.MostrarEdicionPublicacion ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function MostrarEdicionCambios( _
                                            p_ObjColCambios As Scripting.Dictionary, _
                                            p_Edicion As String, _
                                            p_EdicionAnterior As String, _
                                            p_NombreProyecto As String, _
                                            Optional ByRef p_Error As String _
                                            ) As String
    
    Dim m_mensaje As ADODB.stream
    
    On Error GoTo errores
    
    Me.Error = ""
    Set m_mensaje = HTMLEdicionCambios(p_ObjColCambios, p_Edicion, p_EdicionAnterior, p_NombreProyecto, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
     MostrarHTML p_Mensaje:=m_mensaje, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HTML.MostrarEdicionCambios ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Public Function MostrarControlCambiosVersion( _
                                                p_Version As CCVersion, _
                                                Optional p_QueMostrar As EnumControlCambios = EnumControlCambios.Todo, _
                                                Optional ByRef p_Error As String _
                                                ) As String
    
    Dim m_mensaje As ADODB.stream
    
    On Error GoTo errores
    
    Me.Error = ""
    Set m_mensaje = HTMLVersionCambios(p_Version, p_QueMostrar, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
     MostrarHTML p_Mensaje:=m_mensaje, p_EnFormulario:=EnumSiNo.No, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HTML.MostrarControlCambiosVersion ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Public Function MostrarEdicionQuitarPublicacion( _
                                                p_ObjDatos As Object, _
                                                Optional ByRef p_Error As String _
                                                ) As String
    
    Dim m_mensaje As ADODB.stream
    
    On Error GoTo errores
    Me.Error = ""
    m_mensaje = HTMLEdicionQuitarPublicacion(p_ObjDatos, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
     MostrarHTML p_Mensaje:=m_mensaje, p_Error:=p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HTML.MostrarEdicionQuitarPublicacion ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function
Private Function HTMLEdicionCambios( _
                                        p_ObjColCambios As Scripting.Dictionary, _
                                        p_Edicion As String, _
                                        p_EdicionAnterior As String, _
                                        p_NombreProyecto As String, _
                                        Optional ByRef p_Error As String _
                                        ) As ADODB.stream
    
    
    Dim m_HTMLTablaCambios As String
    Dim m_HTMLTablaCambiosExplicaciones As String
    
    On Error GoTo errores
    
    If p_ObjColCambios Is Nothing Then
        Exit Function
    End If
    
    m_HTMLTablaCambios = HTMLCambios(p_ObjColCambios, p_Edicion, p_EdicionAnterior, p_NombreProyecto, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_HTMLTablaCambiosExplicaciones = getHTMLTablaCambiosExplicaciones(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Me.Titulo = "Cambios"
    Set HTMLEdicionCambios = New ADODB.stream
    HTMLEdicionCambios.Open
    HTMLEdicionCambios.WriteText Me.Cabecera & vbNewLine
    
    HTMLEdicionCambios.WriteText m_HTMLTablaCambios & vbNewLine
    HTMLEdicionCambios.WriteText "<br /><br />" & vbNewLine
    HTMLEdicionCambios.WriteText m_HTMLTablaCambiosExplicaciones & vbNewLine
    HTMLEdicionCambios.WriteText "</body>" & vbNewLine
    HTMLEdicionCambios.WriteText "</html>" & vbNewLine
    HTMLEdicionCambios.WriteText m_ObjEntorno.FinHTML & vbNewLine
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HTML.HTMLEdicionCambios ha devuelto el error: " & vbNewLine & Err.Description
    End If
    
End Function

Private Function HTMLEdicionPublicacion( _
                                        p_ObjEdicion As Edicion, _
                                        Optional ByRef p_Error As String _
                                        ) As ADODB.stream
    
    Dim m_Asunto As String
    Dim m_HTMLTablaEdicion As String
    
    On Error GoTo errores
    
    If p_ObjEdicion Is Nothing Then
        p_Error = "Se ha de introducir la Edición"
        Err.Raise 1000
    End If
    Set HTMLEdicionPublicacion = New ADODB.stream
    HTMLEdicionPublicacion.Open
    m_Asunto = "PUBLICACIÓN DE LA Edición " & p_ObjEdicion.Edicion & " DE LA gestión de riesgos " & _
        p_ObjEdicion.Proyecto.NombreProyecto
    
    m_HTMLTablaEdicion = HTMLTablaEdicion(p_ObjEdicion, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Me.Titulo = m_Asunto
    HTMLEdicionPublicacion.WriteText Me.Cabecera & vbNewLine
    HTMLEdicionPublicacion.WriteText m_Asunto & vbNewLine
    HTMLEdicionPublicacion.WriteText "<br /><br />" & vbNewLine
    HTMLEdicionPublicacion.WriteText m_HTMLTablaEdicion & vbNewLine
    HTMLEdicionPublicacion.WriteText "</body>" & vbNewLine
    HTMLEdicionPublicacion.WriteText "</html>" & vbNewLine
    HTMLEdicionPublicacion.WriteText m_ObjEntorno.FinHTML & vbNewLine
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método HTML.HTMLEdicionPublicacion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    
End Function

Private Function HTMLEdicionQuitarPublicacion( _
                                                p_ObjEdicion As Edicion, _
                                                Optional ByRef p_Error As String _
                                                ) As ADODB.stream
        
    Dim m_Asunto As String
    
    Dim m_HTMLTablaEdicion As String
    
    On Error GoTo errores
    
    If p_ObjEdicion Is Nothing Then
        p_Error = "Se ha de introducir la Edición"
        Err.Raise 1000
    End If
    Set HTMLEdicionQuitarPublicacion = New ADODB.stream
    HTMLEdicionQuitarPublicacion.Open
    m_Asunto = "ELIMINADA PUBLICACIÓN DE LA Edición " & p_ObjEdicion.Edicion & " DE LA gestión de riesgos " & _
        p_ObjEdicion.Proyecto.NombreProyecto
    
    m_HTMLTablaEdicion = HTMLTablaEdicion(p_ObjEdicion, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Me.Titulo = m_Asunto
    HTMLEdicionQuitarPublicacion.WriteText Me.Cabecera & vbNewLine
    HTMLEdicionQuitarPublicacion.WriteText m_Asunto & vbNewLine
    HTMLEdicionQuitarPublicacion.WriteText "<br /><br />" & vbNewLine
    HTMLEdicionQuitarPublicacion.WriteText m_HTMLTablaEdicion & vbNewLine
    HTMLEdicionQuitarPublicacion.WriteText "</body>" & vbNewLine
    HTMLEdicionQuitarPublicacion.WriteText "</html>" & vbNewLine
    HTMLEdicionQuitarPublicacion.WriteText m_ObjEntorno.FinHTML & vbNewLine
    Exit Function
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método HTML.HTMLEdicionQuitarPublicacion ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function

Private Function HTMLTablaEdicion( _
                                    p_ObjEdicion As Edicion, _
                                    Optional ByRef p_Error As String _
                                    ) As String
    
    Dim m_IdRiesgo As Variant
    Dim m_ObjRiesgo As riesgo
    Dim m_mensaje As String
    Dim m_Proyecto As Proyecto
    
    On Error GoTo errores
    
    If p_ObjEdicion Is Nothing Then
        p_Error = "Se ha de introducir la Edición"
        Err.Raise 1000
    End If
    Set m_Proyecto = Constructor.getProyecto(p_ObjEdicion.IDProyecto, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_mensaje = "<table>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
        m_mensaje = m_mensaje & "<td colspan='2' class=""ColespanArriba""> DATOS GENERALES</td>"
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""centrado"" > <strong>p_ObjEdicion</strong></td>"
            m_mensaje = m_mensaje & "<td> " & _
                IIf(m_Proyecto.Proyecto = "", "&nbsp;", m_Proyecto.Proyecto) & "</td>" & vbNewLine
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""centrado"" > <strong>Nombre del proyecto</strong></td>"
            m_mensaje = m_mensaje & "<td> " & _
            IIf(m_Proyecto.NombreProyecto = "", "&nbsp;", m_Proyecto.NombreProyecto) & "</td>" & vbNewLine
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""centrado"" > <strong>Jurídica</strong></td>"
            m_mensaje = m_mensaje & "<td> " & _
            IIf(m_Proyecto.Juridica = "", "&nbsp;", m_Proyecto.Juridica) & "</td>" & vbNewLine
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""centrado"" > <strong>Cliente</strong></td>"
            m_mensaje = m_mensaje & "<td> " & _
            IIf(m_Proyecto.Cliente = "", "&nbsp;", m_Proyecto.Cliente) & "</td>" & vbNewLine
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""centrado"" > <strong>En UTE</strong></td>"
            m_mensaje = m_mensaje & "<td> " & _
            IIf(m_Proyecto.EnUTE = "", "&nbsp;", m_Proyecto.EnUTE) & "</td>" & vbNewLine
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""centrado"" > <strong>Edición</strong></td>"
            m_mensaje = m_mensaje & "<td> " & _
            IIf(p_ObjEdicion.Edicion = "", "&nbsp;", p_ObjEdicion.Edicion) & "</td>" & vbNewLine
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""centrado"" > <strong>F.Publicación</strong></td>"
            m_mensaje = m_mensaje & "<td> " & _
            IIf(p_ObjEdicion.FechaPublicacion = "", "&nbsp;", p_ObjEdicion.FechaPublicacion) & "</td>" & vbNewLine
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""centrado"" > <strong>Código Documento</strong></td>"
            m_mensaje = m_mensaje & "<td> " & _
            IIf(m_Proyecto.CodigoDocumento = "", "&nbsp;", m_Proyecto.CodigoDocumento) & "</td>" & vbNewLine
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""centrado"" > <strong>Elabora</strong></td>"
            m_mensaje = m_mensaje & "<td> " & _
            IIf(p_ObjEdicion.Elaborado = "", "&nbsp;", p_ObjEdicion.Elaborado) & "</td>" & vbNewLine
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""centrado"" > <strong>Revisa</strong></td>"
            m_mensaje = m_mensaje & "<td> " & _
            IIf(p_ObjEdicion.Revisado = "", "&nbsp;", p_ObjEdicion.Revisado) & "</td>" & vbNewLine
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""centrado"" > <strong>Aprueba</strong></td>"
            m_mensaje = m_mensaje & "<td> " & _
            IIf(p_ObjEdicion.Aprobado = "", "&nbsp;", p_ObjEdicion.Aprobado) & "</td>" & vbNewLine
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
    m_mensaje = m_mensaje & "</table>" & vbNewLine
    
    m_mensaje = m_mensaje & "<br /><br />" & vbNewLine
    '------------------------------------------------------
    '   GENERACIÓN DE HTML DE LOS RIESGOS
    '------------------------------------------------------
    m_mensaje = m_mensaje & "<table>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
        m_mensaje = m_mensaje & "<td colspan='3' class=""ColespanArriba""> RIESGOS</td>" & vbNewLine
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""centrado"" > <strong>Código</strong></td>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""centrado"" > <strong>Descripción</strong></td>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""centrado"" > <strong>Detectado Por</strong></td>" & vbNewLine
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        If Not p_ObjEdicion.colRiesgos Is Nothing Then
            For Each m_IdRiesgo In p_ObjEdicion.colRiesgos.keys
                Set m_ObjRiesgo = p_ObjEdicion.colRiesgos(m_IdRiesgo)
                m_mensaje = m_mensaje & "<tr>" & vbNewLine
                    m_mensaje = m_mensaje & "<td> " & _
                    IIf(m_ObjRiesgo.CodigoRiesgo = "", "&nbsp;", m_ObjRiesgo.CodigoRiesgo) & "</td>" & vbNewLine
                    m_mensaje = m_mensaje & "<td> " & _
                    IIf(m_ObjRiesgo.Descripcion = "", "&nbsp;", m_ObjRiesgo.Descripcion) & "</td>" & vbNewLine
                    m_mensaje = m_mensaje & "<td> " & _
                    IIf(m_ObjRiesgo.DetectadoPor = "", "&nbsp;", m_ObjRiesgo.DetectadoPor) & "</td>" & vbNewLine
                m_mensaje = m_mensaje & "</tr>" & vbNewLine
                Set m_ObjRiesgo = Nothing
            Next
        End If
    m_mensaje = m_mensaje & "</table>" & vbNewLine
    HTMLTablaEdicion = m_mensaje
    Exit Function
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método HTML.HTMLTablaEdicion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    
End Function
Private Function HTMLCambios( _
                                p_ColCambios As Scripting.Dictionary, _
                                p_Edicion As String, _
                                p_EdicionAnterior As String, _
                                p_NombreProyecto As String, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    Dim m_Cambio As Cambio
    Dim m_IDCambio As Variant
    Dim m_mensaje As String
    Dim m_Titulo As String
    
    On Error GoTo errores
    
    If p_ColCambios Is Nothing Then
        p_Error = "Sin Cambios"
        Err.Raise 1000
    End If
    m_Titulo = "CAMBIO ENTRE EDICIONES " & p_Edicion & " Y " & p_EdicionAnterior & " EN " & UCase(p_NombreProyecto)
    m_mensaje = "<table>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td colspan='2' class=""ColespanArriba""> " & m_Titulo & " </td>"
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""Cabecera"" > <strong>APARTADOS</strong></td>"
            m_mensaje = m_mensaje & "<td class=""Cabecera"" > <strong>DESCRIPCIÓN DEL CAMBIO</strong></td>"
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        For Each m_IDCambio In p_ColCambios
            Set m_Cambio = p_ColCambios(m_IDCambio)
            With m_Cambio
                m_mensaje = m_mensaje & "<tr>" & vbNewLine
                    m_mensaje = m_mensaje & "<td> " & .riesgo & "</td>" & vbNewLine
                    m_mensaje = m_mensaje & "<td> " & .Descripcion & "</td>" & vbNewLine
                    
                m_mensaje = m_mensaje & "</tr>" & vbNewLine
            End With
            Set m_Cambio = Nothing
        Next
        
    m_mensaje = m_mensaje & "</table>" & vbNewLine
    
    HTMLCambios = m_mensaje
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HTML.HTMLCambios ha devuelto el error: " & vbNewLine & Err.Description
    End If
    
End Function

Private Function HTMLTablaVersionCambios( _
                                        p_Version As CCVersion, _
                                        Optional p_QueMostrar As EnumControlCambios = EnumControlCambios.Todo, _
                                        Optional ByRef p_Error As String _
                                        ) As String
    
    Dim m_Cambio As CCCambio
    Dim m_Cambios As Scripting.Dictionary
    Dim m_IDCambio As Variant
    Dim m_mensaje As String
    Dim m_Titulo As String
    Dim m_ColResultante As Scripting.Dictionary
    Dim m_Documentos As Scripting.Dictionary
    Dim m_IDDoc As Variant
    Dim m_Doc As CCDocumentoCambio
    

    On Error GoTo errores
    
    If p_Version Is Nothing Then
        p_Error = "Sin Versión"
        Err.Raise 1000
    End If
    Set m_Cambios = p_Version.Cambios
    If m_Cambios Is Nothing Then
        p_Error = "Sin Cambios"
        Err.Raise 1000
    End If
    
    
    If p_QueMostrar = EnumControlCambios.Todo Then
        Set m_ColResultante = m_Cambios
    ElseIf p_QueMostrar = EnumControlCambios.SoloCalidad Then
        For Each m_IDCambio In m_Cambios
            Set m_Cambio = m_Cambios(m_IDCambio)
            With m_Cambio
                If .MostrarEninformeGeneral = "Sí" Then
                    If m_ColResultante Is Nothing Then
                        Set m_ColResultante = New Scripting.Dictionary
                        m_ColResultante.CompareMode = TextCompare
                    End If
                    If Not m_ColResultante.Exists(m_IDCambio) Then
                        m_ColResultante.Add m_IDCambio, m_Cambio
                    End If
                End If
            End With
            Set m_Cambio = Nothing
        Next
    Else
        p_Error = "No se ha definido qué mostrar"
        Err.Raise 1000
    End If
    
    m_Titulo = "CONTROL DE CAMBIOS VERSIÓN " & p_Version.Version
    m_mensaje = "<table>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td colspan='6' class=""ColespanArriba""> " & m_Titulo & " </td>"
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""Cabecera"" > <strong>CÓDIGO</strong></td>"
            m_mensaje = m_mensaje & "<td class=""Cabecera"" > <strong>CAMBIOS</strong></td>"
            m_mensaje = m_mensaje & "<td class=""Cabecera"" > <strong>ERROR</strong></td>"
            m_mensaje = m_mensaje & "<td class=""Cabecera"" > <strong>MEJORA</strong></td>"
            m_mensaje = m_mensaje & "<td class=""Cabecera"" > <strong>DESCRIPCIÓN DEL CAMBIO</strong></td>"
            m_mensaje = m_mensaje & "<td class=""Cabecera"" > <strong>DOCUMENTOS</strong></td>"
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        If Not m_ColResultante Is Nothing Then
            For Each m_IDCambio In m_ColResultante
                Set m_Cambio = m_ColResultante(m_IDCambio)
                With m_Cambio
                    m_mensaje = m_mensaje & "<tr>" & vbNewLine
                        m_mensaje = m_mensaje & "<td> " & .CodCambio & "</td>" & vbNewLine
                        m_mensaje = m_mensaje & "<td> " & .NombreCambio & "</td>" & vbNewLine
                        m_mensaje = m_mensaje & "<td class=""centrado"" > " & .EsBug & "</td>" & vbNewLine
                        m_mensaje = m_mensaje & "<td class=""centrado"" > " & .EsMejora & "</td>" & vbNewLine
                        m_mensaje = m_mensaje & "<td> " & .DescripcionCambio & "</td>" & vbNewLine
                        
                        Set m_Documentos = .Documentos
                        If Not m_Documentos Is Nothing Then
                            m_mensaje = m_mensaje & "<td class=""centrado"" >"
                            For Each m_IDDoc In m_Documentos
                                Set m_Doc = m_Documentos(m_IDDoc)
                                    m_mensaje = m_mensaje & "<a href='" & m_Doc.URLDocumentoCalculado & "' target='_blank'>" & vbNewLine
                                        m_mensaje = m_mensaje & "<img src='" & m_ObjEntorno.URLIconoPDF & "'>" & vbNewLine
                                    m_mensaje = m_mensaje & "</a>" & vbNewLine
                                    
                                Set m_Doc = Nothing
                            Next
                            m_mensaje = m_mensaje & "</td>" & vbNewLine
                        Else
                            m_mensaje = m_mensaje & "<td> &nbsp;</td>" & vbNewLine
                            
                        End If
                    
                End With
                Set m_Cambio = Nothing
            Next
        End If
        
        
    m_mensaje = m_mensaje & "</table>" & vbNewLine
    
    HTMLTablaVersionCambios = m_mensaje
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HTML.HTMLTablaVersionCambios ha devuelto el error: " & vbNewLine & Err.Description
    End If
    
End Function
Private Function HTMLTablaVersionGeneral( _
                                        p_Version As CCVersion, _
                                        Optional ByRef p_Error As String _
                                        ) As String
    
    Dim m_Descripcion As String
    Dim m_FechaInicio As String
    Dim m_FechaFin As String
    Dim m_mensaje As String
    Dim m_Titulo As String
    
    

    On Error GoTo errores
    
    If p_Version Is Nothing Then
        p_Error = "Sin Versión"
        Err.Raise 1000
    End If
    
    
    m_Titulo = "VERSIÓN " & p_Version.Version
    m_mensaje = "<table>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td colspan='6' class=""ColespanArriba""> " & m_Titulo & " </td>"
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""Cabecera"" > <strong>F.Inicio</strong></td>"
            m_mensaje = m_mensaje & "<td class=""Cabecera"" > <strong>F.Fin</strong></td>"
            m_mensaje = m_mensaje & "<td class=""Cabecera"" > <strong>DESCRIPCIÓN</strong></td>"
            m_mensaje = m_mensaje & "<td class=""Cabecera"" > <strong>CAMBIOS DE ERRORES</strong></td>"
            m_mensaje = m_mensaje & "<td class=""Cabecera"" > <strong>CAMBIOS DE MEJORA</strong></td>"
            m_mensaje = m_mensaje & "<td class=""Cabecera"" > <strong>DOCUMENTO</strong></td>"
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        With p_Version
            m_mensaje = m_mensaje & "<tr>" & vbNewLine
                m_mensaje = m_mensaje & "<td class=""centrado"" > " & .FechaInicio & "</td>" & vbNewLine
                m_mensaje = m_mensaje & "<td class=""centrado"" > " & .fechaFin & "</td>" & vbNewLine
                m_mensaje = m_mensaje & "<td> " & .Descripcion & "</td>" & vbNewLine
                m_mensaje = m_mensaje & "<td class=""centrado"" > " & .CambiosDeErrores & "</td>" & vbNewLine
                m_mensaje = m_mensaje & "<td class=""centrado"" > " & .CambiosDeMejoras & "</td>" & vbNewLine
                If fso.FileExists(.URLRequisitosPreviosCalculada) Then
                    m_mensaje = m_mensaje & "<td class=""centrado"" >"
                        m_mensaje = m_mensaje & "<a href='" & .URLRequisitosPreviosCalculada & "' target='_blank'>" & vbNewLine
                            m_mensaje = m_mensaje & "<img src='" & .URLIcono & "'>" & vbNewLine
                        m_mensaje = m_mensaje & "</a>" & vbNewLine
                    m_mensaje = m_mensaje & "</td>"
                    
                Else
                    m_mensaje = m_mensaje & "<td> &nbsp;</td>" & vbNewLine
                End If
            m_mensaje = m_mensaje & "</tr>" & vbNewLine
        End With
                
        
        
    m_mensaje = m_mensaje & "</table>" & vbNewLine
    
    HTMLTablaVersionGeneral = m_mensaje
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HTML.HTMLTablaVersionGeneral ha devuelto el error: " & vbNewLine & Err.Description
    End If
    
End Function
Public Function MostrarHTML( _
                            p_Mensaje As ADODB.stream, _
                            Optional p_EnFormulario As EnumSiNo = EnumSiNo.Sí, _
                            Optional ByRef p_Error As String _
                            ) As String
    
    
    
    On Error GoTo errores
    p_EnFormulario = EnumSiNo.No
    HTMLENTXT , p_Mensaje, p_EnFormulario, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método HTML.MostrarHTML ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function


Private Function HTMLVersionCambios( _
                                        p_Version As CCVersion, _
                                        Optional p_QueMostrar As EnumControlCambios = EnumControlCambios.Todo, _
                                        Optional ByRef p_Error As String _
                                        ) As ADODB.stream
    
    
    Dim m_HTMLTablaVersionGeneral As String
    Dim m_HTMLTablaVersionCambios As String
    On Error GoTo errores
    
    If p_Version Is Nothing Then
        Exit Function
    End If
    m_HTMLTablaVersionGeneral = HTMLTablaVersionGeneral(p_Version, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_HTMLTablaVersionCambios = HTMLTablaVersionCambios(p_Version, p_QueMostrar, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    
    Me.Titulo = "Control de Cambios"
    Set HTMLVersionCambios = New ADODB.stream
    HTMLVersionCambios.Open
    HTMLVersionCambios.WriteText Me.Cabecera & vbNewLine
    
    HTMLVersionCambios.WriteText m_HTMLTablaVersionGeneral & vbNewLine
    HTMLVersionCambios.WriteText "</BR>" & vbNewLine
    HTMLVersionCambios.WriteText m_HTMLTablaVersionCambios & vbNewLine
    
    HTMLVersionCambios.WriteText "</body>" & vbNewLine
    HTMLVersionCambios.WriteText "</html>" & vbNewLine
    HTMLVersionCambios.WriteText m_ObjEntorno.FinHTML & vbNewLine
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HTML.HTMLVersionCambios ha devuelto el error: " & vbNewLine & Err.Description
    End If
    
End Function



