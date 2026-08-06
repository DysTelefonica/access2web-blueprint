VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Documento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Public IDDocumento As String
Public FechaAlta As String
Public CodExp As String
Public codigo As String
Public Titulo As String
Public Tipo As String
Public Area As String
Public Edicion As String
Public FEdicion As String
Public FENTRADAENVIGOR As String
Public Archivo As String
Public ResponsableEdicion As String
Public URLCarpetaArchivosLocales As String
Public CadenaNombreArchivos As String
Public URLCarpetaAGEDO As String
Public CENTRO As String
Public Clasificacion As String
Public Registrador As String
Public RegistroEntrada As String
Public RegistroSalida As String
Public Versionable As String
Public UltimaVersion As String
Public Estado As String
Public FechaCaducidad As String
Public CLASE As String
Public Observaciones As String
Public FechaDeComienzoNuevaVersion As String
Public IDDocumentoNuevaVersion As String
Public IDDocumentoRevision As String
Public NCPARAINDICADOR As String
Public NCABIERTACERRADA As String
Public AdjuntoEncriptado As String
Private m_objEdicionObj As Edicion
Private m_ObjProyecto As Proyecto
Private m_sURLAdjunto As String
Private m_sIDProyectoCalculado As String
Private m_URLFinal As String


Private m_objColCampos As Collection
Public Error As String
Private m_Error As String


Public Property Get IDDocumentoCalculado() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    m_sIDProyectoCalculado = DameID("TbDocumentos", "IDDOcumento", getdb, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    IDDocumentoCalculado = m_sIDProyectoCalculado
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Documento.IDDocumentoCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property


Public Property Get Proyecto() As Proyecto

    Set Proyecto = m_ObjProyecto

End Property

Public Property Set Proyecto(ByVal objNewValue As Proyecto)

    Set m_ObjProyecto = objNewValue

End Property


Public Property Get EdicionObj() As Edicion

    Set EdicionObj = m_objEdicionObj

End Property

Public Property Set EdicionObj(ByVal objNewValue As Edicion)

    Set m_objEdicionObj = objNewValue

End Property


Public Property Get URLAdjunto() As String
    
    Dim m_URLFinal As String
    Dim m_URLCarpeta As String
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    If m_sURLAdjunto <> "" Then
        URLAdjunto = m_sURLAdjunto
        Exit Property
    End If
    If Me.CadenaNombreArchivos = "" Then
        Exit Property
    End If
    If Me.URLCarpetaAGEDO = "" Then
        Exit Property
    End If
    
    m_URLCarpeta = m_ObjEntorno.URLDirectorioTemporalAGEDO
    
    m_sURLAdjunto = m_URLCarpeta & Me.CadenaNombreArchivos
    URLAdjunto = m_sURLAdjunto
       
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Documento.URLAdjunto ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Function MotivoNoOK( _
                                p_URLArchivoLocal As String, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    Dim m_FechaRef As Date
    
    Me.Error = ""
    On Error GoTo errores
    
    m_FechaRef = Date
    'm_fechaRef = "12/12/2023"
    With Me
        If .IDDocumento = "" Then
            .IDDocumento = .IDDocumentoCalculado
            p_Error = .Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        If .CodExp = "" Then
            
            If Me.Proyecto Is Nothing Then
                p_Error = "No se conoce el expediente del documento"
                Err.Raise 1000
            End If
            .CodExp = Me.Proyecto.Proyecto
        End If
        If .Titulo = "" Then
            .Titulo = "Informe de seguimiento de Riesgos"
        End If
        If .Tipo = "" Then
            .Tipo = "IS"
        End If
        If .Area = "" Then
            .Area = "E"
        End If
        If .FechaAlta = "" Then
            .FechaAlta = CStr(m_FechaRef)
        End If
        If .Clasificacion = "" Then
            .Clasificacion = "SINCLAS"
        End If
        If .CadenaNombreArchivos = "" Then
            p_Error = "Falta el adjunto"
            Err.Raise 1000
        End If
        ' issue-58: NO sobrescribir URLCarpetaAGEDO aqui. El caller (RegistrarEnAGEDO)
        ' ya lo setea a URLDirectorioTemporalAGEDO (la carpeta 1ParaAGEDO real de produccion).
        ' Antes este override apuntaba a URLDirectorioDocumentacion (la carpeta de ANEXOS),
        ' lo que provocaba que la creacion de carpeta y la copia del archivo fueran a paths
        ' distintos, y la copia fallara silenciosamente via On Error Resume Next.
        If .codigo = "" Then
            .codigo = getCodigoDocumento(.CodExp, p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        If .Edicion = "" Then
            If Me.EdicionObj Is Nothing Then
                p_Error = "No se conoce la edición"
                Err.Raise 1000
            End If
            .Edicion = Me.EdicionObj.Edicion
        End If
        If .Versionable = "" Then
            .Versionable = "Sí"
        End If
        If fso.FileExists(p_URLArchivoLocal) Then
            If FicheroAbierto(p_URLArchivoLocal) Then
                p_Error = "No se ha podido grabar el documento en AGEDO al estar abierto el archivo"
                Err.Raise 1000
            End If
        End If
        m_URLFinal = Me.URLCarpetaAGEDO & Me.CadenaNombreArchivos
        If fso.FileExists(m_URLFinal) Then
            If FicheroAbierto(m_URLFinal) Then
                p_Error = "No se ha podido grabar el documento en AGEDO al estar abierto el archivo"
                Err.Raise 1000
            End If
        End If
    End With
    
    Exit Function
     
errores:
     If Err.Number <> 1000 Then
         p_Error = "El método MotivoNoOK ha devuelto el error: " & Err.Description
     End If
                                
End Function
Public Function Registrar( _
                            p_URLArchivoLocal As String, _
                            Optional ByRef p_Error As String _
                            ) As String
                                        
    Dim rcdDatos As DAO.Recordset
     Dim m_SQL As String
    Dim m_MotivoNoOK As String
    Dim m_FechaRef As Date
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    m_FechaRef = Date
    'm_fechaRef = "12/12/2023"
    If Not fso.FileExists(p_URLArchivoLocal) Then
        p_Error = "No se puede alcanzar el archivo"
        Err.Raise 1000
    End If
    m_MotivoNoOK = MotivoNoOK(p_URLArchivoLocal, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    p_Error = m_MotivoNoOK
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    m_SQL = "TbDocumentos"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        .AddNew
            .Fields("IDDocumento") = Me.IDDocumento
            .Fields("FechaAlta") = Me.FechaAlta
            .Fields("CodExp") = Me.CodExp
            .Fields("Codigo") = Me.codigo
            
            .Fields("Edicion") = Me.Edicion
            .Fields("TIpo") = Me.Tipo
            .Fields("Area") = Me.Area
            .Fields("Archivo") = Me.Archivo
            .Fields("Versionable") = Me.Versionable
            .Fields("Titulo") = Me.Titulo
            .Fields("FEdicion") = CStr(m_FechaRef)
            .Fields("FENTRADAENVIGOR") = CStr(m_FechaRef)
            .Fields("Clasificacion") = Me.Clasificacion
            .Fields("CadenaNombreArchivos") = Me.CadenaNombreArchivos
            ' issue-58: usar Me.URLCarpetaAGEDO (que el caller setea a la carpeta AGEDO real)
            ' en vez de la referencia hardcoded a URLDirectorioTemporalAGEDO. Asi el INSERT
            ' en TbDocumentos queda consistente con la copia del archivo.
            .Fields("URLCarpetaAGEDO") = Me.URLCarpetaAGEDO
            .Fields("URLCarpetaArchivosLocales") = Me.URLCarpetaArchivosLocales
            
        .Update
    End With
    If Not fso.FolderExists(Me.URLCarpetaAGEDO) Then
        fso.CreateFolder Me.URLCarpetaAGEDO
    End If
    ' issue-58: copiar a URLCarpetaAGEDO (que el caller setea a URLDirectorioTemporalAGEDO,
    ' la carpeta 1ParaAGEDO real de produccion) y verificar post-copia. Antes se copiaba
    ' a URLDirectorioTemporalAGEDO directo pero la carpeta se creaba en URLDirectorioDocumentacion
    ' (ANEXOS) por el override de MotivoNoOK, y el 'On Error Resume Next' silenciaba cualquier
    ' fallo de copia, por lo que el sistema seguia como si el documento estuviera en AGEDO
    ' cuando en realidad nunca llego alli.
    Dim sURLDestino As String
    sURLDestino = Me.URLCarpetaAGEDO & fso.GetFileName(p_URLArchivoLocal)
    fso.CopyFile p_URLArchivoLocal, sURLDestino, True
    If Not fso.FileExists(sURLDestino) Then
        p_Error = "No se ha podido copiar el documento a la carpeta AGEDO '" & sURLDestino & "'"
        Err.Raise 1000
    End If

    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Documento.Registrar ha devuelto el error: " & Err.Description
    End If
End Function

Public Function Borrar( _
                        Optional ByRef p_Error As String _
                        ) As String
                                     
    Dim m_SQL As String
    On Error GoTo errores
    
    If fso.FileExists(Me.URLAdjunto) Then
        If FicheroAbierto(Me.URLAdjunto) Then
            p_Error = "Se ha de cerrar primero el documento"
            Err.Raise 1000
        End If
    End If
    m_SQL = "DELETE * FROM TbDocumentos " & _
            "WHERE IDDocumento=" & Me.IDDocumento & ";"
    getdb.Execute m_SQL
    If fso.FileExists(Me.URLAdjunto) Then
        fso.DeleteFile Me.URLAdjunto, True
    End If
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método deleteRegistro ha devuelto el error: " & Err.Description
    End If
End Function

Private Function getCodigoDocumento( _
                                    p_CodExp As String, _
                                    Optional ByRef p_Error As String _
                                    ) As String
                                     
    Dim m_Tipo As String
    Dim m_Area As String
    Dim m_Centro As String
    Dim m_PrimerParteExp As String
    Dim m_Año As String
    Dim m_ParteVariable As String
    Dim m_ObjDocumento As Documento
    Dim m_ObjColDocumentos As Scripting.Dictionary
    Dim m_IDDocumento As Variant
    Dim m_EdicionMaxima As Integer
    Dim m_CodigoSinEdicion As String
    Dim dato
    On Error GoTo errores
    
    m_Tipo = "IS"
    m_Area = "E"
    m_Centro = "CCC"
    If InStr(1, p_CodExp, "/") = 0 Then
        p_Error = "El expediente ha de tener la forma ????/??"
        Err.Raise 1000
    End If
    dato = Split(p_CodExp, "/")
    m_PrimerParteExp = CStr(dato(0))
    m_Año = CStr(dato(1))
    m_ParteVariable = m_Año & m_PrimerParteExp & m_Centro
    m_CodigoSinEdicion = m_Tipo & m_ParteVariable & m_Area & "NN"
    Set m_ObjDocumento = getDocumentoByCodigo(m_CodigoSinEdicion, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_ObjColDocumentos Is Nothing Then
        For Each m_IDDocumento In m_ObjColDocumentos.keys
            If IsObject(m_ObjColDocumentos(m_IDDocumento)) Then
                Set m_ObjDocumento = m_ObjColDocumentos(m_IDDocumento)
                If IsNumeric(m_ObjDocumento.Edicion) Then
                    If CInt(m_ObjDocumento.Edicion) > m_EdicionMaxima Then
                        m_EdicionMaxima = CInt(m_ObjDocumento.Edicion)
                    End If
                End If
                Set m_ObjDocumento = Nothing
            End If
        Next
    End If
    getCodigoDocumento = m_CodigoSinEdicion & Format(m_EdicionMaxima + 1, "00")
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getCodigoDocumento ha devuelto el error: " & Err.Description
    End If
    
                                        
End Function

Public Property Get ColCampos() As Collection

    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    If Not m_objColCampos Is Nothing Then
        Set ColCampos = m_objColCampos
        Exit Property
    End If
    Set m_objColCampos = New Collection
    With m_objColCampos
        .Add "IDDocumento"
        .Add "FechaAlta"
        .Add "CodExp"
        .Add "Codigo"
        .Add "Titulo"
        .Add "Tipo"
        .Add "Area"
        .Add "Edicion"
        .Add "FEdicion"
        .Add "FENTRADAENVIGOR"
        .Add "Archivo"
        .Add "ResponsableEdicion"
        .Add "URLCarpetaArchivosLocales"
        .Add "CadenaNombreArchivos"
        .Add "URLCarpetaAGEDO"
        .Add "Centro"
        .Add "Clasificacion"
        .Add "Registrador"
        .Add "RegistroEntrada"
        .Add "RegistroSalida"
        .Add "Versionable"
        .Add "UltimaVersion"
        .Add "Estado"
        .Add "FechaCaducidad"
        .Add "CLASE"
        .Add "Observaciones"
        .Add "FechaDeComienzoNuevaVersion"
        .Add "IDDocumentoNuevaVersion"
        .Add "IDDocumentoRevision"
        .Add "NCPARAINDICADOR"
        .Add "NCABIERTACERRADA"
        .Add "AdjuntoEncriptado"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Documento.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "IDDocumento" Then
        getPropiedad = Me.IDDocumento
        Exit Function
    End If
    If m_NombrePropiedad = "FechaAlta" Then
        getPropiedad = Me.FechaAlta
        Exit Function
    End If
    If m_NombrePropiedad = "CodExp" Then
        getPropiedad = Me.CodExp
        Exit Function
    End If
    If m_NombrePropiedad = "Codigo" Then
        getPropiedad = Me.codigo
        Exit Function
    End If
    If m_NombrePropiedad = "Titulo" Then
        getPropiedad = Me.Titulo
        Exit Function
    End If
    If m_NombrePropiedad = "Tipo" Then
        getPropiedad = Me.Tipo
        Exit Function
    End If
    If m_NombrePropiedad = "Area" Then
        getPropiedad = Me.Area
        Exit Function
    End If
    If m_NombrePropiedad = "Edicion" Then
        getPropiedad = Me.Edicion
        Exit Function
    End If
    If m_NombrePropiedad = "FEdicion" Then
        getPropiedad = Me.FEdicion
        Exit Function
    End If
    If m_NombrePropiedad = "FENTRADAENVIGOR" Then
        getPropiedad = Me.FENTRADAENVIGOR
        Exit Function
    End If
    If m_NombrePropiedad = "Archivo" Then
        getPropiedad = Me.Archivo
        Exit Function
    End If
    If m_NombrePropiedad = "ResponsableEdicion" Then
        getPropiedad = Me.ResponsableEdicion
        Exit Function
    End If
    If m_NombrePropiedad = "URLCarpetaArchivosLocales" Then
        getPropiedad = Me.URLCarpetaArchivosLocales
        Exit Function
    End If
    If m_NombrePropiedad = "CadenaNombreArchivos" Then
        getPropiedad = Me.CadenaNombreArchivos
        Exit Function
    End If
    If m_NombrePropiedad = "URLCarpetaAGEDO" Then
        getPropiedad = Me.URLCarpetaAGEDO
        Exit Function
    End If
    If m_NombrePropiedad = "Centro" Then
        getPropiedad = Me.CENTRO
        Exit Function
    End If
    If m_NombrePropiedad = "Clasificacion" Then
        getPropiedad = Me.Clasificacion
        Exit Function
    End If
    If m_NombrePropiedad = "Registrador" Then
        getPropiedad = Me.Registrador
        Exit Function
    End If
    If m_NombrePropiedad = "RegistroEntrada" Then
        getPropiedad = Me.RegistroEntrada
        Exit Function
    End If
    If m_NombrePropiedad = "RegistroSalida" Then
        getPropiedad = Me.RegistroSalida
        Exit Function
    End If
    If m_NombrePropiedad = "Versionable" Then
        getPropiedad = Me.Versionable
        Exit Function
    End If
    If m_NombrePropiedad = "UltimaVersion" Then
        getPropiedad = Me.UltimaVersion
        Exit Function
    End If
    If m_NombrePropiedad = "Estado" Then
        getPropiedad = Me.Estado
        Exit Function
    End If
    If m_NombrePropiedad = "FechaCaducidad" Then
        getPropiedad = Me.FechaCaducidad
        Exit Function
    End If
    If m_NombrePropiedad = "CLASE" Then
        getPropiedad = Me.CLASE
        Exit Function
    End If
    If m_NombrePropiedad = "Observaciones" Then
        getPropiedad = Me.Observaciones
        Exit Function
    End If
    If m_NombrePropiedad = "FechaDeComienzoNuevaVersion" Then
        getPropiedad = Me.FechaDeComienzoNuevaVersion
        Exit Function
    End If
    If m_NombrePropiedad = "IDDocumentoNuevaVersion" Then
        getPropiedad = Me.IDDocumentoNuevaVersion
        Exit Function
    End If
    If m_NombrePropiedad = "IDDocumentoRevision" Then
        getPropiedad = Me.IDDocumentoRevision
        Exit Function
    End If
    If m_NombrePropiedad = "NCPARAINDICADOR" Then
        getPropiedad = Me.NCPARAINDICADOR
        Exit Function
    End If
    If m_NombrePropiedad = "NCABIERTACERRADA" Then
        getPropiedad = Me.NCABIERTACERRADA
        Exit Function
    End If
    If m_NombrePropiedad = "AdjuntoEncriptado" Then
        getPropiedad = Me.AdjuntoEncriptado
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en Documento.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Documento.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function

Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "IDDocumento" Then
        Me.IDDocumento = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaAlta" Then
        Me.FechaAlta = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CodExp" Then
        Me.CodExp = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Codigo" Then
        Me.codigo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Titulo" Then
        Me.Titulo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Tipo" Then
        Me.Tipo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Area" Then
        Me.Area = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Edicion" Then
        Me.Edicion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FEdicion" Then
        Me.FEdicion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FENTRADAENVIGOR" Then
        Me.FENTRADAENVIGOR = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Archivo" Then
        Me.Archivo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "ResponsableEdicion" Then
        Me.ResponsableEdicion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "URLCarpetaArchivosLocales" Then
        Me.URLCarpetaArchivosLocales = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CadenaNombreArchivos" Then
        Me.CadenaNombreArchivos = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "URLCarpetaAGEDO" Then
        Me.URLCarpetaAGEDO = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Centro" Then
        Me.CENTRO = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Clasificacion" Then
        Me.Clasificacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Registrador" Then
        Me.Registrador = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "RegistroEntrada" Then
        Me.RegistroEntrada = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "RegistroSalida" Then
        Me.RegistroSalida = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Versionable" Then
        Me.Versionable = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "UltimaVersion" Then
        Me.UltimaVersion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Estado" Then
        Me.Estado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaCaducidad" Then
        Me.FechaCaducidad = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CLASE" Then
        Me.CLASE = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Observaciones" Then
        Me.Observaciones = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaDeComienzoNuevaVersion" Then
        Me.FechaDeComienzoNuevaVersion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDDocumentoNuevaVersion" Then
        Me.IDDocumentoNuevaVersion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDDocumentoRevision" Then
        Me.IDDocumentoRevision = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NCPARAINDICADOR" Then
        Me.NCPARAINDICADOR = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NCABIERTACERRADA" Then
        Me.NCABIERTACERRADA = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "AdjuntoEncriptado" Then
        Me.AdjuntoEncriptado = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en Documento.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Documento.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


