VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Anexo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private fso As New FileSystemObject



Public IDAnexo As String
Public IDEVENTO As String
Public IDActividad As String
Public IDMaterial As String
Public IDMaterialSeguimiento As String
Public IDSubContratacion As String
Public IDGasto As String
Public Titulo As String
Public Descripcion As String
Public NombreArchivo As String

Private m_sIDAnexoCalculado As String

Private m_sURLAnexo As String
Private m_ObjEvento As Evento
Private m_ObjActividad As Actividad
Private m_ObjMaterial As Material
Private m_ObjMaterialSeg As MaterialSeguimiento
Private m_ObjSubContratacion As Subcontratacion
Private m_ObjGasto As Gasto

Private m_eCerrado As EnumSino
Private m_eTipoAnexo As EnumTipoAnexo
Private m_objColCampos As Collection
Public Error As String
Private m_Error As String

Public Property Get TipoAnexo() As EnumTipoAnexo
    
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_eTipoAnexo <> 0 Then
        TipoAnexo = m_eTipoAnexo
        Exit Property
    End If
    If Me.IDEVENTO <> "" Then
        m_eTipoAnexo = EnumTipoAnexo.Evento
    ElseIf Me.IDActividad <> "" Then
        m_eTipoAnexo = EnumTipoAnexo.Actividad
    ElseIf Me.IDMaterial <> "" Then
        m_eTipoAnexo = EnumTipoAnexo.Material
    ElseIf Me.IDMaterialSeguimiento <> "" Then
        m_eTipoAnexo = EnumTipoAnexo.MaterialSeg
    ElseIf Me.IDGasto <> "" Then
        m_eTipoAnexo = EnumTipoAnexo.Gasto
    ElseIf Me.IDSubContratacion <> "" Then
        m_eTipoAnexo = EnumTipoAnexo.Subcontratacion
    End If
    TipoAnexo = m_eTipoAnexo
    
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Anexo.TipoAnexo ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Property Get IDAnexoCalculado() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    
    m_sIDAnexoCalculado = DameID("TbAnexos", "IDAnexo")
    IDAnexoCalculado = m_sIDAnexoCalculado
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Anexo.IDAnexoCalculado ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Property Get Cerrado() As EnumSino
    
    Dim m_URL As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    If m_eCerrado <> 0 Then
        Cerrado = m_eCerrado
        Exit Property
    End If
    m_URL = Me.URLAnexo
    m_Error = Me.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If fso.FileExists(m_URL) Then
        If FicheroAbierto(m_URL) Then
            m_eCerrado = EnumSino.No
        Else
            m_eCerrado = EnumSino.Sí
        End If
    Else
        m_eCerrado = EnumSino.Sí
    End If
    Cerrado = m_eCerrado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Anexo.Cerrado ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Get URLAnexo() As String
    
    Dim m_URL As String
    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_sURLAnexo <> "" Then
        URLAnexo = m_sURLAnexo
        Exit Property
    End If
    If Me.NombreArchivo = "" Then
        Exit Property
    End If
    If Me.IDGasto <> "" Then
        m_URL = m_ObjEntorno.URLDirectorioGastos
    Else
        m_URL = m_ObjEntorno.URLDirectorioDocumentacion
    End If
    If m_ObjEntorno.URLDirectorioAnexos = "" Then
        Exit Property
    End If
    
    m_sURLAnexo = m_URL & Me.NombreArchivo
    
    URLAnexo = m_sURLAnexo
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Anexo.URLAnexo ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Property Let URLAnexo(ByVal sNewValue As String)

    m_sURLAnexo = sNewValue

End Property

Public Property Get Evento() As Evento

    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_ObjEvento Is Nothing Then
        Set Evento = m_ObjEvento
        Exit Property
    End If
    If Me.IDEVENTO = "" Then
        Exit Property
    End If
    m_ObjEvento = Constructor.getEvento(Me.IDEVENTO, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Evento = m_ObjEvento
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Anexo.Evento ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error

End Property

Public Property Set Evento(ByVal objNewValue As Evento)

    Set m_ObjEvento = objNewValue

End Property

Public Property Get Actividad() As Actividad

    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_ObjActividad Is Nothing Then
        Set Actividad = m_ObjActividad
        Exit Property
    End If
    If Me.IDActividad = "" Then
        Exit Property
    End If
    m_ObjActividad = Constructor.getActividad(Me.IDActividad, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Actividad = m_ObjActividad
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Anexo.Actividad ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error

End Property

Public Property Set Actividad(ByVal objNewValue As Actividad)

    Set m_ObjActividad = objNewValue

End Property

Public Property Get Material() As Material

    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_ObjMaterial Is Nothing Then
        Set Material = m_ObjMaterial
        Exit Property
    End If
    If Me.IDMaterial = "" Then
        Exit Property
    End If
    m_ObjMaterial = Constructor.getMaterial(Me.IDMaterial, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Material = m_ObjMaterial
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Anexo.Material ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error

End Property

Public Property Set Material(ByVal objNewValue As Material)

    Set m_ObjMaterial = objNewValue

End Property

Public Property Get Subcontratacion() As Subcontratacion

    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_ObjSubContratacion Is Nothing Then
        Set Subcontratacion = m_ObjSubContratacion
        Exit Property
    End If
    If Me.IDSubContratacion = "" Then
        Exit Property
    End If
    m_ObjSubContratacion = Constructor.getSubContratacion(Me.IDSubContratacion, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Subcontratacion = m_ObjSubContratacion
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Anexo.SubContratacion ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error

End Property

Public Property Set Subcontratacion(ByVal objNewValue As Subcontratacion)

    Set m_ObjSubContratacion = objNewValue

End Property

Public Property Get Gasto() As Gasto

    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_ObjGasto Is Nothing Then
        Set Gasto = m_ObjGasto
        Exit Property
    End If
    If Me.IDGasto = "" Then
        Exit Property
    End If
    m_ObjGasto = Constructor.getGasto(Me.IDGasto, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Gasto = m_ObjGasto
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Anexo.Gasto ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error

End Property

Public Property Set Gasto(ByVal objNewValue As Gasto)

    Set m_ObjGasto = objNewValue

End Property

Public Property Get MaterialSeg() As MaterialSeguimiento

    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_ObjMaterialSeg Is Nothing Then
        Set MaterialSeg = m_ObjMaterialSeg
        Exit Property
    End If
    If Me.IDMaterialSeguimiento = "" Then
        Exit Property
    End If
    m_ObjMaterialSeg = Constructor.getMaterialSeguimiento(Me.IDMaterialSeguimiento, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set MaterialSeg = m_ObjMaterialSeg
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Anexo.MaterialSeg ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error

End Property

Public Property Set MaterialSeg(ByVal objNewValue As MaterialSeguimiento)

    Set m_ObjMaterialSeg = objNewValue

End Property



Public Function GetURLAnexoCalculado( _
                                    p_URLLocal As String, _
                                    Optional ByRef p_Error As String _
                                    ) As String
    
    Dim m_URLDir As String
    Dim m_NumeroAnexosPorClase As String
    Dim m_TipoAnexo As EnumTipoAnexo
    Dim m_OrdinalAnexofuturo As String
    Dim m_ExtensionAnexo As String
    Dim m_NombreArchivoFinalSinRuta As String
    Dim m_NombreCampoID As String
    Dim m_ID As String
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    m_NombreCampoID = getNombreCampoID(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_ID = getPropiedad(m_NombreCampoID, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_ID = "" Then
        p_Error = "NO se ha podido obtener el id de la entidad del anexo"
        Err.Raise 1000
    End If
    m_NumeroAnexosPorClase = DameNumeroAnexosPorClase(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_TipoAnexo = Me.TipoAnexo
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_TipoAnexo = EnumTipoAnexo.Gasto Then
        m_URLDir = m_ObjEntorno.URLDirectorioGastos
    Else
        m_URLDir = m_ObjEntorno.URLDirectorioDocumentacion
    End If
    m_OrdinalAnexofuturo = Format(CLng(m_NumeroAnexosPorClase) + 1, "00")
    m_ExtensionAnexo = fso.GetExtensionName(p_URLLocal)
    m_NombreArchivoFinalSinRuta = m_ID & "_" & m_OrdinalAnexofuturo & "." & m_ExtensionAnexo
    m_NombreArchivoFinalSinRuta = Replace(m_NombreArchivoFinalSinRuta, "/", "-")
    GetURLAnexoCalculado = m_URLDir & m_NombreArchivoFinalSinRuta
    
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Anexo.URLAnexoCalculado ha devuelto el error:" & vbNewLine & Err.Description
    End If
    p_Error = Me.Error
End Function

Private Function getID( _
                        Optional ByRef p_Error As String _
                        ) As String
    
    Dim m_TipoAnexo As EnumTipoAnexo
    
    On Error GoTo errores
    
    m_TipoAnexo = Me.TipoAnexo
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    
    If m_TipoAnexo = EnumTipoAnexo.Evento Then
        getID = Me.IDEVENTO
        
    ElseIf m_TipoAnexo = EnumTipoAnexo.Actividad Then
       getID = Me.IDActividad
    ElseIf m_TipoAnexo = EnumTipoAnexo.Material Then
        getID = Me.IDMaterial
    ElseIf m_TipoAnexo = EnumTipoAnexo.MaterialSeg Then
        getID = Me.IDMaterialSeguimiento
    ElseIf m_TipoAnexo = EnumTipoAnexo.Subcontratacion Then
        getID = Me.IDSubContratacion
    ElseIf m_TipoAnexo = EnumTipoAnexo.Gasto Then
       getID = Me.IDGasto
    Else
        p_Error = "Tipo Anexo no reconocido"
        Err.Raise 1000
    End If
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Anexo.getID ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function
Private Function getNombreCampoID( _
                                    Optional ByRef p_Error As String _
                                    ) As String
                
    Dim m_TipoAnexo As EnumTipoAnexo
    
    On Error GoTo errores
    
    m_TipoAnexo = Me.TipoAnexo
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    
    If m_TipoAnexo = EnumTipoAnexo.Evento Then
        getNombreCampoID = "IDEVENTO"
        
    ElseIf m_TipoAnexo = EnumTipoAnexo.Actividad Then
       getNombreCampoID = "IDActividad"
    ElseIf m_TipoAnexo = EnumTipoAnexo.Material Then
        getNombreCampoID = "IDMaterial"
    ElseIf m_TipoAnexo = EnumTipoAnexo.MaterialSeg Then
        getNombreCampoID = "IDMaterialSeguimiento"
    ElseIf m_TipoAnexo = EnumTipoAnexo.Subcontratacion Then
        getNombreCampoID = "IDSubContratacion"
    ElseIf m_TipoAnexo = EnumTipoAnexo.Gasto Then
       getNombreCampoID = "IDGasto"
    Else
        p_Error = "Tipo Anexo no reconocido"
        Err.Raise 1000
    End If
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Anexo.getNombreCampoID ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function
Private Function CampoIDEsNumerico( _
                                    Optional ByRef p_Error As String _
                                    ) As Boolean
    
    
    Dim m_TipoAnexo As EnumTipoAnexo
    On Error GoTo errores
    
    
        
    m_TipoAnexo = Me.TipoAnexo
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_TipoAnexo = EnumTipoAnexo.Subcontratacion Then
        CampoIDEsNumerico = True
    ElseIf m_TipoAnexo = EnumTipoAnexo.Gasto Then
       CampoIDEsNumerico = True
    End If
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Anexo.CampoIDEsNumerico ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function
Private Function DameNumeroAnexosPorClase( _
                                            Optional ByRef p_Error As String _
                                            ) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim intNumero As Long
    Dim m_NombreCampoID As String
    Dim m_CampoIDEsNumerico As Boolean
    Dim m_ID As String
    On Error GoTo errores
    
     
    m_ID = getID(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_CampoIDEsNumerico = CampoIDEsNumerico(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_NombreCampoID = getNombreCampoID(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_CampoIDEsNumerico = False Then
        m_SQL = "SELECT Count(TbAnexos.IDAnexo) AS CuentaDeIDAnexo " & _
                "FROM TbAnexos " & _
                "WHERE " & m_NombreCampoID & "='" & m_ID & "';"
    Else
        m_SQL = "SELECT Count(TbAnexos.IDAnexo) AS CuentaDeIDAnexo " & _
                "FROM TbAnexos " & _
                "WHERE " & m_NombreCampoID & "=" & m_ID & ";"
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            If IsNumeric(Nz(.Fields("CuentaDeIDAnexo"), "")) Then
                intNumero = .Fields("CuentaDeIDAnexo")
            End If
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    DameNumeroAnexosPorClase = CStr(intNumero)
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Anexo.DameNumeroAnexosPorClase ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function

Public Function MotivoNoOK( _
                            p_URLLocal As String, _
                            Optional ByRef p_URLFinal As String, _
                            Optional ByRef p_Error As String _
                            ) As String
    
    
    Dim m_ID As String
    Dim m_TituloRepetido As EnumSino
    Dim m_TipoAnexo As EnumTipoAnexo
    
    On Error GoTo errores
    Me.Error = ""
    m_Error = ""
    p_Error = ""
    With Me
        If Not fso.FileExists(p_URLLocal) Then
            MotivoNoOK = "El anexo local no se encuentra"
            Exit Function
        End If
        If FicheroAbierto(p_URLLocal) Then
            MotivoNoOK = "El anexo local se ha de cerrar"
            Exit Function
        End If
        m_TipoAnexo = .TipoAnexo
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If m_TipoAnexo < 1 Or m_TipoAnexo > 6 Then
            MotivoNoOK = "Tipo Anexo no reconocido"
            Exit Function
        End If

        m_ID = getID(p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If m_ID = "" Then
            MotivoNoOK = "NO se sabe a qué entidad se le ha de anexar el documento"
            Exit Function
        End If
        If Me.IDAnexo = "" Then
            Me.IDAnexo = Me.IDAnexoCalculado
            p_Error = Me.Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If Me.IDAnexo = "" Then
                MotivoNoOK = "NO se ha podido obtener el IDAnexo"
                Exit Function
            End If
        End If
        p_URLFinal = GetURLAnexoCalculado(p_URLLocal)
        If fso.FileExists(p_URLFinal) Then
            If FicheroAbierto(p_URLFinal) Then
                MotivoNoOK = "Existe un fichero con el mismo nombre en el servidor y parece estar abierto"
                Exit Function
            End If
        End If
        If Me.Titulo = "" Then
            MotivoNoOK = "Es Obligatorio el título"
            Exit Function
        End If
        m_TituloRepetido = TituloRepetido(p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If m_TituloRepetido = EnumSino.Sí Then
            MotivoNoOK = "El título ya ha sido usado para esa entidad"
            Exit Function
        End If
    End With
    
    Exit Function
     
errores:
     If Err.Number <> 1000 Then
         p_Error = "El método MotivoNoOK ha devuelto el error: " & Err.Description
     End If
                                
End Function

Public Function Registrar( _
                            ByVal p_URLLocal As String, _
                            Optional ByRef p_Error As String _
                            ) As String
    
    Dim rcdDatos As DAO.Recordset
    
    Dim lngNumeroAnexosPorClase As Long
    Dim m_URLFinal As String
    Dim m_IDAnexo As String
    Dim m_MotivoNoOK As String
    
    On Error GoTo errores
   
    m_MotivoNoOK = MotivoNoOK(p_URLLocal, m_URLFinal, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    p_Error = m_MotivoNoOK
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    m_SQL = "TbAnexos"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        
        .AddNew
            .Fields("IDAnexo") = Me.IDAnexoCalculado
            .Fields("Titulo") = Me.Titulo
            .Fields("NombreArchivo") = fso.GetFileName(m_URLFinal)
            If Me.IDEVENTO <> "" Then
                .Fields("IDEvento") = Me.IDEVENTO
            End If
            If Me.IDActividad <> "" Then
                .Fields("IDActividad") = Me.IDActividad
            End If
            If Me.IDMaterial <> "" Then
                .Fields("IDMaterial") = Me.IDMaterial
            End If
            If Me.IDMaterialSeguimiento <> "" Then
                .Fields("IDMaterialSeguimiento") = Me.IDMaterialSeguimiento
            End If
            If Me.IDSubContratacion <> "" Then
                .Fields("IDSubContratacion") = Me.IDSubContratacion
            End If
            If Me.IDGasto <> "" Then
                .Fields("IDGasto") = Me.IDGasto
            End If
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    
    fso.CopyFile p_URLLocal, m_URLFinal, True
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Anexo.Registrar ha devuelto un error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function
Public Function RegistrarAutorizacionMaterial( _
                                                ByVal p_URLLocal As String, _
                                                Optional ByRef p_Error As String _
                                                ) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_AnexoAutorizacionPrevio As Anexo
    Dim m_URLFinal As String
    
    
    
    On Error GoTo errores
   
    If Me.IDMaterial = "" Then
        p_Error = "No se ha determinado el IDMaterial"
        Err.Raise 1000
    End If
    
    
    Set m_AnexoAutorizacionPrevio = Constructor.getAnexoAutorizacionDeMaterial(Me.IDMaterial, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_AnexoAutorizacionPrevio Is Nothing Then
        m_Linea = "Obteniendo la lista de Materiales............ Anexo NO existente"
        Avance m_Linea
        m_URLFinal = GetURLAnexoCalculado(p_URLLocal)
        m_SQL = "TbAnexos"
        Set rcdDatos = getdb().OpenRecordset(m_SQL)
        With rcdDatos
            
            .AddNew
                Me.IDAnexo = Me.IDAnexoCalculado
                .Fields("IDAnexo") = Me.IDAnexo
                Me.Titulo = "Autorización Materiales"
                .Fields("Titulo") = Me.Titulo
                .Fields("NombreArchivo") = fso.GetFileName(m_URLFinal)
                .Fields("IDMaterial") = Me.IDMaterial
            .Update
        End With
        rcdDatos.Close
        Set rcdDatos = Nothing
        
    Else
        m_Linea = "Obteniendo la lista de Materiales............ Anexo ya existente"
        Avance m_Linea
        m_URLFinal = m_AnexoAutorizacionPrevio.URLAnexo
        
        
    End If
    fso.CopyFile p_URLLocal, m_URLFinal, True
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Anexo.RegistrarAutorizacionMaterial ha devuelto un error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function
Public Function Resubir( _
                        ByVal p_URLLocal As String, _
                        Optional ByRef p_Error As String _
                        ) As String
    
    
    Dim m_URLFinal As String
    Dim m_SQL As String
    Dim m_NombreFinal As String
    
    On Error GoTo errores
   
    m_URLFinal = Me.URLAnexo
    
    m_NombreFinal = fso.GetBaseName(m_URLFinal) & "." & fso.GetExtensionName(p_URLLocal)
    m_URLFinal = fso.GetParentFolderName(m_URLFinal) & "\" & m_NombreFinal
    If fso.FileExists(m_URLFinal) Then
        If FicheroAbierto(m_URLFinal) Then
            p_Error = "El fichero existe y está abierto"
            Err.Raise 1000
        End If
    End If
    If m_NombreFinal <> Me.NombreArchivo Then
        m_SQL = "UPDATE TbAnexos SET NombreArchivo = '" & m_NombreFinal & "' " & _
                "WHERE IDAnexo=" & Me.IDAnexo & ";"
        getdb().Execute m_SQL
    End If
    
    
    fso.CopyFile p_URLLocal, m_URLFinal, True
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Anexo.Resubir ha devuelto un error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function
Private Function TituloRepetido( _
                                Optional ByRef p_Error As String _
                                ) As EnumSino
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_NombreCampoID As String
    Dim m_valorCampoID As String
    
    On Error GoTo errores
    
   
    If Me.Titulo = "" Then
        p_Error = "Se ha de indicar el Título a evaluar"
        Err.Raise 1000
    End If
    m_NombreCampoID = getNombreCampoID(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_valorCampoID = Me.getPropiedad(m_NombreCampoID, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_SQL = "SELECT TbAnexos.Titulo " & _
                "FROM TbAnexos " & _
                "WHERE Titulo='" & Me.Titulo & "' AND " & m_NombreCampoID & "='" & m_valorCampoID & "';"
    On Error Resume Next
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo errores
        m_SQL = "SELECT TbAnexos.Titulo " & _
                "FROM TbAnexos " & _
                "WHERE Titulo='" & Me.Titulo & "' AND " & m_NombreCampoID & "=" & m_valorCampoID & ";"
        Set rcdDatos = getdb().OpenRecordset(m_SQL)
    End If
    With rcdDatos
        If .EOF Then
            TituloRepetido = EnumSino.No
        Else
            TituloRepetido = EnumSino.Sí
        End If
    End With
    
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Anexo.TituloRepetido ha producido el error nº: " & Err.Number & vbCrLf & "Detalle: " & Err.Description
    End If
    
End Function

Public Function Eliminar(Optional ByRef p_Error As String) As String
    
    Dim m_SQL As String
    On Error GoTo errores
    
    If Me.IDAnexo = "" Then
        p_Error = "Se ha de indicar el IDAnexo"
        Err.Raise 1000
    End If
    If fso.FileExists(Me.URLAnexo) Then
        If FicheroAbierto(Me.URLAnexo) Then
            p_Error = "Tiene el anexo abierto"
            Err.Raise 1000
        End If
    End If
    m_SQL = "DELETE * FROM TbAnexos " & _
            "WHERE IDAnexo=" & Me.IDAnexo & ";"
    getdb().Execute m_SQL
    If fso.FileExists(Me.URLAnexo) Then
        fso.DeleteFile Me.URLAnexo, True
    End If
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Anexo.Eliminar ha devuelto el error:" & vbNewLine & Err.Description
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
        .Add "IDAnexo"
        .Add "IDEvento"
        .Add "IDActividad"
        .Add "IDMaterial"
        .Add "IDMaterialSeguimiento"
        .Add "IDSubContratacion"
        .Add "IDGasto"
        .Add "Titulo"
        .Add "Descripcion"
        .Add "NombreArchivo"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Anexo.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "IDAnexo" Then
        getPropiedad = Me.IDAnexo
        Exit Function
    End If
    If m_NombrePropiedad = "IDEvento" Then
        getPropiedad = Me.IDEVENTO
        Exit Function
    End If
    If m_NombrePropiedad = "IDActividad" Then
        getPropiedad = Me.IDActividad
        Exit Function
    End If
    If m_NombrePropiedad = "IDMaterial" Then
        getPropiedad = Me.IDMaterial
        Exit Function
    End If
    If m_NombrePropiedad = "IDMaterialSeguimiento" Then
        getPropiedad = Me.IDMaterialSeguimiento
        Exit Function
    End If
    If m_NombrePropiedad = "IDSubContratacion" Then
        getPropiedad = Me.IDSubContratacion
        Exit Function
    End If
    If m_NombrePropiedad = "IDGasto" Then
        getPropiedad = Me.IDGasto
        Exit Function
    End If
    If m_NombrePropiedad = "Titulo" Then
        getPropiedad = Me.Titulo
        Exit Function
    End If
    If m_NombrePropiedad = "Descripcion" Then
        getPropiedad = Me.Descripcion
        Exit Function
    End If
    If m_NombrePropiedad = "NombreArchivo" Then
        getPropiedad = Me.NombreArchivo
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en Anexo.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Anexo.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function



Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "IDAnexo" Then
        Me.IDAnexo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDEvento" Then
        Me.IDEVENTO = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDActividad" Then
        Me.IDActividad = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDMaterial" Then
        Me.IDMaterial = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDMaterialSeguimiento" Then
        Me.IDMaterialSeguimiento = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDSubContratacion" Then
        Me.IDSubContratacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDGasto" Then
        Me.IDGasto = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Titulo" Then
        Me.Titulo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Descripcion" Then
        Me.Descripcion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NombreArchivo" Then
        Me.NombreArchivo = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en Anexo.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Anexo.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


