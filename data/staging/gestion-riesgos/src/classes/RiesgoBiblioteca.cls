VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RiesgoBiblioteca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public IDRiesgoTipo As String
Public codigo As String

Public Descripcion As String
Public Usuario As String
Public FechaAlta As String
Public FechaModificacion As String
Public Familia As String
Public Activo As String

Private m_sDRiestoTipoCalculado As String
Private m_sCodigoCalculado As String
Private m_eTipoCalculado As EnumTipoRiesgoBiblioteca

Private m_sIDRiesgoTipoCalculado As String
Private m_eRiesgoParaRetipificar As EnumSiNo
Private m_eUsado As EnumSiNo
Private m_objColCampos As Collection
Public Error As String
Private m_Error As String


Public Property Get IDRiesgoTipoCalculado() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    m_sIDRiesgoTipoCalculado = DameID("TbBibliotecaRiesgos", "IDRiesgoTipo", , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    IDRiesgoTipoCalculado = m_sIDRiesgoTipoCalculado
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.IDRiesgoTipoCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get CodigoCalculado() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    
    
    m_sCodigoCalculado = getCodigo(m_Error)
    If m_Error <> "" Then
        Err.Raise 100
    End If
    CodigoCalculado = m_sCodigoCalculado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método RiesgoBiblioteca.CodigoCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Let CodigoCalculado(ByVal sNewValue As String)

    m_sCodigoCalculado = sNewValue

End Property



Public Property Get RiesgoParaRetipificar() As EnumSiNo
    
    On Error GoTo errores
    
    Me.Error = ""
    If m_eRiesgoParaRetipificar <> Empty Then
        RiesgoParaRetipificar = m_eRiesgoParaRetipificar
        Exit Property
    End If
    If Me.codigo = m_ObjEntorno.CodigoBibliotecaParaRetipificar Then
        m_eRiesgoParaRetipificar = EnumSiNo.Sí
    Else
        m_eRiesgoParaRetipificar = EnumSiNo.No
    End If
    
    RiesgoParaRetipificar = m_eRiesgoParaRetipificar
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método RiesgoBiblioteca.RiesgoParaRetipificar ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Let RiesgoParaRetipificar(ByVal eNewValue As EnumSiNo)
    
    m_eRiesgoParaRetipificar = eNewValue

End Property

Public Property Get Usado() As EnumSiNo
    
    On Error GoTo errores
    
    Me.Error = ""
    If m_eUsado <> Empty Then
        Usado = m_eUsado
        Exit Property
    End If
    m_eUsado = RiesgoBibliotecaUsado(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    Usado = m_eUsado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método RiesgoBiblioteca.Usado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Private Function getCodigo( _
                            Optional ByRef p_Error As String _
                            ) As String
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_Particula As String
    Dim m_CodigoMaximo As String
    Dim m_IndiceMax As Integer
    
    On Error GoTo errores
    
    
    
    m_Particula = "PR"

    m_SQL = "SELECT TbBibliotecaRiesgos.CODIGO " & _
            "FROM TbBibliotecaRiesgos " & _
            "WHERE (((TbBibliotecaRiesgos.CODIGO) Like 'R-" & m_Particula & "-??')) " & _
            "ORDER BY TbBibliotecaRiesgos.CODIGO DESC;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Not .EOF Then
            m_CodigoMaximo = .Fields("CODIGO")
            If IsNumeric(Right(m_CodigoMaximo, 2)) Then
                m_IndiceMax = CInt(Right(m_CodigoMaximo, 2))
            End If
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    getCodigo = "R-" & m_Particula & "-" & Format(m_IndiceMax + 1, "00")
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CodigoExistente ha devuelto el error: " & Err.Description
    End If

End Function
Private Function MotivoNoOK( _
                                Optional p_RiesgoBibliotecaAlInicio As RiesgoBiblioteca, _
                                Optional ByRef p_Error As String _
                                ) As String
    
    
    
   
    Dim m_CampoRepetido As EnumSiNo
    Dim blnComprobarCampo As Boolean
    Dim m_AlgunRiesgoDeBibliotecaEnProyectosActivos As EnumSiNo
    Dim m_AlgunRiesgoDeBibliotecaEnProyectos As EnumSiNo
    On Error GoTo errores
    
    
    If EsTecnico = EnumSiNo.Sí Then
        MotivoNoOK = "Sólo los miembros de Calidad pueden dar de alta un Riesgo en la Biblioteca de Riesgos"
        Exit Function
    End If
    With Me
        
        If Me.Familia = "" Then
            MotivoNoOK = "Falta la Familia"
            Exit Function
        End If
        If .Activo <> "Sí" And .Activo <> "No" Then
            MotivoNoOK = "Se ha de indicar si está activo o no"
            Exit Function
        End If
        If .Descripcion = "" Then
            MotivoNoOK = "Falta la descripción"
            Exit Function
        End If
        If p_RiesgoBibliotecaAlInicio Is Nothing Then
            
            blnComprobarCampo = True
        Else
            If p_RiesgoBibliotecaAlInicio.Descripcion <> .Descripcion Then
                blnComprobarCampo = True
            End If
        End If
        If blnComprobarCampo = True Then
            m_CampoRepetido = CampoRepetido("TbBibliotecaRiesgos", "Descripcion", .Descripcion, EnumSiNo.Sí, p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If m_CampoRepetido = EnumSiNo.Sí Then
                MotivoNoOK = "Descripción Repetida"
                Exit Function
            End If
        End If
        blnComprobarCampo = False
        If Not p_RiesgoBibliotecaAlInicio Is Nothing Then
            If .Descripcion <> p_RiesgoBibliotecaAlInicio.Descripcion Then
                m_AlgunRiesgoDeBibliotecaEnProyectos = AlgunRiesgoDeBibliotecaEnProyectos(Me, p_Error)
                If p_Error <> "" Then
                    MotivoNoOK = "Este Tipo de Riesgo ha sido usado al menos en una gestión de riesgos"
                    Exit Function
                End If
            End If
            
        End If
        
        
    End With
    
   
    Exit Function
     
errores:
     If Err.Number <> 1000 Then
         p_Error = "El método MotivoNoOK ha devuelto el error: " & Err.Description
     End If
    
End Function

Private Function ValidacionEliminar( _
                                    Optional ByRef p_Error As String _
                                    ) As String
                                    
    Dim m_Usado As EnumSiNo
    On Error GoTo errores
    'no ha de haber formado parte de ninguna gestión de riesgos
    'no se puede eliminar la que tiene la frase Ninguno de los anteriores
    
    Me.Error = ""
    If Me Is Nothing Then
        p_Error = "No se ha podido obtener el Riesgo del Catálogo"
        Err.Raise 1000
    End If
    If EsTecnico = EnumSiNo.Sí Then
        p_Error = "El perfil de su usuario no tiene autorización para realizar esta operación"
        Err.Raise 1000
    End If
    If Me.RiesgoParaRetipificar = EnumSiNo.Sí Then
        p_Error = "No se puede borrar un tipo de riesgo que tiene al comienzo de su descripción Ninguno de los anteriores"
        Err.Raise 1000
    End If
    If Me.Usado = EnumSiNo.Sí Then
        p_Error = "Ya ha sido usado en al menos un riesgo"
        Err.Raise 1000
    End If
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ValidacionEliminar ha devuelto el error: " & Err.Description
    End If
                                    
End Function

Public Function Borrar( _
                        Optional ByRef p_Error As String _
                        ) As String
                                
    Dim m_SQL As String
    On Error GoTo errores
    
    ValidacionEliminar p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_SQL = "DELETE * FROM TbBibliotecaRiesgos " & _
            "WHERE IDRiesgoTipo=" & Me.IDRiesgoTipo & ";"
    getdb().Execute m_SQL
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Borrar ha devuelto el error: " & Err.Description
    End If
End Function

Public Function Registrar( _
                            Optional p_RiesgoBibliotecaAlInicio As RiesgoBiblioteca, _
                            Optional p_Comprobando As EnumSiNo = EnumSiNo.Sí, _
                            Optional ByRef p_Error As String _
                            ) As String
                            
    Dim rcdDatos As DAO.Recordset
     Dim m_SQL As String
    Dim m_MotivoNoOK As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    If p_Comprobando = EnumSiNo.Sí Then
        m_MotivoNoOK = MotivoNoOK(p_RiesgoBibliotecaAlInicio, p_Error)
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        p_Error = m_MotivoNoOK
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    
    If p_RiesgoBibliotecaAlInicio Is Nothing Then
        Me.IDRiesgoTipo = Me.IDRiesgoTipoCalculado
        Me.codigo = Me.CodigoCalculado
        If Not IsDate(Me.FechaAlta) Then
            Me.FechaAlta = Date
        End If
        m_SQL = "TbBibliotecaRiesgos"
    Else
        m_SQL = "SELECT * " & _
                "FROM TbBibliotecaRiesgos " & _
                "WHERE IDRiesgoTipo=" & Me.IDRiesgoTipo & ";"
        If Not IsDate(Me.FechaModificacion) Then
            Me.FechaModificacion = Date
        End If
    End If
    If Me.Usuario = "" Then
        Me.Usuario = m_ObjUsuarioConectado.UsuarioRed
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If p_RiesgoBibliotecaAlInicio Is Nothing Then
            .AddNew
                .Fields("IDRiesgoTipo") = Me.IDRiesgoTipo
                .Fields("Codigo") = Me.codigo
                .Fields("Familia") = Me.Familia
                .Fields("Descripcion") = Me.Descripcion
                .Fields("Usuario") = Me.Usuario
                .Fields("FechaAlta") = Me.FechaAlta
                .Fields("Activo") = Me.Activo
            .Update
        Else
            .Edit
               .Fields("Codigo") = Me.codigo
                .Fields("Familia") = Me.Familia
                .Fields("Descripcion") = Me.Descripcion
                .Fields("Usuario") = Me.Usuario
                .Fields("FechaModificacion") = Me.FechaModificacion
                .Fields("Activo") = Me.Activo
                
            .Update
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Registrar ha devuelto el error: " & Err.Description
    End If
End Function



Private Function RiesgoBibliotecaUsado( _
                                        Optional ByRef p_Error As String _
                                        ) As EnumSiNo
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    
    On Error GoTo errores
    If Me.codigo = "" Then
        p_Error = "No se sabe el código"
        Exit Function
    End If
    m_SQL = "SELECT * " & _
                "FROM TbRiesgos " & _
                "WHERE CodRiesgoBiblioteca='" & Me.codigo & "';"
   
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            RiesgoBibliotecaUsado = EnumSiNo.No
        Else
            RiesgoBibliotecaUsado = EnumSiNo.Sí
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RiesgoBibliotecaUsado ha devuelto el error: " & Err.Description
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
        .Add "IDRiesgoTipo"
        .Add "CODIGO"
        .Add "Familia"
        .Add "Descripcion"
        .Add "Usuario"
        .Add "FechaAlta"
        .Add "FechaModificacion"
        .Add "Activo"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método RiesgoBiblioteca.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "IDRiesgoTipo" Then
        getPropiedad = Me.IDRiesgoTipo
        Exit Function
    End If
    If m_NombrePropiedad = "CODIGO" Then
        getPropiedad = Me.codigo
        Exit Function
    End If
    
    If m_NombrePropiedad = "Descripcion" Then
        getPropiedad = Me.Descripcion
        Exit Function
    End If
    If m_NombrePropiedad = "Usuario" Then
        getPropiedad = Me.Usuario
        Exit Function
    End If
    If m_NombrePropiedad = "FechaAlta" Then
        getPropiedad = Me.FechaAlta
        Exit Function
    End If
    If m_NombrePropiedad = "FechaModificacion" Then
        getPropiedad = Me.FechaModificacion
        Exit Function
    End If
    If m_NombrePropiedad = "Activo" Then
        getPropiedad = Me.Activo
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en RiesgoBiblioteca.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RiesgoBiblioteca.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function

Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "IDRiesgoTipo" Then
        Me.IDRiesgoTipo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CODIGO" Then
        Me.codigo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Familia" Then
        Me.Familia = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Descripcion" Then
        Me.Descripcion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Usuario" Then
        Me.Usuario = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaAlta" Then
        Me.FechaAlta = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaModificacion" Then
        Me.FechaModificacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Activo" Then
        Me.Activo = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en RiesgoBiblioteca.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RiesgoBiblioteca.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


