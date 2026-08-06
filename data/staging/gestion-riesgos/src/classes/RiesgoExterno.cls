VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RiesgoExterno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public IDRiesgoExt As String
Public CodRiesgo As String
Public IDRiesgo As String
Public Origen As String
Public IDEdicion As String
Public Descripcion As String
Public CausaRaiz As String
Public FechaDetectado As String
Public FechaAltaRegistro As String
Public UsuarioRegistra As String
Public MotivoNoIntegrado As String
Public FechaMotivo As String
Public Trasladar As String
Public Suministrador As String
Public Pedido As String
Public ProveedorPedido As String
Public CausaRiesgoPedido As String
Public RequiereRiesgoDeBiblioteca As String
Public CodRiesgoBiblioteca As String
Public RiesgoPendienteRetipificacion As String


Private m_sIDRiesgoExtCalculado As String
Private m_sCodigoRiesgoExtCalculado As String
Private m_eRequiereRiesgoDeBibliotecaCalculado As EnumSiNo
Private m_sRequiereRiesgoDeBibliotecaCalculadoTexto As String
Private m_eRiesgoParaRetipificar As EnumSiNo
Private m_ObjRiesgoBiblioteca As RiesgoBiblioteca
Private m_ObjRiesgo As riesgo
Private m_objUsuarioRegistraObj As Usuario
Private m_objEdicion As Edicion
Private m_eUsuarioConectadoAutorizado As EnumSiNo
Private m_eTraslasdado As EnumSiNo

Private m_objColCampos As Collection
Public Error As String
Private m_Error As String

Public Property Get IDRiesgoExtCalculado() As String

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    m_sIDRiesgoExtCalculado = DameID("TbRiesgosAIntegrar", "IDRiesgoExt", , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    IDRiesgoExtCalculado = m_sIDRiesgoExtCalculado
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método RiesgoExterno.IDRiesgoExtCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get CodigoRiesgoExtCalculado() As String
    
    Dim m_IdRiesgo As Variant
    Dim m_ObjRiesgoExt As RiesgoExterno
    Dim intOrdinal As Integer
    Dim intOrdinalMax As Integer
    Dim m_CodRiesgo As String
    Dim m_ObjColRiesgosExt As Scripting.Dictionary
    Dim dato
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    
    If Me.Edicion Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_Error = "No se sabe a qué Edición pertence el Riesgo Externo"
        Err.Raise 1000
    End If
    Set Me.Edicion.ColRiesgosExternos = Nothing
    Set m_ObjColRiesgosExt = Me.Edicion.ColRiesgosExternos
    m_Error = Me.Edicion.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_ObjColRiesgosExt Is Nothing Then
        For Each m_IdRiesgo In m_ObjColRiesgosExt.keys
            If Nz(m_IdRiesgo, "") = "" Then
                GoTo SiguienteRiesgoExt
            End If
            Set m_ObjRiesgoExt = m_ObjColRiesgosExt(m_IdRiesgo)
            m_CodRiesgo = m_ObjRiesgoExt.CodRiesgo
            If InStr(1, m_CodRiesgo, "Re") <> 0 Then
                dato = Split(m_CodRiesgo, "Re")
                If IsNumeric(dato(1)) Then
                    intOrdinal = CInt(dato(1))
                    If intOrdinal > intOrdinalMax Then
                        intOrdinalMax = intOrdinal
                    End If
                End If
            End If
            Set m_ObjRiesgoExt = Nothing
SiguienteRiesgoExt:
        Next
    End If
    
    m_sCodigoRiesgoExtCalculado = "Re" & Format(intOrdinalMax + 1, "000")
    CodigoRiesgoExtCalculado = m_sCodigoRiesgoExtCalculado
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.CodigoRiesgoExtCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = ""
End Property


Public Property Get RequiereRiesgoDeBibliotecaCalculado() As EnumSiNo
    
    On Error GoTo errores
    
    m_Error = ""
    Me.Error = ""
    
    m_eRequiereRiesgoDeBibliotecaCalculado = Me.Edicion.Proyecto.RequiereRiesgoDeBibliotecaCalculado
    m_Error = Me.Edicion.Proyecto.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    RequiereRiesgoDeBibliotecaCalculado = m_eRequiereRiesgoDeBibliotecaCalculado
     Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.RequiereRiesgoDeBibliotecaCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Let RequiereRiesgoDeBibliotecaCalculado(ByVal eNewValue As EnumSiNo)

    m_eRequiereRiesgoDeBibliotecaCalculado = eNewValue

End Property

Public Property Get RequiereRiesgoDeBibliotecaCalculadoTexto() As String
    
    On Error GoTo errores
    
    m_Error = ""
    Me.Error = ""
    
    m_sRequiereRiesgoDeBibliotecaCalculadoTexto = Me.Edicion.Proyecto.RequiereRiesgoDeBiblioteca
    m_Error = Me.Edicion.Proyecto.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    RequiereRiesgoDeBibliotecaCalculadoTexto = m_sRequiereRiesgoDeBibliotecaCalculadoTexto
     Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.RequiereRiesgoDeBibliotecaCalculadoTexto ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property
Public Property Let RequiereRiesgoDeBibliotecaCalculadoTexto(ByVal valor As String)

    m_sRequiereRiesgoDeBibliotecaCalculadoTexto = valor

End Property
Public Property Get RiesgoParaRetipificar() As EnumSiNo

    On Error GoTo errores
    
    m_Error = ""
    Me.Error = ""
    If m_eRiesgoParaRetipificar <> 0 Then
        RiesgoParaRetipificar = m_eRiesgoParaRetipificar
        Exit Property
    End If
    If Me.RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.No Then
        m_eRiesgoParaRetipificar = EnumSiNo.No
        RiesgoParaRetipificar = m_eRiesgoParaRetipificar
        Exit Property
    End If
    If Me.RiesgoBiblioteca Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    End If
    If Me.RiesgoBiblioteca.RiesgoParaRetipificar = EnumSiNo.Sí Then
        m_eRiesgoParaRetipificar = EnumSiNo.Sí
    ElseIf Me.RiesgoBiblioteca.RiesgoParaRetipificar = EnumSiNo.No Then
        m_eRiesgoParaRetipificar = EnumSiNo.No
    End If
    RiesgoParaRetipificar = m_eRiesgoParaRetipificar
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.RiesgoParaRetipificar ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Let RiesgoParaRetipificar(ByVal sNewValue As EnumSiNo)

    m_eRiesgoParaRetipificar = sNewValue

End Property


Public Property Get RiesgoBiblioteca() As RiesgoBiblioteca

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjRiesgoBiblioteca Is Nothing Then
        Set RiesgoBiblioteca = m_ObjRiesgoBiblioteca
        Exit Property
    End If
    If Me.CodRiesgoBiblioteca = "" Then
        Exit Property
    End If
    Set m_ObjRiesgoBiblioteca = Constructor.getRiesgoBiblioteca(, Me.CodRiesgoBiblioteca, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set RiesgoBiblioteca = m_ObjRiesgoBiblioteca
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Riesgo.RiesgoBiblioteca ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Set RiesgoBiblioteca(ByVal objNewValue As RiesgoBiblioteca)

    Set m_ObjRiesgoBiblioteca = objNewValue

End Property

Public Property Get riesgo() As riesgo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjRiesgo Is Nothing Then
        Set riesgo = m_ObjRiesgo
        Exit Property
    End If
    If Me.IDRiesgo = "" Then
        Exit Property
    End If
    Set m_ObjRiesgo = Constructor.getRiesgo(Me.IDRiesgo, , , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set riesgo = m_ObjRiesgo
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método RiesgoExterno.Riesgo ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set riesgo(ByVal sNewValue As Variant)

    Set m_ObjRiesgo = sNewValue

End Property

Public Property Get UsuarioRegistraObj() As Usuario
    
    On Error GoTo errores
        
    Me.Error = ""
    m_Error = ""
    If Not m_objUsuarioRegistraObj Is Nothing Then
        Set UsuarioRegistraObj = m_objUsuarioRegistraObj
        Exit Property
    End If
    If Me.UsuarioRegistra = "" Then
        m_Error = "El método RiesgoExterno.UsuarioRegistraObj ha llamado a RiesgoExterno.UsuarioRegistra sin resultado"
        Err.Raise 1000
    End If
    Set m_objUsuarioRegistraObj = Constructor.getUsuario(Me.UsuarioRegistra, , , , m_Error)
    If m_Error <> "" Then
        m_Error = "El método RiesgoExterno.UsuarioRegistraObj ha llamado a UsuarioGestor.getUsuario dando el error:" & vbNewLine & m_Error
        Err.Raise 1000
    End If
    Set UsuarioRegistraObj = m_objUsuarioRegistraObj
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método RiesgoExterno.UsuarioRegistraObj ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
    
End Property

Public Property Set UsuarioRegistraObj(ByVal objNewValue As Usuario)
    
    
    On Error GoTo errores
    If Me.IDRiesgoExt <> "" Then
        m_Error = "El método RiesgoExterno.UsuarioRegistraObj quiere establecer el Proyecto teniendo IDRiesgoExt"
        Err.Raise 1000
    End If
    Set m_objUsuarioRegistraObj = objNewValue
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método RiesgoExterno.UsuarioRegistraObj ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
    
End Property

Public Property Get Edicion() As Edicion
    
    On Error GoTo errores
        
    Me.Error = ""
    m_Error = ""
    If Not m_objEdicion Is Nothing Then
        Set Edicion = m_objEdicion
        Exit Property
    End If
    If Me.IDEdicion = "" Then
        Exit Property
    End If
    Set m_objEdicion = Constructor.getEdicion(Me.IDEdicion, m_Error)
    If m_Error <> "" Then
        m_Error = "El método RiesgoExterno.Edicion ha llamado a EdicionGestor.getEdicion dando el error:" & vbNewLine & m_Error
        Err.Raise 1000
    End If
    Set Edicion = m_objEdicion
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método RiesgoExterno.Edicion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Set Edicion(ByVal objNewValue As Edicion)
    
    On Error GoTo errores
    If Me.IDRiesgoExt <> "" Then
        m_Error = "El método RiesgoExterno.Edicion quiere establecer el Proyecto teniendo IDRiesgoExt"
        Err.Raise 1000
    End If
    Set m_objEdicion = objNewValue
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método RiesgoExterno.Edicion ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
    
End Property

Public Property Get UsuarioConectadoAutorizado() As EnumSiNo
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_eUsuarioConectadoAutorizado <> 0 Then
        UsuarioConectadoAutorizado = m_eUsuarioConectadoAutorizado
        Exit Property
    End If
    m_eUsuarioConectadoAutorizado = UsuarioAutorizado(Me, , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    UsuarioConectadoAutorizado = m_eUsuarioConectadoAutorizado
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método RiesgoExterno.UsuarioConectadoAutorizado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get Traslasdado() As EnumSiNo

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Me.riesgo Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_eTraslasdado = EnumSiNo.No
    Else
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        m_eTraslasdado = EnumSiNo.Sí
    End If
    Traslasdado = m_eTraslasdado
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método RiesgoExterno.Traslasdado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Function DescripcionRepetida( _
                                    p_Descripcion As String, _
                                    Optional p_IDRiesgoExtAExcluir As String, _
                                    Optional ByRef p_Error As String _
                                    ) As EnumSiNo
                               
    Dim m_ID As Variant
    Dim m_ObjRiesgoExt As RiesgoExterno
    Dim m_Descripcion As String
    Dim m_objEdicion As Edicion
    
    On Error GoTo errores
    Set m_objEdicion = Me.Edicion
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_objEdicion Is Nothing Then
        p_Error = "No se sabe a qué edición pertence este Riesgo Externo"
        Err.Raise 1000
    End If
    
    If m_objEdicion.ColRiesgosOferta Is Nothing Then
        DescripcionRepetida = EnumSiNo.No
        Exit Function
    End If
    For Each m_ID In m_objEdicion.ColRiesgosOferta.keys
        If Nz(m_ID, "") = "" Then
            GoTo SiguienteRiesgoExterno
        End If
        Set m_ObjRiesgoExt = m_objEdicion.ColRiesgosOferta(m_ID)
        m_Descripcion = m_ObjRiesgoExt.Descripcion
        If p_Descripcion = m_Descripcion Then
            If p_IDRiesgoExtAExcluir <> "" Then
                If p_IDRiesgoExtAExcluir = m_ObjRiesgoExt.IDRiesgoExt Then
                    Set m_ObjRiesgoExt = Nothing
                    GoTo SiguienteRiesgoExterno
                End If
            End If
            DescripcionRepetida = EnumSiNo.Sí
            Set m_ObjRiesgoExt = Nothing
            Exit Function
        End If
        Set m_ObjRiesgoExt = Nothing
SiguienteRiesgoExterno:
    Next
    
    DescripcionRepetida = EnumSiNo.No
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método DescripcionRepetida ha devuelto el error: " & Err.Description
    End If
End Function
Public Function DescripcionYCausaRaizRepetidas( _
                                                p_Descripcion As String, _
                                                p_CausaRaiz As String, _
                                                Optional p_IDRiesgoExtAExcluir As String, _
                                                Optional ByRef p_Error As String _
                                                ) As EnumSiNo
                               
    Dim m_ID As Variant
    Dim m_ObjRiesgoExt As RiesgoExterno
    Dim m_Descripcion As String
    Dim m_CausaRaiz As String
    Dim m_objEdicion As Edicion
    
    On Error GoTo errores
    Set m_objEdicion = Me.Edicion
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_objEdicion Is Nothing Then
        p_Error = "No se sabe a qué edición pertence este Riesgo Externo"
        Err.Raise 1000
    End If
    
    If m_objEdicion.ColRiesgosOferta Is Nothing Then
        DescripcionYCausaRaizRepetidas = EnumSiNo.No
        Exit Function
    End If
    For Each m_ID In m_objEdicion.ColRiesgosOferta.keys
        If Nz(m_ID, "") = "" Then
            GoTo SiguienteRiesgoExterno
        End If
        Set m_ObjRiesgoExt = m_objEdicion.ColRiesgosOferta(m_ID)
        m_Descripcion = m_ObjRiesgoExt.Descripcion
        m_CausaRaiz = m_ObjRiesgoExt.CausaRaiz
        If p_Descripcion = m_Descripcion And p_CausaRaiz = m_CausaRaiz Then
            If p_IDRiesgoExtAExcluir <> "" Then
                If p_IDRiesgoExtAExcluir = m_ObjRiesgoExt.IDRiesgoExt Then
                    Set m_ObjRiesgoExt = Nothing
                    GoTo SiguienteRiesgoExterno
                End If
            End If
            DescripcionYCausaRaizRepetidas = EnumSiNo.Sí
            Set m_ObjRiesgoExt = Nothing
            Exit Function
        End If
        Set m_ObjRiesgoExt = Nothing
SiguienteRiesgoExterno:
    Next
    
    DescripcionYCausaRaizRepetidas = EnumSiNo.No
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método DescripcionYCausaRaizRepetidas ha devuelto el error: " & Err.Description
    End If
End Function
Public Function CodigoRepetido( _
                                p_Codigo As String, _
                                Optional p_IDRiesgoExtAExcluir As String, _
                                Optional ByRef p_Error As String _
                                ) As EnumSiNo
                               
    Dim m_ID As Variant
    Dim m_ObjRiesgoExt As RiesgoExterno
    Dim m_Codigo As String
    Dim m_objEdicion As Edicion
    
    
    On Error GoTo errores
    Set m_objEdicion = Me.Edicion
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If m_objEdicion Is Nothing Then
        p_Error = "No se sabe a qué edición pertence este Riesgo Externo"
        Err.Raise 1000
    End If
    Set m_objEdicion.ColRiesgosOferta = Nothing
    If m_objEdicion.ColRiesgosOferta Is Nothing Then
        CodigoRepetido = EnumSiNo.No
        Exit Function
    End If
    
    For Each m_ID In m_objEdicion.ColRiesgosOferta.keys
        If Nz(m_ID, "") = "" Then
            GoTo SiguienteRiesgoExterno
        End If
        Set m_ObjRiesgoExt = m_objEdicion.ColRiesgosOferta(m_ID)
        m_Codigo = m_ObjRiesgoExt.CodRiesgo
        If p_Codigo = m_Codigo Then
            If p_IDRiesgoExtAExcluir <> "" Then
                If p_IDRiesgoExtAExcluir = m_ObjRiesgoExt.IDRiesgoExt Then
                    Set m_ObjRiesgoExt = Nothing
                    GoTo SiguienteRiesgoExterno
                End If
            End If
            CodigoRepetido = EnumSiNo.Sí
            Set m_ObjRiesgoExt = Nothing
            Exit Function
        End If
        Set m_ObjRiesgoExt = Nothing
SiguienteRiesgoExterno:
    Next
    
    CodigoRepetido = EnumSiNo.No
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CodigoRepetido ha devuelto el error: " & Err.Description
    End If
End Function

Private Function MotivoNoOK( _
                            p_ObjRiesgoExternoAlInicio As RiesgoExterno, _
                            Optional ByRef p_Error As String _
                            ) As String
    
    
    
    Dim m_CampoRepetido As EnumSiNo
    
    Dim m_IdRiesgo As Variant
    On Error GoTo errores
    
    If EsTecnico = EnumSiNo.Sí Then
        MotivoNoOK = "El perfil de su usuario no tiene autorización para realizar esta operación"
        Exit Function
    End If
    With Me
        
        If .Edicion Is Nothing Then
            MotivoNoOK = "No se conoce la edición del Riesgo Externo a dar de alta"
            Exit Function
        End If
        If .Edicion.Proyecto.RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.Sí Then
            If .RiesgoParaRetipificar = EnumSiNo.Sí Then
                MotivoNoOK = "Un riesgo externo no se puede retipificar después"
                Exit Function
            End If
            If .RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.Sí Then
                If .CodRiesgoBiblioteca = "" Then
                    MotivoNoOK = "La gestión de riesgos requiere Código de Biblioteca"
                    Exit Function
                End If
            End If
        End If
        If .Origen <> "Oferta" And .Origen <> "Suministrador" And .Origen <> "Pedido" Then
            MotivoNoOK = "El origen ha de ser Oferta,Suministrador o Pedido"
            Exit Function
        End If
        If .Origen = "Suministrador" Then
            If .Suministrador = "" Then
                MotivoNoOK = "El origen es de suministrador y falta"
                Exit Function
            End If
        End If
        If .Origen = "Oferta" Then
            If .Edicion.Edicion <> "1" Then
                MotivoNoOK = "El Origen del Riesgo Externo es de Oferta, y sólo se puede añadir o editar las de la " & _
                            "primera edición"
                Exit Function
            End If
        Else
            If .Edicion.EsActivo = EnumSiNo.No Then
                MotivoNoOK = "Sólo se puede añadir o editar las de la primera edición, y ha de ser la activa"
                Exit Function
            End If
        End If
        If .Origen = "Pedido" Then
            If .Pedido = "" Then
                MotivoNoOK = "El Origen del Riesgo Externo es de Pedido y falta el mismo"
                Exit Function
            End If
            If .ProveedorPedido = "" Then
                MotivoNoOK = "El Origen del Riesgo Externo es de Pedido y falta el proveedor"
                Exit Function
            End If
        End If
        If .Descripcion = "" Then
            MotivoNoOK = "Es obligatoria la Descripción"
            Exit Function
        End If
        If .Edicion.Proyecto.RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.No Then
            If p_ObjRiesgoExternoAlInicio Is Nothing Then
                If .DescripcionRepetida(.Descripcion, , p_Error) = EnumSiNo.Sí Then
                    MotivoNoOK = "Esta descripción del Riesgo Externo ya ha sido registrada para la misma gestión de riesgos"
                    Exit Function
                End If
                If .DescripcionYCausaRaizRepetidas(.Descripcion, .CausaRaiz, , p_Error) = EnumSiNo.Sí Then
                    MotivoNoOK = "Esta descripción del Riesgo Externo junto con la causa raíz, ya han sido registradas para la misma gestión de riesgos"
                    Exit Function
                End If
            Else
                If p_ObjRiesgoExternoAlInicio.Descripcion <> .Descripcion Then
                    If .DescripcionRepetida(.Descripcion, , p_Error) = EnumSiNo.Sí Then
                        MotivoNoOK = "Esta descripción del Riesgo Externo ya ha sido registrada para la misma gestión de riesgos"
                        Exit Function
                    End If
                End If
                If p_ObjRiesgoExternoAlInicio.Descripcion <> .Descripcion Or p_ObjRiesgoExternoAlInicio.CausaRaiz <> .CausaRaiz Then
                    If .DescripcionYCausaRaizRepetidas(.Descripcion, .CausaRaiz, , p_Error) = EnumSiNo.Sí Then
                        MotivoNoOK = "Esta descripción del Riesgo Externo junto con la causa raíz, ya han sido registradas para la misma gestión de riesgos"
                        Exit Function
                    End If
                End If
            End If
        End If
        'si la edición tiene riesgos y esta se traslada no puede tener la misma descripción de otro riesgo que pudiera haber
        If .Trasladar = "Sí" And Not .Edicion.colRiesgos Is Nothing Then
            For Each m_IdRiesgo In .Edicion.colRiesgos.keys
                If Nz(m_IdRiesgo, "") <> "" Then
                    Set m_ObjRiesgo = .Edicion.colRiesgos(m_IdRiesgo)
                    If .Edicion.Proyecto.RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.No Then
                        If .Descripcion = m_ObjRiesgo.Descripcion Then
                            MotivoNoOK = "El Riesgo Externo va a ser trasladado pero tiene una descripción idéntica a uno de los riesgos existentes: " & m_ObjRiesgo.CodigoRiesgo
                            Exit Function
                        End If
                    Else
                        If .Descripcion = m_ObjRiesgo.Descripcion And .CausaRaiz = m_ObjRiesgo.CausaRaiz Then
                            MotivoNoOK = "El Riesgo Externo va a ser trasladado pero tiene una descripción y Causa Raíz idéntica a uno de los riesgos existentes: " & m_ObjRiesgo.CodigoRiesgo
                            Exit Function
                        End If
                    End If
                    Set m_ObjRiesgo = Nothing
                End If
            Next
        End If
    
        If Not IsDate(.FechaDetectado) Then
            MotivoNoOK = "Es obligatorio FechaDetectado"
            Exit Function
        End If
        If .CodRiesgo = "" Then
            .CodRiesgo = .CodigoRiesgoExtCalculado
            m_Error = Me.Error
            If m_Error <> "" Then
                Err.Raise 1000
            End If
            If .CodRiesgo = "" Then
                MotivoNoOK = "Es obligatorio CodRiesgo"
                Exit Function
            End If
            
        End If
        If p_ObjRiesgoExternoAlInicio Is Nothing Then
            If .CodigoRepetido(.CodRiesgo, , p_Error) = EnumSiNo.Sí Then
                MotivoNoOK = "Este Código del Riesgo Externo ya ha sido registrado para la misma gestión de riesgos"
                Exit Function
            End If
        Else
            If p_ObjRiesgoExternoAlInicio.CodRiesgo <> .CodRiesgo Then
                If .CodigoRepetido(.CodRiesgo, , p_Error) = EnumSiNo.Sí Then
                    MotivoNoOK = "Este Código del Riesgo Externo ya ha sido registrado para la misma gestión de riesgos"
                    Exit Function
                End If
            End If
        End If
        If .Trasladar <> "Sí" And .Trasladar <> "No" Then
            MotivoNoOK = "Trasladar ha de ser Sí o no"
            Exit Function
        End If
        If .Trasladar = "No" Then
            If .MotivoNoIntegrado = "" Then
                MotivoNoOK = "Si no se traslada hay que indicar algo en MotivoNoIntegrado"
                Exit Function
            End If
            If Not IsDate(.FechaMotivo) Then
                .FechaMotivo = Now()
            End If
        End If
        If .UsuarioRegistra = "" Then
            .UsuarioRegistra = m_ObjUsuarioConectado.UsuarioRed
        End If
        If .Edicion.Proyecto.RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.Sí Then
            If .CausaRaiz = "" Then
                MotivoNoOK = "La causa raíz es obligatoria"
                Exit Function
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
                            Optional p_ObjRiesgoExternoAlInicio As RiesgoExterno, _
                            Optional ByRef p_Error As String _
                            ) As String
                            
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_MotivoNoOK As String
    Dim blnAltaRiesgoExterno As Boolean
    Dim m_ObjUsuario As Usuario
    Dim m_ObjRiesgo As riesgo
    Dim m_FechaRef As Date
    
    On Error GoTo errores
    
    Me.Error = ""
    m_FechaRef = Date
    'm_fechaRef = "12/12/2023"
    
    m_MotivoNoOK = MotivoNoOK(p_ObjRiesgoExternoAlInicio, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    p_Error = m_MotivoNoOK
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    If p_ObjRiesgoExternoAlInicio Is Nothing Then
        Me.IDRiesgoExt = Me.IDRiesgoExtCalculado
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        m_SQL = "TbRiesgosAIntegrar"
    Else
        m_SQL = "SELECT * FROM TbRiesgosAIntegrar " & _
                "WHERE IDRiesgoExt=" & Me.IDRiesgoExt & ";"
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If p_ObjRiesgoExternoAlInicio Is Nothing Then
            .AddNew
                If Me.IDRiesgoExt <> "" Then
                    .Fields("IDRiesgoExt") = Me.IDRiesgoExt
                End If
                If Me.CodRiesgo <> "" Then
                    .Fields("CodRiesgo") = Me.CodRiesgo
                End If
                If Me.FechaDetectado <> "" Then
                    .Fields("FechaDetectado") = Me.FechaDetectado
                End If
                If Me.IDRiesgo <> "" Then
                    .Fields("IDRiesgo") = Me.IDRiesgo
                End If
                If Me.Origen <> "" Then
                    .Fields("Origen") = Me.Origen
                End If
                If Me.IDEdicion <> "" Then
                    .Fields("IDEdicion") = Me.IDEdicion
                End If
                If Me.Descripcion <> "" Then
                    .Fields("Descripcion") = Me.Descripcion
                End If
                If Me.CausaRaiz <> "" Then
                    .Fields("CausaRaiz") = Me.CausaRaiz
                End If
                Me.FechaAltaRegistro = Now()
                If Me.FechaAltaRegistro <> "" Then
                    .Fields("FechaAltaRegistro") = Me.FechaAltaRegistro
                End If
                Me.UsuarioRegistra = m_ObjUsuarioConectado.UsuarioRed
                If Me.UsuarioRegistra <> "" Then
                    .Fields("UsuarioRegistra") = Me.UsuarioRegistra
                End If
                If Me.MotivoNoIntegrado <> "" Then
                    .Fields("MotivoNoIntegrado") = Me.MotivoNoIntegrado
                End If
                If Me.FechaMotivo <> "" Then
                    .Fields("FechaMotivo") = Me.FechaMotivo
                End If
                If Me.Trasladar <> "" Then
                    .Fields("Trasladar") = Me.Trasladar
                End If
                If Me.Suministrador <> "" Then
                    .Fields("Suministrador") = Me.Suministrador
                End If
                If Me.Pedido <> "" Then
                    .Fields("Pedido") = Me.Pedido
                End If
                If Me.ProveedorPedido <> "" Then
                    .Fields("ProveedorPedido") = Me.ProveedorPedido
                End If
                If Me.CausaRiesgoPedido <> "" Then
                    .Fields("CausaRiesgoPedido") = Me.CausaRiesgoPedido
                End If
                Me.RequiereRiesgoDeBiblioteca = Me.Edicion.Proyecto.RequiereRiesgoDeBiblioteca
                If Me.RequiereRiesgoDeBiblioteca <> "" Then
                    .Fields("RequiereRiesgoDeBiblioteca") = Me.RequiereRiesgoDeBiblioteca
                End If
                If Me.CodRiesgoBiblioteca <> "" Then
                    .Fields("CodRiesgoBiblioteca") = Me.CodRiesgoBiblioteca
                End If
                If Me.RiesgoPendienteRetipificacion <> "" Then
                    .Fields("RiesgoPendienteRetipificacion") = Me.RiesgoPendienteRetipificacion
                End If
                
               
            .Update
        Else
            If .EOF Then
                p_Error = "No está registrado el Riesgo externo"
                Err.Raise 1000
            End If
            .Edit
               If Me.CodRiesgo <> "" Then
                    .Fields("CodRiesgo") = Me.CodRiesgo
                Else
                    .Fields("CodRiesgo") = Null
                End If
                If Me.FechaDetectado <> "" Then
                    .Fields("FechaDetectado") = Me.FechaDetectado
                Else
                    .Fields("FechaDetectado") = Null
                End If
                If Me.IDRiesgo <> "" Then
                    .Fields("IDRiesgo") = Me.IDRiesgo
                Else
                    .Fields("IDRiesgo") = Null
                End If
                If Me.Origen <> "" Then
                    .Fields("Origen") = Me.Origen
                Else
                    .Fields("Origen") = Null
                End If
                If Me.IDEdicion <> "" Then
                    .Fields("IDEdicion") = Me.IDEdicion
                Else
                    .Fields("IDEdicion") = Null
                End If
                If Me.Descripcion <> "" Then
                    .Fields("Descripcion") = Me.Descripcion
                Else
                    .Fields("Descripcion") = Null
                End If
                If Me.CausaRaiz <> "" Then
                    .Fields("CausaRaiz") = Me.CausaRaiz
                Else
                    .Fields("CausaRaiz") = Null
                End If
                
                Me.UsuarioRegistra = m_ObjUsuarioConectado.UsuarioRed
                If Me.UsuarioRegistra <> "" Then
                    .Fields("UsuarioRegistra") = Me.UsuarioRegistra
                End If
                If Me.MotivoNoIntegrado <> "" Then
                    .Fields("MotivoNoIntegrado") = Me.MotivoNoIntegrado
                Else
                    .Fields("MotivoNoIntegrado") = Null
                End If
                If Me.FechaMotivo <> "" Then
                    .Fields("FechaMotivo") = Me.FechaMotivo
                Else
                    .Fields("FechaMotivo") = Null
                End If
                If Me.Trasladar <> "" Then
                    .Fields("Trasladar") = Me.Trasladar
                Else
                    .Fields("Trasladar") = Null
                End If
                If Me.Suministrador <> "" Then
                    .Fields("Suministrador") = Me.Suministrador
                Else
                    .Fields("Suministrador") = Null
                End If
                If Me.Pedido <> "" Then
                    .Fields("Pedido") = Me.Pedido
                Else
                    .Fields("Pedido") = Null
                End If
                If Me.ProveedorPedido <> "" Then
                    .Fields("ProveedorPedido") = Me.ProveedorPedido
                Else
                    .Fields("ProveedorPedido") = Null
                End If
                If Me.CausaRiesgoPedido <> "" Then
                    .Fields("CausaRiesgoPedido") = Me.CausaRiesgoPedido
                Else
                    .Fields("CausaRiesgoPedido") = Null
                End If
                Me.RequiereRiesgoDeBiblioteca = Me.Edicion.Proyecto.RequiereRiesgoDeBiblioteca
                If Me.RequiereRiesgoDeBiblioteca <> "" Then
                    .Fields("RequiereRiesgoDeBiblioteca") = Me.RequiereRiesgoDeBiblioteca
                Else
                    .Fields("RequiereRiesgoDeBiblioteca") = Null
                End If
                If Me.CodRiesgoBiblioteca <> "" Then
                    .Fields("CodRiesgoBiblioteca") = Me.CodRiesgoBiblioteca
                Else
                    .Fields("CodRiesgoBiblioteca") = Null
                End If
                If Me.RiesgoPendienteRetipificacion <> "" Then
                    .Fields("RiesgoPendienteRetipificacion") = Me.RiesgoPendienteRetipificacion
                Else
                    .Fields("RiesgoPendienteRetipificacion") = Null
                End If
            .Update
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If p_ObjRiesgoExternoAlInicio Is Nothing Then
        blnAltaRiesgoExterno = True
    End If
    
    
    
    
    
    '---------------------------------------------------------------------------------
    ' ahora el alta de los riesgos de los que se trasladen a la gestión de riesgos
    '---------------------------------------------------------------------------------
    If Me.Trasladar = "Sí" Then
        
        
        Set m_ObjRiesgo = New riesgo
        With m_ObjRiesgo
            .IDEdicion = Me.IDEdicion
            If Me.Suministrador <> "" Then
                .EntidadDetecta = "Empresa"
                .DetectadoPor = Me.Suministrador
            ElseIf Me.ProveedorPedido <> "" Then
                .EntidadDetecta = "Persona"
                .DetectadoPor = m_ObjUsuarioConectado.Nombre
            End If
            .CodRiesgoBiblioteca = Me.CodRiesgoBiblioteca
            .Descripcion = Me.Descripcion
            If Me.Edicion.Proyecto.RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.Sí Then
                .CausaRaiz = Me.CausaRaiz
            End If
            .Estado = .ESTADOCalculadoTexto
            p_Error = .Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            .FechaDetectado = Me.FechaDetectado
            .FechaEstado = CStr(m_FechaRef)
            
            .CodigoRiesgo = .CodigoRiesgoCalculado
            .Origen = "Oferta"
            p_Error = .Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If Me.Edicion.Proyecto.RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.No Then
                .Priorizacion = .PriorizacionCalculado
            End If
            
            .RiesgoParaRetipificar = EnumSiNo.No
            
            .CodigoUnico = .CodigoUnicoCalculado
            p_Error = .Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            .AltaDesdeRiesgoExterno p_Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            
            Me.IDRiesgo = .IDRiesgo
            
            m_SQL = "UPDATE TbRiesgosAIntegrar SET TbRiesgosAIntegrar.IDRiesgo =" & .IDRiesgo & " " & _
                    "WHERE IDRiesgoExt=" & Me.IDRiesgoExt & ";"
            getdb().Execute m_SQL
            If Me.Edicion.Proyecto.RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.Sí And Me.Edicion.Proyecto.ParaInformeAvisos <> "No" Then
                If .RiesgoParaRetipificar = EnumSiNo.Sí Then
                    EnviarCorreoRiesgoRequiereRetipificacion m_ObjRiesgo, m_Error
                    If m_Error <> "" Then
                        Err.Raise 1000
                    End If
                End If
            End If
        End With
    Else
        Set m_ObjRiesgo = Me.riesgo
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If Not m_ObjRiesgo Is Nothing Then
            If m_ObjRiesgo.Edicion.Edicion = 1 Then
                If m_ObjRiesgo.EsActivo = EnumSiNo.Sí Then
                    m_ObjRiesgo.Borrar p_Error
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                End If
            End If
        End If
    End If
    With Me.Edicion
        Set .ColRiesgosExternos = Nothing
        Set .ColRiesgosOferta = Nothing
        

       
    End With
    With Me.Edicion.Proyecto
        Set .ColRiesgosExternos = Nothing
        Set .ColRiesgosOferta = Nothing
        Set .ColRiesgosSubContratista = Nothing
        Set .ColRiesgosPedidos = Nothing
        
    End With
    

    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Registrar ha devuelto el error: " & Err.Description
    End If
    If blnAltaRiesgoExterno = True Then
        If Not Me Is Nothing Then
            If Me.IDRiesgoExt <> "" Then
                Me.Borrar
            End If
        End If
        If Not m_ObjRiesgo Is Nothing Then
            If m_ObjRiesgo.IDRiesgo <> "" Then
                m_ObjRiesgo.Borrar
            End If
        End If
    End If
End Function

Public Function GenerarRiesgo( _
                                Optional ByRef p_Error As String _
                                ) As String
                                        
    Dim m_IdRiesgo As String
    Dim m_ObjRiesgo As riesgo
    Dim m_FechaRef As Date
    Dim m_SQL As String
    On Error GoTo errores
    m_FechaRef = Date
    'm_fechaRef = "12/12/2023"
    Set m_ObjRiesgo = New riesgo
    With m_ObjRiesgo
        .IDEdicion = Me.IDEdicion
        .DetectadoPor = Me.UsuarioRegistraObj.Nombre
        .Descripcion = Me.Descripcion
        If Me.Edicion.Proyecto.RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.Sí Then
            .CausaRaiz = Me.CausaRaiz
        End If
        If Me.Origen = "Suministrador" Then
            .EntidadDetecta = "Empresa"
            .DetectadoPor = Me.Suministrador
        Else
            .EntidadDetecta = "Persona"
        End If
        
        .FechaDetectado = CStr(m_FechaRef)
        .FechaEstado = CStr(m_FechaRef)
        .CodigoRiesgo = .CodigoRiesgoCalculado
        If Me.Edicion.Proyecto.RequiereRiesgoDeBibliotecaCalculado = EnumSiNo.No Then
            .Priorizacion = .PriorizacionCalculado
        End If
        .Estado = "Incompleto"
        p_Error = .Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        .CodigoUnico = .CodigoUnicoCalculado
        p_Error = .Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        .CodRiesgoBiblioteca = Me.CodRiesgoBiblioteca
        
        .RiesgoParaRetipificar = Me.RiesgoParaRetipificar
        .AltaDesdeRiesgoExterno p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        m_IdRiesgo = .IDRiesgo
        'ahora hay que editar el riesgo externo para meterle la IDRiesgo trasladado
        m_SQL = "UPDATE TbRiesgosAIntegrar SET TbRiesgosAIntegrar.IDRiesgo =" & .IDRiesgo & " " & _
                "WHERE IDRiesgoExt=" & Me.IDRiesgoExt & ";"
        getdb().Execute m_SQL
    End With
    With Me.Edicion
        If .ColRiesgosExternos Is Nothing Then
            Set .ColRiesgosExternos = New Scripting.Dictionary
            .ColRiesgosExternos.CompareMode = TextCompare
        End If
        Set .ColRiesgosExternos(Me.IDRiesgoExt) = Me
        If .ColRiesgosOferta Is Nothing Then
            Set .ColRiesgosOferta = New Scripting.Dictionary
            .ColRiesgosOferta.CompareMode = TextCompare
        End If
        Set .ColRiesgosOferta(Me.IDRiesgoExt) = Me
    End With
    With Me.Edicion.Proyecto
        Set .ColRiesgosExternos = Nothing
        Set .ColRiesgosDeJP = Nothing
        Set .ColRiesgosOferta = Nothing
        
    End With
                                    
    Exit Function
    
errores:

    If Err.Number <> 1000 Then
        p_Error = "El método GenerarRiesgo ha devuelto el error: " & Err.Description
    End If
End Function
Public Function QuitarRiesgoGenerado( _
                                        Optional ByRef p_Error As String _
                                        ) As String
                                        
    Dim m_SQL As String
    Dim m_Origen As String
    On Error GoTo errores
    'puede que el riesgo hubiera estado trasladado y ahora no se quiere
        
    If Not Me.riesgo Is Nothing Then
        Me.IDRiesgo = ""
    End If
    m_SQL = "UPDATE TbRiesgosAIntegrar SET TbRiesgosAIntegrar.IDRiesgo = Null " & _
                "WHERE IDRiesgoExt=" & Me.IDRiesgoExt & ";"
    getdb().Execute m_SQL
    
    If Not Me.riesgo Is Nothing Then
        'se ha de borrar el riesgo trasladado
        Me.riesgo.Borrar p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Set Me.riesgo = Nothing
        
    End If
    m_Origen = Me.Origen
    With Me.Edicion.Proyecto
        Set .ColRiesgosExternos = Nothing
        Set .ColRiesgosDeJP = Nothing
        Set .ColRiesgosOferta = Nothing
        Set .ColRiesgosSubContratista = Nothing
    End With
    

    Exit Function
    
errores:

    If Err.Number <> 1000 Then
        p_Error = "El método QuitarRiesgoGenerado ha devuelto el error: " & Err.Description
    End If
End Function

Private Function ValidacionEliminar( _
                                    Optional ByRef p_Error As String _
                                    ) As String

    'sólo se puede eliminar si
    '1) Siempre si no se ha trasladado el riesgo
    '2) Si se ha trasladado el riesgo sólo cuando este es de una edición aún no publicada
    
    On Error GoTo errores
    
    
    If EsTecnico = EnumSiNo.Sí Then
        p_Error = "El perfil de su usuario no tiene autorización para realizar esta operación"
        Err.Raise 1000
    End If
    If Me.riesgo Is Nothing Then
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Function
    End If
    
    If Me.riesgo.Edicion Is Nothing Then
        p_Error = "No se sabe de qué edición es el Riesgo trasladado"
        Err.Raise 1000
    End If
    If Me.Edicion.Proyecto.EdicionUltima.EsActivo Then
        p_Error = Me.Edicion.Proyecto.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        If Me.Edicion.Proyecto.EdicionUltima.IDEdicion <> Me.riesgo.Edicion.IDEdicion Then
            p_Error = "El Riesgo Externo ya se ha trasladado en una edición que no es la activa en estos momentos"
            Err.Raise 1000
        End If
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
                                
    Dim m_IDRiesgoExt As String
    Dim m_Origen As String
    Dim m_ObjRiesgo As riesgo
    Dim m_SQL As String
    On Error GoTo errores
    
    Me.Error = ""
    Set m_ObjRiesgo = Me.riesgo
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    ValidacionEliminar p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    If Not m_ObjRiesgo Is Nothing Then
        m_ObjRiesgo.Borrar p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    m_IDRiesgoExt = Me.IDRiesgoExt
    m_Origen = Me.Origen
    m_SQL = "DELETE * FROM TbRiesgosAIntegrar " & _
            "WHERE IDRiesgoExt=" & Me.IDRiesgoExt & ";"
    getdb().Execute m_SQL
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Borrar ha devuelto el error: " & Err.Description
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
        .Add "IDRiesgoExt"
        .Add "CodRiesgo"
        .Add "IDRiesgo"
        .Add "Origen"
        .Add "IDEdicion"
        .Add "Descripcion"
        .Add "CausaRaiz"
        .Add "FechaDetectado"
        .Add "FechaAltaRegistro"
        .Add "UsuarioRegistra"
        .Add "MotivoNoIntegrado"
        .Add "FechaMotivo"
        .Add "Trasladar"
        .Add "Suministrador"
        .Add "Pedido"
        .Add "ProveedorPedido"
        .Add "CausaRiesgoPedido"
        .Add "RequiereRiesgoDeBiblioteca"
        .Add "CodRiesgoBiblioteca"
        .Add "RiesgoPendienteRetipificacion"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método RiesgosAIntegrar.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "IDRiesgoExt" Then
        getPropiedad = Me.IDRiesgoExt
        Exit Function
    End If
    If m_NombrePropiedad = "CodRiesgo" Then
        getPropiedad = Me.CodRiesgo
        Exit Function
    End If
    If m_NombrePropiedad = "IDRiesgo" Then
        getPropiedad = Me.IDRiesgo
        Exit Function
    End If
    If m_NombrePropiedad = "Origen" Then
        getPropiedad = Me.Origen
        Exit Function
    End If
    If m_NombrePropiedad = "IDEdicion" Then
        getPropiedad = Me.IDEdicion
        Exit Function
    End If
    If m_NombrePropiedad = "Descripcion" Then
        getPropiedad = Me.Descripcion
        Exit Function
    End If
    If m_NombrePropiedad = "CausaRaiz" Then
        getPropiedad = Me.CausaRaiz
        Exit Function
    End If
    If m_NombrePropiedad = "FechaDetectado" Then
        getPropiedad = Me.FechaDetectado
        Exit Function
    End If
    If m_NombrePropiedad = "FechaAltaRegistro" Then
        getPropiedad = Me.FechaAltaRegistro
        Exit Function
    End If
    If m_NombrePropiedad = "UsuarioRegistra" Then
        getPropiedad = Me.UsuarioRegistra
        Exit Function
    End If
    If m_NombrePropiedad = "MotivoNoIntegrado" Then
        getPropiedad = Me.MotivoNoIntegrado
        Exit Function
    End If
    If m_NombrePropiedad = "FechaMotivo" Then
        getPropiedad = Me.FechaMotivo
        Exit Function
    End If
    If m_NombrePropiedad = "Trasladar" Then
        getPropiedad = Me.Trasladar
        Exit Function
    End If
    If m_NombrePropiedad = "Suministrador" Then
        getPropiedad = Me.Suministrador
        Exit Function
    End If
    If m_NombrePropiedad = "Pedido" Then
        getPropiedad = Me.Pedido
        Exit Function
    End If
    If m_NombrePropiedad = "ProveedorPedido" Then
        getPropiedad = Me.ProveedorPedido
        Exit Function
    End If
    If m_NombrePropiedad = "CausaRiesgoPedido" Then
        getPropiedad = Me.CausaRiesgoPedido
        Exit Function
    End If
    If m_NombrePropiedad = "RequiereRiesgoDeBiblioteca" Then
        getPropiedad = Me.RequiereRiesgoDeBiblioteca
        Exit Function
    End If
    If m_NombrePropiedad = "CodRiesgoBiblioteca" Then
        getPropiedad = Me.CodRiesgoBiblioteca
        Exit Function
    End If
    If m_NombrePropiedad = "RiesgoPendienteRetipificacion" Then
        getPropiedad = Me.RiesgoPendienteRetipificacion
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en RiesgosAIntegrar.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RiesgosAIntegrar.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function

Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "IDRiesgoExt" Then
        Me.IDRiesgoExt = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CodRiesgo" Then
        Me.CodRiesgo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDRiesgo" Then
        Me.IDRiesgo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Origen" Then
        Me.Origen = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDEdicion" Then
        Me.IDEdicion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Descripcion" Then
        Me.Descripcion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CausaRaiz" Then
        Me.CausaRaiz = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaDetectado" Then
        Me.FechaDetectado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaAltaRegistro" Then
        Me.FechaAltaRegistro = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "UsuarioRegistra" Then
        Me.UsuarioRegistra = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "MotivoNoIntegrado" Then
        Me.MotivoNoIntegrado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaMotivo" Then
        Me.FechaMotivo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Trasladar" Then
        Me.Trasladar = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Suministrador" Then
        Me.Suministrador = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Pedido" Then
        Me.Pedido = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "ProveedorPedido" Then
        Me.ProveedorPedido = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CausaRiesgoPedido" Then
        Me.CausaRiesgoPedido = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "RequiereRiesgoDeBiblioteca" Then
        Me.RequiereRiesgoDeBiblioteca = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "CodRiesgoBiblioteca" Then
        Me.CodRiesgoBiblioteca = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "RiesgoPendienteRetipificacion" Then
        Me.RiesgoPendienteRetipificacion = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en RiesgosAIntegrar.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método RiesgosAIntegrar.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


