VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Observacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public IDObservacion As String
Public ID As String
Public Fecha As String
Public Observacion As String
Public TIPO As String

Public ParaAlta As EnumSiNo
Private m_sIDNuevo As String
Private m_ObjUsuario As UsuarioHPS
Private m_objColCampos As Collection
Public Error As String
Private m_Error As String

Public Property Get Usuario() As UsuarioHPS
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjUsuario Is Nothing Then
        Set Usuario = m_ObjUsuario
        Exit Property
    End If
    If Me.ID = "" Then
        Exit Property
    End If
    Set m_ObjUsuario = Constructor.getUsuarioHPS(Me.ID, , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Usuario = m_ObjUsuario
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Observacion.Usuario ha devuelto el error: " & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    Me.Error = m_Error
End Property



Public Property Get IDNuevo() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    m_sIDNuevo = getID("TbObservaciones", "IDObservacion", m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    IDNuevo = m_sIDNuevo
    
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método Observacion.IDNuevo ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Function MotivoNoOK( _
                            Optional ByRef p_Error As String _
                                ) As String
   
    
    On Error GoTo errores
    If Application.TempVars("SoloLectura") = "Sí" Then
        p_Error = "El Sistema está en modo sólo lectura"
        Err.Raise 1000
    End If
    With Me
        If .ParaAlta = Empty Then
            MotivoNoOK = "No se sabe si es alta o edición"
            Exit Function
            
        End If
        
        If .Observacion = "" Then
            MotivoNoOK = "El Observacion es obligatorio"
            Exit Function
        End If
        
        If .TIPO <> "Curso" And .TIPO <> "General" Then
            MotivoNoOK = "El Tipo es obligatorio y ha de ser Curso o General"
            Exit Function
        End If
       
        If .Usuario Is Nothing Then
            p_Error = .Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            MotivoNoOK = "El usuario es obligatorio"
            Exit Function
        Else
            p_Error = .Error
            If p_Error <> "" Then
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

Private Function ValidacionEliminar( _
                                    Optional ByRef p_Error As String _
                                    ) As String
                                    
   
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Application.TempVars("SoloLectura") = "Sí" Then
        p_Error = "El Sistema está en modo sólo lectura"
        Err.Raise 1000
    End If
    
        
        
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método UsuarioHPS.ValidacionEliminar ha devuelto el error: " & Err.Description
    End If
End Function

Public Function Registrar( _
                        Optional ByRef p_Error As String _
                        ) As String
                                
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_MotivoNoOK As String
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    
    If Me.ParaAlta = 0 Then
        p_Error = "No se sabe si es alta o edición"
        Err.Raise 1000
    End If
    
    m_MotivoNoOK = MotivoNoOK(p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    p_Error = m_MotivoNoOK
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Me.ParaAlta = EnumSiNo.Sí Then
        m_SQL = "TbObservaciones"
        Me.IDObservacion = Me.IDNuevo
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    Else
        m_SQL = "SELECT * FROM TbObservaciones " & _
            "WHERE IDObservacion=" & Me.IDObservacion & ";"
        
    End If
    If Not IsDate(Me.Fecha) Then
        Me.Fecha = Date
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If Me.ParaAlta = EnumSiNo.No Then
            If .EOF Then
                p_Error = "No se ha encontrado el registro"
                Err.Raise 1000
            End If
        End If
        If Me.ParaAlta = EnumSiNo.Sí Then
            .AddNew
                .Fields("IDObservacion") = Me.IDObservacion
                .Fields("ID") = Me.ID
                .Fields("Fecha") = Me.Fecha
                .Fields("Observacion") = Me.Observacion
                .Fields("Tipo") = Me.TIPO
            .Update
        Else
            .Edit
                .Fields("ID") = Me.ID
                .Fields("Fecha") = Me.Fecha
                .Fields("Observacion") = Me.Observacion
                .Fields("Tipo") = Me.TIPO
            .Update
        
        End If
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Editar ha devuelto el error: " & Err.Description
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
   
    
    m_SQL = "DELETE * FROM TbObservaciones " & _
            "WHERE IDObservacion=" & Me.IDObservacion & ";"
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
        .Add "IDObservacion"
        .Add "ID"
        .Add "Fecha"
        .Add "Observacion"
        .Add "Tipo"
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Observacion.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "IDObservacion" Then
        getPropiedad = Me.IDObservacion
        Exit Function
    End If
    If m_NombrePropiedad = "ID" Then
        getPropiedad = Me.ID
        Exit Function
    End If
    If m_NombrePropiedad = "Fecha" Then
        getPropiedad = Me.Fecha
        Exit Function
    End If
    If m_NombrePropiedad = "Observacion" Then
        getPropiedad = Me.Observacion
        Exit Function
    End If
    If m_NombrePropiedad = "Tipo" Then
        getPropiedad = Me.TIPO
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en Observacion.getPropiedad"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Observacion.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "IDObservacion" Then
        Me.IDObservacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "ID" Then
        Me.ID = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Fecha" Then
        Me.Fecha = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Observacion" Then
        Me.Observacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Tipo" Then
        Me.TIPO = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en Observacion.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Observacion.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


