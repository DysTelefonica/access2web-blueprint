VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CarenciasExplicacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public ID As String
Public Apartado As String
Public Explicacion As String


Private m_sIDCalculado As String

Private m_objColCampos As Collection
Public Error As String
Private m_Error As String


Public Property Get IDCalculado() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    
    m_sIDCalculado = DameID("TbExplicacionCarencias", "ID", , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    IDCalculado = m_sIDCalculado
    
    Exit Property
    
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método CarenciasExplicacion.IDCalculado ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Function MotivoNoOK( _
                            p_ObjCarenciaExpliacionAlInicio As CarenciasExplicacion, _
                            Optional ByRef p_Error As String _
                            ) As String
    
    Dim m_CampoRepetido As EnumSiNo
    Dim blnComprobarCampo As Boolean
    On Error GoTo errores
    
     With Me
        If .Apartado = "" Then
            MotivoNoOK = "El apartado no puede estar vacío"
            Exit Function
        End If
        If .Explicacion = "" Then
            MotivoNoOK = "La explicación no puede estar vacía"
            Exit Function
        End If
        If p_ObjCarenciaExpliacionAlInicio Is Nothing Then
            blnComprobarCampo = True
        Else
            If .Apartado <> p_ObjCarenciaExpliacionAlInicio.Apartado Then
                blnComprobarCampo = True
            End If
        End If
        If blnComprobarCampo = True Then
            m_CampoRepetido = CampoRepetido("TbExplicacionCarencias", "Apartado", .Apartado, EnumSiNo.Sí, p_Error)
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            If m_CampoRepetido = EnumSiNo.Sí Then
                MotivoNoOK = "El Apartado ya ha sido usado"
                Exit Function
            End If
        End If
        
        
    End With
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CarenciasExplicacion.MotivoNoOK ha devuelto el error: " & Err.Description
    End If
End Function

Public Function Registrar( _
                            p_ObjCarenciaExpliacionAlInicio As CarenciasExplicacion, _
                            Optional ByRef p_Error As String _
                            ) As String
                            
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_MotivoNoOK As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
    
    m_MotivoNoOK = MotivoNoOK(p_ObjCarenciaExpliacionAlInicio, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    p_Error = m_MotivoNoOK
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If p_ObjCarenciaExpliacionAlInicio Is Nothing Then
        Me.ID = Me.IDCalculado
        p_Error = Me.Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
        m_SQL = "TbExplicacionCarencias"
    Else
        m_SQL = "SELECT * " & _
                "FROM TbExplicacionCarencias " & _
                "WHERE ID=" & Me.ID & ";"
    End If
    
   
    
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If p_ObjCarenciaExpliacionAlInicio Is Nothing Then
            .AddNew
                .Fields("ID") = Me.ID
                If Me.Apartado <> "" Then
                    .Fields("Apartado") = Me.Apartado
                End If
                If Me.Explicacion <> "" Then
                    .Fields("Explicacion") = Me.Explicacion
                End If
            .Update
        Else
            If .EOF Then
                p_Error = "No hay registro con ese ID"
                Err.Raise 1000
            End If
            .Edit
                
                If Me.Apartado <> "" Then
                    .Fields("Apartado") = Me.Apartado
                Else
                    .Fields("Apartado") = Null
                End If
                If Me.Explicacion <> "" Then
                    .Fields("Explicacion") = Me.Explicacion
                Else
                    .Fields("Explicacion") = Null
                End If
            .Update
        End If
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CarenciasExplicacion.Registrar ha devuelto el error: " & Err.Description
    End If
End Function

Public Function Borrar( _
                        Optional ByRef p_Error As String _
                        ) As String

    Dim m_SQL As String
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    m_SQL = "DELETE * " & _
                "FROM TbExplicacionCarencias " & _
                "WHERE ID=" & Me.ID & ";"
    
    getdb().Execute m_SQL
    
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CarenciasExplicacion.Borrar ha devuelto el error: " & Err.Description
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
        .Add "ID"
        .Add "Apartado"
        .Add "Explicacion"
        
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método CarenciasExplicacion.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property

Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    m_Error = ""
    
    If m_NombrePropiedad = "ID" Then
        getPropiedad = Me.ID
        Exit Function
    End If
    If m_NombrePropiedad = "Apartado" Then
        getPropiedad = Me.Apartado
        Exit Function
    End If
    If m_NombrePropiedad = "Explicacion" Then
        getPropiedad = Me.Explicacion
        Exit Function
    End If
    
    p_Error = "Propiedad no reconocida en CarenciasExplicacion.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CarenciasExplicacion.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function

Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "ID" Then
        Me.ID = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Apartado" Then
        Me.Apartado = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Explicacion" Then
        Me.Explicacion = m_ValorPropiedad
        Exit Function
    End If
    
    p_Error = "Propiedad no reconocida en CarenciasExplicacion.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método CarenciasExplicacion.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


