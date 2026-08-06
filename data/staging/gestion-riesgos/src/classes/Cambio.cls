VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cambio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Public IDCambio As String
Public IDProyecto As String
Public EdicionFinal As String
Public EdicionInicial As String
Public riesgo As String
Public Descripcion As String
Public NombreCampo As String
Public FechaRegistro As String

Private m_ObjProyecto As Proyecto
Private m_objEdicionInicialObj As Edicion
Private m_objEdicionFinalObj As Edicion
Private m_sIDCambioCalculado As String


Private m_objColCampos As Collection
Public Error As String
Private m_Error As String


Public Property Get Proyecto() As Proyecto

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjProyecto Is Nothing Then
        Set Proyecto = m_ObjProyecto
        Exit Property
    End If
    If Me.IDProyecto = "" Then
        m_Error = "No se conoce el IDProyecto"
        Err.Raise 1000
    End If
    Set m_ObjProyecto = Constructor.getProyecto(Me.IDProyecto, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set Proyecto = m_ObjProyecto
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Cambio.Proyecto ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get EdicionInicialObj() As Edicion

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objEdicionInicialObj Is Nothing Then
        Set EdicionInicialObj = m_objEdicionInicialObj
        Exit Property
    End If
    If Me.EdicionInicial = "" Then
        Exit Property
    End If
    If Me.IDProyecto = "" Then
        Exit Property
    End If
    Set m_objEdicionInicialObj = Constructor.getEdicionV2(Me.IDProyecto, Me.EdicionInicial, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set EdicionInicialObj = m_objEdicionInicialObj
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Cambio.EdicionInicialObj ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get EdicionFinalObj() As Edicion

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_objEdicionFinalObj Is Nothing Then
        Set EdicionFinalObj = m_objEdicionFinalObj
        Exit Property
    End If
    If Me.EdicionFinal = "" Then
        Exit Property
    End If
    If Me.IDProyecto = "" Then
        Exit Property
    End If
    Set m_objEdicionFinalObj = Constructor.getEdicionV2(Me.IDProyecto, Me.EdicionFinalObj, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set EdicionFinalObj = m_objEdicionFinalObj
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Cambio.EdicionFinalobj ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property



Public Property Get IDCambioCalculado() As String
    
    On Error GoTo errores
    
    m_Error = ""
    Me.Error = ""
    
    m_sIDCambioCalculado = DameID("TbCambios", "IDCambio", getdb(), m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    IDCambioCalculado = m_sIDCambioCalculado
    
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Cambio.IDCambioCalculado ha devuelto el error " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Function Registrar( _
                            Optional ByRef p_Error As String _
                            ) As String
    
    
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_IDCambio As String
    
    On Error GoTo errores
    
    With Me
        
        If .IDProyecto = "" Then
            p_Error = "Falta IDProyecto"
            Err.Raise 1000
        End If
        If Not IsNumeric(.EdicionFinal) Then
            p_Error = "Falta EdicionFinal"
            Err.Raise 1000
        End If
        If Not IsNumeric(.EdicionInicial) Then
            If CInt(.EdicionFinal > 1) Then
                .EdicionInicial = CStr(CInt(.EdicionFinal) - 1)
            End If
        End If
        
        
    End With
    m_SQL = "tbCambios"

    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        .AddNew
            Me.IDCambio = Me.IDCambioCalculado
            .Fields("IDCambio") = Me.IDCambio
            .Fields("IDProyecto") = Me.IDProyecto
            If IsNumeric(Me.EdicionInicial) Then
                .Fields("EdicionInicial") = Me.EdicionInicial
            End If
            .Fields("EdicionFinal") = Me.EdicionFinal
            .Fields("Riesgo") = Me.riesgo
            .Fields("Descripcion") = Me.Descripcion
            If Me.NombreCampo <> "" Then
                .Fields("NombreCampo") = Me.NombreCampo
            End If
            .Fields("FechaRegistro") = Now()
            
        .Update
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método GrabarCambio ha devuelto el error: " & vbNewLine & Err.Description
    End If
    
End Function

Public Function Borrar( _
                            Optional p_IDCambio As String, _
                            Optional p_Edicion As String, _
                            Optional p_IDProyecto As String, _
                            Optional ByRef p_Error As String _
                            ) As String
    
    Dim m_SQL As String
    On Error GoTo errores
    
    If p_IDCambio = "" And p_Edicion = "" And p_IDProyecto = "" Then
        Exit Function
    End If
    If p_IDCambio <> "" Then
        m_SQL = "DELETE * " & _
                "FROM tbCambios " & _
                "WHERE IDCambio=" & p_IDCambio & ";"
        getdb().Execute m_SQL
        Exit Function
    End If
    
    If p_Edicion <> "" And p_IDProyecto <> "" Then
        m_SQL = "DELETE * " & _
                "FROM tbCambios " & _
                "WHERE EdicionFinal=" & p_Edicion & " AND IDProyecto=" & p_IDProyecto & ";"
        getdb().Execute m_SQL
        Exit Function
    End If
    If p_Edicion = "" And p_IDProyecto <> "" Then
        m_SQL = "DELETE * " & _
                "FROM tbCambios " & _
                "WHERE IDProyecto=" & p_IDProyecto & ";"
        getdb().Execute m_SQL
        Exit Function
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Borrar ha devuelto el error: " & vbNewLine & Err.Description
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
        .Add "IDCambio"
        .Add "IDProyecto"
        .Add "EdicionInicial"
        .Add "EdicionFinal"
        .Add "Riesgo"
        .Add "NombreCampo"
        .Add "Descripcion"
        .Add "FechaRegistro"
        
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método Cambio.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "IDCambio" Then
        getPropiedad = Me.IDCambio
        Exit Function
    End If
    If m_NombrePropiedad = "IDProyecto" Then
        getPropiedad = Me.IDProyecto
        Exit Function
    End If
    If m_NombrePropiedad = "EdicionInicial" Then
        getPropiedad = Me.EdicionInicial
        Exit Function
    End If
    If m_NombrePropiedad = "EdicionFinal" Then
        getPropiedad = Me.EdicionFinal
        Exit Function
    End If
    If m_NombrePropiedad = "Riesgo" Then
        getPropiedad = Me.riesgo
        Exit Function
    End If
    If m_NombrePropiedad = "NombreCampo" Then
        getPropiedad = Me.NombreCampo
        Exit Function
    End If
    If m_NombrePropiedad = "Descripcion" Then
        getPropiedad = Me.Descripcion
        Exit Function
    End If
    
    If m_NombrePropiedad = "FechaRegistro" Then
        getPropiedad = Me.FechaRegistro
        Exit Function
    End If
    
    p_Error = "Propiedad no reconocida en Cambio.getPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Cambio.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function

Public Function SetPropiedad(m_NombrePropiedad As Variant, m_ValorPropiedad As Variant, Optional ByRef p_Error As String) As String
    On Error GoTo errores

    Me.Error = ""
    p_Error = ""
    If m_NombrePropiedad = "IDCambio" Then
        Me.IDCambio = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDProyecto" Then
        Me.IDProyecto = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "EdicionInicial" Then
        Me.EdicionInicial = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "EdicionFinal" Then
        Me.EdicionFinal = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Riesgo" Then
        Me.riesgo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "NombreCampo" Then
        Me.NombreCampo = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Descripcion" Then
        Me.Descripcion = m_ValorPropiedad
        Exit Function
    End If
    
    If m_NombrePropiedad = "FechaRegistro" Then
        Me.FechaRegistro = m_ValorPropiedad
        Exit Function
    End If
    
    p_Error = "Propiedad no reconocida en Cambio.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método Cambio.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function


