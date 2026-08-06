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

Private Function HTMLTablaAperturas( _
                                        p_Col As Scripting.Dictionary, _
                                        p_Titulo As String, _
                                        Optional ByRef p_Error As String _
                                        ) As String
    
    Dim m_FechaApertura As String
    Dim m_FechaCierre As String
    Dim m_mensaje As String
    
    
    Dim m_ID As Variant
    Dim m_Apertura As AplicacionApertura
    

    On Error GoTo errores
    
    If p_Col Is Nothing Then
        Exit Function
    End If
    m_mensaje = "<table>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td colspan='9' class=""ColespanArriba""> " & p_Titulo & " </td>"
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        m_mensaje = m_mensaje & "<tr>" & vbNewLine
            m_mensaje = m_mensaje & "<td class=""Cabecera"" > <strong>Aplicación</strong></td>"
            m_mensaje = m_mensaje & "<td class=""Cabecera"" > <strong>Usuario</strong></td>"
            m_mensaje = m_mensaje & "<td class=""Cabecera"" > <strong>Apertura</strong></td>"
            m_mensaje = m_mensaje & "<td class=""Cabecera"" > <strong>Cierre</strong></td>"
            m_mensaje = m_mensaje & "<td class=""Cabecera"" > <strong>En Oficina</strong></td>"
            m_mensaje = m_mensaje & "<td class=""Cabecera"" > <strong>Versión</strong></td>"
            m_mensaje = m_mensaje & "<td class=""Cabecera"" > <strong>Usuario Máquina</strong></td>"
            m_mensaje = m_mensaje & "<td class=""Cabecera"" > <strong>Máquina</strong></td>"
            m_mensaje = m_mensaje & "<td class=""Cabecera"" > <strong>Observaciones</strong></td>"
            
        m_mensaje = m_mensaje & "</tr>" & vbNewLine
        For Each m_ID In p_Col
            m_mensaje = m_mensaje & "<tr>" & vbNewLine
            Set m_Apertura = p_Col(m_ID)
            With m_Apertura
                m_mensaje = m_mensaje & "<td class=""centrado"" > " & .NombreAplicacion & "</td>" & vbNewLine
                m_mensaje = m_mensaje & "<td class=""centrado"" > " & .NombreUsuario & "</td>" & vbNewLine
                If Not IsDate(.HoraApertura) Then
                    m_FechaApertura = .FechaApertura
                Else
                    m_FechaApertura = .FechaApertura & " " & .HoraApertura
                End If
                m_mensaje = m_mensaje & "<td class=""centrado"" > " & m_FechaApertura & "</td>" & vbNewLine
                If Not IsDate(.FechaCierre) Then
                    m_FechaCierre = "&nbsp;"
                Else
                    m_FechaCierre = .FechaCierre & " " & .HoraCierre
                End If
                m_mensaje = m_mensaje & "<td class=""centrado"" > " & m_FechaCierre & "</td>" & vbNewLine
                m_mensaje = m_mensaje & "<td class=""centrado"" > " & .EnOficina & "</td>" & vbNewLine
                m_mensaje = m_mensaje & "<td class=""centrado"" > " & .VersionAplicacion & "</td>" & vbNewLine
                m_mensaje = m_mensaje & "<td class=""centrado"" > " & .UsuarioMaquina & "</td>" & vbNewLine
                m_mensaje = m_mensaje & "<td class=""centrado"" > " & .NombreMaquina & "</td>" & vbNewLine
                If .Observaciones <> "" Then
                    m_mensaje = m_mensaje & "<td class=""centrado"" > " & .Observaciones & "</td>" & vbNewLine
                Else
                    m_mensaje = m_mensaje & "<td class=""centrado"" >&nbsp;</td>" & vbNewLine
                End If
            End With
            Set m_Apertura = Nothing
            m_mensaje = m_mensaje & "</tr>" & vbNewLine
        Next
    m_mensaje = m_mensaje & "</table>" & vbNewLine
    
    HTMLTablaAperturas = m_mensaje
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HTML.HTMLTablaAperturas ha devuelto el error: " & vbNewLine & Err.Description
    End If
    
End Function

Private Function HTMLAperturas( _
                                p_Col As Scripting.Dictionary, _
                                p_Titulo As String, _
                                Optional ByRef p_Error As String _
                                ) As ADODB.stream
    
    
    Dim m_HTMLTablaAperturas As String
    
    On Error GoTo errores
    
    If p_Col Is Nothing Then
        Exit Function
    End If
    m_HTMLTablaAperturas = HTMLTablaAperturas(p_Col, p_Titulo, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    
    
    Me.Titulo = p_Titulo
    Set HTMLAperturas = New ADODB.stream
    HTMLAperturas.Open
    HTMLAperturas.WriteText Me.Cabecera & vbNewLine
    
    HTMLAperturas.WriteText m_HTMLTablaAperturas & vbNewLine
    
    HTMLAperturas.WriteText "</body>" & vbNewLine
    HTMLAperturas.WriteText "</html>" & vbNewLine
    HTMLAperturas.WriteText m_ObjEntorno.FinHTML & vbNewLine
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HTML.HTMLAperturas ha devuelto el error: " & vbNewLine & Err.Description
    End If
    
End Function
Public Function MostrarAperturas( _
                                    p_Col As Scripting.Dictionary, _
                                    p_Titulo As String, _
                                    Optional ByRef p_Error As String _
                                    ) As String
    
    Dim m_mensaje As ADODB.stream
    
    On Error GoTo errores
    
    Me.Error = ""
    Set m_mensaje = HTMLAperturas(p_Col, p_Titulo, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    MostrarHTML m_mensaje, p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método HTML.MostrarAperturas ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function


Public Function MostrarHTML( _
                            p_Mensaje As ADODB.stream, _
                            Optional ByRef p_Error As String _
                            ) As String
    
    
    
    On Error GoTo errores
    HTMLENTXT , p_Mensaje, , p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Function
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método HTML.MostrarHTML ha devuelto el error: " & vbNewLine & Err.Description
    End If
End Function





