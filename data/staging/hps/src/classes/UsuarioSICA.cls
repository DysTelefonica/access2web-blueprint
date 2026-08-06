VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UsuarioSICA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Public ID As String
Public IDHPS As String
Public IDHPSHistorico As String
Public Nombre As String
Public Apellido_1 As String
Public Apellido_2 As String
Public TramitacionExterna As String
Public UsuarioSICATSOL As String
Public UsuarioSICATDE As String
Public FechaCreacion As String
Public FechaModificacion As String
Public Correo_e As String
Public Observaciones As String
Public IDEmpresaTramitadora As String
Public FAltaTSOL As String
Public FAltaTdE As String
Public FbajaTSOL As String
Public FBajaTdE As String

Public ParaAlta As EnumSiNo
Private m_sURLAnexo As String
Private m_ObjAnexoUsuarioSICA As AnexoUsuarioSICA
Private m_sError As String
Private m_sNombreCompleto As String
Private m_objUsuarioHPS As UsuarioHPS
Private m_ObjUsuarioHistorico As UsuarioHistorico
Private m_ObjColCamposParaCambios As Scripting.Dictionary
Private m_objEmpresaTramitadoraCalculada As Suministrador
Private m_sIDNuevo As String


Private m_objColCampos As Collection
Public Error As String
Private m_Error As String




Public Property Get EmpresaTramitadoraCalculada() As Suministrador

    On Error GoTo errores
    
    Me.Error = ""
     m_Error = ""
    If Not m_objEmpresaTramitadoraCalculada Is Nothing Then
        Set EmpresaTramitadoraCalculada = m_objEmpresaTramitadoraCalculada
        Exit Property
    End If
    If Me.IDEmpresaTramitadora = "" Then
        Exit Property
    End If
    Set m_objEmpresaTramitadoraCalculada = Constructor.getSuministrador(Me.IDEmpresaTramitadora, , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set EmpresaTramitadoraCalculada = m_objEmpresaTramitadoraCalculada
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioSICA.EmpresaTramitadoraCalculada ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Get UsuarioHPS() As UsuarioHPS
    
    On Error GoTo errores
    
    Me.Error = ""
     m_Error = ""
    If Not m_objUsuarioHPS Is Nothing Then
        Set UsuarioHPS = m_objUsuarioHPS
        Exit Property
    End If
    If Me.IDHPS = "" Then
        Exit Property
    End If
    
    Set m_objUsuarioHPS = getUsuarioHPS(Me.IDHPS, , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set UsuarioHPS = m_objUsuarioHPS
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioSICA.UsuarioHPS ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get UsuarioHistorico() As UsuarioHistorico
    
    On Error GoTo errores
    
    Me.Error = ""
     m_Error = ""
    If Not m_ObjUsuarioHistorico Is Nothing Then
        Set UsuarioHistorico = m_ObjUsuarioHistorico
        Exit Property
    End If
    If Me.IDHPSHistorico = "" Then
        Exit Property
    End If
    
    Set m_ObjUsuarioHistorico = getUsuarioHistorico(Me.IDHPSHistorico, , m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set UsuarioHistorico = m_ObjUsuarioHistorico
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioSICA.UsuarioHistorico ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get AnexoSICA() As AnexoUsuarioSICA

    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If Not m_ObjAnexoUsuarioSICA Is Nothing Then
        Set AnexoSICA = m_ObjAnexoUsuarioSICA
        Exit Property
    End If
    If Me.ID = "" Then
        Exit Property
    End If
    Set m_ObjAnexoUsuarioSICA = Constructor.getAnexoUsuarioSICA(, Me.ID, m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    Set AnexoSICA = m_ObjAnexoUsuarioSICA
    
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioSICA.AnexoSICA ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Set AnexoSICA(m_sNew As AnexoUsuarioSICA)

    Set m_ObjAnexoUsuarioSICA = m_sNew

End Property

Public Property Get URLAnexo() As String
    
    
    On Error GoTo errores
    
    Me.Error = ""
     m_Error = ""
    If m_sURLAnexo <> "" Then
        URLAnexo = m_sURLAnexo
        Exit Property
    End If
    If Me.AnexoSICA Is Nothing Then
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
        Exit Property
    Else
        m_Error = Me.Error
        If m_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    m_sURLAnexo = Me.AnexoSICA.URLAnexo
    m_Error = Me.AnexoSICA.Error
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    URLAnexo = m_sURLAnexo
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioSICA.URLAnexo ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error

End Property

Public Property Let URLAnexo(ByVal sNewValue As String)

    m_sURLAnexo = sNewValue

End Property


Public Property Get NombreCompleto() As String
    
    On Error GoTo errores
    
    Me.Error = ""
    m_Error = ""
    If m_sNombreCompleto <> "" Then
        NombreCompleto = m_sNombreCompleto
        Exit Property
    End If
    If Me.Nombre = "" And Me.Apellido_1 = "" And Me.Apellido_2 = "" Then
        Exit Property
    End If
    m_sNombreCompleto = Me.Nombre & " " & Me.Apellido_1 & " " & Me.Apellido_2
    NombreCompleto = m_sNombreCompleto
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioSICA.NombreCompleto ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property

Public Property Get IDNuevo() As String
    
    On Error GoTo errores
    
    Me.Error = ""
     m_Error = ""
    m_sIDNuevo = getID(m_Error)
    If m_Error <> "" Then
        Err.Raise 1000
    End If
    
    IDNuevo = m_sIDNuevo
    
    Exit Property
errores:
    If m_Error <> "" Then
        m_Error = "El método UsuarioSICA.IDNuevo ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Me.Error = m_Error
End Property


Public Function PasarAHistorico( _
                                p_IDHistorico As String, _
                                Optional ByRef p_Error As String _
                                ) As String
                        
    Dim m_SQL As String
   
    On Error GoTo errores
    
    If p_IDHistorico = "" Then
        Exit Function
    End If
    If Me.ID = "" Then
        Exit Function
    End If
    m_SQL = "UPDATE TbUsuariosSICA SET TbUsuariosSICA.IDHPS = Null, TbUsuariosSICA.IDHPSHistorico =" & p_IDHistorico & " " & _
            "WHERE (((TbUsuariosSICA.ID)='" & Me.ID & "'));"

    getdb().Execute m_SQL
    Call RefreshSicaLocalCaches(Me.ID, p_Error:=p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método PasarAHistorico ha devuelto el error: " & Err.Description
    End If
End Function

Public Function PasarAActual( _
                                p_ID As String, _
                                Optional ByRef p_Error As String _
                                ) As String
                        
    Dim m_SQL As String
   
    On Error GoTo errores
    
    If p_ID = "" Then
        Exit Function
    End If
    If Me.IDHPSHistorico = "" Then
        Exit Function
    End If
    m_SQL = "UPDATE TbUsuariosSICA SET TbUsuariosSICA.IDHPS =" & p_ID & ", TbUsuariosSICA.IDHPSHistorico = Null " & _
            "WHERE (((TbUsuariosSICA.ID)='" & Me.ID & "'));"

    getdb().Execute m_SQL
    Call RefreshSicaLocalCaches(Me.ID, p_Error:=p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método PasarAActual ha devuelto el error: " & Err.Description
    End If
End Function

Private Function getID( _
                        Optional ByRef p_Error As String _
                        ) As String
                        
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_ID As Long
    On Error GoTo errores
    
    
    m_SQL = "SELECT Max(Right([ID],4)) AS MáxID " & _
            "FROM TbUsuariosSICA;"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    
    With rcdDatos
        If Not .EOF Then
            If IsNumeric(Nz(.Fields("MáxID"), "")) Then
                m_ID = .Fields("MáxID")
            End If
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    
    getID = "SICA" & Format(m_ID + 1, "0000")
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método getID ha devuelto el error: " & Err.Description
    End If
End Function
Public Function MotivoNoOK( _
                                Optional ByRef p_ObjUsuarioSICAAlInicio As UsuarioSICA, _
                                Optional p_URLocal As String, _
                                Optional ByRef p_Error As String _
                                ) As String
    Dim blnValorRepetido As Boolean
    
    On Error GoTo errores
        
    With Me
        If .ParaAlta = Empty Then
            MotivoNoOK = "No se sabe si es alta o edición"
            Exit Function
            
        End If
        If p_URLocal <> "" Then
            If Not fso.FileExists(p_URLocal) Then
                MotivoNoOK = "El Anexo no es alcanzable"
                Exit Function
            End If
            If FicheroAbierto(p_URLocal) Then
                MotivoNoOK = "El Anexo parece estar abierto"
                Exit Function
            End If
            If Not p_ObjUsuarioSICAAlInicio Is Nothing Then
                If Not p_ObjUsuarioSICAAlInicio.AnexoSICA Is Nothing Then
                    If fso.FileExists(p_ObjUsuarioSICAAlInicio.AnexoSICA.URLAnexo) Then
                        If FicheroAbierto(p_ObjUsuarioSICAAlInicio.AnexoSICA.URLAnexo) Then
                            MotivoNoOK = "El Usuario sica ya tenía un anexo y éste parece que está abierto o es de sólo lectura"
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
        
        
        If .ParaAlta = EnumSiNo.No Then
            
            If p_ObjUsuarioSICAAlInicio Is Nothing Then
                Set p_ObjUsuarioSICAAlInicio = Constructor.getUsuarioSICA(Me.IDHPS, , , p_Error)
                If p_Error <> "" Then
                    Err.Raise 1000
                End If
            End If
            
            If p_ObjUsuarioSICAAlInicio Is Nothing Then
                MotivoNoOK = "Es una edición de usuario SICA y no se ha encontrado registrado"
                Exit Function
            End If
        End If
        If Me.TramitacionExterna <> "Sí" And Me.TramitacionExterna <> "No" Then
            MotivoNoOK = "La TramitacionExterna es obligatoria"
            Exit Function
        End If
        If Me.TramitacionExterna = "Sí" Then
            If Me.IDHPS <> "" Or Me.IDHPSHistorico <> "" Then
                MotivoNoOK = "Si la tramitación es externa no se puede usar un usuario HPS, ni histórico"
                Exit Function
            End If
        Else
            If Me.IDHPS = "" And Me.IDHPSHistorico = "" Then
                MotivoNoOK = "Si la tramitación es interna Se ha de indicar el usuario HPS o el histórico"
                Exit Function
            End If
            If Me.IDHPS <> "" And Me.IDHPSHistorico <> "" Then
                MotivoNoOK = "Si la tramitación es interna Se ha de indicar el usuario HPS o el histórico. Sólo uno"
                Exit Function
            End If
        End If
        
        If .Nombre = "" Then
            MotivoNoOK = "El Nombre es obligatorio"
            Exit Function
        End If
        If .Apellido_1 = "" Then
            MotivoNoOK = "El Apellido_1 es obligatorio"
            Exit Function
        End If
        
        If .UsuarioSICATSOL = "" And UsuarioSICATDE = "" Then
            MotivoNoOK = "Uno de los dos usuario del sistema acreditado es obligatorio"
            Exit Function
        End If
        If .UsuarioSICATSOL <> "" Then
            If Not IsDate(.FAltaTSOL) Then
                MotivoNoOK = "Se ha de indicar la fecha de alta del usuario SICATSOL"
                Exit Function
            End If
            If IsDate(.FbajaTSOL) Then
                If CDate(.FbajaTSOL) < CDate(.FAltaTSOL) Then
                    MotivoNoOK = "La fecha de baja de SICATSOL no puede ser anterior al alta"
                    Exit Function
                End If
            End If
        End If
        If .UsuarioSICATDE <> "" Then
            If Not IsDate(.FAltaTdE) Then
                MotivoNoOK = "Se ha de indicar la fecha de alta del usuario SICATDE"
                Exit Function
            End If
            If IsDate(.FBajaTdE) Then
                If CDate(.FBajaTdE) < CDate(.FAltaTdE) Then
                    MotivoNoOK = "La fecha de baja de SICATDE no puede ser anterior al alta"
                    Exit Function
                End If
            End If
        End If
        If .EmpresaTramitadoraCalculada Is Nothing Then
            p_Error = .Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
            MotivoNoOK = "La empresa tramitadora es obligatoria"
            Exit Function
        Else
            p_Error = .Error
            If p_Error <> "" Then
                Err.Raise 1000
            End If
        End If
        
        If InStr(1, .Correo_e, "@") = 0 Then
            MotivoNoOK = "Falta el Correo electrónico o no tiene la forma adecuada"
            Exit Function
        End If
        If .ParaAlta = EnumSiNo.No Then
            If .Correo_e <> p_ObjUsuarioSICAAlInicio.Correo_e Then
                If Me.IDHPS <> "" Then
                     blnValorRepetido = ValorRepetido("TbUsuariosSICA", "Correo_e", .Correo_e, getdb(), p_Error)
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                Else
                    blnValorRepetido = ValorRepetido("TbUsuariosHistoricos", "Correo_e", .Correo_e, getdb(), p_Error)
                    If p_Error <> "" Then
                        Err.Raise 1000
                    End If
                End If
               
                If blnValorRepetido Then
                    MotivoNoOK = "El Correo_e ha sido registrado para otro usuario"
                    Exit Function
                End If
            End If
            If .UsuarioSICATSOL <> "" Then
                If p_ObjUsuarioSICAAlInicio.UsuarioSICATSOL <> .UsuarioSICATSOL Then
                    If Me.IDHPS <> "" Then
                        blnValorRepetido = ValorRepetido("TbUsuariosSICA", "UsuarioSICATSOL", .UsuarioSICATSOL, getdb(), p_Error)
                        If p_Error <> "" Then
                            Err.Raise 1000
                        End If
                    Else
                        blnValorRepetido = ValorRepetido("TbUsuariosHistoricos", "UsuarioSICATSOL", .UsuarioSICATSOL, getdb(), p_Error)
                        If p_Error <> "" Then
                            Err.Raise 1000
                        End If
                    End If
                    
                    If blnValorRepetido Then
                        MotivoNoOK = "El UsuarioSICATSOL ha sido registrado para otro usuario"
                        Exit Function
                    End If
                End If
                
            End If
            If .UsuarioSICATDE <> "" Then
                If p_ObjUsuarioSICAAlInicio.UsuarioSICATDE <> .UsuarioSICATDE Then
                
                    If Me.IDHPS <> "" Then
                        blnValorRepetido = ValorRepetido("TbUsuariosSICA", "UsuarioSICATDE", .UsuarioSICATDE, getdb(), p_Error)
                        If p_Error <> "" Then
                            Err.Raise 1000
                        End If
                    Else
                        blnValorRepetido = ValorRepetido("TbUsuariosHistoricos", "UsuarioSICATDE", .UsuarioSICATDE, getdb(), p_Error)
                        If p_Error <> "" Then
                            Err.Raise 1000
                        End If
                    End If
                    If blnValorRepetido Then
                        MotivoNoOK = "El UsuarioSICATDE ha sido registrado para otro usuario"
                        Exit Function
                    End If
                End If
                
            End If
        End If
        
        
        
    End With
    
    Exit Function
     
errores:
     If Err.Number <> 1000 Then
         p_Error = "El método UsuarioSICA.MotivoNoOK ha devuelto el error: " & Err.Description
     End If
                                
End Function



Private Function ValidacionEliminar( _
                                    Optional ByRef p_Error As String _
                                    ) As String
                                        
    Dim m_URLAnexo As String
    Dim m_AnexoUsuarioSICA As AnexoUsuarioSICA
    On Error GoTo errores
    
    If Application.TempVars("SoloLectura") = "Sí" Then
        p_Error = "El Sistema está en modo sólo lectura"
        Err.Raise 1000
    End If
    Set m_AnexoUsuarioSICA = Me.AnexoSICA
    If Not m_AnexoUsuarioSICA Is Nothing Then
        m_URLAnexo = m_AnexoUsuarioSICA.URLAnexo
        If fso.FileExists(m_URLAnexo) Then
            If FicheroAbierto(m_URLAnexo) Then
                p_Error = "Tiene un anexo que está abierto o es de sólo lectura"
                Err.Raise 1000
            End If
           
        End If
    End If
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método UsuarioSICA.ValidacionEliminar ha devuelto el error: " & Err.Description
    End If
End Function

Public Function Edicion( _
                            ByRef p_ObjUsuarioSICAAlInicio As UsuarioSICA, _
                            Optional ByRef p_Error As String _
                            ) As String
                                        
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_MotivoNoOK As String
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    If p_ObjUsuarioSICAAlInicio Is Nothing Then
        p_Error = "Faltan datos iniciales"
        Err.Raise 1000
    End If
    
    m_MotivoNoOK = MotivoNoOK(p_ObjUsuarioSICAAlInicio, , p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    p_Error = m_MotivoNoOK
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_SQL = "SELECT * FROM TbUsuariosSICA " & _
            "WHERE ID='" & Me.ID & "';"
        
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If .EOF Then
            p_Error = "No se ha encontrado el registro"
            Err.Raise 1000
        End If
        
        .Edit
            If Me.IDHPS <> "" Then
                .Fields("IDHPS") = Me.IDHPS
            Else
                .Fields("IDHPS") = Null
            End If
            If Me.IDHPSHistorico <> "" Then
                .Fields("IDHPSHistorico") = Me.IDHPSHistorico
            Else
                .Fields("IDHPSHistorico") = Null
            End If
            .Fields("Nombre") = Me.Nombre
            .Fields("Apellido_1") = Me.Apellido_1
            If Me.Apellido_2 <> "" Then
                .Fields("Apellido_2") = Me.Apellido_2
            Else
                .Fields("Apellido_2") = Null
            End If
            .Fields("TramitacionExterna") = Me.TramitacionExterna
            .Fields("FechaModificacion") = Now()
            .Fields("Correo_e") = Me.Correo_e
            .Fields("IDEmpresaTramitadora") = Me.IDEmpresaTramitadora
            If Me.UsuarioSICATSOL <> "" Then
                .Fields("UsuarioSICATSOL") = Me.UsuarioSICATSOL
                .Fields("FAltaTSOL") = Me.FAltaTSOL
                If IsDate(Me.FbajaTSOL) Then
                    .Fields("FbajaTSOL") = Me.FbajaTSOL
                Else
                    .Fields("FbajaTSOL") = Null
                End If
            Else
                .Fields("UsuarioSICATSOL") = Null
                .Fields("FAltaTSOL") = Null
                .Fields("FbajaTSOL") = Null
            End If
            If Me.UsuarioSICATDE <> "" Then
                .Fields("UsuarioSICATDE") = Me.UsuarioSICATDE
                .Fields("FAltaTdE") = Me.FAltaTdE
                If IsDate(Me.FBajaTdE) Then
                    .Fields("FBajaTdE") = Me.FBajaTdE
                End If
            Else
                .Fields("UsuarioSICATDE") = Null
                .Fields("FAltaTdE") = Null
                 .Fields("FBajaTdE") = Null
            End If
            If Me.Observaciones <> "" Then
                .Fields("Observaciones") = Me.Observaciones
            Else
                .Fields("Observaciones") = Null
            End If
        .Update
    
       
        
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método UsuarioSICA.registrar ha devuelto el error: " & Err.Description
    End If
End Function
Public Function Anexar( _
                            ByRef p_URLLocal As String, _
                            Optional ByRef p_Error As String _
                            ) As String
                                        
    Dim m_Anexo As AnexoUsuarioSICA
    Dim m_AnexoAlInicio As AnexoUsuarioSICA
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    If Not fso.FileExists(p_URLLocal) Then
        p_Error = "No es alcanzable el archivo a anexar"
        Err.Raise 1000
    End If
    If FicheroAbierto(p_URLLocal) Then
        p_Error = "Ha de cerrar el anexo local"
        Err.Raise 1000
    End If
    Set m_AnexoAlInicio = Me.AnexoSICA
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_AnexoAlInicio Is Nothing Then
        If fso.FileExists(m_AnexoAlInicio.URLAnexo) Then
            If FicheroAbierto(m_AnexoAlInicio.URLAnexo) Then
                p_Error = "El Usuario SICA ya tenía un adjunto y parece estar abierto o de sólo lectura"
                Err.Raise 1000
            End If
        End If
        m_AnexoAlInicio.Borrar p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    Set m_Anexo = New AnexoUsuarioSICA
    With m_Anexo
        .IDUsuarioSICA = Me.ID
        .URLLocal = p_URLLocal
        .Registrar p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End With
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método UsuarioSICA.Anexar ha devuelto el error: " & Err.Description
    End If
End Function

Public Function QuitarAnexo( _
                            Optional ByRef p_Error As String _
                            ) As String
                                        
    Dim m_Anexo As AnexoUsuarioSICA
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    Set m_Anexo = Me.AnexoSICA
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_Anexo Is Nothing Then
        m_Anexo.Borrar p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método UsuarioSICA.QuitarAnexo ha devuelto el error: " & Err.Description
    End If
End Function
Public Function Alta( _
                        Optional p_URLocal As String, _
                        Optional ByRef p_Error As String _
                        ) As String
                                        
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_MotivoNoOK As String
    
    On Error GoTo errores
    
    Me.Error = ""
    p_Error = ""
    
    m_MotivoNoOK = MotivoNoOK(, p_URLocal, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    p_Error = m_MotivoNoOK
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    m_SQL = "TbUsuariosSICA"
    Me.ID = Me.IDNuevo
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        .AddNew
            
            .Fields("ID") = Me.ID
            If Me.IDHPS <> "" Then
                .Fields("IDHPS") = Me.IDHPS
            End If
            If Me.IDHPSHistorico <> "" Then
                .Fields("IDHPSHistorico") = Me.IDHPSHistorico
            End If
            .Fields("Nombre") = Me.Nombre
            .Fields("Apellido_1") = Me.Apellido_1
            If Me.Apellido_2 <> "" Then
                .Fields("Apellido_2") = Me.Apellido_2
            End If
            .Fields("TramitacionExterna") = Me.TramitacionExterna
            .Fields("FechaCreacion") = Now()
            .Fields("FechaModificacion") = Now()
            .Fields("Correo_e") = Me.Correo_e
            .Fields("IDEmpresaTramitadora") = Me.IDEmpresaTramitadora
            If Me.UsuarioSICATSOL <> "" Then
                .Fields("UsuarioSICATSOL") = Me.UsuarioSICATSOL
                .Fields("FAltaTSOL") = Me.FAltaTSOL
                If IsDate(Me.FbajaTSOL) Then
                    .Fields("FbajaTSOL") = Me.FbajaTSOL
                End If
            End If
            If Me.UsuarioSICATDE <> "" Then
                .Fields("UsuarioSICATDE") = Me.UsuarioSICATDE
                .Fields("FAltaTdE") = Me.FAltaTdE
                If IsDate(Me.FBajaTdE) Then
                    .Fields("FBajaTdE") = Me.FBajaTdE
                End If
            End If
            If Me.Observaciones <> "" Then
                .Fields("Observaciones") = Me.Observaciones
            End If
        .Update
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    If p_URLocal <> "" Then
        Anexar p_URLocal, p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    
    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método UsuarioSICA.Alta ha devuelto el error: " & Err.Description
    End If
End Function
Public Function Borrar( _
                        Optional ByRef p_Error As String _
                        ) As String
    Dim m_SQL As String
    Dim m_Anexo As AnexoUsuarioSICA
    On Error GoTo errores
    
    
    
    ValidacionEliminar p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    Set m_Anexo = Me.AnexoSICA
    p_Error = Me.Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    If Not m_Anexo Is Nothing Then
        m_Anexo.Borrar p_Error
        If p_Error <> "" Then
            Err.Raise 1000
        End If
    End If
    m_SQL = "DELETE * FROM TbUsuariosSICA " & _
            "WHERE ID='" & Me.ID & "';"
    getdb().Execute m_SQL
    Call DeleteSicaLocalCaches(Me.ID, p_Error:=p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    

    Exit Function
    
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método UsuarioSICA.Borrar ha devuelto el error: " & Err.Description
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
        .Add "IDHPS"
        .Add "IDHPSHistorico"
        .Add "Nombre"
        .Add "Apellido_1"
        .Add "Apellido_2"
        .Add "TramitacionExterna"
        .Add "UsuarioSICATSOL"
        .Add "UsuarioSICATDE"
        .Add "FechaCreacion"
        .Add "FechaModificacion"
        .Add "Correo_e"
        .Add "Observaciones"
        .Add "IDEmpresaTramitadora"
        .Add "FAltaTSOL"
        .Add "FAltaTdE"
        .Add "FbajaTSOL"
        .Add "FBajaTdE"
        
    End With
    Set ColCampos = m_objColCampos
    Exit Property
errores:
    If Err.Number <> 1000 Then
        m_Error = "El método UsuariosSICA.ColCampos ha devuelto el error:" & vbNewLine & Err.Description
    End If
    m_Error = Me.Error
End Property


Public Function getPropiedad(m_NombrePropiedad As Variant, Optional ByRef p_Error As String) As Variant

    On Error GoTo errores

    Me.Error = ""
    If m_NombrePropiedad = "ID" Then
        getPropiedad = Me.ID
        Exit Function
    End If
    If m_NombrePropiedad = "IDHPS" Then
        getPropiedad = Me.IDHPS
        Exit Function
    End If
    If m_NombrePropiedad = "IDHPSHistorico" Then
        getPropiedad = Me.IDHPSHistorico
        Exit Function
    End If
    If m_NombrePropiedad = "Nombre" Then
        getPropiedad = Me.Nombre
        Exit Function
    End If
    If m_NombrePropiedad = "Apellido_1" Then
        getPropiedad = Me.Apellido_1
        Exit Function
    End If
    If m_NombrePropiedad = "Apellido_2" Then
        getPropiedad = Me.Apellido_2
        Exit Function
    End If
    If m_NombrePropiedad = "TramitacionExterna" Then
        getPropiedad = Me.TramitacionExterna
        Exit Function
    End If
    If m_NombrePropiedad = "UsuarioSICATSOL" Then
        getPropiedad = Me.UsuarioSICATSOL
        Exit Function
    End If
    If m_NombrePropiedad = "UsuarioSICATDE" Then
        getPropiedad = Me.UsuarioSICATDE
        Exit Function
    End If
    If m_NombrePropiedad = "FechaCreacion" Then
        getPropiedad = Me.FechaCreacion
        Exit Function
    End If
    If m_NombrePropiedad = "FechaModificacion" Then
        getPropiedad = Me.FechaModificacion
        Exit Function
    End If
    If m_NombrePropiedad = "Correo_e" Then
        getPropiedad = Me.Correo_e
        Exit Function
    End If
    
    If m_NombrePropiedad = "Observaciones" Then
        getPropiedad = Me.Observaciones
        Exit Function
    End If
    If m_NombrePropiedad = "IDEmpresaTramitadora" Then
        getPropiedad = Me.IDEmpresaTramitadora
        Exit Function
    End If
    If m_NombrePropiedad = "FAltaTSOL" Then
        getPropiedad = Me.FAltaTSOL
        Exit Function
    End If
    If m_NombrePropiedad = "FAltaTdE" Then
        getPropiedad = Me.FAltaTdE
        Exit Function
    End If
    If m_NombrePropiedad = "FbajaTSOL" Then
        getPropiedad = Me.FbajaTSOL
        Exit Function
    End If
    If m_NombrePropiedad = "FBajaTdE" Then
        getPropiedad = Me.FBajaTdE
        Exit Function
    End If
    
    p_Error = "Propiedad no reconocida en UsuariosSICA.getPropiedad"
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método UsuariosSICA.getPropiedad ha devuelto el error:" & vbNewLine & Err.Description
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
    If m_NombrePropiedad = "IDHPS" Then
        Me.IDHPS = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDHPSHistorico" Then
        Me.IDHPSHistorico = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Nombre" Then
        Me.Nombre = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Apellido_1" Then
        Me.Apellido_1 = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Apellido_2" Then
        Me.Apellido_2 = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "TramitacionExterna" Then
        Me.TramitacionExterna = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "UsuarioSICATSOL" Then
        Me.UsuarioSICATSOL = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "UsuarioSICATDE" Then
        Me.UsuarioSICATDE = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaCreacion" Then
        Me.FechaCreacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FechaModificacion" Then
        Me.FechaModificacion = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "Correo_e" Then
        Me.Correo_e = m_ValorPropiedad
        Exit Function
    End If
    
    If m_NombrePropiedad = "Observaciones" Then
        Me.Observaciones = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "IDEmpresaTramitadora" Then
        Me.IDEmpresaTramitadora = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FAltaTSOL" Then
        Me.FAltaTSOL = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FAltaTdE" Then
        Me.FAltaTdE = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FbajaTSOL" Then
        Me.FbajaTSOL = m_ValorPropiedad
        Exit Function
    End If
    If m_NombrePropiedad = "FBajaTdE" Then
        Me.FBajaTdE = m_ValorPropiedad
        Exit Function
    End If
    p_Error = "Propiedad no reconocida en UsuariosSICA.SetPropiedad"
    Err.Raise 1000
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método UsuariosSICA.SetPropiedad ha devuelto el error:" & vbNewLine & Err.Description
    End If
End Function







