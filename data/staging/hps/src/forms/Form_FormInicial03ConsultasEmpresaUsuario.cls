Attribute VB_Name = "Form_FormInicial03ConsultasEmpresaUsuario"

Option Compare Database
Option Explicit

Private m_Error As String

' --- Eventos de Botones de Limpieza ---

Private Sub ComandoLimpiarEmpresaUsuario_Click()
    Me.IDEmpresaUsuario = Null
    IDEmpresaUsuario_AfterUpdate
End Sub

Private Sub ComandoLimpiarEstadoHPS_Click()
    Me.ComboEstadoHPS = Null
    ComboEstadoHPS_AfterUpdate
End Sub

' --- Eventos de Cambio de Valor ---

Private Sub IDEmpresaUsuario_AfterUpdate()
    On Error GoTo errores
    m_Error = ""
    Filtrar
    If m_Error <> "" Then Err.Raise 1000
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al IDEmpresaUsuario ha devuelto el error:" & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub ComboEstadoHPS_AfterUpdate()
    On Error GoTo errores
    m_Error = ""
    Filtrar
    If m_Error <> "" Then Err.Raise 1000
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al ComboEstadoHPS ha devuelto el error:" & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

' --- Lógica Principal ---

Private Sub Filtrar()
    Dim m_Suministrador As Suministrador
    
    On Error GoTo errores
    
    If m_ObjColParaFiltros Is Nothing Then
        Set m_ObjColParaFiltros = New Scripting.Dictionary
        m_ObjColParaFiltros.CompareMode = TextCompare
    End If
    
    ' 1. Filtro de Estado
    m_ObjColParaFiltros("Estado") = Nz(Me.ComboEstadoHPS, "")
    
    ' 2. Filtro de Empresa (CORREGIDO: Busca el Nombre Real por ID)
    ' Esto arregla el fallo de búsqueda: usamos el ID del combo para obtener el Nombre exacto
    ' que está guardado en la tabla de datos local.
    If Nz(Me.IDEmpresaUsuario, "") <> "" Then
        Set m_Suministrador = Constructor.getSuministrador(p_ID:=Me.IDEmpresaUsuario)
        If Not m_Suministrador Is Nothing Then
            m_ObjColParaFiltros("EmpresaUsuario") = m_Suministrador.Nombre
        Else
            m_ObjColParaFiltros("EmpresaUsuario") = ""
        End If
    Else
        m_ObjColParaFiltros("EmpresaUsuario") = ""
    End If
    
    ' 3. Aplicar filtro
    EstablecerFiltroUsuarios Me, m_Error
    If m_Error <> "" Then Err.Raise 1000
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Filtrar ha devuelto el error:" & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
End Sub

Private Sub Form_Load()
    On Error GoTo errores
    
    m_Error = ""
    Set m_ObjColParaFiltros = Nothing
    
    ' Configuración visual inicial
    EstablecerControlesDatos Me, m_Error
    If m_Error <> "" Then Err.Raise 1000
    
    ' Carga del Combo de Empresas (NUEVA LÓGICA)
    CargarEmpresasUsadas
    
    ' Carga del Combo de Estados
    RellenarComboEstado Me.ComboEstadoHPS, m_Error
    If m_Error <> "" Then Err.Raise 1000
    
    Exit Sub
errores:
    If Err.Number <> 1000 Then
        m_Error = "Al Form_Load ha devuelto el error:" & vbNewLine & Err.Description
        CorreoAlAdministrador m_Error
    End If
    pregunta = MsgBox(m_Error, vbCritical, "Error")
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

' --- Rutina de Carga Optimizada y Unificada ---
Private Sub CargarEmpresasUsadas()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim dictIDs As Object ' Scripting.Dictionary
    Dim vID As Variant
    Dim sNombre As String
    Dim sSQL As String
    
    On Error GoTo errores
    
    Set dictIDs = CreateObject("Scripting.Dictionary")
    Set db = getdb() ' Base de datos de Usuarios
    
    ' 1. Obtener IDs únicos de empresas en USUARIOS ACTIVOS
    ' Filtramos nulos para evitar registros vacíos
    sSQL = "SELECT DISTINCT IDEmpresaUsuario FROM TbUsuarios WHERE IDEmpresaUsuario IS NOT NULL"
    Set rs = db.OpenRecordset(sSQL)
    Do While Not rs.EOF
        dictIDs(CStr(rs(0))) = 1
        rs.MoveNext
    Loop
    rs.Close
    
    ' 2. Obtener IDs únicos de empresas en USUARIOS HISTÓRICOS
    sSQL = "SELECT DISTINCT IDEmpresaUsuario FROM TbUsuariosHistoricos WHERE IDEmpresaUsuario IS NOT NULL"
    Set rs = db.OpenRecordset(sSQL)
    Do While Not rs.EOF
        dictIDs(CStr(rs(0))) = 1
        rs.MoveNext
    Loop
    rs.Close
    Set db = Nothing
    
    ' 3. Rellenar el Combo cruzando los IDs con la tabla de Suministradores
    ' Usamos getdbExpedientes donde está la tabla TbSuministradores
    Set db = getdbExpedientes()
    Me.IDEmpresaUsuario.RowSource = ""
    
    ' Traemos Nombre y CIF para mostrar info completa
    sSQL = "SELECT IDSuministrador, Nombre, CIF FROM TbSuministradores ORDER BY Nombre"
    Set rs = db.OpenRecordset(sSQL)
    
    Do While Not rs.EOF
        vID = CStr(rs("IDSuministrador"))
        
        ' Solo añadimos si la empresa está siendo usada por algún usuario (Activo o Histórico)
        If dictIDs.Exists(vID) Then
            sNombre = Nz(rs("Nombre"), "")
            
            ' Evitamos añadir nombres vacíos al combo
            If sNombre <> "" Then
                ' Formato: ID ; Nombre (CIF)
                ' Esto asegura que el usuario vea el nombre real y no haya filas en blanco
                Me.IDEmpresaUsuario.AddItem rs("IDSuministrador") & ";" & sNombre & " (" & Nz(rs("CIF"), "") & ")"
            End If
        End If
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Set dictIDs = Nothing
    
    Exit Sub

errores:
    m_Error = "Error al CargarEmpresasUsadas: " & Err.Description
    Err.Raise 1000
End Sub