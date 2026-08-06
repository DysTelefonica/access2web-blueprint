Attribute VB_Name = "Form_FormAnexos"
Option Compare Database
Option Explicit
' --------------------------------------------------------------------------------
' MÓDULO: Form_F_Anexos
' OBJETIVO: Gestión de anexos con flexibilidad de subida y validación de duplicados
' ARQUITECTURA: HPS (Pragmática)
' --------------------------------------------------------------------------------

Private m_objEdicion As Edicion
Private m_ObjProyecto As Proyecto
Private m_ObjRiesgo As riesgo

Private m_Error As String

Private blnPermitidoEscribir As Boolean
Private m_ObjAnexoSeleccionado As Anexo
Private m_ParaProyecto As EnumSiNo
Private m_ParaEdicion As EnumSiNo
Private m_ParaRiesgo As EnumSiNo
Private m_EnUTE As String

' Eventos personalizados
Public Event AnexoAñadido(ByVal p_ObjAnexo As Anexo)
Public Event AnexoEliminado(ByVal IDAnexo As String)

' Variables internas
Private m_ObjAnexo As Anexo
Private m_ObjDELAnexo As Object
Private m_ObjColAnexos As Scripting.Dictionary
Dim m_URLAnexoSeleccionado As String

' Constantes de diseño
Const ALTO_FORM_CONUTE As Long = 9145
Const ALTO_FORM_SINUTE As Long = 7836
Const SUPERIOR_ANEXO_CONUTE As Long = 1679
Const SUPERIOR_ANEXO_SINUTE As Long = 1049
Const ALTO_CUADRO_PPAL_CONUTE As Long = 7753
Const ALTO_CUADRO_PPAL_SINUTE As Long = 6464
Const SUPERIOR_LISTA_CONUTE As Long = 2384
Const SUPERIOR_LISTA_SINUTE As Long = 959
Const SUPERIOR_TITULO_CONUTE As Long = 7709
Const SUPERIOR_TITULO_SINUTE As Long = 6284
Const SUPERIOR_CUADRO2_CONUTE As Long = 8459
Const SUPERIOR_CUADRO2_SINUTE As Long = 7150
Const SUPERIOR_SALIR_CONUTE As Long = 8518
Const SUPERIOR_SALIR_SINUTE As Long = 7228

' =================================================================================
' 1. BOTONES PRINCIPALES: EXAMINAR, ACEPTAR Y ACTUALIZAR
' =================================================================================

Private Sub btnExaminar_Click()
    ' --------------------------------------------------------------------------------
    ' BOTÓN: EXAMINAR
    ' Objetivo: Seleccionar archivo y sugerir nombre si está vacío.
    ' --------------------------------------------------------------------------------
    Dim fDialog As Object
    Dim strPath As String
    Dim fso As Object
    
    On Error GoTo errores
    
    Set fDialog = Application.FileDialog(3) ' 3 = msoFileDialogFilePicker
    
    With fDialog
        .Title = "Seleccionar documento para anexar"
        .AllowMultiSelect = False
        .filters.Clear
        .filters.Add "Todos los archivos", "*.*"
        
        If .Show = -1 Then
            strPath = .SelectedItems(1)
            Me.txtRuta.value = strPath
            
            ' Sugerir título si está vacío y no es Evidencia UTE
            If Nz(Me.EvidenciaUTE, "") <> "Sí" Then
                If Nz(Me.TituloDoc.value, "") = "" Then
                    Set fso = CreateObject("Scripting.FileSystemObject")
                    Me.TituloDoc.value = fso.GetBaseName(strPath)
                    Set fso = Nothing
                End If
            End If
        End If
    End With
    
    Set fDialog = Nothing
    Exit Sub

errores:
    MsgBox "Error al seleccionar archivo: " & Err.Description, vbCritical
End Sub

Private Sub btnAceptar_Click()
    ' --------------------------------------------------------------------------------
    ' BOTÓN: ACEPTAR
    ' Objetivo: Validar duplicados, instanciar objeto y guardar.
    ' --------------------------------------------------------------------------------
    Dim m_URL As String
    Dim m_Titulo As String
    Dim pregunta As Variant
    
    On Error GoTo errores
    
    ' 1. Validaciones UI
    If Nz(Me.txtRuta.value, "") = "" Then
        MsgBox "Debe seleccionar un archivo primero.", vbExclamation
        Me.btnExaminar.SetFocus
        Exit Sub
    End If
    
    m_Titulo = Nz(Me.TituloDoc.value, "")
    If Nz(Me.EvidenciaUTE, "") <> "Sí" And m_Titulo = "" Then
        MsgBox "El título es obligatorio.", vbExclamation
        Me.TituloDoc.SetFocus
        Exit Sub
    End If
    
    ' 2. Validación de Duplicados (Nuevo)
    ' Pasamos ID "0" para indicar que es nuevo
    If ExisteNombreEnAmbito(m_Titulo, "0") Then
        MsgBox "Ya existe un anexo con el título '" & m_Titulo & "'." & vbNewLine & _
               "Por favor, cambie el nombre antes de continuar.", vbExclamation, "Título Duplicado"
        Me.TituloDoc.SetFocus
        Exit Sub
    End If
    
    VBA.DoEvents
    DoCmd.Hourglass True
    m_Error = ""
    
    ' 3. Creación del Objeto
    Set m_ObjAnexo = New Anexo
    m_URL = Me.txtRuta.value
    
    With m_ObjAnexo
        .EvidenciaUTE = Nz(Me.EvidenciaUTE, "")
        
        ' Lógica UTE
        If m_ParaEdicion = EnumSiNo.Sí And Me.EvidenciaUTE = "Sí" Then
             m_Titulo = "ED " & m_objEdicion.Edicion & " Acuerdo tratamiento Riesgos con miembros de UTE"
             Me.TituloDoc = m_Titulo
        End If
        .Titulo = m_Titulo
        
        ' Asignación de IDs
        If m_ParaEdicion = EnumSiNo.Sí Then
            If Not m_objEdicion Is Nothing Then .IDEdicion = m_objEdicion.IDEdicion
            
            ' Control de sustitución UTE
            If .EvidenciaUTE = "Sí" And Not m_objEdicion.AnexoEvidenciaUTE Is Nothing Then
                DoCmd.Hourglass False
                pregunta = MsgBox("Ya existe evidencia UTE. ¿Desea sustituirla?", vbQuestion + vbYesNo)
                If pregunta <> vbYes Then Exit Sub
                m_objEdicion.AnexoEvidenciaUTE.EliminarAnexo
                DoCmd.Hourglass True
            End If
        End If
        
        If m_ParaProyecto = EnumSiNo.Sí And Not m_ObjProyecto Is Nothing Then .IDProyecto = m_ObjProyecto.IDProyecto
        If m_ParaRiesgo = EnumSiNo.Sí And Not m_ObjRiesgo Is Nothing Then .IDRiesgo = m_ObjRiesgo.IDRiesgo
        
        ' 4. Persistencia
        .Registrar m_ObjDELAnexo, m_URL, m_Error
        If m_Error <> "" Then Err.Raise 1000
    End With
    
    ' 5. Finalización
    RaiseEvent AnexoAñadido(m_ObjAnexo)
    EstablecerLista m_Error
    
    ' Limpiar UI
    Me.txtRuta.value = ""
    Me.TituloDoc.value = ""
    Me.TituloDoc.SetFocus
    
    DoCmd.Hourglass False
    MsgBox "Anexo guardado correctamente.", vbInformation
    Exit Sub

errores:
    DoCmd.Hourglass False
    If Err.Number <> 1000 Then
        MsgBox "Error: " & Err.Description, vbCritical
    Else
        MsgBox m_Error, vbExclamation
    End If
End Sub

Private Sub btnActualizarNombreArchivo_Click()
    ' --------------------------------------------------------------------------------
    ' BOTÓN: ACTUALIZAR NOMBRE (Renombrar)
    ' --------------------------------------------------------------------------------
    Dim strNuevoTitulo As String
    On Error GoTo errores
    
    If m_ObjAnexoSeleccionado Is Nothing Then Exit Sub
    
    strNuevoTitulo = Trim(Nz(Me.TituloDoc.value, ""))
    If strNuevoTitulo = "" Then
        MsgBox "El título no puede estar vacío.", vbExclamation
        Exit Sub
    End If
    
    If strNuevoTitulo = m_ObjAnexoSeleccionado.Titulo Then Exit Sub
    
    ' Validar duplicados excluyendo el propio ID
    If ExisteNombreEnAmbito(strNuevoTitulo, m_ObjAnexoSeleccionado.IDAnexo) Then
        MsgBox "Ya existe otro anexo con este nombre.", vbExclamation
        Exit Sub
    End If
    
    DoCmd.Hourglass True
    
    ' Actualizar Objeto y Persistencia
    m_ObjAnexoSeleccionado.CambiarNombre strNuevoTitulo, m_Error
    If m_Error <> "" Then Err.Raise 1000
    
    EstablecerLista m_Error
    Me.TituloDoc.SetFocus
    DoCmd.Hourglass False
    MsgBox "Nombre actualizado.", vbInformation
    Exit Sub

errores:
    DoCmd.Hourglass False
    MsgBox "Error al actualizar: " & Err.Description, vbCritical
End Sub

' =================================================================================
' 2. FUNCIONES DE VALIDACIÓN Y AYUDA
' =================================================================================

Private Function ExisteNombreEnAmbito(ByVal p_Titulo As String, ByVal p_IDExcluir As String) As Boolean
    ' Verifica si el título ya existe en la colección cargada (m_ObjColAnexos)
    Dim vItem As Variant
    Dim oAnexoTemp As Anexo
    
    ExisteNombreEnAmbito = False
    If m_ObjColAnexos Is Nothing Then Exit Function
    
    For Each vItem In m_ObjColAnexos.items
        If Not vItem Is Nothing Then
            Set oAnexoTemp = vItem
            ' Comparar Título ignorando mayúsculas
            If StrComp(Trim(oAnexoTemp.Titulo), Trim(p_Titulo), vbTextCompare) = 0 Then
                ' Verificar que no sea el mismo ID (para ediciones)
                If CStr(oAnexoTemp.IDAnexo) <> CStr(p_IDExcluir) Then
                    ExisteNombreEnAmbito = True
                    Exit Function
                End If
            End If
        End If
    Next
End Function

' =================================================================================
' 3. LÓGICA DE INTERFAZ Y EVENTOS EXISTENTES
' =================================================================================

Public Function EstablecerLista(Optional ByRef p_Error As String) As String
    Dim lst As ListBox
    Dim m_Filas As Collection
    Dim m_Fila As Variant

    On Error GoTo errores

    Set lst = Me.ListaDocumentos
    Me.Titulo = ""
    lst.RowSource = "Tipo;IDAnexo;Título;Riesgo;Edic_Anexado;Fecha"
    Me.ComandoEliminarAnexo.Enabled = False
    Me.ImagenAnexo.Visible = False

    ' B2 / Punto 15 — cache-first: ColAnexosTotales delega en
    ' modFormRiesgoDocumentosHelper.getAnexosDeEdicionCached (Edicion mode) o
    ' getAnexosDeRiesgoCached (Riesgo mode).
    Set m_ObjColAnexos = m_ObjDELAnexo.ColAnexosTotales

    p_Error = m_ObjDELAnexo.Error
    If p_Error <> "" Then Err.Raise 1000
    If m_ObjColAnexos Is Nothing Then Exit Function

    ' THIN FORM: toda la logica de transformacion vive en
    ' modAnexosListPresenter (testeable sin abrir el form).
    ' Esta funcion solo hace wiring UI: helper -> render.
    Set m_Filas = modAnexosListPresenter.BuildAnexosListRows(m_ObjColAnexos)
    For Each m_Fila In m_Filas
        lst.AddItem CStr(m_Fila)
    Next m_Fila
    lst.Requery
    ListaDocumentos_Click
    Exit Function

errores:
    If Err.Number <> 1000 Then p_Error = Err.Description
End Function

Private Sub ListaDocumentos_Click()
    Dim m_IDAnexoSeleccionado As String
    On Error GoTo errores
    
    VBA.DoEvents
    m_Error = ""
    
    Me.ComandoEliminarAnexo.Enabled = False
    Me.ImagenAnexo.Visible = False
    Me.btnActualizarNombreArchivo.Visible = False
    
    m_IDAnexoSeleccionado = Nz(Me.ListaDocumentos.Column(1), "")
    
    If m_IDAnexoSeleccionado = "" Then
        m_URLAnexoSeleccionado = ""
        Set m_ObjAnexoSeleccionado = Nothing
        Me.TituloDoc.value = ""
        Exit Sub
    End If
    
    Set m_ObjAnexoSeleccionado = Constructor.getAnexo(m_IDAnexoSeleccionado, m_Error)
    
    If m_Error <> "" Or m_ObjAnexoSeleccionado Is Nothing Then
        Me.TituloDoc.value = ""
        Exit Sub
    End If
    
    ' Visualización
    Me.TituloDoc.value = m_ObjAnexoSeleccionado.Titulo
    m_URLAnexoSeleccionado = m_ObjAnexoSeleccionado.URLAnexo
    
    If m_URLAnexoSeleccionado <> "" Then Me.ImagenAnexo.Visible = True
    
    ' Habilitar edición solo si está permitido
    If blnPermitidoEscribir Then
        Me.btnActualizarNombreArchivo.Visible = True
        Me.ComandoEliminarAnexo.Enabled = True
    End If
    
    Exit Sub
errores:
    MsgBox "Error en selección: " & Err.Description, vbCritical
End Sub

Private Sub ComandoEliminarAnexo_Click()
    Dim pregunta As Variant
    On Error GoTo errores
    
    If m_ObjAnexoSeleccionado Is Nothing Then Exit Sub
    
    pregunta = MsgBox("¿Desea eliminar el anexo seleccionado?", vbQuestion + vbYesNo + vbDefaultButton2)
    If pregunta <> vbYes Then Exit Sub
    
    DoCmd.Hourglass True
    m_ObjAnexoSeleccionado.EliminarAnexo m_Error
    If m_Error <> "" Then Err.Raise 1000
    
    RaiseEvent AnexoEliminado(m_ObjAnexoSeleccionado.IDAnexo)
    Set m_ObjAnexoSeleccionado = Nothing
    
    ' Recargar Padre para refrescar colección
    If m_ParaProyecto = EnumSiNo.Sí Then
        Set m_ObjDELAnexo = GetCachedProyecto(m_ObjProyecto.IDProyecto, m_Error)
    ElseIf m_ParaEdicion = EnumSiNo.Sí Then
        Set m_ObjDELAnexo = m_objEdicion ' La edición suele mantenerse en memoria viva
    ElseIf m_ParaRiesgo = EnumSiNo.Sí Then
        Set m_ObjDELAnexo = m_ObjRiesgo
    End If
    
    EstablecerLista m_Error
    DoCmd.Hourglass False
    Exit Sub
errores:
    DoCmd.Hourglass False
    MsgBox "Error al eliminar: " & IIf(m_Error <> "", m_Error, Err.Description), vbCritical
End Sub

Private Sub Form_Open(Cancel As Integer)
    Dim m_Argumento As String
    On Error GoTo errores
    
    DoCmd.Hourglass True
    m_Error = ""
    AjustarTamaño Me
    
    Me.EvidenciaUTE.Enabled = False
    Me.ComandoEliminarAnexo.Enabled = False
    Me.btnAceptar.Enabled = False
    
    m_Argumento = Nz(Me.openArgs(), "")
    
    ' Carga de contexto (Proyecto / Edición / Riesgo)
    If m_Argumento = "Proyecto" Then
        Set m_ObjProyecto = m_ObjProyectoActivo
        Set m_ObjDELAnexo = m_ObjProyecto
        m_ParaProyecto = EnumSiNo.Sí
        Me.lblTitulo1.Caption = "ANEXOS PROYECTO: " & UCase(m_ObjProyecto.Proyecto)
        
    ElseIf m_Argumento = "Edición" Then
        Set m_objEdicion = m_ObjEdicionActiva
        Set m_ObjProyecto = m_objEdicion.Proyecto
        Set m_ObjDELAnexo = m_objEdicion
        If m_ObjProyecto.EnUTE = "Sí" Then
            m_EnUTE = "Sí"
            Me.EvidenciaUTE.Enabled = True
        End If
        m_ParaEdicion = EnumSiNo.Sí
        Me.lblTitulo1.Caption = "ANEXOS EDICIÓN: " & m_objEdicion.Edicion
        
    ElseIf m_Argumento = "Riesgo" Then
        Set m_ObjRiesgo = m_ObjRiesgoActivo
        Set m_objEdicion = m_ObjRiesgo.Edicion
        Set m_ObjProyecto = m_objEdicion.Proyecto
        Set m_ObjDELAnexo = m_ObjRiesgo
        m_ParaRiesgo = EnumSiNo.Sí
        Me.lblTitulo1.Caption = "ANEXOS RIESGO: " & m_ObjRiesgo.CodigoRiesgo
    End If
    
    ' Permisos
    If m_ObjDELAnexo.EsActivo = EnumSiNo.Sí Then
        blnPermitidoEscribir = True
        Me.btnAceptar.Enabled = True
    Else
        blnPermitidoEscribir = False
    End If
    
    AdaptarTamañoFormulario m_Error
    EstablecerLista m_Error
    
    DoCmd.Hourglass False
    Exit Sub
errores:
    DoCmd.Hourglass False
    MsgBox "Error al abrir: " & Err.Description, vbCritical
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Function AdaptarTamañoFormulario(Optional ByRef p_Error As String) As String
    If m_EnUTE = "Sí" Then
        Me.lblEvidenciaUTE.Visible = True
        Me.EvidenciaUTE.Visible = True
    Else
        Me.lblEvidenciaUTE.Visible = False
        Me.EvidenciaUTE.Visible = False
    End If
    AdaptarTamañoFormulario = "OK"
End Function

Private Sub ImagenAnexo_Click()
    On Error Resume Next
    If m_URLAnexoSeleccionado <> "" Then
        AbrirEnLocal m_URLAnexoSeleccionado, Me.hWnd, m_Error
    End If
End Sub

Private Sub CmdSalir_Click()
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub ListaDocumentos_DblClick(Cancel As Integer)
    ImagenAnexo_Click
End Sub
