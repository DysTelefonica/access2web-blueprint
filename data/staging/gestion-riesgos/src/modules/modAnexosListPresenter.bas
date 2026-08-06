Attribute VB_Name = "modAnexosListPresenter"
Option Compare Database
Option Explicit

' ============================================================
' modAnexosListPresenter
'   B2 / Punto 15 - UI row formatting para Form_FormAnexos.
'   Skill: access-vba-e2e-methodology (form thin, helper testable).
'
'   Convierte un Dictionary<Anexo> (poblado por cache-first) en una
'   Collection<String> de filas formateadas para ListBox.AddItem.
'
'   Formato por fila: "Tipo;IDAnexo;Titulo;CodRiesgo;Edic_Anexado;Fecha"
'     - Tipo:           derivado: "E" si IDRiesgo vacio, "R" si IDRiesgo lleno
'                       (contrato: anexo directo de edicion no tiene riesgo padre;
'                        anexo de riesgo hijo tiene IDRiesgo lleno).
'                       NO se usa p_ObjAnexo.Tipo porque ese campo no se popula
'                       del SQL (issue #83 pre-existente, Anexo.Tipo vacio siempre).
'     - IDAnexo:        p_ObjAnexo.IDAnexo
'     - Titulo:         p_ObjAnexo.Titulo
'     - CodRiesgo:      solo si Tipo="R" y objAnexo.riesgo existe
'     - Edic_Anexado:   solo si objAnexo.IDEdicion <> "" (directos)
'     - Fecha:          dd/mm/yyyy si FechaAnexo es fecha, sino ""
'
'   Honest signature: solo la coleccion de Anexos. Sin DAO.Database
'   (los Anexos ya estan cargados en memoria por el caller via cache-first).
'   NO requiere UI: los atoms TDD llaman al helper y assertan sobre el
'   Collection<String> resultante, sin abrir ningun formulario.
'
'   Contrato verificado por Test_modAnexosListPresenter (4 atoms):
'     - BuildAnexosListRows_Happy_5Anexos_5FilasOrdenadas
'     - BuildAnexosListRows_Empty_Nothing_RetornaCollectionVacia
'     - BuildAnexosListRows_DirectoSinEdicion_EdicAnexadoVacio
'     - BuildAnexosListRows_Riesgo_CodRiesgoPoblado
' ============================================================

Public Function BuildAnexosListRows(ByVal p_ColAnexos As Scripting.Dictionary) As Collection
    Dim colFilas As New Collection
    
    If p_ColAnexos Is Nothing Then
        Set BuildAnexosListRows = colFilas
        Exit Function
    End If
    
    ' Pre-cargar display values por ID unico para evitar N+1.
    ' Cada Anexo lazy-loads su padre via Constructor (1 SELECT por ID unico);
    ' la instancia queda cacheada en m_objEdicion / m_objRiesgo del Anexo.
    ' Sin este preload, B2 fixture pagaba 5 SELECTs; con el preload, paga 3
    ' (1 Edicion + 2 Riesgos unicos). En formularios grandes (decenas de
    ' anexos directos al mismo Edicion), la mejora es lineal con la duplicacion.
    Dim m_DictEdicionDisplay As Scripting.Dictionary
    Dim m_DictRiesgoCodigo As Scripting.Dictionary
    Set m_DictEdicionDisplay = PreloadEdicionDisplays(p_ColAnexos)
    Set m_DictRiesgoCodigo = PreloadRiesgoCodigos(p_ColAnexos)
    
    Dim m_IDAnexo As Variant
    Dim m_ObjAnexo As Anexo
    
    For Each m_IDAnexo In p_ColAnexos.keys
        If LenB(Nz(CStr(m_IDAnexo), "")) > 0 Then
            If p_ColAnexos.Exists(m_IDAnexo) Then
                Set m_ObjAnexo = p_ColAnexos(m_IDAnexo)
                If Not m_ObjAnexo Is Nothing Then
                    colFilas.Add BuildAnexoListRow(m_ObjAnexo, m_DictEdicionDisplay, m_DictRiesgoCodigo)
                End If
            End If
        End If
    Next
    
    Set BuildAnexosListRows = colFilas
End Function

' ----------------------------------------------------------------------------
' BuildAnexoListRow (privada interna) - formatea UNA fila
'   Usa los maps pre-cargados (O(1) lookup) en lugar de llamar a
'   .Edicion.Edicion / .riesgo.CodigoRiesgo por cada anexo.
' ----------------------------------------------------------------------------
Private Function BuildAnexoListRow( _
        ByVal p_ObjAnexo As Anexo, _
        ByVal p_DictEdicionDisplay As Scripting.Dictionary, _
        ByVal p_DictRiesgoCodigo As Scripting.Dictionary) As String
    Dim m_EdicionDeAnexo As String
    Dim m_FechaAnexo As String
    Dim m_Tipo As String
    Dim m_CodRiesgo As String
    
    ' Edic_Anexado: solo si el anexo esta asociado a una edicion (directo).
    ' O(1) lookup en el map pre-cargado.
    Dim m_IDEdicion As String
    m_IDEdicion = Nz(p_ObjAnexo.IDEdicion, "")
    If LenB(m_IDEdicion) > 0 Then
        If Not p_DictEdicionDisplay Is Nothing Then
            If p_DictEdicionDisplay.Exists(m_IDEdicion) Then
                m_EdicionDeAnexo = CStr(p_DictEdicionDisplay(m_IDEdicion))
            End If
        End If
    End If
    
    ' Fecha: dd/mm/yyyy si es fecha valida, sino ""
    m_FechaAnexo = FormatFechaAnexo(p_ObjAnexo.FechaAnexo)
    
    ' Tipo: derivado del flag IDRiesgo (NO de p_ObjAnexo.Tipo).
    ' Ver DeriveTipoFromAnexo para justificacion.
    m_Tipo = DeriveTipoFromAnexo(p_ObjAnexo)

    ' CodRiesgo: solo si Tipo="R" (anexo de riesgo hijo).
    ' O(1) lookup en el map pre-cargado.
    If m_Tipo = "R" Then
        Dim m_IDRiesgo As String
        m_IDRiesgo = Nz(p_ObjAnexo.IDRiesgo, "")
        If LenB(m_IDRiesgo) > 0 Then
            If Not p_DictRiesgoCodigo Is Nothing Then
                If p_DictRiesgoCodigo.Exists(m_IDRiesgo) Then
                    m_CodRiesgo = CStr(p_DictRiesgoCodigo(m_IDRiesgo))
                End If
            End If
        End If
    End If
    
    BuildAnexoListRow = m_Tipo & ";" & Nz(p_ObjAnexo.IDAnexo, "") & ";" & _
                         Nz(p_ObjAnexo.Titulo, "") & ";" & m_CodRiesgo & ";" & _
                         m_EdicionDeAnexo & ";" & m_FechaAnexo
End Function

' ----------------------------------------------------------------------------
' PreloadEdicionDisplays (privada interna)
'   Itera los anexos UNA vez. Por cada IDEdicion unico (no visto antes),
'   dispara la lazy-load via .Edicion (que cachea la instancia en el Anexo
'   y resuelve via Constructor.getEdicion con cache de Edicion). El map
'   resultante es Dict<IDEdicion, DisplayValue> (String -> String).
' ----------------------------------------------------------------------------
Private Function PreloadEdicionDisplays(ByVal p_ColAnexos As Scripting.Dictionary) As Scripting.Dictionary
    Dim m_Dict As New Scripting.Dictionary
    m_Dict.CompareMode = TextCompare
    
    Dim m_IDAnexo As Variant
    Dim m_ObjAnexo As Anexo
    Dim m_Edicion As Edicion
    Dim m_IDEdicion As String
    
    For Each m_IDAnexo In p_ColAnexos.keys
        If Not p_ColAnexos.Exists(m_IDAnexo) Then GoTo NextAnexo
        Set m_ObjAnexo = p_ColAnexos(m_IDAnexo)
        If m_ObjAnexo Is Nothing Then GoTo NextAnexo
        
        m_IDEdicion = Nz(m_ObjAnexo.IDEdicion, "")
        If LenB(m_IDEdicion) = 0 Then GoTo NextAnexo
        If m_Dict.Exists(m_IDEdicion) Then GoTo NextAnexo
        
        ' Lazy-load: Anexo.Edicion cachea la instancia internamente.
        ' Si Edicion ya esta cacheado en el Anexo, NO hay SQL adicional.
        Set m_Edicion = m_ObjAnexo.Edicion
        If Not m_Edicion Is Nothing Then
            m_Dict.Add m_IDEdicion, Nz(m_Edicion.Edicion, "")
        End If
NextAnexo:
    Next m_IDAnexo
    
    Set PreloadEdicionDisplays = m_Dict
End Function

' ----------------------------------------------------------------------------
' PreloadRiesgoCodigos (privada interna) - mismo patron que Edicion
' ----------------------------------------------------------------------------
Private Function PreloadRiesgoCodigos(ByVal p_ColAnexos As Scripting.Dictionary) As Scripting.Dictionary
    Dim m_Dict As New Scripting.Dictionary
    m_Dict.CompareMode = TextCompare
    
    Dim m_IDAnexo As Variant
    Dim m_ObjAnexo As Anexo
    Dim m_Riesgo As Riesgo
    Dim m_IDRiesgo As String
    
    For Each m_IDAnexo In p_ColAnexos.keys
        If Not p_ColAnexos.Exists(m_IDAnexo) Then GoTo NextAnexo
        Set m_ObjAnexo = p_ColAnexos(m_IDAnexo)
        If m_ObjAnexo Is Nothing Then GoTo NextAnexo
        
        m_IDRiesgo = Nz(m_ObjAnexo.IDRiesgo, "")
        If LenB(m_IDRiesgo) = 0 Then GoTo NextAnexo
        If m_Dict.Exists(m_IDRiesgo) Then GoTo NextAnexo
        
        Set m_Riesgo = m_ObjAnexo.riesgo
        If Not m_Riesgo Is Nothing Then
            m_Dict.Add m_IDRiesgo, Nz(m_Riesgo.CodigoRiesgo, "")
        End If
NextAnexo:
    Next m_IDAnexo
    
    Set PreloadRiesgoCodigos = m_Dict
End Function

' ----------------------------------------------------------------------------
' FormatFechaAnexo (privada interna) - dd/mm/yyyy o vacio
' ----------------------------------------------------------------------------
Private Function FormatFechaAnexo(ByVal p_FechaAnexo As Variant) As String
    If IsDate(p_FechaAnexo) Then
        FormatFechaAnexo = Format(p_FechaAnexo, "dd/mm/yyyy")
    End If
End Function

' ----------------------------------------------------------------------------
' DeriveTipoFromAnexo (privada interna)
'   Tipo NO se popula del SQL (issue #83), pero el contrato del dominio es
'   claro: un anexo tiene IDRiesgo si esta asociado a un riesgo (hijo), y
'   NO tiene IDRiesgo si esta asociado directamente a una edicion.
'   Mapeamos: IDRiesgo lleno -> "R"; vacio -> "E".
' ----------------------------------------------------------------------------
Private Function DeriveTipoFromAnexo(ByVal p_ObjAnexo As Anexo) As String
    If LenB(Nz(p_ObjAnexo.IDRiesgo, "")) > 0 Then
        DeriveTipoFromAnexo = "R"
    Else
        DeriveTipoFromAnexo = "E"
    End If
End Function

' --- Reset del presenter (idempotencia entre tests, si lo necesitara) ---
' (No tiene estado mutable propio; cache vive en modFormRiesgoDocumentosHelper)