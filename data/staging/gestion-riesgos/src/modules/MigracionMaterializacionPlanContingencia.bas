Attribute VB_Name = "MigracionMaterializacionPlanContingencia"
Option Compare Database
Option Explicit

Public Function EnsureMaterializacionPlanContingenciaField( _
    Optional ByRef p_Error As String _
) As EnumSiNo
    EnsureMaterializacionPlanContingenciaField = _
        MigracionMaterializacionPlanContingencia_AsegurarCampo(p_Error)
End Function

Public Function MigracionMaterializacionPlanContingencia_AsegurarCampo( _
    Optional ByRef p_Error As String _
) As EnumSiNo
    Dim db As DAO.Database
    Dim m_NombreTabla As String

    On Error GoTo errores

    p_Error = ""
    Set db = getdb(p_Error)
    If p_Error <> "" Then Err.Raise 1000, "MigracionMaterializacionPlanContingencia_AsegurarCampo", p_Error
    If db Is Nothing Then
        p_Error = "No se pudo abrir la base de datos para ejecutar la migración"
        Err.Raise 1000, "MigracionMaterializacionPlanContingencia_AsegurarCampo", p_Error
    End If

    m_NombreTabla = MigracionMaterializacionPlanContingencia_ResolverTabla(db, p_Error)
    If p_Error <> "" Then Err.Raise 1000, "MigracionMaterializacionPlanContingencia_AsegurarCampo", p_Error

    If Not MigracionMaterializacionPlanContingencia_ExisteCampo(db, m_NombreTabla, "IDPlanContingencia") Then
        db.Execute "ALTER TABLE [" & m_NombreTabla & "] ADD COLUMN [IDPlanContingencia] LONG", dbFailOnError
    End If

    MigracionMaterializacionPlanContingencia_AsegurarCampo = EnumSiNo.Sí
    Exit Function

errores:
    MigracionMaterializacionPlanContingencia_AsegurarCampo = EnumSiNo.No
    If Err.Number <> 1000 Then
        p_Error = "El método MigracionMaterializacionPlanContingencia_AsegurarCampo ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Err.Raise Err.Number, "MigracionMaterializacionPlanContingencia_AsegurarCampo", p_Error
End Function

Private Function MigracionMaterializacionPlanContingencia_ResolverTabla( _
    ByVal p_db As DAO.Database, _
    Optional ByRef p_Error As String _
) As String
    Dim m_Candidatas As Variant
    Dim m_Tabla As Variant

    On Error GoTo errores

    p_Error = ""
    m_Candidatas = Array("TbRiesgosMaterializaciones", "TbRiesgosMaterializados")

    For Each m_Tabla In m_Candidatas
        If MigracionMaterializacionPlanContingencia_ExisteTabla(p_db, CStr(m_Tabla)) Then
            MigracionMaterializacionPlanContingencia_ResolverTabla = CStr(m_Tabla)
            Exit Function
        End If
    Next m_Tabla

    p_Error = "No se encontró la tabla de materializaciones para aplicar la migración"
    Err.Raise 1000, "MigracionMaterializacionPlanContingencia_ResolverTabla", p_Error

errores:
    If Err.Number <> 1000 Then
        p_Error = "El método MigracionMaterializacionPlanContingencia_ResolverTabla ha devuelto el error: " & vbNewLine & Err.Description
    End If
    Err.Raise Err.Number, "MigracionMaterializacionPlanContingencia_ResolverTabla", p_Error
End Function

Private Function MigracionMaterializacionPlanContingencia_ExisteTabla( _
    ByVal p_db As DAO.Database, _
    ByVal p_NombreTabla As String _
) As Boolean
    Dim tdf As DAO.TableDef

    On Error GoTo errores

    For Each tdf In p_db.TableDefs
        If StrComp(tdf.Name, p_NombreTabla, vbTextCompare) = 0 Then
            MigracionMaterializacionPlanContingencia_ExisteTabla = True
            Exit Function
        End If
    Next tdf

    MigracionMaterializacionPlanContingencia_ExisteTabla = False
    Exit Function

errores:
    MigracionMaterializacionPlanContingencia_ExisteTabla = False
    Err.Clear
End Function

Private Function MigracionMaterializacionPlanContingencia_ExisteCampo( _
    ByVal p_db As DAO.Database, _
    ByVal p_NombreTabla As String, _
    ByVal p_NombreCampo As String _
) As Boolean
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field

    On Error GoTo errores

    Set tdf = p_db.TableDefs(p_NombreTabla)
    For Each fld In tdf.Fields
        If StrComp(fld.Name, p_NombreCampo, vbTextCompare) = 0 Then
            MigracionMaterializacionPlanContingencia_ExisteCampo = True
            Exit Function
        End If
    Next fld

    MigracionMaterializacionPlanContingencia_ExisteCampo = False
    Exit Function

errores:
    MigracionMaterializacionPlanContingencia_ExisteCampo = False
    Err.Clear
End Function
