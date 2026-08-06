VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConfiguracionOperaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Function CamposCambian( _
                                p_Configuracion As Configuracion, _
                                Optional ByRef p_Error As String _
                                ) As String
     
    
    Dim m_Campos As String
    
    On Error GoTo errores
    If p_Configuracion Is Nothing Then
        Exit Function
    End If
    If m_ObjEntorno.Configuracion Is Nothing Then
        Exit Function
    End If
    With m_ObjEntorno.Configuracion
        If .DiasParaRecordatorioExcel1 <> p_Configuracion.DiasParaRecordatorioExcel1 Then
            m_Campos = "DiasParaRecordatorioExcel1"
        End If
        If .DiasParaRecordatorioExcel2 <> p_Configuracion.DiasParaRecordatorioExcel2 Then
            If m_Campos = "" Then
                m_Campos = "DiasParaRecordatorioExcel2"
            Else
                m_Campos = m_Campos & ";" & "DiasParaRecordatorioExcel2"
            End If
        End If
        If .DiasCancelacionPreMARGA <> p_Configuracion.DiasCancelacionPreMARGA Then
            If m_Campos = "" Then
                m_Campos = "DiasCancelacionPreMARGA"
            Else
                m_Campos = m_Campos & ";" & "DiasCancelacionPreMARGA"
            End If
        End If
        If .DiasParaRecordatorioRellenoMarga1 <> p_Configuracion.DiasParaRecordatorioRellenoMarga1 Then
            If m_Campos = "" Then
                m_Campos = "DiasParaRecordatorioRellenoMarga1"
            Else
                m_Campos = m_Campos & ";" & "DiasParaRecordatorioRellenoMarga1"
            End If
        End If
        If .DiasParaRecordatorioRellenoMarga2 <> p_Configuracion.DiasParaRecordatorioRellenoMarga2 Then
            If m_Campos = "" Then
                m_Campos = "DiasParaRecordatorioRellenoMarga2"
            Else
                m_Campos = m_Campos & ";" & "DiasParaRecordatorioRellenoMarga2"
            End If
        End If
        
        If .DiasCancelacionMARGA <> p_Configuracion.DiasCancelacionMARGA Then
            If m_Campos = "" Then
                m_Campos = "DiasCancelacionMARGA"
            Else
                m_Campos = m_Campos & ";" & "DiasCancelacionMARGA"
            End If
        End If
        If .EmailDirectorSeguridad <> p_Configuracion.EmailDirectorSeguridad Then
            If m_Campos = "" Then
                m_Campos = "EmailDirectorSeguridad"
            Else
                m_Campos = m_Campos & ";" & "EmailDirectorSeguridad"
            End If
        End If
        
        If .AutocancelacionPreMARGA <> p_Configuracion.AutocancelacionPreMARGA Then
            If m_Campos = "" Then
                m_Campos = "AutocancelacionPreMARGA"
            Else
                m_Campos = m_Campos & ";" & "AutocancelacionPreMARGA"
            End If
        End If
       
        If .AutocancelacionMARGA <> p_Configuracion.AutocancelacionMARGA Then
            If m_Campos = "" Then
                m_Campos = "AutocancelacionMARGA"
            Else
                m_Campos = m_Campos & ";" & "AutocancelacionMARGA"
            End If
        End If
        If .VersionPlantillasHTML <> p_Configuracion.VersionPlantillasHTML Then
            If m_Campos = "" Then
                m_Campos = "VersionPlantillasHTML"
            Else
                m_Campos = m_Campos & ";" & "VersionPlantillasHTML"
            End If
        End If
        If .VersionPlantillasEXCEL <> p_Configuracion.VersionPlantillasEXCEL Then
            If m_Campos = "" Then
                m_Campos = "VersionPlantillasEXCEL"
            Else
                m_Campos = m_Campos & ";" & "VersionPlantillasEXCEL"
            End If
        End If
    End With
    CamposCambian = m_Campos
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ConfiguracionOperaciones.CamposCambian ha producido el error num " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function
Public Function Registrar( _
                            p_Configuracion As Configuracion, _
                            Optional ByRef p_Error As String _
                            ) As String
     
    Dim rcdDatos As DAO.Recordset
    Dim m_SQL As String
    Dim m_MotivoNoOK As String
    
    
    On Error GoTo errores
    If p_Configuracion Is Nothing Then
        p_Error = "No hay datos que guardar"
        Err.Raise 1000
    End If
    
    Avance "Comprobando datos de la factura"
    m_MotivoNoOK = MotivoNoOK(p_Configuracion, p_Error)
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    p_Error = m_MotivoNoOK
    If p_Error <> "" Then
        Err.Raise 1000
    End If
    
    m_SQL = "TbConfiguracion"
    Set rcdDatos = getdb().OpenRecordset(m_SQL)
    With rcdDatos
        If m_ObjEntorno.Configuracion Is Nothing Then
            p_Configuracion.id = p_Configuracion.ID_Calculado
            .AddNew
                .Fields("ID") = p_Configuracion.id
                If p_Configuracion.CorreosAutomaticos <> "" Then
                    .Fields("CorreosAutomaticos") = p_Configuracion.CorreosAutomaticos
                End If
                If p_Configuracion.BuzonSeguridad <> "" Then
                    .Fields("BuzonSeguridad") = p_Configuracion.BuzonSeguridad
                End If
                If IsNumeric(p_Configuracion.DiasParaRecordatorioExcel1) Then
                    .Fields("DiasParaRecordatorioExcel1") = p_Configuracion.DiasParaRecordatorioExcel1
                End If
                If p_Configuracion.EmailDirectorSeguridad <> "" Then
                    .Fields("EmailDirectorSeguridad") = p_Configuracion.EmailDirectorSeguridad
                End If
                If IsNumeric(p_Configuracion.DiasParaRecordatorioExcel2) Then
                    .Fields("DiasParaRecordatorioExcel2") = p_Configuracion.DiasParaRecordatorioExcel2
                End If
                
                If IsNumeric(p_Configuracion.DiasCancelacionPreMARGA) Then
                    .Fields("DiasCancelacionPreMARGA") = p_Configuracion.DiasCancelacionPreMARGA
                End If
                If IsNumeric(p_Configuracion.DiasParaRecordatorioRellenoMarga1) Then
                    .Fields("DiasParaRecordatorioRellenoMarga1") = p_Configuracion.DiasParaRecordatorioRellenoMarga1
                End If
                If IsNumeric(p_Configuracion.DiasParaRecordatorioRellenoMarga2) Then
                    .Fields("DiasParaRecordatorioRellenoMarga2") = p_Configuracion.DiasParaRecordatorioRellenoMarga2
                End If
                If p_Configuracion.AutocancelacionPreMARGA <> "" Then
                    .Fields("AutocancelacionPreMARGA") = p_Configuracion.AutocancelacionPreMARGA
                End If
                If p_Configuracion.AutocancelacionMARGA <> "" Then
                    .Fields("AutocancelacionMARGA") = p_Configuracion.AutocancelacionMARGA
                End If
                If IsNumeric(p_Configuracion.DiasCancelacionPreMARGA) Then
                    .Fields("DiasCancelacionPreMARGA") = p_Configuracion.DiasCancelacionPreMARGA
                End If
                If IsNumeric(p_Configuracion.DiasCancelacionMARGA) Then
                    .Fields("DiasCancelacionMARGA") = p_Configuracion.DiasCancelacionMARGA
                End If
                If p_Configuracion.VersionPlantillasHTML <> "" Then
                    .Fields("VersionPlantillasHTML") = p_Configuracion.VersionPlantillasHTML
                End If
                If p_Configuracion.VersionPlantillasEXCEL <> "" Then
                    .Fields("VersionPlantillasEXCEL") = p_Configuracion.VersionPlantillasEXCEL
                End If
            .Update
            
        Else
            .Edit
                If p_Configuracion.CorreosAutomaticos <> "" Then
                    .Fields("CorreosAutomaticos") = p_Configuracion.CorreosAutomaticos
                Else
                    .Fields("CorreosAutomaticos") = Null
                End If
                If p_Configuracion.BuzonSeguridad <> "" Then
                    .Fields("BuzonSeguridad") = p_Configuracion.BuzonSeguridad
                Else
                    .Fields("BuzonSeguridad") = Null
                End If
                If IsNumeric(p_Configuracion.DiasParaRecordatorioExcel1) Then
                    .Fields("DiasParaRecordatorioExcel1") = p_Configuracion.DiasParaRecordatorioExcel1
                Else
                    .Fields("DiasParaRecordatorioExcel1") = Null
                End If
                If p_Configuracion.EmailDirectorSeguridad <> "" Then
                    .Fields("EmailDirectorSeguridad") = p_Configuracion.EmailDirectorSeguridad
                Else
                    .Fields("EmailDirectorSeguridad") = Null
                End If
                If IsNumeric(p_Configuracion.DiasParaRecordatorioExcel2) Then
                    .Fields("DiasParaRecordatorioExcel2") = p_Configuracion.DiasParaRecordatorioExcel2
                Else
                    .Fields("DiasParaRecordatorioExcel2") = Null
                End If
                
                If IsNumeric(p_Configuracion.DiasCancelacionPreMARGA) Then
                    .Fields("DiasCancelacionPreMARGA") = p_Configuracion.DiasCancelacionPreMARGA
                Else
                    .Fields("DiasCancelacionPreMARGA") = Null
                End If
                If IsNumeric(p_Configuracion.DiasParaRecordatorioRellenoMarga1) Then
                    .Fields("DiasParaRecordatorioRellenoMarga1") = p_Configuracion.DiasParaRecordatorioRellenoMarga1
                Else
                    .Fields("DiasParaRecordatorioRellenoMarga1") = Null
                End If
                If IsNumeric(p_Configuracion.DiasParaRecordatorioRellenoMarga2) Then
                    .Fields("DiasParaRecordatorioRellenoMarga2") = p_Configuracion.DiasParaRecordatorioRellenoMarga2
                Else
                    .Fields("DiasParaRecordatorioRellenoMarga2") = Null
                End If
                If p_Configuracion.AutocancelacionPreMARGA <> "" Then
                    .Fields("AutocancelacionPreMARGA") = p_Configuracion.AutocancelacionPreMARGA
                Else
                    .Fields("AutocancelacionPreMARGA") = Null
                End If
                If p_Configuracion.AutocancelacionMARGA <> "" Then
                    .Fields("AutocancelacionMARGA") = p_Configuracion.AutocancelacionMARGA
                Else
                    .Fields("AutocancelacionMARGA") = Null
                End If
                If IsNumeric(p_Configuracion.DiasCancelacionPreMARGA) Then
                    .Fields("DiasCancelacionPreMARGA") = p_Configuracion.DiasCancelacionPreMARGA
                Else
                    .Fields("DiasCancelacionPreMARGA") = Null
                End If
                If IsNumeric(p_Configuracion.DiasCancelacionMARGA) Then
                    .Fields("DiasCancelacionMARGA") = p_Configuracion.DiasCancelacionMARGA
                Else
                    .Fields("DiasCancelacionMARGA") = Null
                End If
                If p_Configuracion.VersionPlantillasHTML <> "" Then
                    .Fields("VersionPlantillasHTML") = p_Configuracion.VersionPlantillasHTML
                Else
                    .Fields("VersionPlantillasHTML") = Null
                End If
                If p_Configuracion.VersionPlantillasEXCEL <> "" Then
                    .Fields("VersionPlantillasEXCEL") = p_Configuracion.VersionPlantillasEXCEL
                Else
                    .Fields("VersionPlantillasEXCEL") = Null
                End If
            .Update
            
        End If
    End With
    rcdDatos.Close
    Set rcdDatos = Nothing
    m_FaltaAlgunParametroEnConfiguracion = DatosConfInsuficientes(p_Configuracion, EnumSiNo.no, p_Error)
    If p_Error <> "" Then
       Err.Raise 1000
    End If
    Set m_ObjEntorno = Nothing
    EVE p_Error
    If p_Error <> "" Then
        Err.Raise 1000
    End If

    'RegistrarLog Log
    Set m_ObjEntorno.Configuracion = Nothing
    Exit Function
errores:
    If Err.Number <> 1000 Then
        p_Error = "El método ConfiguracionOperaciones.Registrar ha producido el error num " & Err.Number & vbNewLine & "Detalle: " & Err.Description
    End If
    
End Function

Public Function MotivoNoOK( _
                            p_Configuracion As Configuracion, _
                            Optional ByRef p_Error As String _
                            ) As String
    
    Dim m_URLFinal As String
    On Error GoTo errores
    
    With p_Configuracion
        If .DiasParaRecordatorioExcel1 <> "" Then
            If Not IsNumeric(.DiasParaRecordatorioExcel1) Then
                MotivoNoOK = "DiasParaRecordatorioExcel1 ha de ser un número entero positivo"
                Exit Function
            End If
            If CDbl(.DiasParaRecordatorioExcel1) < 0 Then
                MotivoNoOK = "DiasParaRecordatorioExcel1 ha de ser un número entero positivo"
                Exit Function
            End If
        End If
        If .DiasParaRecordatorioExcel2 <> "" Then
            If Not IsNumeric(.DiasParaRecordatorioExcel2) Then
                MotivoNoOK = "DiasParaRecordatorioExcel2 ha de ser un número entero positivo"
                Exit Function
            End If
            If CDbl(.DiasParaRecordatorioExcel2) < 0 Then
                MotivoNoOK = "DiasParaRecordatorioExcel2 ha de ser un número entero positivo"
                Exit Function
            End If
        End If
        If .DiasCancelacionPreMARGA <> "" Then
            If Not IsNumeric(.DiasCancelacionPreMARGA) Then
                MotivoNoOK = "DiasCancelacionPreMARGA ha de ser un número entero positivo"
                Exit Function
            End If
            If CDbl(.DiasCancelacionPreMARGA) < 0 Then
                MotivoNoOK = "DiasParaAutocancelacion ha de ser un número entero positivo"
                Exit Function
            End If
        End If
        If .DiasParaRecordatorioRellenoMarga1 <> "" Then
            If Not IsNumeric(.DiasParaRecordatorioRellenoMarga1) Then
                MotivoNoOK = "DiasParaRecordatorioRellenoMarga1 ha de ser un número entero positivo"
                Exit Function
            End If
            If CDbl(.DiasParaRecordatorioRellenoMarga1) < 0 Then
                MotivoNoOK = "DiasParaRecordatorioRellenoMarga1 ha de ser un número entero positivo"
                Exit Function
            End If
        End If
        If .DiasParaRecordatorioRellenoMarga2 <> "" Then
            If Not IsNumeric(.DiasParaRecordatorioRellenoMarga2) Then
                MotivoNoOK = "DiasParaRecordatorioRellenoMarga2 ha de ser un número entero positivo"
                Exit Function
            End If
            If CDbl(.DiasParaRecordatorioRellenoMarga2) < 0 Then
                MotivoNoOK = "DiasParaRecordatorioRellenoMarga2 ha de ser un número entero positivo"
                Exit Function
            End If
        End If
        If .DiasCancelacionMARGA <> "" Then
            If Not IsNumeric(.DiasCancelacionMARGA) Then
                MotivoNoOK = "DiasCancelacionMARGA ha de ser un número entero positivo"
                Exit Function
            End If
            If CDbl(.DiasCancelacionMARGA) < 0 Then
                MotivoNoOK = "DiasCancelacionMARGA ha de ser un número entero positivo"
                Exit Function
            End If
        End If
        If .BuzonSeguridad <> "" Then
            If Not EsCorreoValido(.BuzonSeguridad) Then
                MotivoNoOK = "BuzonSeguridad no tiene un formato correcto de email"
                Exit Function
            End If
            
        End If
        If .EmailDirectorSeguridad <> "" Then
            If Not EsCorreoValido(.EmailDirectorSeguridad) Then
                MotivoNoOK = "EmailDirectorSeguridad no tiene un formato correcto de email"
                Exit Function
            End If
        End If
       
        If .AutocancelacionPreMARGA <> "Sí" And .AutocancelacionPreMARGA <> "No" Then
            MotivoNoOK = "AutocancelacionPreMARGA ha de ser Sí o No"
            Exit Function
        End If
        If .AutocancelacionMARGA <> "Sí" And .AutocancelacionMARGA <> "No" Then
            MotivoNoOK = "AutocancelacionMARGA ha de ser Sí o No"
            Exit Function
        End If
        If .VersionPlantillasHTML = "" Then
            MotivoNoOK = "Se ha de indicar la versión de las plantillas HTML (Vxx)"
            Exit Function
        End If
        If .VersionPlantillasEXCEL = "" Then
            MotivoNoOK = "Se ha de indicar la versión de las plantillas EXCEL (Vxx)"
            Exit Function
        End If
        
    End With
    
    Exit Function
     
errores:
     If Err.Number <> 1000 Then
         p_Error = "El método ConfiguracionOperaciones.MotivoNoOK ha devuelto el error: " & Err.Description
     End If
                                
End Function

