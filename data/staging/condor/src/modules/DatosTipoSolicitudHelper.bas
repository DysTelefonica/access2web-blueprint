Attribute VB_Name = "DatosTipoSolicitudHelper"
Option Compare Database
Option Explicit

' ============================================================================
' DatosTipoSolicitudHelper — pure RowSource builder for cboTipoSolicitud
'
' Skill:    access-vba-e2e-methodology (thin form, pure helper)
' SDD/CAP:  CAP-001 (PCSUB), delta alta-solicitud (Spec-002)
' Branch:   staging
' Project:  condor (Dysflow projectId)
'
' Why pure:
'   Form frmAltaSolicitud must stay thin (UI wiring only). The RowSource
'   decision is delegated here so atoms can test it without opening a
'   form, without DAO, and without Screen.
'
' Contract (from design.md):
'   Principal detection MUST trim and be case-insensitive for "Sí".
'   Safe default for empty/unknown role: subcontratista list
'     (PC, PC_SUB, CD_CA_SUB).
' ============================================================================

' --- Canonical rowSource tokens (single source of truth, used by form and
'     also referenced by the atoms in Test_DatosTipoSolicitudHelper_PCSUB_Strict) ---
Private Const ROW_PRINCIPAL_PC As String = "'PC';'PC - Propuesta de Cambio'"
Private Const ROW_PRINCIPAL_CDCA As String = "'CD_CA';'CD/CA - Concesión/Desviación'"
Private Const ROW_SUB_PC As String = "'PC';'PC - Propuesta de Cambio'"
Private Const ROW_SUB_PCSUB As String = "'PC_SUB';'PC-SUB - Subcontratista'"
Private Const ROW_SUB_CDCASUB As String = "'CD_CA_SUB';'CD/CA-SUB - Subcontratista'"

' ============================================================================
' Helpers
' ============================================================================
Private Function EsPrincipal(ByVal valor As String) As Boolean
    ' Case-insensitive + trim compare against canonical "Sí".
    ' UCase handles "SÍ", "Sì", "sí", "SI"; Trim handles leading/trailing ws.
    EsPrincipal = (UCase$(Trim$(Nz(valor, ""))) = "SÍ")
End Function

' ============================================================================
' Public API
' ============================================================================
Public Function ConstruirRowSourceComboTipo(ByVal contratistaPrincipal As String) As String
    ' Returns the Access RowSource string for cboTipoSolicitud.
    ' Pure function: no DAO, no controls, no DB. Only string assembly.
    '
    ' Principal   -> PC + CD_CA
    ' Subcontratista (else, including empty/unknown) -> PC + PC_SUB + CD_CA_SUB
    '
    ' Params:
    '   contratistaPrincipal: expediente.ContratistaPrincipal ("Sí" / "No" / empty / unknown)
    '
    ' Returns:
    '   RowSource string formatted as ';'-delimited quoted pairs.
    Dim partes As String

    If EsPrincipal(contratistaPrincipal) Then
        partes = ROW_PRINCIPAL_PC & ";" & ROW_PRINCIPAL_CDCA
    Else
        ' Safe default for subcontratista (incl. empty and unknown)
        partes = ROW_SUB_PC & ";" & ROW_SUB_PCSUB & ";" & ROW_SUB_CDCASUB
    End If

    ConstruirRowSourceComboTipo = partes
End Function