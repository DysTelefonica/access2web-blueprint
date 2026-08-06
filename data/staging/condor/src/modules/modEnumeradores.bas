Attribute VB_Name = "modEnumeradores"

Option Compare Database
Option Explicit

' --- Constantes Globales ---
Public Const CONST_ID_APLICACION As String = "23"
Public Const wdFormatPDF As Long = 17 ' Constante de Word para Late Binding

' --- Enumeradores ---
Public Enum rol
    Administrador = 1
    Calidad = 2
    Tecnico = 3
End Enum

Public Enum EnumSiNo
    sí = 1
    no = 2
End Enum

Public Enum EnumDecision
    Decision_Aprobado = 1
    Decision_Rechazado = 2
    Decision_Pendiente = 0
End Enum



Public Enum TipoGrafico
    GraficoAnillos
    GraficoBarras
    GraficoTimeline
End Enum

Public Enum enumEstados
    estadoPreregistro = 1
    estadoRegistro = 2
    estadoDesarrolloTecnico = 3
    estadoModificacion = 4
    estadoValidacion = 5
    estadoRevision = 6
    estadoFormalizacion = 7
    estadoAprobada = 8
    estadoRechazada = 9
End Enum

Public Enum enumBloqueFormulario
    Bloque_Generales = 1
    Bloque_Propuesta = 2
    Bloque_Impacto = 3
    Bloque_AprobacionSuministrador = 4
    Bloque_DictamenRAC = 5
    Bloque_DecisionFinal = 6
End Enum

Public Enum enumTipoFormulario
    TipoForm_PC = 1
    TipoForm_CDCA = 2
    TipoForm_CDCASUB = 3
    TipoForm_PCSUB = 4
End Enum
