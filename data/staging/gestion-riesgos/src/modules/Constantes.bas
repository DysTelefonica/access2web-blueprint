Attribute VB_Name = "Constantes"
Option Compare Database
Option Explicit

Public Const intColorEditable As Long = 16446961
Public Const intColorNOEditable As Long = 15720403
Public Const ETIQUETA_ESTADO_RIESGO_ACEPTADO_RECHAZADO_MATERIALIZADO = 2366701
Public Const ETIQUETA_ESTADO_RIESGO_NO_ACEPTADO_RECHAZADO_MATERIALIZADO = 2646607
Public Const MITIGACION_ACEPTAR_JUSTIFICADA_RECHAZADA As Long = 2366701
Public Const MITIGACION_ACEPTAR_JUSTIFICADA_NO_APROBADA As Long = intColorEditable
Public Const MITIGACION_ACEPTAR_JUSTIFICADA_POR_APROBAR As Long = 9486586
Public Const MITIGACION_ACEPTAR_JUSTIFICADA_APROBADA As Long = 5880731
Public Const MITIGACION_ACEPTAR_POR_JUSTIFICAR As Long = 13382655

Public Const RIESGO_RETIRADO_JUSTIFICADO_RECHAZADO As Long = 2366701
Public Const RIESGO_RETIRADO_JUSTIFICADO_NO_APROBADO As Long = intColorEditable
Public Const RIESGO_RETIRADO_JUSTIFICADO_POR_APROBAR As Long = 9486586
Public Const RIESGO_RETIRADO_JUSTIFICADO_APROBADO As Long = 5880731
Public Const RIESGO_RETIRADO_POR_JUSTIFICAR As Long = 13382655


Public Const ETIQUETA_ESTADO_COLOR_VERDE = 2646607
Public Const ETIQUETA_ESTADO_COLOR_ROJO = 2366701

Public Const PESTAÑA_TODOS_DATOS_OK_SI = 4210752
Public Const PESTAÑA_TODOS_DATOS_OK_NO = 2366701


Public Const ANCHO_MINIMO_LBL_PX As Long = 3053
Public Const ANCHO_MAXIMO_LBL_PX As Long = 3299

Public Const ANCHO_MINIMO_LBL_PXAcciones As Long = 1500
Public Const ANCHO_MAXIMO_LBL_PXAcciones As Long = 4500


Public Const COLOR_BORDE_CAMPO_NORELLENO As Long = 1643706
Public Const COLOR_BORDE_CAMPO_RELLENO As Long = 14136213
Public Const ANCHO_BORDE_CAMPO_RELLENO As Long = 0
Public Const ANCHO_BORDE_CAMPO_NORELLENO As Long = 4
'----------------------------------
'   Constantes de Excel
'------------------------------------------
Public Const xlDiagonalDown As Long = 5
Public Const xlDiagonalUp As Long = 6
Public Const xlEdgeLeft As Long = 7
Public Const xlEdgeTop As Long = 8
Public Const xlEdgeBottom As Long = 9
Public Const xlEdgeRight As Long = 10

Public Const xlInsideVertical As Long = 11
Public Const xlInsideHorizontal As Long = 12
Public Const xlContinuous As Long = 1
Public Const xlThin As Long = 2
Public Const xlMedium As Long = -4138
Public Const xlAutomatic As Long = -4105
Public Const xlNone As Long = -4142
Public Const xlDouble As Long = -4119
Public Const xlThick As Long = 4
Public Const xlToLeft As Long = -4159

Public Const xlLineMarkersStacked As Long = 66
Public Const xlLineMarkers As Long = 65
Public Const xlValue As Long = 2
Public Const xlLinear As Long = -4132
Public Const xlCategory As Long = 1
Public Const xlSquare As Long = 1
Public Const xlHairline As Long = 1
Public Const xlUnderlineStyleNone As Long = -4142
Public Const xlCenter As Long = -4108
Public Const xlContext As Long = -5002
Public Const xlLabelPositionAbove As Long = 0
Public Const xlHorizontal As Long = -4128
Public Const xlBottom As Long = -4107
Public Const xlColumnClustered As Long = 51
Public Const xlSolid As Long = 1
Public Const xlPrintNoComments As Long = -4142
Public Const xlLandscape As Long = 2
Public Const xlPortrait As Long = 1
Public Const xlPaperA4 As Long = 9
Public Const xlDownThenOver As Long = 1
Public Const xlPrintErrorsDisplayed As Long = 0
Public Const xlGeneral As Long = 1
Public Const xlTop As Long = -4160
Public Const xlLegendPositionBottom As Long = -4107


Public Const xlLeft As Long = -4131
Public Const xlRight As Long = -4152
Public Const xlUp As Long = -4162

Public Const xlTypePDF As Long = 0
Public Const msoFileDialogFilePicker As Long = 3
Public Const msoFileDialogFolderPicker As Long = 4
Public Const msoFileDialogOpen As Long = 1
Public Const msoFileDialogSaveAs As Long = 2

Public Const m_ColorComplementado As Long = 5026082
Public Const m_ColorNoComplementado As Long = 2366701

Public Const m_ColorOK As Long = 4210752
Public Const m_ColorNOOK As Long = 2366701
'-------------------------------------------
' VARIABLES QUE VA A USAR TODA LA BASE DE DATOS
'----------------------------------------------
'The structures returned by the API call GetIpAddrTable...
Public Type IPINFO
    dwAddr As Long          ' IP address
    dwIndex As Long         ' interface index
    dwMask As Long          ' subnet mask
    dwBCastAddr As Long     ' broadcast address
    dwReasmSize  As Long    ' assembly size
    Reserved1 As Integer
    Reserved2 As Integer
End Type
Public Enum EnumSiNo
    Sí = 1
    No = 2
End Enum
Public Enum EnumTipoInformePublicacion
    Excel = 1
    html = 2
End Enum
Public Enum EnumTipoNodo
    Edicion = 1
    riesgo = 2
    PM = 3
    PC = 4
    PMA = 5
    PCA = 6
End Enum
Public Enum EnumElemento
    Primero = 1
    Ultimo = 2
    Ordinal = 3
End Enum
Public Enum EnumTipoObjeto
    Proyecto = 1
    Edicion = 2
    riesgo = 3
    PlanMitigacion = 4
    PlanContingencia = 5
    PlanMitigacionAccion = 6
    PlanContingenciaAccion = 7
    LogObj = 8
    ProyectoJP = 9
    ProyectoRespCalidad = 10
    RiesgoExterno = 11
    Anexo = 12
    CORREO = 13
    Documento = 14
    RiesgoBiblioteca = 15
    RiesgoNC = 16
End Enum
Public Enum EnumRiesgoEstado
    Detectado = 1
    Planificado = 2
    Activo = 3
    Materializado = 4
    AceptadoSinJustificar = 5
    AceptadoSinVisar = 6
    AceptadoRechazado = 7
    Aceptado = 8
    RetiradoSinJustificar = 9
    RetiradoSinVisar = 10
    RetiradoRechazado = 11
    Retirado = 12
    Cerrado = 13
    Incompleto = 14 'viene de Riesgos externos y no se ha rellenado aún toda la ficha
End Enum
Public Enum EnumPlanEstado
    SinAcciones = 1
    FaltanDatos = 2
    Definido = 3
    Planificado = 4
    Activo = 5
    Finalizado = 6
End Enum
Public Enum EnumPlanAccionEstado
    FaltanDatos = 1
    Definido = 2
    Planificado = 3
    Activo = 4
    Finalizado = 5
End Enum
Public Enum EnumRiesgoValoracion
    MuyBajo = 1
    Bajo = 2
    Medio = 3
    Alto = 4
    MuyAlto = 5
End Enum
Public Enum EnumMitigacionValores
    Aceptar = 1
    Evitar = 2
    Reducir = 3
    Transferir = 4
End Enum

Public Enum EnumTipoPlan
    Mitigacion = 1
    Contingencia = 2
End Enum
Public Enum EnumTipoTareas
    ParaAceptadosRetirados = 1
    ParaMaterializados = 2
    ParaProyectos = 3
    ParaPorRetipificar = 4
    ParaRetipificados = 5
    ParaPreparadasParaPublicar = 6
End Enum


Public Enum EnumAccionesDeTareas
    RiesgoAlta = 1
    RiesgoAltaConRetipificacion = 2
    RiesgoAceptado = 3
    RiesgoAceptadoQuitar = 4
    RiesgoAceptadoVisado = 5
    RiesgoAceptadoVisadoQuitar = 6
    RIESGOACEPTADORECHAZADO = 7
    RiesgoAceptadoRechazadoQuitar = 8
    
    RiesgoRetirado = 9
    RiesgoRetiradoQuitar = 10
    RiesgoRetiradoVisado = 11
    RiesgoRetiradoVisadoQuitar = 12
    RIESGORETIRADORECHAZADO = 13
    RiesgoRetiradoRechazadoQuitar = 14
    
    RiesgoParaRetipificar = 15
    RiesgoParaRetipificarQuitar = 16
    
    ProyectoAPuntoDeCaducarQuitar = 17
    ProyectoCaducadoQuitar = 18
    
End Enum
Public Enum EnumPerfilUsuario
    Administrador = 1
    Calidad = 2
    Tecnico = 3
End Enum
Public Enum EnumOrigenRiesgoExterno
    Oferta = 1
    Subcontratista = 2
    Pedido = 3
End Enum

Public Enum EnumValoresPlazoCalidadCosteVulnerabilidad
    MuyBajo = 1
    Bajo = 2
    Medio = 3
    Alto = 4
    MuyAlto = 5
End Enum
Public Enum EnumRiesgosOrdenadosPor
    ID = 1
    Prioridad = 2
End Enum
Public Enum EnumActuarSobrePrioridad
    Subir = 1
    bajar = 2
End Enum
Public Enum EnumOperacionesAceptacionRetiro
    JustificarAceptacion = 1
    QuitarJustificacionAceptacion = 2
    JustificarRetiro = 3
    QuitarJustificacionRetiro = 4
    AceptacionAprobar = 5
    AceptacionAprobarQuitar = 6
    RetiroAprobar = 7
    RetiroAprobacionQuitar = 8
    RechazarAceptacion = 9
    QuitarRechazoAceptacion = 10
    RechazarRetiro = 11
    QuitarRechazoRetiro = 12
End Enum
Public Enum EnumTipoOperacion
    Alta = 1
    Baja = 2
    Edicion = 3
End Enum
Public Enum EnumAreaImpacto
    Calidad = 1
    Coste = 2
    Plazo = 3
End Enum
Public Enum EnumTipoRiesgoBiblioteca
    Proyecto = 1
    Pedido = 2
End Enum
Public Enum EnumEstadoMaterializado
    NoNC = 1
    SiNC = 2
    Pendiente = 3
    NoMaterializado = 4
End Enum
Public Enum EnumTipoNodoTareasCalidad
    EDICIONESCADUCADAS = 13
    EDICIONCADUCADA = 14
    EDICIONESAPUNTODECADUCAR = 1
    EDICIONAPUNTODECADUCAR = 2
    EDICIONESPREPARADASPARAPUBLICAR = 3
    EDICIONPREPARADAPARAPUBLICAR = 4
    RIESGOSPENDIENTESRETIPIFICAR = 5
    RIESGOPENDIENTERETIPIFICAR = 6
    RIESGOSACEPTADOSPORVISAR = 7
    RIESGOACEPTADOPORVISAR = 8
    RIESGOSRETIRADOSPORVISAR = 9
    RIESGORETIRADOPORVISAR = 10
    RIESGOSMATERIALIZADOSPORDECIDIR = 11
    RIESGOMATERIALIZADOPORDECIDIR = 12
    
End Enum

Public Enum EnumTipoNodoTareasTecnico
    EDICIONESAPUNTODECADUCARSINPROPUESTA = 1
    EDICIONAPUNTODECADUCARSINPROPUESTA = 2
    RIESGOSACEPTADOSRECHAZADOS = 3
    RIESGOACEPTADORECHAZADO = 4
    RIESGOSRETIRADOSRECHAZADOS = 5
    RIESGORETIRADORECHAZADO = 6
    EDICIONESPROPUESTASRECHAZADAS = 7
    EDICIONPROPUESTARECHAZADA = 8
    
End Enum
Public Enum EnumTipoRiesgoTarea
    Aceptados = 1
    Retirados = 2
    Retipificados = 3
    
End Enum
Public Enum EnumTipoCeldaAlto
    RiesgoCabeceraPpal = 1 '---> 21,75
    RiesgoCabeceraApartado = 2 '---> 18,75
    RiesgoDatos = 3 '-->38,25
    RiesgoDescripcion = 4 '---> 57
    RiesgoPAccionesCabecera = 5 '---> 25,5
    RiesgoPAcciones = 6 '---> variable
    PortadaCambios = 7
    RiesgoPlanDisparador = 8 '---> variable
    RiesgoCausaRaiz = 9
End Enum
Public Enum EnumTipoRecuadro
    TodoGruesoSinInterior = 1
    TodoGruesoConInteriorVertical = 2
    TodoGruesoConInteriorVerticalYHorizontal = 3
    TodoFinoSinInterior = 4
    TodoFinoConInteriorrVertical = 5
    TodoFinoConInteriorrVerticalYHorizontal = 6
    ArribaGruesoAbajoMedioConInterior = 7
End Enum
Public Enum EnumAnchoLinea
    Gruesa = xlThick
    Mediana = xlMedium
    fina = xlThin
    SinLinea = xlNone
    
End Enum
Public Enum EnumEstados
    Oferta = 1
    Adjudicada = 2
    EnEjecucion = 3
    EnGarantia = 4
    Cerrado = 5
    Desestimado = 6
    Perdido = 7
    NoAplica = 8
    Desconocido = 9
End Enum
Public Enum EnumTipoExpediente
    AM = 1
    Lote = 2
    BasadoDeAM = 3
    BasadoDeLote = 4
    EXPIndividual = 5
    EXPHPS = 6
End Enum
Public Enum EnumControlCambios
    Todo = 1
    SoloCalidad = 2
End Enum
Public Enum EnumControlCambiosAlcance
    EnumControlCambiosAlcanceResumen3 = 1
    EnumControlCambiosAlcanceCompleto = 2
End Enum
Public Enum EnumMensaje
    Error = vbCritical
    Warning = vbExclamation
    Information = vbInformation
    
End Enum
