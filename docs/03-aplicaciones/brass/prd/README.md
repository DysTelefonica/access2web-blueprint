# PRD — Brass

> **Estado:** fase 1 (fijación e inventario). Los formularios, módulos y clases quedan **pendientes**; las fichas, la matriz, la coexistencia, las capacidades y las issues propuestas llegan en las PR siguientes de la cadena.

| Campo | Valor |
|---|---|
| Repositorio | DysTelefonica/BRASS |
| Commit fijado | `3f77d586814096c04f24bd2b6c229831cfbfcd34` (`main`, 2026-10-14, «sync: BRASS main src and binary from production baseline (#4)») |
| Índice CodeGraph-VBA | 218 ficheros indexados, 12 626 nodos, versión 2.1.1, sobre clon de solo lectura de ese commit |
| Estudio previo | [epic.md](../epic.md), [forms.md](../forms.md), [capabilities.md](../capabilities.md), [data-model.md](../data-model.md), [security-rules.md](../security-rules.md), [integrations-automation.md](../integrations-automation.md), [migration-matrix.md](../migration-matrix.md) (DRAFT), walkthroughs `walkthrough-*.json`, [ERD/schema.sql](../ERD/schema.sql) |
| Snapshot comparado | `data/staging/brass/src` (2026-08-06) |
| Issue | #1078 (método y plantilla de #1069) |

## Alcance actual

Fija el commit y concilia el inventario. Los recuentos de líneas, controles y procedimientos son mecánicos (ver [Método del recuento](#método-del-recuento)) y no sustituyen al análisis.

## Inventario reconciliado

### Recuentos por tipo

| Tipo | `3f77d58` | Snapshot | Diferencia |
|---|---|---|---|
| Formularios (`src/forms/Form_*.cls`, con su `.form.txt`) | 87 | 87 | +0 |
| Módulos estándar (`src/modules/*.bas`) | 16 | 16 | +0 |
| Clases (`src/classes/*.cls`) | 28 | 28 | +0 |

La comparación normaliza las líneas `Attribute VB_`, los comentarios, las líneas en blanco y las diferencias de espacios y mayúsculas. Diferencia textual cero no demuestra que el estudio previo sea correcto: solo que el código no ha cambiado desde el snapshot.

### Delta de código frente al snapshot

- **Distintos del snapshot:** 1 módulo (`Variables Globales`), sin cambios en formularios ni clases.
- **Nuevos en el commit (sin equivalente en el snapshot):** 0 formularios, 0 módulos y 0 clases.
- **Del snapshot ausentes en el commit:** 0 formularios, 0 módulos y 0 clases.

### Formularios

Líneas: `.cls` del formulario. Controles: bloques `Begin` del `.form.txt`. Procedimientos: `Sub`, `Function` y `Property` del `.cls`. «Frente al snapshot»: **igual** si el `.cls` coincide tras la normalización.

| Objeto | Ruta | Líneas | Controles | Proc. | Frente al snapshot | Estado | Ficha o motivo |
|---|---|---|---|---|---|---|---|
| Form_Form0BDInformes | src/forms/Form_Form0BDInformes.cls | 107 | 21 | 16 | igual | analizado | [formularios/Form_Form0BDInformes.md](formularios/Form_Form0BDInformes.md) |
| Form_Form0BDOpciones | src/forms/Form_Form0BDOpciones.cls | 175 | 38 | 25 | igual | analizado | [formularios/Form_Form0BDOpciones.md](formularios/Form_Form0BDOpciones.md) |
| Form_Form0BDOpcionesOtras | src/forms/Form_Form0BDOpcionesOtras.cls | 144 | 40 | 29 | igual | analizado | [formularios/Form_Form0BDOpcionesOtras.md](formularios/Form_Form0BDOpcionesOtras.md) |
| Form_FormActividadAlta | src/forms/Form_FormActividadAlta.cls | 251 | 38 | 7 | igual | analizado | [formularios/Form_FormActividadAlta.md](formularios/Form_FormActividadAlta.md) |
| Form_FormActividadDeEventoPorLotes | src/forms/Form_FormActividadDeEventoPorLotes.cls | 383 | 31 | 10 | igual | analizado | [formularios/Form_FormActividadDeEventoPorLotes.md](formularios/Form_FormActividadDeEventoPorLotes.md) |
| Form_FormActividadEdicion | src/forms/Form_FormActividadEdicion.cls | 803 | 43 | 13 | igual | analizado | [formularios/Form_FormActividadEdicion.md](formularios/Form_FormActividadEdicion.md) |
| Form_FormActividadGestion | src/forms/Form_FormActividadGestion.cls | 1065 | 39 | 26 | igual | analizado | [formularios/Form_FormActividadGestion.md](formularios/Form_FormActividadGestion.md) |
| Form_FormActividadTipoAlta | src/forms/Form_FormActividadTipoAlta.cls | 148 | 23 | 4 | igual | analizado | [formularios/Form_FormActividadTipoAlta.md](formularios/Form_FormActividadTipoAlta.md) |
| Form_FormActividadTipoEdicion | src/forms/Form_FormActividadTipoEdicion.cls | 323 | 23 | 6 | igual | analizado | [formularios/Form_FormActividadTipoEdicion.md](formularios/Form_FormActividadTipoEdicion.md) |
| Form_FormActividadTipoGestion | src/forms/Form_FormActividadTipoGestion.cls | 180 | 19 | 7 | igual | analizado | [formularios/Form_FormActividadTipoGestion.md](formularios/Form_FormActividadTipoGestion.md) |
| Form_FormAnexos | src/forms/Form_FormAnexos.cls | 606 | 39 | 12 | igual | analizado | [formularios/Form_FormAnexos.md](formularios/Form_FormAnexos.md) |
| Form_FormAnexosNoAlcanzables | src/forms/Form_FormAnexosNoAlcanzables.cls | 219 | 28 | 6 | igual | analizado | [formularios/Form_FormAnexosNoAlcanzables.md](formularios/Form_FormAnexosNoAlcanzables.md) |
| Form_FormBUIGestion | src/forms/Form_FormBUIGestion.cls | 229 | 18 | 7 | igual | analizado | [formularios/Form_FormBUIGestion.md](formularios/Form_FormBUIGestion.md) |
| Form_FormCalibracionAlta | src/forms/Form_FormCalibracionAlta.cls | 228 | 35 | 4 | igual | analizado | [formularios/Form_FormCalibracionAlta.md](formularios/Form_FormCalibracionAlta.md) |
| Form_FormCalibracionEdicion | src/forms/Form_FormCalibracionEdicion.cls | 346 | 38 | 7 | igual | analizado | [formularios/Form_FormCalibracionEdicion.md](formularios/Form_FormCalibracionEdicion.md) |
| Form_FormCalibracionEquipoAlta | src/forms/Form_FormCalibracionEquipoAlta.cls | 188 | 31 | 4 | igual | analizado | [formularios/Form_FormCalibracionEquipoAlta.md](formularios/Form_FormCalibracionEquipoAlta.md) |
| Form_FormCalibracionEquipoEdicion | src/forms/Form_FormCalibracionEquipoEdicion.cls | 407 | 34 | 8 | igual | analizado | [formularios/Form_FormCalibracionEquipoEdicion.md](formularios/Form_FormCalibracionEquipoEdicion.md) |
| Form_FormCalibracionEquipoGestion | src/forms/Form_FormCalibracionEquipoGestion.cls | 800 | 38 | 26 | igual | analizado | [formularios/Form_FormCalibracionEquipoGestion.md](formularios/Form_FormCalibracionEquipoGestion.md) |
| Form_FormCalibracionGestion | src/forms/Form_FormCalibracionGestion.cls | 625 | 36 | 12 | igual | analizado | [formularios/Form_FormCalibracionGestion.md](formularios/Form_FormCalibracionGestion.md) |
| Form_FormEquipoAlta | src/forms/Form_FormEquipoAlta.cls | 329 | 31 | 6 | igual | analizado | [formularios/Form_FormEquipoAlta.md](formularios/Form_FormEquipoAlta.md) |
| Form_FormEquipoEdicion | src/forms/Form_FormEquipoEdicion.cls | 602 | 33 | 11 | igual | analizado | [formularios/Form_FormEquipoEdicion.md](formularios/Form_FormEquipoEdicion.md) |
| Form_FormEquipoGestion | src/forms/Form_FormEquipoGestion.cls | 838 | 38 | 27 | igual | analizado | [formularios/Form_FormEquipoGestion.md](formularios/Form_FormEquipoGestion.md) |
| Form_FormEquipoMedida | src/forms/Form_FormEquipoMedida.cls | 206 | 37 | 6 | igual | analizado | [formularios/Form_FormEquipoMedida.md](formularios/Form_FormEquipoMedida.md) |
| Form_FormEquipoMedidaCalibracion | src/forms/Form_FormEquipoMedidaCalibracion.cls | 203 | 36 | 6 | igual | analizado | [formularios/Form_FormEquipoMedidaCalibracion.md](formularios/Form_FormEquipoMedidaCalibracion.md) |
| Form_FormEquipoMedidaCalibracionesGestion | src/forms/Form_FormEquipoMedidaCalibracionesGestion.cls | 331 | 24 | 10 | igual | analizado | [formularios/Form_FormEquipoMedidaCalibracionesGestion.md](formularios/Form_FormEquipoMedidaCalibracionesGestion.md) |
| Form_FormEquiposMedidaGestion | src/forms/Form_FormEquiposMedidaGestion.cls | 471 | 29 | 14 | igual | analizado | [formularios/Form_FormEquiposMedidaGestion.md](formularios/Form_FormEquiposMedidaGestion.md) |
| Form_FormEventoAlta | src/forms/Form_FormEventoAlta.cls | 880 | 65 | 19 | igual | analizado | [formularios/Form_FormEventoAlta.md](formularios/Form_FormEventoAlta.md) |
| Form_FormEventoCopia | src/forms/Form_FormEventoCopia.cls | 455 | 20 | 7 | igual | analizado | [formularios/Form_FormEventoCopia.md](formularios/Form_FormEventoCopia.md) |
| Form_FormEventoEdicion | src/forms/Form_FormEventoEdicion.cls | 1308 | 72 | 25 | igual | analizado | [formularios/Form_FormEventoEdicion.md](formularios/Form_FormEventoEdicion.md) |
| Form_FormEventoEqMedida | src/forms/Form_FormEventoEqMedida.cls | 283 | 23 | 7 | igual | analizado | [formularios/Form_FormEventoEqMedida.md](formularios/Form_FormEventoEqMedida.md) |
| Form_FormEventoEqMedidaEdicion | src/forms/Form_FormEventoEqMedidaEdicion.cls | 307 | 23 | 7 | igual | analizado | [formularios/Form_FormEventoEqMedidaEdicion.md](formularios/Form_FormEventoEqMedidaEdicion.md) |
| Form_FormEventoFranqueo | src/forms/Form_FormEventoFranqueo.cls | 103 | 19 | 3 | igual | analizado | [formularios/Form_FormEventoFranqueo.md](formularios/Form_FormEventoFranqueo.md) |
| Form_FormEventoGestion | src/forms/Form_FormEventoGestion.cls | 2892 | 74 | 63 | igual | analizado | [formularios/Form_FormEventoGestion.md](formularios/Form_FormEventoGestion.md) |
| Form_FormEventoSLA | src/forms/Form_FormEventoSLA.cls | 450 | 49 | 6 | igual | analizado | [formularios/Form_FormEventoSLA.md](formularios/Form_FormEventoSLA.md) |
| Form_FormEventoTipoGestion | src/forms/Form_FormEventoTipoGestion.cls | 222 | 18 | 7 | igual | analizado | [formularios/Form_FormEventoTipoGestion.md](formularios/Form_FormEventoTipoGestion.md) |
| Form_FormEventosAltaMasivaPorNS | src/forms/Form_FormEventosAltaMasivaPorNS.cls | 629 | 55 | 17 | igual | analizado | [formularios/Form_FormEventosAltaMasivaPorNS.md](formularios/Form_FormEventosAltaMasivaPorNS.md) |
| Form_FormFacturaAlta | src/forms/Form_FormFacturaAlta.cls | 248 | 37 | 6 | igual | analizado | [formularios/Form_FormFacturaAlta.md](formularios/Form_FormFacturaAlta.md) |
| Form_FormFacturaResultado | src/forms/Form_FormFacturaResultado.cls | 337 | 48 | 6 | igual | analizado | [formularios/Form_FormFacturaResultado.md](formularios/Form_FormFacturaResultado.md) |
| Form_FormFacturacionGestion | src/forms/Form_FormFacturacionGestion.cls | 666 | 32 | 18 | igual | analizado | [formularios/Form_FormFacturacionGestion.md](formularios/Form_FormFacturacionGestion.md) |
| Form_FormGastoAlta | src/forms/Form_FormGastoAlta.cls | 250 | 34 | 8 | igual | analizado | [formularios/Form_FormGastoAlta.md](formularios/Form_FormGastoAlta.md) |
| Form_FormGastoEdicion | src/forms/Form_FormGastoEdicion.cls | 558 | 36 | 13 | igual | analizado | [formularios/Form_FormGastoEdicion.md](formularios/Form_FormGastoEdicion.md) |
| Form_FormGastoGestion | src/forms/Form_FormGastoGestion.cls | 683 | 36 | 23 | igual | analizado | [formularios/Form_FormGastoGestion.md](formularios/Form_FormGastoGestion.md) |
| Form_FormInformeSLA | src/forms/Form_FormInformeSLA.cls | 753 | 27 | 13 | igual | analizado | [formularios/Form_FormInformeSLA.md](formularios/Form_FormInformeSLA.md) |
| Form_FormInformesDesdeLista | src/forms/Form_FormInformesDesdeLista.cls | 272 | 26 | 11 | igual | analizado | [formularios/Form_FormInformesDesdeLista.md](formularios/Form_FormInformesDesdeLista.md) |
| Form_FormInformesParaRAC | src/forms/Form_FormInformesParaRAC.cls | 330 | 24 | 10 | igual | analizado | [formularios/Form_FormInformesParaRAC.md](formularios/Form_FormInformesParaRAC.md) |
| Form_FormInformesPorFactura | src/forms/Form_FormInformesPorFactura.cls | 312 | 28 | 11 | igual | analizado | [formularios/Form_FormInformesPorFactura.md](formularios/Form_FormInformesPorFactura.md) |
| Form_FormMaterialAltaReparacion | src/forms/Form_FormMaterialAltaReparacion.cls | 594 | 69 | 10 | igual | analizado | [formularios/Form_FormMaterialAltaReparacion.md](formularios/Form_FormMaterialAltaReparacion.md) |
| Form_FormMaterialEleccion | src/forms/Form_FormMaterialEleccion.cls | 163 | 18 | 3 | igual | analizado | [formularios/Form_FormMaterialEleccion.md](formularios/Form_FormMaterialEleccion.md) |
| Form_FormMaterialIntervencionAlta | src/forms/Form_FormMaterialIntervencionAlta.cls | 258 | 39 | 6 | igual | analizado | [formularios/Form_FormMaterialIntervencionAlta.md](formularios/Form_FormMaterialIntervencionAlta.md) |
| Form_FormMaterialIntervencionEdicion | src/forms/Form_FormMaterialIntervencionEdicion.cls | 735 | 41 | 8 | igual | analizado | [formularios/Form_FormMaterialIntervencionEdicion.md](formularios/Form_FormMaterialIntervencionEdicion.md) |
| Form_FormMaterialReparacionEdicion | src/forms/Form_FormMaterialReparacionEdicion.cls | 667 | 61 | 9 | igual | analizado | [formularios/Form_FormMaterialReparacionEdicion.md](formularios/Form_FormMaterialReparacionEdicion.md) |
| Form_FormMaterialRepuestoAlta | src/forms/Form_FormMaterialRepuestoAlta.cls | 247 | 58 | 7 | igual | analizado | [formularios/Form_FormMaterialRepuestoAlta.md](formularios/Form_FormMaterialRepuestoAlta.md) |
| Form_FormMaterialRepuestoEdicion | src/forms/Form_FormMaterialRepuestoEdicion.cls | 631 | 63 | 10 | igual | analizado | [formularios/Form_FormMaterialRepuestoEdicion.md](formularios/Form_FormMaterialRepuestoEdicion.md) |
| Form_FormMaterialSeguimiento | src/forms/Form_FormMaterialSeguimiento.cls | 803 | 54 | 12 | igual | analizado | [formularios/Form_FormMaterialSeguimiento.md](formularios/Form_FormMaterialSeguimiento.md) |
| Form_FormMaterialesGestion | src/forms/Form_FormMaterialesGestion.cls | 1393 | 51 | 44 | igual | pendiente | — |
| Form_FormNodoBUIGestion | src/forms/Form_FormNodoBUIGestion.cls | 321 | 26 | 13 | igual | pendiente | — |
| Form_FormNodoGestion | src/forms/Form_FormNodoGestion.cls | 235 | 18 | 7 | igual | pendiente | — |
| Form_FormOriginadorGestion | src/forms/Form_FormOriginadorGestion.cls | 218 | 18 | 7 | igual | pendiente | — |
| Form_FormParteGestion | src/forms/Form_FormParteGestion.cls | 657 | 30 | 15 | igual | pendiente | — |
| Form_FormPlanificacionAlta | src/forms/Form_FormPlanificacionAlta.cls | 208 | 35 | 7 | igual | pendiente | — |
| Form_FormPlanificacionAltaEquipo | src/forms/Form_FormPlanificacionAltaEquipo.cls | 86 | 21 | 3 | igual | pendiente | — |
| Form_FormPlanificacionAnexos | src/forms/Form_FormPlanificacionAnexos.cls | 351 | 25 | 7 | igual | pendiente | — |
| Form_FormPlanificacionCalendarios | src/forms/Form_FormPlanificacionCalendarios.cls | 200 | 18 | 4 | igual | pendiente | — |
| Form_FormPlanificacionCambioIDEvento | src/forms/Form_FormPlanificacionCambioIDEvento.cls | 282 | 31 | 6 | igual | pendiente | — |
| Form_FormPlanificacionCierre | src/forms/Form_FormPlanificacionCierre.cls | 242 | 44 | 8 | igual | pendiente | — |
| Form_FormPlanificacionEdicionEquipo | src/forms/Form_FormPlanificacionEdicionEquipo.cls | 287 | 22 | 7 | igual | pendiente | — |
| Form_FormPlanificacionEquipos | src/forms/Form_FormPlanificacionEquipos.cls | 862 | 38 | 25 | igual | pendiente | — |
| Form_FormPlanificacionGestion | src/forms/Form_FormPlanificacionGestion.cls | 955 | 46 | 19 | igual | pendiente | — |
| Form_FormPlanificacionReprogramacion | src/forms/Form_FormPlanificacionReprogramacion.cls | 278 | 47 | 11 | igual | pendiente | — |
| Form_FormPlanificacionesBusqueda | src/forms/Form_FormPlanificacionesBusqueda.cls | 1058 | 52 | 32 | igual | pendiente | — |
| Form_FormSubContratacionAlta | src/forms/Form_FormSubContratacionAlta.cls | 161 | 30 | 4 | igual | pendiente | — |
| Form_FormSubContratacionEdicion | src/forms/Form_FormSubContratacionEdicion.cls | 495 | 32 | 9 | igual | pendiente | — |
| Form_FormSubContratacionGestion | src/forms/Form_FormSubContratacionGestion.cls | 762 | 35 | 24 | igual | pendiente | — |
| Form_FormSubSistemaAlta | src/forms/Form_FormSubSistemaAlta.cls | 146 | 23 | 4 | igual | pendiente | — |
| Form_FormSubSistemaEdicion | src/forms/Form_FormSubSistemaEdicion.cls | 288 | 23 | 6 | igual | pendiente | — |
| Form_FormSubSistemaGestion | src/forms/Form_FormSubSistemaGestion.cls | 384 | 26 | 12 | igual | pendiente | — |
| Form_FormTecnico | src/forms/Form_FormTecnico.cls | 270 | 34 | 7 | igual | pendiente | — |
| Form_FormTecnicoAusencias | src/forms/Form_FormTecnicoAusencias.cls | 216 | 29 | 6 | igual | pendiente | — |
| Form_FormTecnicoAusenciasGestion | src/forms/Form_FormTecnicoAusenciasGestion.cls | 491 | 26 | 16 | igual | pendiente | — |
| Form_FormTecnicoConsultaHoras | src/forms/Form_FormTecnicoConsultaHoras.cls | 450 | 25 | 5 | igual | pendiente | — |
| Form_FormTecnicoFiestas | src/forms/Form_FormTecnicoFiestas.cls | 79 | 19 | 3 | igual | pendiente | — |
| Form_FormTecnicoFiestasGestion | src/forms/Form_FormTecnicoFiestasGestion.cls | 316 | 22 | 11 | igual | pendiente | — |
| Form_FormTecnicoGestion | src/forms/Form_FormTecnicoGestion.cls | 645 | 34 | 21 | igual | pendiente | — |
| Form_FormTecnicoTipo | src/forms/Form_FormTecnicoTipo.cls | 112 | 25 | 3 | igual | pendiente | — |
| Form_FormTipoTecnicoPrecio | src/forms/Form_FormTipoTecnicoPrecio.cls | 170 | 26 | 5 | igual | pendiente | — |
| Form_FormTipoTecnicoPreciosGestion | src/forms/Form_FormTipoTecnicoPreciosGestion.cls | 448 | 24 | 14 | igual | pendiente | — |
| Form_FormUbicacionGestion | src/forms/Form_FormUbicacionGestion.cls | 220 | 19 | 7 | igual | pendiente | — |

### Módulos estándar

| Objeto | Ruta | Líneas | Controles | Proc. | Frente al snapshot | Estado | Ficha o motivo |
|---|---|---|---|---|---|---|---|
| Constructor | src/modules/Constructor.bas | 5698 | — | 96 | igual | analizado | [modulos.md](modulos.md) |
| Funciones Generales | src/modules/Funciones Generales.bas | 7942 | — | 106 | igual | analizado | [modulos.md](modulos.md) |
| HTML | src/modules/HTML.bas | 258 | — | 4 | igual | analizado | [modulos.md](modulos.md) |
| JsonConverter | src/modules/JsonConverter.bas | 1399 | — | 31 | igual | analizado | [modulos.md](modulos.md) |
| Módulo1 | src/modules/Módulo1.bas | 181 | — | 4 | igual | analizado | [modulos.md](modulos.md) |
| Módulo2 | src/modules/Módulo2.bas | 448 | — | 6 | igual | analizado | [modulos.md](modulos.md) |
| Pruebas | src/modules/Pruebas.bas | 41 | — | 1 | igual | no aplicable | módulo de pruebas; ver [modulos.md](modulos.md) |
| PruebasSLA | src/modules/PruebasSLA.bas | 1697 | — | 55 | igual | no aplicable | módulo de pruebas; ver [modulos.md](modulos.md) |
| SLAHTMLService | src/modules/SLAHTMLService.bas | 1163 | — | 11 | igual | analizado | [modulos.md](modulos.md) |
| SLAReportService | src/modules/SLAReportService.bas | 629 | — | 6 | igual | analizado | [modulos.md](modulos.md) |
| SLAValidator | src/modules/SLAValidator.bas | 166 | — | 2 | igual | analizado | [modulos.md](modulos.md) |
| Test_SLAReportService | src/modules/Test_SLAReportService.bas | 411 | — | 13 | igual | no aplicable | módulo de pruebas; ver [modulos.md](modulos.md) |
| Test_SLAValidator | src/modules/Test_SLAValidator.bas | 515 | — | 13 | igual | no aplicable | módulo de pruebas; ver [modulos.md](modulos.md) |
| Variables Globales | src/modules/Variables Globales.bas | 583 | — | 5 | distinto | analizado | [modulos.md](modulos.md) |
| modMigracionesSLA | src/modules/modMigracionesSLA.bas | 199 | — | 5 | igual | analizado | [modulos.md](modulos.md) |
| modTestingSLA | src/modules/modTestingSLA.bas | 382 | — | 6 | igual | no aplicable | módulo de pruebas; ver [modulos.md](modulos.md) |

### Clases

| Objeto | Ruta | Líneas | Controles | Proc. | Frente al snapshot | Estado | Ficha o motivo |
|---|---|---|---|---|---|---|---|
| Actividad | src/classes/Actividad.cls | 623 | — | 15 | igual | analizado | [clases-1.md](clases-1.md) |
| Anexo | src/classes/Anexo.cls | 1083 | — | 31 | igual | analizado | [clases-1.md](clases-1.md) |
| Entorno | src/classes/Entorno.cls | 1207 | — | 37 | igual | analizado | [clases-1.md](clases-1.md) |
| Equipo | src/classes/Equipo.cls | 189 | — | 5 | igual | analizado | [clases-1.md](clases-1.md) |
| EquipoMedida | src/classes/EquipoMedida.cls | 671 | — | 16 | igual | analizado | [clases-1.md](clases-1.md) |
| EquipoMedidaCalibracion | src/classes/EquipoMedidaCalibracion.cls | 299 | — | 7 | igual | analizado | [clases-1.md](clases-1.md) |
| Evento | src/classes/Evento.cls | 3015 | — | 60 | igual | analizado | [clases-1.md](clases-1.md) |
| EventoBusqueda | src/classes/EventoBusqueda.cls | 21 | — | 0 | igual | analizado | [clases-1.md](clases-1.md) |
| EventoEquipoMedida | src/classes/EventoEquipoMedida.cls | 200 | — | 8 | igual | analizado | [clases-1.md](clases-1.md) |
| EventoEquipoMedidaAntes | src/classes/EventoEquipoMedidaAntes.cls | 185 | — | 7 | igual | analizado | [clases-1.md](clases-1.md) |
| Factura | src/classes/Factura.cls | 6827 | — | 98 | igual | analizado | [clases-1.md](clases-1.md) |
| FacturaPerfiles | src/classes/FacturaPerfiles.cls | 324 | — | 9 | igual | analizado | [clases-1.md](clases-1.md) |
| Fiesta | src/classes/Fiesta.cls | 202 | — | 6 | igual | analizado | [clases-1.md](clases-1.md) |
| Gasto | src/classes/Gasto.cls | 433 | — | 12 | igual | analizado | [clases-1.md](clases-1.md) |
| GestorInforme | src/classes/GestorInforme.cls | 281 | — | 8 | igual | analizado | [clases-2.md](clases-2.md) |
| Informe | src/classes/Informe.cls | 987 | — | 11 | igual | analizado | [clases-2.md](clases-2.md) |
| LIbranza | src/classes/LIbranza.cls | 398 | — | 9 | igual | analizado | [clases-2.md](clases-2.md) |
| Material | src/classes/Material.cls | 714 | — | 17 | igual | analizado | [clases-2.md](clases-2.md) |
| MaterialSeguimiento | src/classes/MaterialSeguimiento.cls | 378 | — | 9 | igual | analizado | [clases-2.md](clases-2.md) |
| Parte | src/classes/Parte.cls | 645 | — | 16 | igual | analizado | [clases-2.md](clases-2.md) |
| Subcontratacion | src/classes/Subcontratacion.cls | 374 | — | 10 | igual | analizado | [clases-2.md](clases-2.md) |
| Tecnico | src/classes/Tecnico.cls | 703 | — | 13 | igual | analizado | [clases-2.md](clases-2.md) |
| TecnicoHoras | src/classes/TecnicoHoras.cls | 528 | — | 15 | igual | analizado | [clases-2.md](clases-2.md) |
| TipoDia | src/classes/TipoDia.cls | 186 | — | 5 | igual | analizado | [clases-2.md](clases-2.md) |
| TipoTecnico | src/classes/TipoTecnico.cls | 526 | — | 17 | igual | analizado | [clases-2.md](clases-2.md) |
| TipoTecnicoPrecio | src/classes/TipoTecnicoPrecio.cls | 892 | — | 24 | igual | analizado | [clases-2.md](clases-2.md) |
| Usuario | src/classes/Usuario.cls | 412 | — | 6 | igual | analizado | [clases-2.md](clases-2.md) |
| UsuarioAplicacionPermisos | src/classes/UsuarioAplicacionPermisos.cls | 402 | — | 19 | igual | analizado | [clases-2.md](clases-2.md) |

Totales: 93 analizados · 5 no aplicables · 33 pendientes.

## Discrepancias abiertas

| ID | Discrepancia | Evidencia | Estado |
|---|---|---|---|
| D-BRA-1 | ERD de origen: falta la extracción de solo lectura del `.accdb` (workflow `dysflow-erd`, #1144) | `ERD/` del legacy | pendiente del propietario |

## Método del recuento

Clon de solo lectura bajo `.atl/sources/brass/` (ruta ignorada) con HEAD desacoplado en el commit fijado e índice CodeGraph-VBA (`codegraph init -i`). Los recuentos salen de un script que cuenta líneas, bloques `Begin` y declaraciones de procedimiento; la comparación con el snapshot usa la normalización descrita arriba.
