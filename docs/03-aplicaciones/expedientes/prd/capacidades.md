# Capacidades de Expedientes RC1 frente al estudio previo

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Issue:** #1073 · **Fase:** 5 del [roadmap](../../../07-migracion/roadmap-descubrimiento-prd.md) · **PRD:** [README](README.md)

Clasifica las capacidades del [ledger normativo](../capability-ledger.md) (EXP-CAP-001 a 063) como **añadida**, **modificada**, **retirada** o **sin cambios** respecto al snapshot del 2026-08-06, con evidencia, y enlaza cada una con las fichas de formulario que la implementan en RC1.

## Método y límites

- Cada capacidad se asocia a los objetos de lógica de RC1 (clases y módulos) que la implementan, tras leer las 45 fichas. La asociación es **juicio del estudio**, no dato del ledger: puede discutirse fila a fila.
- El estado sale de comparar cada objeto con el mismo fichero del snapshot (`data/staging/expedientes/src`) tras quitar las líneas `Attribute VB_`, los comentarios, las líneas en blanco y las diferencias de espacios y mayúsculas. «Distinto» significa que cambió el texto del código, **no** que cambiara el comportamiento: el cambio puede ser un refactor. La ficha del formulario es el lugar donde se discute el efecto.
- Los 42 formularios comunes con el snapshot difieren todos, también tras esa normalización ([D-EXP-3](README.md#discrepancias-abiertas)); por eso el estado se mide sobre la lógica y no sobre el formulario.
- «Retirados» son módulos del snapshot que RC1 ya no tiene. El snapshot los trataba como ayudantes (`Helper_*`, `E2EExportService`, `E2EManagementService`, `GuardadoAutomaticoHelper`); RC1 concentra esa lógica en las clases `*Operaciones`, `Expediente` y los módulos E2E nuevos.
- Las capacidades 057 a 063 del ledger son mecanismos legacy a retirar, no capacidades de negocio; se listan sin clasificar.
- No se ha ejecutado nada en Access: «sin cambios» es igualdad de texto normalizado.

## Resumen

| Estado | Capacidades |
|---|---|
| sin cambios en la lógica | 6 |
| modificada | 32 |
| añadida | 2 |
| retirada | 2 |
| solo formulario (sin lógica aislada) | 1 |
| sin equivalente comprobado en Access | 13 |
| mecanismo legacy a retirar | 7 |
| **Total** | **63** |

## Clasificación

| Capacidad | Descripción | Disposición del ledger | Estado frente al snapshot | Objetos de lógica | Fichas de formulario |
|---|---|---|---|---|---|
| EXP-CAP-001 | Alta transaccional del expediente, incluida el alta iniciada por sistemas autorizados. | PRESERVAR | **modificada** | distintos: `Expediente.cls`, `ExpedienteOperaciones.cls`; retirados: `Helper_ExpedienteAlta.bas` | [Expediente](formularios/Form_FormExpediente.md), [ExpedienteAltaTipo](formularios/Form_FormExpedienteAltaTipo.md), [ExpedienteAltaParaHPS](formularios/Form_FormExpedienteAltaParaHPS.md) |
| EXP-CAP-002 | Edición del agregado con invariantes y concurrencia explícitas. | PRESERVAR | **modificada** | distintos: `Expediente.cls`, `ExpedienteOperaciones.cls`; retirados: `Helper_ExpedienteEdicion.bas` | [Expediente](formularios/Form_FormExpediente.md), [ExpedienteGeneral](formularios/Form_FormExpedienteGeneral.md), [ExpedienteFechas](formularios/Form_FormExpedienteFechas.md) |
| EXP-CAP-003 | Baja condicionada, auditable y sin pérdida accidental de relacionados. | PRESERVAR | **modificada** | distintos: `ExpedienteOperaciones.cls`; retirados: `Helper_ExpedienteEliminacion.bas` | [ExpedientesGestion](formularios/Form_FormExpedientesGestion.md) |
| EXP-CAP-004 | Cambio de tipo con validación de elegibilidad y efectos derivados. | PRESERVAR | **modificada** | distintos: `ExpedienteOperaciones.cls`, `ExpedienteCompleto.cls` | [ExpedienteCambioTipo](formularios/Form_FormExpedienteCambioTipo.md), [ExpedientesParaCambioTipo](formularios/Form_FormExpedientesParaCambioTipo.md) |
| EXP-CAP-005 | Estado calculado y transiciones con fechas, garantía y motivos. | PRESERVAR | **modificada** | distintos: `Expediente.cls`, `FUNCIONES UTILES.bas`; retirados: `Helper_ExpedienteFechasLogic.bas` | [ExpedienteFechas](formularios/Form_FormExpedienteFechas.md) |
| EXP-CAP-006 | Jerarquía de acuerdos marco, lotes y basados, incluida integridad padre-hijo. | PRESERVAR | **modificada** | distintos: `Expediente.cls` | [ExpedienteAltaTipo](formularios/Form_FormExpedienteAltaTipo.md), [ExpedientesParaCambioTipo](formularios/Form_FormExpedientesParaCambioTipo.md) |
| EXP-CAP-007 | Derivación ordinal funcional dentro de la jerarquía. | PRESERVAR | **modificada** | distintos: `Expediente.cls` | [ExpedienteGeneral](formularios/Form_FormExpedienteGeneral.md) |
| EXP-CAP-008 | Hitos y sus reglas temporales. | PRESERVAR | **modificada** | distintos: `ExpedienteOperaciones.cls` | [ExpedienteHitos](formularios/Form_FormExpedienteHitos.md) |
| EXP-CAP-009 | Modificados, cambios de campo e historial reciente. | PRESERVAR | **sin cambios en la lógica** | — | [ExpedienteModificados](formularios/Form_FormExpedienteModificados.md), [Modificado](formularios/Form_FormModificado.md) |
| EXP-CAP-010 | Anexos y referencias documentales con eliminación y retención controladas. | MODERNIZAR | **modificada** | distintos: `ExpedienteOperaciones.cls`; retirados: `Helper_ExpedienteAnexos.bas` | [ExpedienteDocumentacion](formularios/Form_FormExpedienteDocumentacion.md) |
| EXP-CAP-011 | Anualidades vinculadas al expediente. | PRESERVAR | **modificada** | distintos: `ExpedienteOperaciones.cls` | [ExpedienteEntidades](formularios/Form_FormExpedienteEntidades.md) |
| EXP-CAP-012 | Responsables por rol y jefaturas. | PRESERVAR | **modificada** | distintos: `ExpedienteResponsable.cls`; añadidos: `modResponsablesEdit.bas`, `clsResponsablesEditState.cls`; retirados: `modResponsablePorRolService.bas` | [ExpedientesResponsables](formularios/Form_FormExpedientesResponsables.md), [ExpedienteEntidades](formularios/Form_FormExpedienteEntidades.md) |
| EXP-CAP-013 | Relaciones con entidades, jurídicas y cadena de entidades. | PRESERVAR | **modificada** | distintos: `ExpedienteEntidad.cls`, `ExpedienteEntidadOperaciones.cls`; retirados: `Helper_EntidadCRUD.bas`, `Helper_ExpedienteEntidadesTabla.cls` | [ExpedienteEntidades](formularios/Form_FormExpedienteEntidades.md) |
| EXP-CAP-014 | Suministradores, contratistas, subcontratistas, UTE y jerarquía «el árbol manda». | MODERNIZAR | **modificada** | distintos: `ExpedienteSuministradorServicio.cls`, `ExpedienteSuministradorRepositorio.bas`; retirados: `Helper_ExpedienteSuministradoresArbol.cls` | [ExpedienteSuministradores](formularios/Form_FormExpedienteSuministradores.md) |
| EXP-CAP-015 | Comerciales. | PRESERVAR | **modificada** | distintos: `ComercialOperaciones.cls` | [Comercial](formularios/Form_FormComercial.md), [ComercialesGestion](formularios/Form_FormComercialesGestion.md) |
| EXP-CAP-016 | CPV. | PRESERVAR | **sin cambios en la lógica** | — | [CPV](formularios/Form_FormCPV.md), [CPVsGestion](formularios/Form_FormCPVsGestion.md) |
| EXP-CAP-017 | Ejércitos. | PRESERVAR | **modificada** | distintos: `EjercitoOperaciones.cls` | [Ejercito](formularios/Form_FormEjercito.md), [EjercitosGestion](formularios/Form_FormEjercitosGestion.md) |
| EXP-CAP-018 | Suministradores. | PRESERVAR | **modificada** | distintos: `SuministradorOperaciones.cls` | [Suministrador](formularios/Form_FormSuministrador.md), [SuministradoresGestion](formularios/Form_FormSuministradoresGestion.md) |
| EXP-CAP-019 | Lugares de ejecución. | PRESERVAR | **modificada** | distintos: `LugarEjecucionOperaciones.cls` | [LugarEjecucion](formularios/Form_FormLugarEjecucion.md), [LugarEjecucionGestion](formularios/Form_FormLugarEjecucionGestion.md) |
| EXP-CAP-020 | PECAL. | PRESERVAR | **modificada** | distintos: `PECALOperaciones.cls` | [PECAL](formularios/Form_FormPECAL.md), [PECALESGestion](formularios/Form_FormPECALESGestion.md) |
| EXP-CAP-021 | RAC. | PRESERVAR | **modificada** | distintos: `RACOperaciones.cls` | [RAC](formularios/Form_FormRAC.md), [RACSGestion](formularios/Form_FormRACSGestion.md) |
| EXP-CAP-022 | Grados de clasificación. | PRESERVAR | **modificada** | distintos: `GradoClasificacionOperaciones.cls` | [GradoClasificacion](formularios/Form_FormGradoClasificacion.md), [GradosClasificacionGestion](formularios/Form_FormGradosClasificacionGestion.md) |
| EXP-CAP-023 | Órganos de contratación. | PRESERVAR | **modificada** | distintos: `OrganoContratacionOperaciones.cls` | [OrganoContratacion](formularios/Form_FormOrganoContratacion.md), [OrganoContratacionGestion](formularios/Form_FormOrganoContratacionGestion.md) |
| EXP-CAP-024 | Oficinas de programa. | PRESERVAR | **modificada** | distintos: `OficinaProgramaOperaciones.cls` | [OficinaPrograma](formularios/Form_FormOficinaPrograma.md), [OficinasProgramaGestion](formularios/Form_FormOficinasProgramaGestion.md) |
| EXP-CAP-025 | Bandeja paginada con filtros y detalle de solo lectura. | MODERNIZAR | **modificada** | distintos: `ExpedienteCompleto.cls`, `Entorno.cls`; retirados: `Helper_ExpedienteConsultas.bas` | [ExpedientesGestion](formularios/Form_FormExpedientesGestion.md) |
| EXP-CAP-026 | Búsqueda avanzada con semántica de filtros preservada. | PRESERVAR | **modificada** | distintos: `ExpedienteBusqueda.cls`; añadidos: `modExpedienteSearch.bas` | [ExpedientesGestion](formularios/Form_FormExpedientesGestion.md) |
| EXP-CAP-027 | Búsqueda técnica sujeta a permiso efectivo. | PRESERVAR | **sin cambios en la lógica** | — | [ExpedientesGestionTecnica](formularios/Form_FormExpedientesGestionTecnica.md) |
| EXP-CAP-028 | Exportación Excel reproducible y autorizada. | PRESERVAR | **modificada** | añadidos: `modInformeExpedientesHelper.bas`; retirados: `Helper_ExpedientesExportExcel.bas` | [EleccionTipoConsulta](formularios/Form_FormEleccionTipoConsulta.md) |
| EXP-CAP-029 | Tareas calculadas por reglas del expediente, sin inventar una tabla de tareas. | MODERNIZAR | **modificada** | distintos: `Entorno.cls`, `FUNCIONES UTILES.bas` | [Tareas](formularios/Form_FormTareas.md) |
| EXP-CAP-030 | Autosave de datos generales y fechas con transacción explícita. | MODERNIZAR | **retirada** | retirados: `GuardadoAutomaticoHelper.bas` | — |
| EXP-CAP-031 | Autosave de hitos, modificados y restantes secciones editables. | MODERNIZAR | **retirada** | retirados: `GuardadoAutomaticoHelper.bas` | — |
| EXP-CAP-032 | Idempotencia, prevención de doble envío y feedback accesible de progreso/error. | MODERNIZAR | **sin equivalente comprobado en Access** | concepto de la web (sin objeto equivalente identificado en RC1) | — |
| EXP-CAP-033 | DTO de dominio estable para intercambio. | PRESERVAR | **sin cambios en la lógica** | — | — |
| EXP-CAP-034 | Exportación JSON completa y determinista. | PRESERVAR | **modificada** | distintos: `ExpedienteJsonExporter.bas`, `JsonConverter.bas`; retirados: `ExpedienteE2EUtils.bas` | — |
| EXP-CAP-035 | Gestión y ejecución batch E2E. | MODERNIZAR | **modificada** | añadidos: `ManualE2EExport.bas`; retirados: `E2EExportService.bas`, `E2EManagementService.bas` | [E2ECola](formularios/Form_FormE2ECola.md) |
| EXP-CAP-036 | Hash de contenido con algoritmo y canonicalización versionados. | PRESERVAR | **modificada** | distintos: `E2EHashService.bas` | — |
| EXP-CAP-037 | Generación del paquete de exportación E2E. | PRESERVAR | **modificada** | añadidos: `ManualE2EExport.bas`; retirados: `E2EExportService.bas` | — |
| EXP-CAP-038 | Trazabilidad entre expediente, batch, artefacto y resultado. | PRESERVAR | **añadida** | añadidos: `E2EColaHtmlSnapshot.bas`, `modE2EColaPanelHelper.bas`, `modNotifyChangeE2E.bas` | [E2ECola](formularios/Form_FormE2ECola.md) |
| EXP-CAP-039 | Selección manual de expedientes para batch. | PRESERVAR | **añadida** | añadidos: `FormE2EGestionEnlacesLogic.bas`, `E2EEnlacesService.bas` | [E2EGestionEnlaces](formularios/Form_FormE2EGestionEnlaces.md) |
| EXP-CAP-040 | Destino de exportación por usuario, sin rutas locales como contrato. | MODERNIZAR | **sin equivalente comprobado en Access** | concepto de la web (sin objeto equivalente identificado en RC1) | — |
| EXP-CAP-041 | Sesión de exportación, reanudación y cierre. | PRESERVAR | **sin equivalente comprobado en Access** | concepto de la web (sin objeto equivalente identificado en RC1) | — |
| EXP-CAP-042 | Ordinal E2E único y consistente en dominio y trazabilidad. | PRESERVAR | **modificada** | distintos: `ExpedienteEntidadOperaciones.cls`; añadidos: `M_E2EMigration.bas` | — |
| EXP-CAP-043 | Autorización deny-by-default por capacidad. | MODERNIZAR | **sin equivalente comprobado en Access** | concepto de la web (sin objeto equivalente identificado en RC1) | — |
| EXP-CAP-044 | Login y principal actual consumidos desde Lanzadera. | MODERNIZAR | **sin equivalente comprobado en Access** | concepto de la web (sin objeto equivalente identificado en RC1) | — |
| EXP-CAP-045 | Estado de usuario y sesión sin globals mutables. | MODERNIZAR | **sin equivalente comprobado en Access** | concepto de la web (sin objeto equivalente identificado en RC1) | — |
| EXP-CAP-046 | Auditoría de accesos, cambios, exportaciones e integraciones. | MODERNIZAR | **sin equivalente comprobado en Access** | concepto de la web (sin objeto equivalente identificado en RC1) | — |
| EXP-CAP-047 | Arranque y readiness con fallos diagnosticables. | MODERNIZAR | **sin equivalente comprobado en Access** | concepto de la web (sin objeto equivalente identificado en RC1) | — |
| EXP-CAP-048 | Configuración tipada por entorno e inyección de dependencias. | MODERNIZAR | **sin equivalente comprobado en Access** | concepto de la web (sin objeto equivalente identificado en RC1) | — |
| EXP-CAP-049 | Caché invalidable y read-models reconstruibles. | MODERNIZAR | **sin equivalente comprobado en Access** | concepto de la web (sin objeto equivalente identificado en RC1) | — |
| EXP-CAP-050 | Binding de backend por despliegue, sin selector interactivo ni rutas embebidas. | MODERNIZAR | **sin equivalente comprobado en Access** | concepto de la web (sin objeto equivalente identificado en RC1) | — |
| EXP-CAP-051 | Alta/consulta desde HPS mediante contrato autenticado e idempotente. | MODERNIZAR | **solo formulario (sin lógica aislada)** | — | [ExpedienteAltaParaHPS](formularios/Form_FormExpedienteAltaParaHPS.md) |
| EXP-CAP-052 | Intercambio con AGEDYS con ownership y errores definidos. | PRESERVAR | **sin cambios en la lógica** | — | — |
| EXP-CAP-053 | Enlace con Gestión de Riesgos por identificador estable. | PRESERVAR | **sin equivalente comprobado en Access** | concepto de la web (sin objeto equivalente identificado en RC1) | — |
| EXP-CAP-054 | Enlace con No Conformidades y su estado asociado. | PRESERVAR | **sin equivalente comprobado en Access** | concepto de la web (sin objeto equivalente identificado en RC1) | — |
| EXP-CAP-055 | Notificaciones por correo mediante puerto y trazabilidad. | MODERNIZAR | **modificada** | distintos: `Correo.cls` | — |
| EXP-CAP-056 | Documentos en SharePoint u object storage mediante referencia opaca, retención y autorización. | MODERNIZAR | **sin cambios en la lógica** | — | [ExpedienteDocumentacion](formularios/Form_FormExpedienteDocumentacion.md) |
| EXP-CAP-057 | Resolución de red Win32. | RETIRAR | **mecanismo legacy a retirar** | no es capacidad de negocio | — |
| EXP-CAP-058 | Gestión de procesos Win32. | RETIRAR | **mecanismo legacy a retirar** | no es capacidad de negocio | — |
| EXP-CAP-059 | Drag & drop OLE y controles ActiveX. | RETIRAR | **mecanismo legacy a retirar** | no es capacidad de negocio | — |
| EXP-CAP-060 | Globals y singletons mutables de sesión/entorno. | RETIRAR | **mecanismo legacy a retirar** | no es capacidad de negocio | — |
| EXP-CAP-061 | Cambio interactivo de backend y rutas embebidas. | RETIRAR | **mecanismo legacy a retirar** | no es capacidad de negocio | — |
| EXP-CAP-062 | Popup modal anti-spam/busy. | RETIRAR | **mecanismo legacy a retirar** | no es capacidad de negocio | — |
| EXP-CAP-063 | Menú JSON-hub y navegación acoplada a formularios. | RETIRAR | **mecanismo legacy a retirar** | no es capacidad de negocio | — |

## Capacidades de RC1 sin entrada propia en el ledger

Candidatas a incorporar al ledger o a confirmar como subsumidas. Son propuestas del estudio, no decisiones:

| Candidata | Evidencia en RC1 | Entrada del ledger más cercana |
|---|---|---|
| Enlaces E2E por expediente (alta y baja de enlace, historial y familia) | `TbE2EEnlaces` ([definidas en src](matriz-campos/definidas-en-src.md)); [Form_FormE2EGestionEnlaces](formularios/Form_FormE2EGestionEnlaces.md) | EXP-CAP-038, 039 y 042 |
| Cola de cambios E2E con panel HTML y verificación de importación (`EstadoImportacion`, `FechaVerificacion`, `MotivoFallo`) | `TbE2ECola` ([definidas en src](matriz-campos/definidas-en-src.md)); [Form_FormE2ECola](formularios/Form_FormE2ECola.md) | EXP-CAP-035 y 038 |
| Menú principal por rol y menú de perfil técnico | [Form_Form0BDOpciones](formularios/Form_Form0BDOpciones.md), [Form_Form0BDOpcionesTecnicos](formularios/Form_Form0BDOpcionesTecnicos.md), [Form_Form0BDGestorEntidades](formularios/Form_Form0BDGestorEntidades.md) | EXP-CAP-043 y 063 |
| Arranque con identificación del usuario | [Form_frmSplash](formularios/Form_frmSplash.md) | EXP-CAP-044 y 047 |
| Recálculo y reparación de la caché `TbExpedientesConEntidades` por expediente | botón «Actualizar» de [Form_FormEleccionTipoConsulta](formularios/Form_FormEleccionTipoConsulta.md) y `ComandoActualizarCompleto` de [Form_FormExpediente](formularios/Form_FormExpediente.md) | EXP-CAP-049 |

## Formularios añadidos y retirados respecto al snapshot

La relación con capacidades es hipótesis del estudio, a verificar con el equipo de Expedientes:

| Formulario | Cambio | Capacidad probable | Evidencia |
|---|---|---|---|
| `Form_FormE2ECola` | añadido | EXP-CAP-035 y 038 | [ficha](formularios/Form_FormE2ECola.md); ya no se abre: la opción 12 del gestor abre un panel HTML |
| `Form_FormE2EGestionEnlaces` | añadido | EXP-CAP-039 | [ficha](formularios/Form_FormE2EGestionEnlaces.md) |
| `Form_FormExpedientesResponsables` | añadido | EXP-CAP-012 | [ficha](formularios/Form_FormExpedientesResponsables.md) |
| `Form_FormE2EGestionBatch` | retirado | EXP-CAP-035 | las cuatro tablas E2E del snapshot no se citan en RC1 (H-EXP-81 del [README](README.md)) |
| `Form_FormExpedienteAlta` | retirado | EXP-CAP-001 | relevo probable: `Form_FormExpedienteAltaTipo` |
| `Form_FormUsuariosGestion` | retirado | EXP-CAP-043 y 044 | sin relevo identificado |
| `Form_frmBusy` | retirado | EXP-CAP-062 | sin relevo identificado |

## Pendientes

- Confirmar con el equipo de Expedientes la asociación capacidad-objeto de la tabla y los relevos de los formularios retirados.
- Ejecutar el contraste semántico (no solo textual) de las capacidades marcadas «modificada», empezando por el alta, la edición y la baja.
- Decidir si las candidatas de la sección anterior entran en el ledger.
