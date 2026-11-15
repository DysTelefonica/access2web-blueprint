# Fase 0: Brass (`Gestion_Brass_Gestion`)

Issue #1574. Épica `docs/07-migracion/epica-coexistencia-postgresql.md` (fase 0).
Base de análisis: rama `2026-004.rc1`, SHA `cb2795128748fa91e79b5cf0a0c2501902692e3f`, worktree `00_BRASS-worktrees/2026-004.rc1`.
Fuentes: ERD del front y del backend con `scripts/dysflow-erd/extract-erd.ps1` (solo lectura, sobre copia) y trazado del código con `read`/`grep` sobre `src/`. CodeGraph no estaba operativo en este entorno; se declara el fallback.

## Base de análisis

| Hecho | Evidencia |
|---|---|
| La rama base es `2026-004.rc1` | `docs/brass/roadmap.md:20, 66, 71-74` («Rama base y destino de integración \| `2026-004.rc1`») |
| La UAT está pendiente, no superada | `docs/brass/roadmap.md:53` («sin resultados de ejecución ni firma para este ciclo»), `:21` (próxima compuerta `uat_handoff`), `:93` («Sin compilación manual, tests Access, CI ni firma UAT») |
| `feat/rc1-invoice-sla-gate` no es base de estudio | La rama apunta al commit inicial del clon; su trabajo (`SLAInvoiceGate.bas`, `SLAInvoiceTestResult.bas`, `Test_SLAInvoiceGate.bas`) está sin commitear y «no se ha integrado en rc1 ni sincronizado en un binario» (`docs/brass/roadmap.md:88`) |
| `main` es `3f77d58` y su `src` no coincide con la RC1 | La RC1 borra `Módulo2.bas` para alinear `src` con el binario y añade `SLAExcelService.bas` y `Utilidades.bas`; `src/queries/` y `src/reports/` solo existen en la RC1 |

Consecuencia: el inventario de abajo describe la RC1, no `main`. Las diferencias con `main` no son un error del inventario.

## Estructura del front

| Hecho | Valor |
|---|---|
| Tablas del front | 61: 4 propias y 57 vinculadas |
| Tablas propias del front | `TbActividadesLotes`, `TbAux`, `TbAuxEventosParaInforme`, `TbAuxNombreCamposParaInforme` |
| Destino de los 57 vínculos | Todos a `Gestion_Brass_Gestion_Datos.accdb` en `\\datoste\Aplicaciones_dys\Aplicaciones PpD\BRASS\`, con credencial en la propia cadena de conexión |
| Vínculos a `Lanzadera_Datos` en el front | Ninguno. El acceso a lanzadera es por código, no por vínculo |
| Formularios con `RecordSource` | 0 |
| Controles con `RowSourceType = Table/Query` | Decenas de combos y listas sobre tablas del backend (inventario en el anexo) |
| Controles con `ControlSource` | 10, todos con el valor `IDActividad` y sin `RecordSource` de formulario: inertes |

La excepción de Brass se confirma en su forma exacta: **no hay formularios unidos a recordset**; lo conectado son combos y listas con `RowSource`, más tablas locales de trabajo.

## Aperturas directas

| Apertura | Código | Observación |
|---|---|---|
| Backend propio | `src/modules/Variables Globales.bas:548-573` (`getdb()`) | Resuelve la ruta leyendo `TableDef.Connect` de los vínculos (`:477-518`) y, si falla, cae a `…\BRASS\Gestion_Brass_Gestion_Datos.accdb` (`:564-566`). La credencial va literal en el código (`:573`) |
| Lanzadera | `src/modules/Variables Globales.bas:520-539` (`getdbLanzadera()`) | Rutas `…000datoslocal\Lanzadera_Datos.accdb` (`:528`) y `…0Lanzadera\Lanzadera_Datos.accdb` (`:530`). Consumida en `Constructor.bas:71, 110, 165` y `Funciones Generales.bas:7623, 7628, 7672` |
| `DoCmd.OpenDatabase` o `TransferDatabase` | Ninguno | El único `DoCmd.OpenTable` es `forms/Form_FormInformesDesdeLista.cls:15` sobre `TbAuxEventosParaInforme`, tabla del propio front |

## Tablas de lanzadera que consume el front

`TbUsuariosAplicaciones` (`Constructor.bas:44-45, 71`), `TbUsuariosAplicacionesPermisos` (`Constructor.bas:163-165`) y `TbAplicacionesAperturas` (`Funciones Generales.bas:7623, 7627, 7668`). Las tres están bien ubicadas: existen en `Lanzadera_Datos` y son la autoridad de identidad y permisos.

## Backend `Gestion_Brass_Gestion_Datos`

| Hecho | Valor |
|---|---|
| Tablas | 57 propias, 0 vinculadas |
| Relaciones | 28 |
| Consultas guardadas | 0 |
| Tablas sin clave primaria | 5: `TbAuxActividad`, `TbAuxManteniminetosPreventivosCalendario`, `TbAuxMateriales`, `TbAuxPlanificacion`, `TbBuiIDEvento` |

Los vínculos del front apuntan a la ruta de producción (`\\datoste\…`). El trabajo con bases reales se hace sobre copia; los vínculos del front no se pudieron abrir en este entorno (57 errores de vínculo registrados), por lo que el inventario de vínculos procede de la definición del vínculo, no de la lectura de los datos de origen.

## Veredicto de propiedad

**Se quedan en el backend**: las 43 tablas que el VBA del front usa en SQL o VBA, entre ellas `TbEventos`, `TbActividades`, `TbMaterial`, `TbAnexos`, `TbTecnicos`, `TbPlanificacion`, `TbEquipos`, `TbFactura*`, `TbPartes*`, `Tb0FiltroGestion` y las cuatro auxiliares `TbAuxActividad`, `TbAuxMateriales`, `TbAuxPlanificacion`, `TbAuxEventos` (esta última se abre con `CurrentDb()` en `Funciones Generales.bas:3713`).

**No las usa el front**: `TbAuxManteniminetosPreventivosCalendario`, `TbCausaFin`, `TbGuiaConciliaciones`, `TbRepuestosTipo`, `TbResultadoVerificacion`, `TbSistema`, `TbTipoAsistencia`. Cero referencias en `src/` de la RC1 y de `main`. Requieren comprobar otros consumidores antes de proponer nada (tarea T6).

**Las usa el front pero no existen ni en el front ni en el backend**: `TbDocumentos` (`Funciones Generales.bas:1336-1340`), `TbIDEquipos` (`Form_FormPlanificacionEquipos.form.txt:1143, 1245`), `TbEquiposEquipoMedida` (`EventoEquipoMedida.cls:34`), `TbSellados` (`Form_FormActividadGestion.form.txt:396-400`), `TbPlanificacionAñadida` (`Form_FormPlanificacionCierre.form.txt:975-977`), `TbImportacionPpal`, `TbImportacionDetalle` y `TbMantenimientoPreventivoEquipos` (solo en consultas de análisis). Son referencias muertas o tablas que faltan: un fallo en runtime las delataría.

## Propuestas para después del descubrimiento global

Ninguna se aplica: la puerta de descubrimiento global sigue cerrada (#1576).

| Id | Propuesta | Grupo (fase 1) | Comprobación previa | Reversión |
|---|---|---|---|---|
| B1 | Poner en cuarentena las 7 tablas del backend sin uso en el front: renombrar a `zzBAJA_<tabla>`. | Imprescindible | Confirmar que ningún otro front ni la web las leen (tarea T6). Verificar en la copia de producción el recuento y que ninguna consulta guardada las usa (el backend no tiene consultas guardadas). | Quitar el prefijo. |
| B2 | Las 4 tablas que viven en el front (`TbActividadesLotes`, `TbAux`, `TbAuxEventosParaInforme`, `TbAuxNombreCamposParaInforme`) incumplen la regla de que las tablas del aplicativo viven en su backend. Son tablas de trabajo e informe. Propuesta: mover las que el backend pueda sostener y declarar el resto como temporales del front. | Imprescindible | Medir su uso real: `TbAuxEventosParaInforme` la abre `Form_FormInformesDesdeLista.cls:15`; `TbAuxNombreCamposParaInforme` la usa `Entorno.cls:1175-1176`. Comprobar que el backend puede alojar la tabla con el mismo esquema. | Devolver la tabla al front. |
| B3 | Declarar clave primaria en las 5 tablas auxiliares sin ella: `TbAuxActividad`, `TbAuxManteniminetosPreventivosCalendario`, `TbAuxMateriales`, `TbAuxPlanificacion`, `TbBuiIDEvento`. | Imprescindible | Verificar duplicados en cada una antes de fijar la clave; `TbBuiIDEvento` solo aparece como argumento de `Dame(...)` (`Funciones Generales.bas:6205`), sin referencia de columna. | Quitar la clave. |
| B4 | Retirar la credencial literal del código del front (`Variables Globales.bas:573`) y leerla de la configuración. | Imprescindible | Es la misma credencial que abre todos los binarios del ecosistema. Comprobar que ninguna otra copia del literal queda en el front. | Restaurar la línea original. |
| B5 | Revisar las referencias a tablas inexistentes (`TbDocumentos`, `TbIDEquipos`, `TbEquiposEquipoMedida`, `TbSellados`, `TbPlanificacionAñadida`): o se recuperan en el backend o se elimina la referencia. | Imprescindible | Reproducir en runtime el recorrido afectado: cierre de planificación, alta de actividad y búsqueda de planificación. | Restaurar la tabla o la referencia. |
| B6 | Revisar que el front no escriba en el share de producción desde desarrollo: los vínculos y la ruta de reserva apuntan a `\\datoste\…`. | Imprescindible | Confirmar con el propietario el entorno de cada copia y aislar los vínculos de desarrollo. | Restituir la configuración original. |

Lo que **no** se toca ahora: los 57 vínculos del front (mientras la web no escriba en esas tablas), las 4 tablas locales del front más allá de la propuesta B2, y las 28 relaciones del backend.

## Anexo: controles con origen de datos (RC1, idéntico en `main`)

`RowSourceType = Table/Query` en combos y listas:

| Formulario | Tablas |
|---|---|
| `Form_FormActividadEdicion` | `TbUbicacion`, `TbCodActividad`, `TbEventos` |
| `Form_FormActividadGestion` | `TbActividades`, `TbEventos`, `TbFacturaActividadesInvolucradas` |
| `Form_FormAnexosNoAlcanzables` | `TbEventos`, `TbEquipos`, `TbBUI` |
| `Form_FormBUIGestion`, `Form_FormCalibracionEquipoAlta/Edicion/Gestion` | `TbBUI`, `TbEquiposCalibrables`, `TbEquiposCalibrablesFechas` |
| `Form_FormEquipoAlta/Edicion/Gestion` | `TbEquipos`, `TbSubsistemaBui`, `TbNodos`, `TbNodoBUI` |
| `Form_FormEventoAlta/Edicion/Gestion`, `Form_FormEventosAltaMasivaPorNS` | `TbNodoBUI`, `TbSubsistemaBui`, `TbTipoEvento`, `TbCriticidad`, `TbOriginador`, `TbEventos` |
| `Form_FormGastoAlta/Edicion/Gestion` | `TbGastos`, `TbTecnicos` |
| `Form_FormMaterial*` | `TbMaterial`, `TbTipoAccion`, `TbMaterialSeguimiento` |
| `Form_FormNodoBUIGestion`, `Form_FormNodoGestion`, `Form_FormOriginadorGestion`, `Form_FormUbicacionGestion` | `TbNodoBUI`, `TbNodos`, `TbOriginador`, `TbUbicacion` |
| `Form_FormPlanificacion*` | `TbEventos`, `TbPlanificacion`, `TbPlanificacionAñadida`, `TbEquipos`, `TbBUI`, `TbIDEquipos` |
| `Form_FormSubContratacion*`, `Form_FormSubSistema*` | `TbSubcontrataciones` |

Varios `RowSource` son listas literales (`Form_FormTecnicoGestion.form.txt:514`: `ALIAS;NOMBRE;TIPO;…`) y no dependen de una tabla.

## Lo que no se pudo verificar

1. Si las 7 tablas sin uso en el front las consume otro aplicativo: exige el grafo de vínculos de todos los backends (tarea T6).
2. Estado de UAT en el tracker: la evidencia es el roadmap del repo, no la issue de UAT.
3. Diferencias `src/` línea a línea entre RC1 y `main`: se comprobó paridad en tablas, `RowSource` y `ControlSource`, no un diff completo de los módulos.
4. El trabajo sin commitear de `feat/rc1-invoice-sla-gate` no se comparó módulo a módulo.
