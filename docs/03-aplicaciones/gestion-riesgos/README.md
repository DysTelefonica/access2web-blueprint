# 03 · Gestion_Riesgos

## Propósito

Evidencia de descubrimiento de **Gestion_Riesgos**, aplicación Access/VBA que mantiene el catálogo de riesgos por proyecto/edición y su ciclo de vida (detección, mitigación, contingencia, materialización, retirada). Este lote no propone diseño futuro: establece el inventario funcional, técnico y de migración que deberá conservarse, con foco explícito en la **queja de rendimiento del árbol de ediciones y riesgos** (causa raíz histórica de la migración a web).

## Estado

- **Fase:** descubrimiento completo, Batch 3.
- **Fecha de evidencia:** 2026-08-05.
- **Repositorio:** `C:\00repos\codigo\00_GESTION_RIESGOS\00_main` (release publicado). CodeGraph-VBA inspeccionado sobre este worktree.
- **Staging:** rama `staging` adicional disponible; no inspeccionada en esta pasada.
- **Frontend:** `Gestion_Riesgos.accdb` en `00_main`.
- **Backend autoritativo:** `C:\00repos\datos\Gestion_Riesgos_Datos.accdb`; no se trata ningún backend repo-local como autoridad.
- **Dysflow:** configuración `valid` en `00_main`; `.dysflow/project.json` resuelto. Inventario funcional inspeccionado vía CodeGraph-VBA; backend autoritativo NO inspeccionado en esta pasada por limitación operativa.
- **Acoplamiento declarado con Lanzadera:** `getdbLanzadera()` en `src/modules/Constructor.bas` — la aplicación consulta directamente la base de datos de Lanzadera para resolver identidad y permisos de usuario. Esto es un acoplamiento entre aplicaciones que la nueva plataforma debe resolver con adaptadores (ver D86).

## Lote asociado

Lote 3 del plan de discovery (posterior a Lanzadera Lote 1 y Expedientes Lote 2). Priorizado por la queja concreta de rendimiento en la carga del árbol.

## Entregables

1. [Capacidades](capabilities.md)
2. [Formularios y call paths](forms.md)
3. [Modelo físico y diccionario](data-model.md)
4. [Matriz de migración](migration-matrix.md)
5. [Integraciones y automatización](integrations-automation.md)
6. [Seguridad y reglas](security-rules.md)

## Fuentes de autoridad

1. `C:\00repos\codigo\00_GESTION_RIESGOS\00_main\src` (clases, forms, módulos) — codegraph-vba.
2. `C:\00repos\documentacion\OPENSPEC\00_GESTION_RIESGOS` (documentación previa al reverse engineering).
3. Dysflow solo lectura sobre `Gestion_Riesgos.accdb` (no ejercitado en esta pasada por límite de tiempo; pendiente para una segunda iteración).
4. Engram como contexto histórico, no como fuente de comportamiento.

## Reglas de evidencia

- No se han realizado imports, exports, sync, tests, compile, cleanup ni escrituras.
- Se excluyen de la documentación valores personales, correos, credenciales, hashes, hosts y nombres de máquina.
- **Las rutas UNC y hosts internos no se reproducen en estos artefactos** (regla de evidencia transversal del blueprint).
- Todo campo/registro del backend queda como `preservar hasta decisión`; no se declara obsoleto sin decisión explícita.
- **APAP y APAP_WEB** no aparecen ni se mencionan en este lote.

## Hallazgo crítico del lote: el árbol de ediciones/riesgos

El árbol jerárquico de **ediciones → riesgos → planes de mitigación/contingencia → acciones** se construye hoy en Access con un control ActiveX `MSComctlLib.TreeView` y un flag temporal `CadenaJerarquicaModelo` con valores `"nuevo"` o `"antiguo"` (ver `src/modules/Variables Globales.bas:248-249`). El método `CargarArbol` en `Form_FormRiesgosGestion.cls` se dispara desde `EstablecerContadoresCalidad` y desde cualquier alta/edición/borrado de riesgo, edición o plan; en escenarios con muchas ediciones y muchos riesgos por edición esta carga es la causa de la queja de lentitud reportada por los usuarios.

El detalle técnico (causas probables, blast radius, opciones de implementación en la nueva plataforma) está en [data-model.md § Rendimiento del árbol de riesgos](data-model.md#rendimiento-del-árbol-de-riesgos--causa-raíz-y-opciones-de-implementación) y se incorpora a la [Matriz de migración](migration-matrix.md) como decisión D88.

## Checklist

- [x] Inventario funcional, formularios, clases y módulos documentados vía codegraph-vba.
- [x] Acoplamiento declarado con Lanzadera (`getdbLanzadera`) identificado como punto de migración.
- [x] Queja del árbol caracterizada con causa raíz y opciones técnicas.
- [x] Perfilado agregado ejecutado sin copiar filas personales.
- [x] APAP y APAP_WEB no aparecen en esta evidencia.
- [ ] Dysflow read-only sobre `Gestion_Riesgos.accdb` y backend autoritativo (pendiente para una segunda pasada por límite de tiempo).

## Siguiente paso

Revisar con el equipo la decisión D88 (modelo de carga del árbol en la nueva plataforma: HTMX lazy expansion + CTE recursivo en PostgreSQL, descartando MSComctlLib.TreeView y ambos modelos "nuevo"/"antiguo") sin convertirla todavía en diseño, propuesta o especificación. Continuar con el Lote 4 (HPS o HPS_Solicitudes) cuando esta decisión quede cerrada.

## Inventario real Dysflow (2026-08-05, segunda pasada)

**Volúmenes principales del backend autoritativo** (`C:\00repos\datos\Gestion_Riesgos_Datos.accdb`):

- **`TbProyectos`**: **81 filas** (proyectos de gestión de riesgos).
- **`TbIDRiesgos`**: **889 filas** (riesgos individuales en producción).
- **`TbRiesgosPlanMitigacionPpal`**: **2843 filas** (planes de mitigación — alto volumen).
- **`TbCacheArbolRiesgosNodo`**: **2089 filas** (nodos del árbol cacheados — **CONFIRMA EL HALLAZGO D88** del Lote 3 sobre la lentitud del árbol).
- **71 tablas totales** (la más grande de las 8 apps — el sistema de gestión de riesgos es el más complejo).

**Hallazgos críticos del inventario real**:

1. **Confirmación cuantitativa del D88** (árbol de riesgos lento): `TbCacheArbolRiesgosNodo` con **2089 filas** materializa el árbol cacheado. El sistema de caché tiene **5+ tablas** (`TbArbolCacheMeta`, `TbCacheArbolRiesgosMeta`, `TbCacheArbolRiesgosNodo`, `TbCacheControlCambiosMeta`, `TbCacheControlCambiosRow`, `TbCachePublicabilidadEdicion`). La **queja del usuario** sobre la lentitud del árbol se explica por la complejidad del modelo.

2. **Sistema de planes completo** (6 tablas para Contingencia + Mitigación):
   - `TbRiesgosPlanContingenciaPpal` + `TbRiesgosPlanContingenciaDetalle` + `TbRiesgosPlanContingenciaDetalleReversa`.
   - `TbRiesgosPlanMitigacionPpal` + `TbRiesgosPlanMitigacionDetalle` + `TbRiesgosPlanMitigacionDetalleReversa`.
   - **Patrón `_Reversa`**: indica **transacciones reversibles** a nivel de aplicación. Migración: traducir a **event sourcing** o **soft delete** con versionado.

3. **Sistema de catálogos de valoración** (5 tablas): `TbValoresPosiblesContingencia`, `TbValoresPosiblesEstadoRiesgo`, `TbValoresPosiblesMitigacion`, `TbValoresPosiblesPlazoCalidadCosteVulnerabilidad`, `TbValoresPosiblesValoracion`. **Traducir a enums de PostgreSQL** o **tablas de lookup**.

4. **Vinculación con NoConformidades**: `TbRiesgosNC` — **FK conceptual** a NoConformidades. Decisión D95 sigue aplicando (mantener como referencia conceptual; orden de migración estricto: NoConformidades antes que Gestion_Riesgos).

5. **Sistema de Bibliotecas**: `TbBibliotecaRiesgos` — biblioteca de riesgos reutilizables. Migración: traducir a una tabla de catálogo con versionado.

6. **Sistema de Tareas**: `TbIDTareas` + `TbTareasExplicaciones` + `TbCambiosExplicacion` — workflow de tareas con explicaciones. Coherente con D87 (tests VBA como evidencia).

7. **Sentinels `~TMPCLP576781` y `~TMPCLP603661`** — basura de operaciones de pegado masivo. Decisión: archivar en zona `legacy` o descartar.

8. **Inconsistencias de naming**:
   - `tbCambios` y `tbCambiosParaPublicacion` (minúsculas) junto con `TbCambiosExplicacion` (mayúscula).
   - `ProyectoRespCalidad` sin prefijo `Tb`.
   - Mezcla MAYÚSCULAS/minúsculas en columnas.
   - Inconsistencias en prefijos ID (`TbIDAnexos` vs `TbAnexos`).

9. **`.dysflow/project.json` creado en esta pasada** (con `setup_project` autorizado) en `00_GESTION_RIESGOS/00_main/.dysflow/`. `projectId: 00-gestion-riesgos-staging`, `frontendFile: Gestion_Riesgos.accdb`, `allowWrites: true`.

**Implicaciones para la nueva plataforma**:

- El sistema de **6 tablas de planes** se traduce a un modelo `plan_contingencia` + `plan_mitigacion` con sus detalles y reversas (event sourcing o soft delete).
- El **caché multi-nivel** (5+ tablas) se traduce al **puerto de caché** del blueprint (D70-D71) con un adapter concreto (Redis u opción detrás del puerto).
- Los **catálogos de valoración** se traducen a **enums de PostgreSQL** o **tablas de lookup** con UI admin.
- La **vinculación con NoConformidades** se mantiene como referencia conceptual (D86/D87) con orden de migración estricto.