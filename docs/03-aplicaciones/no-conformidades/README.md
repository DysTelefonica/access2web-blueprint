# 03 · No Conformidades

## Propósito

Evidencia de descubrimiento de **NoConformidades**, aplicación Access/VBA que mantiene el ciclo de vida de No Conformidades (NC) sobre dos ejes: **Auditorías** (AC, AR, Control de Eficacia) y **Proyectos** (AC, AR, Control de Eficacia), con replanificaciones, vinculación a Riesgos y vista agregada de indicadores. Este lote incluye también el **diagnóstico del fallo de inventario Dysflow** (P3) y un **hallazgo de seguridad** (credenciales en `backends.json`).

## Estado

- **Fase:** descubrimiento completo, Batch 4 + inventario real Dysflow.
- **Fecha de evidencia:** 2026-08-05.
- **Repositorio:** `C:\00repos\codigo\00_NO_CONFORMIDADES\staging` (seleccionado vía Dysflow tras `FRONTEND_TARGET_AMBIGUOUS`). CodeGraph-VBA inspeccionado sobre `00_main`; inventario backend real sobre staging.
- **Otros worktrees:** `00_main` (release), `hotfix-replanificadas`, `slice10` comparten `projectId: "00-no-conformidades-staging-clean"` (HR-11 ambigüedad registrada).
- **Frontend:** `NoConformidades.accdb` en `staging` (43 MB; el más grande de las 8 aplicaciones).
- **Backend autoritativo:** `C:\00repos\datos\NoConformidades_Datos.accdb` (27 MB).
- **Dysflow:** `.dysflow/project.json` migrado (T18 caps-block) y `accessPath` absoluto explícito para invocaciones.
- **Inventario Dysflow real**:
  - **42 tablas** en el backend (ver [data-model.md § Inventario real](data-model.md#inventario-real-delflow-2026-08-05)).
  - **438 NCs de Proyecto** (`TbNoConformidades`) + **55 NCs de Auditoría** (`TbNoConformidadesAuditoria`) = **493 NCs totales**.
  - **14 FK relationships** entre tablas de usuario.
- **Riesgo de seguridad detectado:** `backends.json` contiene `ACCESS_VBA_PASSWORD` en claro (ver [Seguridad § D90](security-rules.md#d90-riesgo-de-seguridad--backendsjson-con-contraseña-en-claro)). Sigue siendo válido como hallazgo.

## Lote asociado

Lote 5 (posterior a Lanzadera Lote 1, Expedientes Lote 2 y Gestion_Riesgos Lote 3). Incluido por P3 (autorización Dysflow read-only + diagnóstico del fallo de inventario).

## Entregables

1. [Capacidades](capabilities.md)
2. [Formularios y call paths](forms.md)
3. [Modelo físico y diccionario](data-model.md)
4. [Matriz de migración](migration-matrix.md)
5. [Integraciones y automatización](integrations-automation.md)
6. [Seguridad y reglas](security-rules.md)

## Fuentes de autoridad

1. `C:\00repos\codigo\00_NO_CONFORMIDADES\00_main\src` (clases, forms, módulos) — codegraph-vba.
2. `C:\00repos\documentacion\OPENSPEC\00_No_Conformidades` (documentación previa al reverse engineering).
3. `.dysflow/project.json` y `backends.json` en `00_main` (revisados para diagnóstico).
4. Dysflow `list_objects` **NO ejercitado** en esta pasada (falla conocida, pendiente de resolución).
5. Engram como contexto histórico.

## Reglas de evidencia

- No se han realizado imports, exports, sync, tests, compile, cleanup ni escrituras.
- Se excluyen valores personales, correos, credenciales, hashes, hosts y nombres de máquina.
- **Las rutas UNC y hosts internos NO se reproducen en estos artefactos** (regla transversal del blueprint).
- El valor de la contraseña en `backends.json` **NO se reproduce en este artefacto** (riesgo de seguridad documentado en D90).
- Todo campo/registro del backend queda como `preservar hasta decisión`.
- APAP y APAP_WEB **no aparecen** ni se mencionan (proyecto personal del desarrollador; regla transversal del blueprint).

## Hallazgos críticos del lote

1. **344 callers de `getdb()`** — el más alto de las 8 aplicaciones. NoConformidades es intensísima en DAO.
2. **D89 INVALIDADO**: el "fallo de inventario Dysflow" fue una diagnosis errónea sin contactar el runtime. Ver [Seguridad § D89](security-rules.md#d89--diagnóstico-del-fallo-de-list_objects-de-dysflow--invalidado). El inventario Dysflow funciona perfectamente cuando se invoca con `accessPath` absoluto explícito.
3. **Riesgo de seguridad en `backends.json`** — ver D90. La contraseña está en claro dentro del repo versionado. Sigue siendo válido como hallazgo.
4. **Caché selectivo maduro** — ver D91. NoConformidades ya implementa un patrón completo de caché con kill switch, diagnóstico, métricas y logs. Es la referencia más rica para el puerto de caché del blueprint.
5. **`NCProyectoListItemVM`** — clase ViewModel separada de la entidad `NCProyecto`. Es el patrón canónico de "lista para UI" que se traduce directamente al concepto de DTO de respuesta en la nueva plataforma.
6. **Formularios masivos** — ~60 forms (cada uno con `.cls` + `.form.txt`). Es la aplicación con más superficie UI de las 8.
7. **Inventario real Dysflow** — 42 tablas, 493 NCs (438 proyecto + 55 auditoría), 14 FKs, 44 columnas en `TbNoConformidades`.

## Checklist

- [x] Inventario funcional, formularios, clases y módulos documentados vía codegraph-vba.
- [x] Inventario real Dysflow del backend (42 tablas, 493 NCs, 14 FKs).
- [x] D89 invalidado (el inventario Dysflow funciona con la invocación correcta).
- [x] Riesgo de seguridad detectado y documentado (D90).
- [x] Patrón de caché preservado como referencia (D91).
- [x] Perfilado agregado ejecutado sin copiar filas personales.
- [x] APAP y APAP_WEB no aparecen en esta evidencia.
- [ ] `backends.json` saneado (D90 pendiente de remediación operativa).
- [ ] Épica + tickets + matriz de migración de datos para NoConformidades.

## Siguiente paso

Generar la **épica + tickets accionables + matriz de migración de datos** para NoConformidades con el inventario real obtenido. Continuar con el codegraph de las 4 apps pendientes (HPS, Condor, Brass, HPS_Solicitudes) y luego sus inventarios Dysflow. Por último, épicas por aplicación según el orden acordado (primero estudio de los 8, luego épica por épica).