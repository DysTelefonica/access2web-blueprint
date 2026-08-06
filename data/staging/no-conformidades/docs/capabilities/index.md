<!--
ÍNDICE DE CAPACIDADES — registro maestro para navegación.
Vive en docs/capabilities/index.md. Una fila por capacidad.
Una IA lee ESTO primero para encontrar una capacidad y luego abre su documento.
Mantenlo sincronizado siempre que se añada, retire o cambie de tier/estado/confianza una capacidad.
Idioma: castellano de España. Los enums (tier/status/source/confianza) se mantienen tal cual.
-->

# Índice de capacidades

> Recordatorio de fuente de verdad: código + tests de Dysflow. Este índice es un mapa, no evidencia.

## Tabla principal

| ID capacidad | Nombre | Dominio | Tier | Estado | Source | Confianza global | ¿Pruebas en verde? | Última release de producción | Documento |
|---|---|---|---|---|---|---|---|---|---|
| CAP-NCP-LC | Ciclo de vida de NC Proyecto | NC-Proyecto-Lifecycle | critical | active | hybrid | mixed | 5/9 | Pendiente | [enlace](./nc-proyecto-lifecycle.md) |
| CAP-NCA-LC | Ciclo de vida de NC Auditoría | NC-Auditoria-Lifecycle | critical | active | hybrid | mixed | 5/7 | Pendiente | [enlace](./nc-auditoria-lifecycle.md) |
| CAP-NCP-AF | Acciones y seguimiento de NC Proyecto | NC-Proyecto-Acciones-Seguimiento | critical | active | hybrid | mixed | 2/7 | Pendiente | [enlace](./nc-proyecto-actions-follow-up.md) |
| CAP-NCA-AF | Acciones y seguimiento de NC Auditoría | NC-Auditoria-Acciones-Seguimiento | critical | active | reverse-engineered | mixed | 3/5 | Pendiente | [enlace](./nc-auditoria-actions-follow-up.md) |
| CAP-CE | Flujo de control de eficacia | Control-Eficacia | critical | active | hybrid | partial | 4/6 | Pendiente | [enlace](./control-eficacia-workflow.md) |
| CAP-IND | Cuadro de mando de indicadores | Indicadores | critical | active | hybrid | mixed | 7/8 | Pendiente | [enlace](./indicators-dashboard.md) |
| CAP-CAT | Maestros y catálogos | Maestros-Catalogos | standard | active | hybrid | mixta | 0/7 | Pendiente | [enlace](./master-data-catalogues.md) |
| CAP-DGE | Documentos y evidencia generada | Documentos-Evidencia | standard | active | hybrid | low-to-mixed | 0/6 | Pendiente | [enlace](./documents-generated-evidence.md) |
| CAP-EXP | Expedientes, riesgos y responsables | Expedientes-Riesgos | standard | active | reverse-engineered | mixta | 0/7 | Pendiente | [enlace](./expedientes-riesgos-responsables.md) |
| CAP-CFG | Configuración, backends y runtime local | Configuracion-Backends | critical | active | hybrid | mixta | 0/6 | Pendiente | [enlace](./configuration-backends-runtime.md) |
| CAP-UPN | Usuarios, permisos y navegación | Usuarios-Permisos | critical | active | reverse-engineered | mixta | 0/7 | Pendiente | [enlace](./users-permissions-navigation.md) |
| CAP-XCUT | Soporte transversal | Soporte-Cross-Cutting | standard | active | hybrid | mixed | 0/7 | Pendiente | [enlace](./cross-cutting-support.md) |
| CAP-COM | Comunicaciones, informes y exportaciones | Comunicacion-Informes-Exportaciones | standard | active | hybrid | mixta | 0/8 | Pendiente | [enlace](./communications-reports-exports.md) |
| CAP-REL | Trazabilidad UAT, release y rollback | Release-UAT-Rollback | standard | active | sdd | Intended | 0/6 | Pendiente | [enlace](./release-uat-rollback-traceability.md) |

> Notas sobre la tabla principal:
> - Los `CAP-ID` siguen la convención corta `CAP-NNN` usada en `openspec/changes/issue-67-feature-tdd-coverage/apply-progress.md` §2. Cada doc de capacidad define además un ID largo propio (por ejemplo `CAP-NCP-LIFECYCLE`); ambos referencian la misma capacidad.
> - "¿Pruebas en verde?" cuenta reglas con `Verified-runtime` (incluyendo `Verified-runtime focused`) sobre el total de reglas de la tabla §2 del doc. Los `Verified-static` figuran como deuda; ver sección de lagunas.
> - "Última release de producción" se mantiene en `Pendiente` para todas las capacidades: la épica `issue-67-feature-tdd-coverage` (Fase 0) aún no tiene tag UAT aprobado ni release de producción; la fila de release/UAT se completará en Fase 3.

## Lagunas de cobertura (obligaciones abiertas)

> Reglas que aún no están en `Verified-runtime`. Cada una es un test que crear con `access-vba-tdd`.
> La confianza actual se extrae de §2 y §7 de cada doc de capacidad; la acción por defecto es `Crear test con access-vba-tdd` salvo que el doc indique otra cosa.

| Capacidad | Regla | Confianza actual | Acción | Tracker GH |
|---|---|---|---|---|
| CAP-CE | BR-CE-5 | Intended | Crear test con access-vba-tdd | #71 (BR-CE-5/6) |
| CAP-CE | BR-CE-6 | Intended | Crear test con access-vba-tdd | #71 (BR-CE-5/6) |
| CAP-IND | BR-IND-8 | Intended | Crear test con access-vba-tdd | #73 |
| CAP-REL | BR-REL-1 | Verified-static | Crear check documental automatizable (no test VBA runtime) | n/a (sin issue abierta) |
| CAP-REL | BR-REL-2 | Verified-static | Crear check documental automatizable (no test VBA runtime) | n/a (sin issue abierta) |
| CAP-REL | BR-REL-3 | Verified-static | Crear check documental automatizable (no test VBA runtime) | n/a (sin issue abierta) |
| CAP-REL | BR-REL-4 | Verified-static | Crear check documental automatizable (no test VBA runtime) | n/a (sin issue abierta) |
| CAP-REL | BR-REL-5 | Intended | Crear check documental automatizable (no test VBA runtime) | n/a (sin issue abierta) |
| CAP-UPN | BR-UPN-1 | Verified-static | Crear test con access-vba-tdd | n/a (sin issue abierta) |
| CAP-UPN | BR-UPN-2 | Verified-static | Crear test con access-vba-tdd | n/a (sin issue abierta) |
| CAP-UPN | BR-UPN-3 | Verified-static | Crear test con access-vba-tdd | n/a (sin issue abierta) |
| CAP-UPN | BR-UPN-4 | Verified-static | Crear test con access-vba-tdd | n/a (sin issue abierta) |
| CAP-UPN | BR-UPN-5 | Verified-static | Crear test con access-vba-tdd | n/a (sin issue abierta) |
| CAP-UPN | BR-UPN-6 | Verified-static | Crear test con access-vba-tdd | n/a (sin issue abierta) |
| CAP-UPN | BR-UPN-7 | Intended | Crear test con access-vba-tdd (post-decisión producto) | #72 |
| CAP-XCUT | BR-XCUT-6 | Intended | Crear test con access-vba-tdd | #82 |
| CAP-DGE | BR-DOC-1 | Intended | Crear test con access-vba-tdd (post-decisión producto) | #80 (mapea a BR-DOC-1/2, ver nota) |
| CAP-DGE | BR-DOC-2 | Intended | Crear test con access-vba-tdd (post-decisión producto) | #80 (mapea a BR-DOC-1/2, ver nota) |
| CAP-EXP | BR-EXP-6 | Verified-static | Crear test con access-vba-tdd | n/a (sin issue abierta) |
| CAP-EXP | BR-EXP-7 | Intended | Crear test con access-vba-tdd (post-decisión producto) | #81 |
| CAP-NCA-AF | BR-NCA-AF-4 | Intended | Crear test con access-vba-tdd (post-decisión producto) | #74 |
| CAP-NCA-AF | BR-NCA-AF-5 | Intended | Crear test con access-vba-tdd (post-decisión producto) | #74 |
| CAP-COM | BR-COM-3 | Intended | Crear test con access-vba-tdd (post-decisión producto) | #75 |
| CAP-COM | BR-COM-5 | Intended | Crear test con access-vba-tdd (post-decisión producto) | #76 |
| CAP-COM | BR-COM-6 | Intended | Crear test con access-vba-tdd (post-decisión producto) | #77 |
| CAP-COM | BR-COM-7 | Intended | Crear test con access-vba-tdd (post-decisión producto) | #78 |
| CAP-CAT | BR-CAT-6 | Intended | Crear test con access-vba-tdd (post-decisión producto) | #79 |
| CAP-CAT | BR-CAT-7 | Intended | Crear test con access-vba-tdd (post-decisión producto) | #79 |

> Notas sobre las lagunas:
> - En el doc de `users-permissions-navigation` solo existen las reglas `BR-UPN-1`..`BR-UPN-7`. La lista de huecos citaba `BR-UPN-8` pero esa regla no aparece en el doc; se omite para no inventar contenido.
> - En el doc de `documents-generated-evidence` el prefijo de las reglas es `BR-DOC-` (no `BR-DGE-` como aparece en la lista de huecos y como titula la issue #80). Se mantiene el prefijo del doc para no corromper las referencias cruzadas; el mapeo entre `BR-DGE-*` (issue) y `BR-DOC-*` (doc) queda como discrepancia nominal a resolver cuando el producto apruebe la regla.

## Reglas con decisión de producto pendiente (issues abiertas GH)

> Issues GH abiertas que bloquean `Verified-runtime` por necesitar decisión de producto (no de ingeniería). Cada una está mapeada a su BR-* en la tabla anterior y al documento de capacidad correspondiente. Cuando producto aprueba la regla: (1) se actualiza la autoridad en §2 del doc de capacidad, (2) se reemplaza la confianza por `Intended` confirmada o `Verified-static` con test, (3) se cierra la issue con `gh issue close <n> --comment` siguiendo la regla `gentle-ai:issue-closure-traceability` (commit SHA + test reference).

| Issue GH | Título | BR-* | Doc de capacidad | Decisión que necesita producto |
|---|---|---|---|---|
| #71 | BR-CE-5/6: matriz transiciones CE + EficaciaOK=No | BR-CE-5, BR-CE-6 | [control-eficacia-workflow.md](./control-eficacia-workflow.md) | Matriz completa de transiciones (aprobado/fallido/no requerido/replanificado/evidencia) + comportamiento de `DatosGeneralesOK(p_MenosCef)` en `Form_FormNCAuditoriaGeneral.ComandoControlEficaciaDatos_Click`. |
| #72 | BR-UPN-7: matriz permisos perfil × acción × dominio | BR-UPN-7 | [users-permissions-navigation.md](./users-permissions-navigation.md) | Matriz de permisos para acciones sensibles (cerrar, eliminar, rehabilitar, documento, acción, informe, configuración) por perfil × dominio (Proyecto/Auditoría). |
| #73 | BR-IND-8: aprobar buckets cuadro de mando | BR-IND-8 | [indicators-dashboard.md](./indicators-dashboard.md) | Aprobación de los buckets del cuadro de mando (qué métricas caen en cada categoría). |
| #74 | BR-NCA-AF-4/5: contrato acciones seguimiento NC auditoría | BR-NCA-AF-4, BR-NCA-AF-5 | [nc-auditoria-actions-follow-up.md](./nc-auditoria-actions-follow-up.md) | Contrato de acciones de seguimiento de NCs de auditoría (cuáles son obligatorias, qué datos requieren, quién las puede cerrar). |
| #75 | BR-COM-3: contrato BCC notificaciones auditoría | BR-COM-3 | [communications-reports-exports.md](./communications-reports-exports.md) | Contrato de BCC en notificaciones de auditoría (qué destinatarios, en qué eventos). |
| #76 | BR-COM-5: regla operativa BCC default | BR-COM-5 | [communications-reports-exports.md](./communications-reports-exports.md) | Regla operativa: ¿BCC por defecto a un buzón interno de auditoría en todos los emails? |
| #77 | BR-COM-6: formato plantillas email | BR-COM-6 | [communications-reports-exports.md](./communications-reports-exports.md) | Formato de plantillas de email (firma, asunto, cuerpo) — plantillas canónicas a usar. |
| #78 | BR-COM-7: contrato export Excel | BR-COM-7 | [communications-reports-exports.md](./communications-reports-exports.md) | Contrato de exportación a Excel en `ComandoExportarAExcel` (columnas, formato, filtros aplicados). |
| #79 | BR-CAT-6/7: catálogos maestros activos | BR-CAT-6, BR-CAT-7 | [master-data-catalogues.md](./master-data-catalogues.md) | Confirmar catálogos maestros activos y campos obligatorios (tipología NC, motivos no requiere CE, etc.). |
| #80 | BR-DGE-1/2: obligatoriedad evidencia al cierre | BR-DOC-1, BR-DOC-2 (mapeo nominal) | [documents-generated-evidence.md](./documents-generated-evidence.md) | Obligatoriedad de evidencia documental al cierre de NC (qué tipo, cuántas, en qué casos). |
| #81 | BR-EXP-7: ciclo de vida canónico de riesgos | BR-EXP-7 | [expedientes-riesgos-responsables.md](./expedientes-riesgos-responsables.md) | Ciclo de vida canónico de riesgos (estados, transiciones, eventos, responsables). |
| #82 | BR-XCUT-6: matriz permisos cross-link | BR-XCUT-6 | [cross-cutting-support.md](./cross-cutting-support.md) | Matriz de permisos con cross-link a BR-UPN-7 D1 (cómo interactúan ambos modelos). |

> **Acción recomendada**: agendar sesión con la autoridad de producto (`Confirmación pendiente` en cada doc de capacidad) para revisar las 12 issues. El cierre debe hacerse con `gh issue close <n> --comment "..."` con la regla `issue-closure-traceability` (commit SHA del fix + test reference).

## Divergencias pendientes de revisión humana

| Capacidad | Hallazgo | Detectada |
|---|---|---|
| CAP-UPN | BR-UPN-7 — la matriz completa de permisos por acción sensible (cerrar, eliminar, rehabilitar, documento, acción, informe, configuración) está aprobada por producto como `Intended`, pero las reglas de autorización siguen embebidas en formularios (`Form_Form0BDOpciones`, `Form_Form0BDTecnicos`, `Form_FormNCAuditoriaGestion`) sin trazabilidad de producto. El producto debe aprobar la matriz; hasta entonces no se puede ascender de `Intended`. | 2026-06-15 |
| CAP-CE | BR-CE-5 y BR-CE-6 — el botón general de auditoría (`Form_FormNCAuditoriaGeneral.ComandoControlEficaciaDatos_Click`) y el flujo completo de resultados de eficacia (motivo no requerido, eficacia fallida, replanificación, evidencia) están documentados como `Intended`. El comportamiento diferido del botón con `DatosGeneralesOK(p_MenosCef)` queda abierto: el spec no coincide con código probado, hace falta prueba de costura helper/servicio y validación de producto sobre el bypass previsto. | 2026-06-15 |
