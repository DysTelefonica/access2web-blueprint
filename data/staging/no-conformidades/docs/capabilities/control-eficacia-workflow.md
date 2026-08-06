# Capacidad: flujo de control eficacia

## Â§0 Identidad
- **ID de capacidad**: `CAP-CONTROL-EFICACIA`
- **Nivel**: critical
- **Estado**: active / documentaciÃ³n alineada con v2; BR-CE-1..4 tienen verde runtime fresco 2026-06-15; flujo completo de resultados pendiente
- **Fuente**: hybrid (documento de funcionalidad de cumplimiento + inventario de fuente + confirmaciÃ³n de negocio pendiente)
- **Responsable / autoridad de producto**: ConfirmaciÃ³n pendiente â€” Cumplimiento / Calidad
- **Ãšltima verificaciÃ³n**: 2026-06-15 actualizaciÃ³n solo documental sobre un run runtime fresco de Dysflow del mismo dÃ­a; esta tarea no ejecutÃ³ Dysflow/Access, solo reescribiÃ³ el estado documental a partir de evidencia fresca aportada
- **Confianza global**: partial â€” BR-CE-1..4 con `Verified-runtime` (13/13 PASS, 7 filtros verdes 2026-06-15); BR-CE-5/6 y reglas de aprobado/no aprobado/no requerido/replanificaciÃ³n/evidencia siguen `Intended` / `FALTA â†’ crear mediante access-vba-tdd`

## Â§1 IntenciÃ³n de negocio
- **PropÃ³sito**: Controlar si el trabajo corrector/de resoluciÃ³n fue eficaz antes del cierre final de NC sin bloquear creaciÃ³n/ediciÃ³n temprana.
- **Usuarios / personas**: Equipo de calidad, responsables de proyecto/auditorÃ­a, revisores UAT, desarrolladores/agentes IA.
- **Problema que resuelve**: Evita cierres prematuros o no conformes mientras permite datos incompletos en etapas tempranas durante `Alta`/`Edicion`.
- **Valor de negocio / por quÃ© existe**: El control de eficacia es un gate de cierre/cumplimiento y debe ser auditable.
- **No objetivos**: Esta pÃ¡gina no define almacenamiento documental completo ni comportamiento del cuadro de mando.
- **Fuente de intenciÃ³n**: Documento existente de funcionalidad de cumplimiento y nombres de fuente; resultados de negocio completos pendientes de confirmaciÃ³n.
- **Referencia tracker de origen**: Issue #45 / issue-19.

## Â§2 Contrato de comportamiento

### Escenarios (Given / When / Then)
- **GIVEN** un usuario crea o edita NC de Proyecto/AuditorÃ­a **WHEN** `FechaPrevistaControlEficacia` estÃ¡ vacÃ­a **THEN** guardar/abrir no debe bloquearse solo por esa fecha ausente.
- **GIVEN** una NC abierta requiere control de eficacia **WHEN** el usuario la cierra sin datos FE obligatorios **THEN** el cierre debe bloquearse.
- **GIVEN** escenarios de bypass FE para alta/edicion/auditoria **WHEN** se ejecutan **THEN** se preserva la invariancia de `EficaciaOK`.
- **GIVEN** se registra un resultado de eficacia no requerido, fallido o replanificado **WHEN** se guarda **THEN** estado, evidencia e impacto en cierre deben seguir reglas de negocio confirmadas.

### Reglas de negocio
| ID de regla | Enunciado (previsto) | Autoridad | Â¿Aplicada en cÃ³digo? | Prueba (evidencia) | Confianza |
|---|---|---|---|---|---|
| BR-CE-1 | `FechaPrevistaControlEficacia` no debe bloquear `Alta` ni `Edicion`. | Documento de funcionalidad de cumplimiento Issue #19 | SÃ­ segÃºn docs en mÃ³dulos de operaciones | Runtime fresco Dysflow 2026-06-15: `tests/tests.vba.json` filtros `Issue19_CE_Alta_` 5/5 (`Test_Issue19_CE_Alta_Si_SinDetalle_NoBloquea`, `Test_Issue19_CE_Alta_Si_ConDetalle_Pasa`, `Test_Issue19_CE_Alta_No_IgnoraDetalle`, `Test_Issue19_CE_Alta_MotivoAlta_Bypass_Si`, `Test_Issue19_CE_Alta_MotivoAlta_Bypass_BlankRequereCE`). Ancla histÃ³rica `8cb7f0a` | Verified-runtime |
| BR-CE-2 | El gate de `FechaPrevistaControlEficacia` se aplica al cierre de NC. | Documento de funcionalidad de cumplimiento Issue #19 | SÃ­ segÃºn docs | Runtime fresco Dysflow 2026-06-15: `tests/tests.vba.json` filtro `Issue19_CE_Cierre_` 2/2 (`Test_Issue19_CE_Cierre_SinDetalle_Bloquea`, `Test_Issue19_CE_Cierre_ConDetalle_PermiteCierre`). Ancla histÃ³rica `8cb7f0a` | Verified-runtime |
| BR-CE-3 | Los escenarios de bypass para alta, edicion y auditoria siguen soportados. | Documento de funcionalidad de cumplimiento Issue #19 | SÃ­ segÃºn docs | Runtime fresco Dysflow 2026-06-15: `tests/tests.vba.json` filtros `Issue19_CE_Edicion_` 1/1 (`Test_Issue19_CE_Edicion_MotivoDatosUnicos_Bypass`) y `Issue19_CE_Auditoria_` 1/1 (`Test_Issue19_CE_Auditoria_MotivoDatosUnicos_Bypass`). Ancla histÃ³rica `8cb7f0a` | Verified-runtime |
| BR-CE-4 | La invariancia de `EficaciaOK` se preserva al mover el gate al cierre. | Documento de funcionalidad de cumplimiento Issue #19 | SÃ­ segÃºn docs | Runtime fresco Dysflow 2026-06-15: `tests/tests.vba.json` filtros `Issue19_CE_EstadoCalculado_` 2/2 (`Test_Issue19_CE_EstadoCalculado_Pendiente`, `Test_Issue19_CE_EstadoCalculado_SinPendiente`) + `Issue19_CE_EficaciaOK_` 1/1 (`Test_Issue19_CE_EficaciaOK_SinCambios`) + `Issue19_Paridad_UI` 1/1 (`Test_Issue19_Paridad_UI_Dominio`). Ancla histÃ³rica `8cb7f0a` | Verified-runtime |
| BR-CE-5 `#71 (BR-CE-5/6)` | Reglas de motivo no requerido, eficacia fallida, replanificaciÃ³n y evidencia son explÃ­citas. | Autoridad de producto pendiente | Desconocido | FALTA â†’ crear mediante access-vba-tdd tras confirmar esquema/reglas | Intended |
| BR-CE-6 `#71 (BR-CE-5/6)` | El botÃ³n de control-eficacia general de auditorÃ­a respeta el bypass previsto (`DatosGeneralesOK(p_MenosCef)`). | DecisiÃ³n abierta de funcionalidad | Pendiente/diferido | FALTA â†’ crear mediante access-vba-tdd contra costura helper/servicio | Intended |

### Validaciones
- Fecha prevista FE no obligatoria en crear/editar.
- Fecha prevista FE obligatoria al cerrar cuando el control es requerido.
- La invariancia de `EficaciaOK` debe sobrevivir a rutas de bypass y cierre.
- Pendientes reglas de completitud de motivo no requerido y requisitos de resultado/evidencia.

### Transiciones de estado
- `New/editing NC without FE planned date` --(`Alta`/`Edicion`)--> `Saved NC`.
- `NC abierta` --(`Cerrar sin datos FE obligatorios`)--> `Cierre bloqueado`.
- `NC abierta con datos FE obligatorios` --(`Cerrar`)--> `NC cerrada`.
- `Eficacia pendiente` --(`Registrar resultado`)--> `Eficaz / no eficaz / no requerida / replanificada` â€” pendiente de confirmaciÃ³n.

### Caminos lÃ­mite y de error
- Las reglas de AuditorÃ­a y Proyecto pueden ser idÃ©nticas o intencionalmente distintas; no estÃ¡ confirmado.
- El comportamiento de botones de formulario debe probarse mediante costura helper/servicio antes de afirmar protecciÃ³n contra regresiÃ³n.

### SeÃ±ales de aceptaciÃ³n / presencia
- El bypass de crear/editar y la exigencia en cierre tienen pruebas enfocadas en verde.
- El flujo completo tiene pruebas para aprobado, fallido, motivo no requerido, replanificaciÃ³n y evidencia.

## Â§3 Mapa de implementaciÃ³n
- **Puntos de entrada UI**: `Form_FormNCProyectoControlEficacia`, `Form_FormNCProyectoControlEficaciaAlta`, `Form_FormNCAuditoriaControlEficacia`, `Form_FormNCAuditoriaControlEficaciaAlta`, `Form_FormMotivosNoRequiereControlEficacia`, `Form_FormNCAuditoriaGeneral.ComandoControlEficaciaDatos_Click`.
- **Puntos de entrada de fuente**: `NCProyectoOperaciones`, `NCaUDITORIAOperaciones`, `Test_Issue19_CEGating`; procedimientos exactos de cierre/UI pendientes de confirmaciÃ³n directa.
- **Datos tocados**: `FechaPrevistaControlEficacia`, `EficaciaOK`, registros NC Proyecto/AuditorÃ­a, catÃ¡logo de motivos no requeridos, evidencia/documentos.
- **Salidas**: validaciÃ³n de cierre, estado/resultado de eficacia, evidencia/documentos.
- **Dependencias e integraciones**: ciclo de vida de Proyecto, ciclo de vida de AuditorÃ­a, documentos/evidencia, permisos de cierre.
- **SincronizaciÃ³n fuenteâ†”binario**: no comprobada en esta tarea solo documental.
- **EvaluaciÃ³n de diseÃ±o (as-built vs ideal)**: mover la validaciÃ³n FE al cierre es la costura de negocio correcta; las reglas de resultados restantes deberÃ­an centralizarse en un servicio en lugar de validaciones dispersas de formulario.

## Â§4 Receta de reconstrucciÃ³n
1. Confirmar estados canÃ³nicos de eficacia, motivos no requeridos, permisos por rol, requisitos de evidencia y diferencias Proyecto vs AuditorÃ­a.
2. Mapear operaciones de cierre y eventos de formulario de control-eficacia a costuras helper/servicio.
3. Crear pruebas con esquema primero para bypass de crear/editar, bloqueo/paso de cierre, no requerido, eficacia fallida, replanificaciÃ³n y comportamiento del botÃ³n de auditorÃ­a.
4. Las pruebas deben usar filas sandbox controladas, `DAO.Database` explÃ­cito, retorno JSON y cardinalidad para cualquier mutaciÃ³n.
5. Cambios futuros de fuente: importaciÃ³n Dysflow â†’ compilaciÃ³n manual del usuario â†’ pruebas Dysflow.

## Â§5 Evidencia y trazabilidad
- **Pruebas (ancla runtime fresca)**: `tests/tests.vba.json` corrido con Dysflow 2026-06-15, 13 procedimientos Ãºnicos verdes repartidos en 7 filtros: `Issue19_CE_Alta_` 5/5, `Issue19_CE_Cierre_` 2/2, `Issue19_CE_EstadoCalculado_` 2/2, `Issue19_Paridad_UI` 1/1, `Issue19_CE_Edicion_` 1/1, `Issue19_CE_Auditoria_` 1/1, `Issue19_CE_EficaciaOK_` 1/1. Esta tarea no reejecutÃ³ Dysflow; el run fue aportado como evidencia externa del mismo dÃ­a.
- **Pruebas (ancla histÃ³rica)**: `openspec/changes/archive/2026-06-06-ce-fecha-obligatoria-postponement/archive-report.md` â€” `tests/tests.vba.json` (`filter=issue-19`) 13/13 PASS en `8cb7f0a` el 2026-06-06. Se mantiene como referencia archivada; el ancla vigente de runtime es el run fresco 2026-06-15.
- **Caveat de runner (histÃ³rico, ya resuelto)**: el intento fresco previo contra `tests/tests.vba.json` con filtro `Issue19_CE_Alta` quedÃ³ interrumpido y dejÃ³ la operaciÃ³n `dysflow-a54c004e-9ac8-41d8-93b9-087e5326c31d` en `status=starting`, `accessPid=null`. Tratarlo Ãºnicamente como nota histÃ³rica del runner; no afecta la promociÃ³n a `Verified-runtime` del run sano 2026-06-15.

### Procedimientos Issue19 cubiertos por la evidencia runtime fresca 2026-06-15

- **Slice A â€” baseline CE-gating behavior (8/8)**: `Test_Issue19_CE_Alta_Si_SinDetalle_NoBloquea`, `Test_Issue19_CE_Alta_Si_ConDetalle_Pasa`, `Test_Issue19_CE_Alta_No_IgnoraDetalle`, `Test_Issue19_CE_Cierre_SinDetalle_Bloquea`, `Test_Issue19_CE_Cierre_ConDetalle_PermiteCierre`, `Test_Issue19_CE_EstadoCalculado_Pendiente`, `Test_Issue19_CE_EstadoCalculado_SinPendiente`, `Test_Issue19_Paridad_UI_Dominio`.
- **Slice B â€” bypass/follow-up (5/5)**: `Test_Issue19_CE_Alta_MotivoAlta_Bypass_Si`, `Test_Issue19_CE_Alta_MotivoAlta_Bypass_BlankRequereCE`, `Test_Issue19_CE_Edicion_MotivoDatosUnicos_Bypass`, `Test_Issue19_CE_Auditoria_MotivoDatosUnicos_Bypass`, `Test_Issue19_CE_EficaciaOK_SinCambios`.

| Elemento (funcionalidad o arreglo) | Ref. tracker | VersiÃ³n staging (UAT) | Estado UAT | Release de producciÃ³n | Fecha en prod | Nota |
|---|---|---|---|---|---|---|
| Gate FE pospuesto al cierre | Issue #45 / issue-19 | `PRUEBAS-001` (v1.0.0) | `pending` (ronda UAT abierta 2026-06-15) | Pendiente | Pendiente | Web de aceptaciÃ³n: `docs/uat/PRUEBAS-001/uat-acceptance.html` (5 casos DADO/CUANDO/ENTONCES); recipient `andres.romandelperal@telefonica.com`. Ancla runtime previa: 13/13 PASS Dysflow 2026-06-15, 7 filtros; ancla histÃ³rica `8cb7f0a` (2026-06-06). |
| Ronda UAT `PRUEBAS-001` (Issue #19) | GH #45, #46 | `PRUEBAS-001` | `pending` (oficina, fecha de maÃ±ana 2026-06-16) | n/a (gate) | n/a (gate) | 5 casos: UAT-1 Alta sin fecha, UAT-2 Edicion sin fecha, UAT-3 Cierre sin fecha bloquea, UAT-4 Cierre con fecha permite, UAT-5 Bypass MotivoAlta. Criterios checksum computado al cargar. Update del Â§5 ledger con `passed`/`rejected` al recibir el correo de `andres.romandelperal@telefonica.com`. |
| Flujo completo de resultados de eficacia | Pendiente | Pendiente | pending | Pendiente | Pendiente | Faltan pruebas/reglas (BR-CE-5/6) |

| SÃ­ntoma | Causa probable | ComprobaciÃ³n (Dysflow) | Ancla documental |
|---|---|---|---|
| Alta/Edicion bloqueada por fecha FE | El gate regresÃ³ a validaciÃ³n temprana | `tests/tests.vba.json` filtro `Issue19_CE_Alta_` | BR-CE-1 |
| Cierre permite ausencia de fecha FE | Bypass de validaciÃ³n de cierre | `tests/tests.vba.json` filtro `Issue19_CE_Cierre_` | BR-CE-2 |
| Bypass incorrecto del botÃ³n de auditorÃ­a | Ruta diferida de auditorÃ­a sin resolver | Crear prueba de costura helper/servicio (FALTA) | BR-CE-6 `#71 (BR-CE-5/6)` |
| Estado de resultado/no requerido poco claro | Falta contrato de negocio | Confirmar reglas + crear pruebas (FALTA) | BR-CE-5 `#71 (BR-CE-5/6)` |

## Â§6 Notas de migraciÃ³n web

### Â§6.1 Conservar (comportamiento de negocio que debe sobrevivir)
- La regla "no exigir `FechaPrevistaControlEficacia` en `Alta`/`Edicion`" (BR-CE-1): la web debe permitir crear/editar NC sin fecha prevista FE; el campo puede ser nulo o vacÃ­o y el guardado no debe bloquearse solo por esa fecha ausente.
- El gate de `FechaPrevistaControlEficacia` se aplica exclusivamente al cierre (BR-CE-2): el endpoint de cierre de la web debe rechazar la operaciÃ³n cuando el control es requerido y la fecha es vacÃ­a o futura mÃ¡s allÃ¡ de la fecha de cierre, replicando `Test_Issue19_CE_Cierre_SinDetalle_Bloquea` y `Test_Issue19_CE_Cierre_ConDetalle_PermiteCierre`.
- La invariancia de `EficaciaOK` sobrevive a todas las rutas de bypass (BR-CE-4): cualquier camino de bypass para alta/edicion/auditoria debe preservar el cÃ¡lculo `EficaciaOK` (`Pendiente` / `SinPendiente` / `EficaciaOK_SinCambios`), igual que validan `Issue19_CE_EstadoCalculado_` y `Issue19_CE_EficaciaOK_SinCambios`.
- El bypass de auditorÃ­a por motivo de datos Ãºnicos (BR-CE-3): el botÃ³n `ComandoControlEficaciaDatos_Click` de `Form_FormNCAuditoriaGeneral` debe seguir permitiendo el bypass previsto vÃ­a `DatosGeneralesOK(p_MenosCef)`, con su test contractual (BR-CE-6).
- Los motivos de "no requiere control de eficacia" como vocabulario de dominio (BR-CE-5): `Form_FormMotivosNoRequiereControlEficacia` y su evento `MotivoRegistrado` deben seguir siendo la fuente de los motivos que luego se persisten en NC Proyecto y NC AuditorÃ­a, sin convertirse en texto libre.

### Â§6.2 Transformar (mecanismo legacy que se reformula)
- Sustituir el patrÃ³n "validaciÃ³n en `Form_Load`/`BeforeUpdate` de cada formulario" por un servicio/gate de cierre Ãºnico invocado por el comando de cerrar, tanto en Proyecto como en AuditorÃ­a.
- Convertir la comprobaciÃ³n de paridad UI/dominio en un middleware que valide `m_ObjEntorno` o equivalente antes de la operaciÃ³n, en lugar de un evento `OnTimer` con flag `m_CargaInicialIndicadoresPendiente`.
- Reemplazar la lectura directa de `EficaciaOK` desde el formulario de seguimiento por una API REST que devuelva el estado calculado y un payload con motivo, fecha y resultado, para que la UI solo pinte.
- Trasladar la lÃ³gica de bypass por motivo (`MotivoAlta`, `MotivoDatosUnicos`) a una tabla versionada y consultable, no como checks dispersos en el `.cls` de cada formulario.
- Modelar las replanificaciones de eficacia como un recurso explÃ­cito con estados (`replanificada`, `aceptada`, `cancelada`), no como un flag dentro del registro NC.

### Â§6.3 NO copiar (deuda legacy de Access que no debe portarse)
- No portar el acoplamiento TempVars/eventos de formulario para `EficaciaOK`: el estado debe vivir en la entidad de dominio NC y ser calculado por un servicio idempotente, no por la UI.
- No duplicar la lÃ³gica de "no requiere CE" en cada formulario de alta/ediciÃ³n: la web debe tener un Ãºnico punto que evalÃºe `RequiereControlEficacia` y centralice la decisiÃ³n.
- No migrar la separaciÃ³n fÃ­sica de "formularios de alta" vs "formularios de ediciÃ³n" como regla de UI: el backend debe aceptar ambas rutas y validar lo mismo.
- No reutilizar `OnTimer` + `TimerInterval = 100` como patrÃ³n de carga diferida de indicadores: la web puede usar un endpoint asÃ­ncrono o un skeleton explÃ­cito, no un timer del lado cliente.
- No exponer la fecha prevista FE como obligatoria en la API de `Crear`/`Editar`; la API de cierre es la Ãºnica que debe rechazarla cuando aplica.

### Â§6.4 Preguntas abiertas al product owner
- Â¿CuÃ¡les son los motivos canÃ³nicos de "no requiere control de eficacia" (BR-CE-5) y quÃ© campos de dominio los describen? Confirmar el catÃ¡logo completo.
- Â¿QuÃ© reglas de aprobado/no aprobado/no requerido/replanificaciÃ³n/evidencia se consideran obligatorias para release (BR-CE-5)? Â¿La replanificaciÃ³n tiene un nÃºmero mÃ¡ximo de veces o es indefinida?
- Â¿El bypass del botÃ³n de auditorÃ­a `ComandoControlEficaciaDatos_Click` (BR-CE-6) es por rol (calidad/auditor) o por motivo? Confirmar la regla de negocio antes de migrar.
- Â¿La evidencia documental es obligatoria en todos los cierres con control de eficacia o solo cuando el resultado es "no eficaz"? (BR-CE-5)
- Â¿La diferencia entre NC Proyecto y NC AuditorÃ­a en el gate de eficacia se mantiene o se unifica en la web? Hoy la regla es equivalente pero el botÃ³n de auditorÃ­a tiene bypass explÃ­cito.
- Â¿CÃ³mo se notifica al responsable cuando `EficaciaOK` cambia o cuando una replanificaciÃ³n vence? Confirmar canal y plantilla de notificaciÃ³n.

## Â§7 Libro de confianza
| Hecho | Confianza | Evidencia | Fecha |
|---|---|---|---|
| BR-CE-1 â€” `FechaPrevistaControlEficacia` no debe bloquear `Alta` ni `Edicion`. | Verified-runtime | Dysflow 2026-06-15: `tests/tests.vba.json` filtros `Issue19_CE_Alta_` 5/5 (`Test_Issue19_CE_Alta_Si_SinDetalle_NoBloquea`, `Test_Issue19_CE_Alta_Si_ConDetalle_Pasa`, `Test_Issue19_CE_Alta_No_IgnoraDetalle`, `Test_Issue19_CE_Alta_MotivoAlta_Bypass_Si`, `Test_Issue19_CE_Alta_MotivoAlta_Bypass_BlankRequereCE`); ancla histÃ³rica `8cb7f0a` | 2026-06-15 |
| BR-CE-2 â€” El gate de `FechaPrevistaControlEficacia` se aplica al cierre de NC. | Verified-runtime | Dysflow 2026-06-15: `tests/tests.vba.json` filtro `Issue19_CE_Cierre_` 2/2 (`Test_Issue19_CE_Cierre_SinDetalle_Bloquea`, `Test_Issue19_CE_Cierre_ConDetalle_PermiteCierre`); ancla histÃ³rica `8cb7f0a` | 2026-06-15 |
| BR-CE-3 â€” Los escenarios de bypass para alta, edicion y auditoria siguen soportados. | Verified-runtime | Dysflow 2026-06-15: `tests/tests.vba.json` filtros `Issue19_CE_Edicion_` 1/1 (`Test_Issue19_CE_Edicion_MotivoDatosUnicos_Bypass`) y `Issue19_CE_Auditoria_` 1/1 (`Test_Issue19_CE_Auditoria_MotivoDatosUnicos_Bypass`); ancla histÃ³rica `8cb7f0a` | 2026-06-15 |
| BR-CE-4 â€” La invariancia de `EficaciaOK` se preserva al mover el gate al cierre. | Verified-runtime | Dysflow 2026-06-15: `tests/tests.vba.json` filtros `Issue19_CE_EstadoCalculado_` 2/2 + `Issue19_CE_EficaciaOK_` 1/1 + `Issue19_Paridad_UI` 1/1; ancla histÃ³rica `8cb7f0a` | 2026-06-15 |
| BR-CE-5 â€” Reglas de motivo no requerido, eficacia fallida, replanificaciÃ³n y evidencia son explÃ­citas. | Intended | FALTA â†’ crear mediante access-vba-tdd tras confirmar esquema/reglas | 2026-06-15 |
| BR-CE-6 â€” El botÃ³n de control-eficacia general de auditorÃ­a respeta el bypass previsto (`DatosGeneralesOK(p_MenosCef)`). | Intended | FALTA â†’ crear mediante access-vba-tdd contra costura helper/servicio | 2026-06-15 |
| La fecha FE estÃ¡ prevista como gate solo de cierre. | Verified-runtime | Dysflow 2026-06-15: `tests/tests.vba.json` 13/13 PASS, 7 filtros verdes; ancla histÃ³rica `8cb7f0a` (2026-06-06) | 2026-06-15 |
| La invariancia de `EficaciaOK` forma parte del contrato de gate. | Verified-runtime | Dysflow 2026-06-15: `Issue19_CE_EficaciaOK_` 1/1 + `Issue19_CE_EstadoCalculado_` 2/2 + `Issue19_Paridad_UI` 1/1; ancla histÃ³rica `8cb7f0a` | 2026-06-15 |
| El flujo completo de resultados de eficacia estÃ¡ documentado y probado. | Intended | Reglas/pruebas pendientes (BR-CE-5/6) â€” `FALTA â†’ crear mediante access-vba-tdd` | 2026-06-15 |

**âš ï¸ Divergencias (intenciÃ³n SDD â‰  realidad del cÃ³digo)**
- Sin divergencia confirmada. Hueco sospechado: el comportamiento diferido del botÃ³n general de auditorÃ­a sigue documentado como abierto en lugar de probado.
