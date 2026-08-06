# Capacidad: documentos y evidencia generada

## Â§0 Identidad
- **ID de capacidad**: `CAP-DOC-EVIDENCE`
- **Nivel**: standard
- **Estado**: active / documentaciÃ³n alineada con v2; reglas de evidencia documental pendientes de prueba runtime
- **Fuente**: hybrid (inventario de fuente + documento de funcionalidad de auditorÃ­a + adyacencia de capacidades)
- **Responsable / autoridad de producto**: ConfirmaciÃ³n pendiente â€” Calidad / gestiÃ³n de evidencias
- **Ãšltima verificaciÃ³n**: 2026-06-15 migraciÃ³n solo documental; no se ejecutÃ³ Dysflow/Access
- **Confianza global**: low-to-mixed â€” la ruta de informe de auditorÃ­a estÃ¡ documentada como `Verified-static`; el comportamiento de adjuntos documentales, almacenamiento y trazabilidad UAT sigue sin prueba runtime

## Â§1 IntenciÃ³n de negocio
- **PropÃ³sito**: Mantener adjuntos, documentos generados y evidencia de informes para flujos de Proyecto, AuditorÃ­a, AC y AR.
- **Usuarios / personas**: Equipo de calidad, usuarios de auditorÃ­a/proyecto, revisores UAT, desarrolladores/agentes IA.
- **Problema que resuelve**: Garantiza que la evidencia de negocio siga vinculada al dominio correcto y pueda recuperarse o regenerarse para decisiones de ciclo de vida y prueba de release/UAT.
- **Valor de negocio / por quÃ© existe**: Los flujos de no conformidad necesitan evidencia auditable, no solo campos de estado.
- **No objetivos**: Esta pÃ¡gina no define arquitectura de filesystem/almacenamiento, permisos, nomenclatura/versionado ni polÃ­tica de retenciÃ³n hasta confirmar reglas de producto/IT.
- **Fuente de intenciÃ³n**: Borrador de capacidad existente + nombres de fuente; reglas documentales mayoritariamente pendientes de confirmaciÃ³n.
- **Referencia tracker de origen**: Issue #67 documentaciÃ³n de capacidades; evidencia de regresiÃ³n de informe de auditorÃ­a 2026-06-14.

## Â§2 Contrato de comportamiento

### Escenarios (Given / When / Then)
- **GIVEN** un padre Proyecto, AuditorÃ­a, AC o AR **WHEN** un usuario adjunta/revisa/elimina evidencia **THEN** la evidencia permanece vinculada al tipo de padre e ID de padre correctos.
- **GIVEN** una NC de auditorÃ­a seleccionada en `Form_FormNCAuditoriaGestion` **WHEN** se ejecuta el comando de informe **THEN** usa `EnsureNCAuditoriaGestionSelected` y `constructor.getNCAuditoria`, no la ruta de constructor de proyecto.
- **GIVEN** se requiere evidencia para cierre/UAT **WHEN** el usuario intenta cerrar o liberar **THEN** la evidencia obligatoria ausente debe bloquearse o informarse segÃºn reglas confirmadas.

### Reglas de negocio
| ID de regla | Enunciado (previsto) | Autoridad | Â¿Aplicada en cÃ³digo? | Prueba (evidencia) | Confianza |
|---|---|---|---|---|---|
| BR-DOC-1 `#80 (BR-DGE-1/2 en tracker, mapeo nominal â€” ver index.md)` | Los enlaces de evidencia deben incluir el contexto correcto de dominio/padre: NC Proyecto, NC AuditorÃ­a, AuditorÃ­a, AC o AR. | Contrato de capacidad | Desconocido / nombres de fuente indican superficies separadas | FALTA â†’ crear mediante access-vba-tdd tras inspecciÃ³n de esquema | Intended |
| BR-DOC-2 `#80 (BR-DGE-1/2 en tracker, mapeo nominal â€” ver index.md)` | Las rutas documentales de Proyecto y AuditorÃ­a no deben confundirse por servicios compartidos. | Inventario de fuente + necesidad de capacidad | Desconocido | FALTA â†’ crear mediante access-vba-tdd; pruebas de enrutamiento de dominio | Intended |
| BR-DOC-3 | La generaciÃ³n de informe de auditorÃ­a usa ruta de resolver/constructor de auditorÃ­a, no constructor de NC Proyecto. | Documento de funcionalidad de auditorÃ­a | SÃ­ â€” fuente actual documentada con `EnsureNCAuditoriaGestionSelected` / `constructor.getNCAuditoria` | Referencia archivada: `tests/tests.vba.audit-gestion-helper.json` / `Test_AuditGestionForm_ReportConstructorPath_Characterization`; FALTA â†’ reejecutar mediante access-vba-tdd para elevar a `Verified-runtime` | Verified-static |
| BR-DOC-4 | Las reglas de permisos para aÃ±adir/eliminar documentos, obligatoriedad, nomenclatura/versionado y retenciÃ³n son explÃ­citas. | Autoridad de producto pendiente | Desconocido | FALTA â†’ crear mediante access-vba-tdd tras confirmar reglas; aÃ±adir UAT cuando proceda | Intended |
| BR-DOC-5 | La evidencia generada es trazable a UAT/release antes de afirmarla como pasada. | EstÃ¡ndar capability-doc | Actualmente solo documentaciÃ³n | FALTA â†’ crear mediante access-vba-tdd donde sea automatizable; aÃ±adir filas de evidencia release/UAT | Intended |
| BR-DOC-6 | Las salidas Word/Excel y correos se enlazan con documentos/evidencias sin perder contrato de dominio, filtros ni privacidad. | Capacidad hermana `CAP-COMMS-REPORTS` | Parcial / pendiente de contrato | FALTA â†’ crear mediante access-vba-tdd para contrato Excel, send-vs-queue y BCC/privacidad | Intended |

### Validaciones
- Registro/dominio padre conocido antes de adjuntar evidencia.
- La NC seleccionada para informe de auditorÃ­a resuelve por ruta helper de auditorÃ­a (`EnsureNCAuditoriaGestionSelected`).
- Reglas de documento requerido, esquema de almacenamiento, tipo de archivo, tamaÃ±o, nomenclatura/versionado, retenciÃ³n y permisos pendientes de confirmaciÃ³n.

### Transiciones de estado
- `Sin documento` --(`Adjuntar`)--> `Documento enlazado al padre`.
- `Documento enlazado` --(`Sustituir/actualizar`)--> `Evidencia actualizada` â€” reglas de versionado pendientes.
- `Documento enlazado` --(`Eliminar`)--> `Evidencia eliminada/inactiva` â€” permisos/retenciÃ³n pendientes.
- `NC de auditorÃ­a seleccionada` --(`Generar informe`)--> `Informe de auditorÃ­a generado` â€” ruta de auditorÃ­a obligatoria.

### Caminos lÃ­mite y de error
- La mezcla de dominios entre evidencia Proyecto/AuditorÃ­a/AR es un sÃ­ntoma de riesgo de release.
- Las reglas de evidencia obligatoria ausente no pueden inferirse de los documentos actuales.

### SeÃ±ales de aceptaciÃ³n / presencia
- Las operaciones de evidencia preservan tipo de padre + ID.
- La ruta de informe de auditorÃ­a no puede instanciar NC de proyecto para selecciones de auditorÃ­a.
- El comportamiento de evidencia obligatoria estÃ¡ cubierto por pruebas/UAT explÃ­citos antes de afirmar release.

## Â§3 Mapa de implementaciÃ³n
- **Puntos de entrada UI**: `Form_FormNCProyectoDocumentos`, `Form_FormNCAuditoriaDocumentos`, `Form_FormAuditoriaDocumentos`, `Form_FormARProyectoDocumentos`, `Form_FormARAuditoriaDocumentos`, `Form_FormNCAuditoriaGestion.ComandoInforme_Click`.
- **Puntos de entrada de fuente**: `DocumentoService`, `DocumentoProyecto`, `DocumentoProyectoOperaciones`, `DocumentoAuditoria`, `DocumentoAuditoriaOperaciones`, `Informe` (que cubre tanto proyecto como auditorÃ­a vÃ­a `GenerarWordNoConformidades(p_EsDeProyecto)`). El mÃ³dulo `InformeNCAuditorias` fue retirado el 2026-06-15 como dead-code marker.
- **Datos tocados**: registros de documento/adjunto (esquema exacto pendiente), registros padre NC Proyecto/NC AuditorÃ­a/AuditorÃ­a/AR, selecciones de informe generado.
- **Salidas**: adjuntos, informes generados de NC de auditorÃ­a, documentos Word, exportaciones Excel, correos/Ã³rdenes de correo y paquete de evidencia UAT/release.
- **Dependencias e integraciones**: ciclo de vida de Proyecto, ciclo de vida de AuditorÃ­a, control eficacia, informes, almacenamiento/filesystem.
- **SincronizaciÃ³n fuenteâ†”binario**: no comprobada en esta tarea solo documental.
- **EvaluaciÃ³n de diseÃ±o (as-built vs ideal)**: formularios/clases separados sugieren separaciÃ³n de dominio, pero la ausencia de pruebas documentadas de almacenamiento/retenciÃ³n/permisos es un riesgo significativo de migraciÃ³n.

## Â§4 Receta de reconstrucciÃ³n
1. Confirmar esquema documental, ubicaciÃ³n de almacenamiento, claves padre, permisos, nomenclatura/versionado, retenciÃ³n y reglas de evidencia obligatoria.
2. Para generaciÃ³n de informes, mantener la lÃ³gica de selecciÃ³n de informe en una costura helper/servicio; probar la costura en lugar del comportamiento directo de formulario.
3. Crear pruebas fixture con esquema primero para enrutamiento de dominio padre y selecciÃ³n de constructor de informe generado.
4. AÃ±adir filas de evidencia UAT/release para flujos documentales obligatorios y salidas generadas.
5. Importar cualquier cambio futuro de fuente solo mediante Dysflow MCP; el usuario compila manualmente; despuÃ©s ejecutar pruebas Dysflow.

## Â§5 Evidencia y trazabilidad
- **Pruebas**: el documento de funcionalidad de auditorÃ­a existente cita `tests/tests.vba.audit-gestion-helper.json` / `Test_AuditGestionForm_ReportConstructorPath_Characterization`. No hubo ejecuciÃ³n reciente en esta tarea y no se afirma `Verified-runtime` nuevo.
- **Candidatas para prueba runtime futura**:
  - `tests/tests.vba.audit-gestion-helper.json`: `Test_AuditGestionForm_ReportConstructorPath_Characterization`, `Test_AuditListadoHelper_RowAndReportContracts_RED`.
  - `tests/tests.vba.seguimiento-tareas-helper.json`: `Test_TareasHelper_DeterministicOrder_ExportInput`, `Test_TareasForm_Delegates_FilterPaths`.
  - `tests/tests.vba.cache-e2e.json` / `tests/tests.vba.cache-warmup.json`: `Test_E2E_Cache_PrecalentarSincronizar_LogEvidence_Atomic`.
- **Enlaces cruzados**: ver `docs/capabilities/nc-auditoria-lifecycle.md` y `docs/features/audit/audit-backend-list-cache.md` para la ruta de informe de auditorÃ­a; ver `docs/capabilities/communications-reports-exports.md` para correo, Word y Excel.

| Elemento (funcionalidad o arreglo) | Ref. tracker | VersiÃ³n staging (UAT) | Estado UAT | Release de producciÃ³n | Fecha en prod | Nota |
|---|---|---|---|---|---|---|
| El informe de auditorÃ­a usa ruta de constructor de auditorÃ­a | Evidencia de regresiÃ³n 2026-06-14 | Pendiente | pending | Pendiente | Pendiente | `Verified-static`: fuente documentada con `EnsureNCAuditoriaGestionSelected` / `constructor.getNCAuditoria`; runtime fresco pendiente. |
| Comportamiento general de adjuntos/evidencia documental | Issue #67 | Pendiente | pending | Pendiente | Pendiente | Faltan pruebas de capacidad y reglas de negocio. |
| Contrato de salidas Word/Excel/correo como evidencia | Issue #67 | Pendiente | pending | Pendiente | Pendiente | FALTA â†’ contrato Excel, send-vs-queue, BCC/privacidad y trazabilidad UAT/release. |

| SÃ­ntoma | Causa probable | ComprobaciÃ³n (Dysflow) | Ancla documental |
|---|---|---|---|
| El informe de auditorÃ­a abre dominio incorrecto | RegresiÃ³n de ruta de informe | Reejecutar prueba helper/ruta de informe de auditorÃ­a | BR-DOC-3 |
| El documento aparece bajo padre incorrecto | Faltan pruebas de enrutamiento de dominio | Crear pruebas de enrutamiento documental | BR-DOC-1..2 |
| Falta evidencia de cierre/UAT | Reglas de evidencia obligatoria desconocidas | Confirmar reglas y despuÃ©s crear pruebas/UAT | BR-DOC-4..5 |
| ExportaciÃ³n/correo no trazable como evidencia | Contrato de salida y privacidad no definido | Crear pruebas de contrato Excel/correo tras inspecciÃ³n de esquema | BR-DOC-6 |

## Â§6 Notas de migraciÃ³n web

### Â§6.1 Conservar (comportamiento de negocio que debe sobrevivir)
- El enlace de cada documento/evidencia al `parentType` y `parentId` correctos: NC Proyecto, NC AuditorÃ­a, AuditorÃ­a, AC o AR (BR-DOC-1). La API REST de la web debe exigir un discriminador de tipo de padre y rechazar adjuntos huÃ©rfanos o con padre de tipo incorrecto.
- El enrutamiento especÃ­fico de dominio para Proyecto vs AuditorÃ­a: una evidencia subida a una NC de Proyecto no debe poder consultarse desde una NC de AuditorÃ­a ni viceversa (BR-DOC-2). La separaciÃ³n de `DocumentoProyecto` y `DocumentoAuditoria` debe sobrevivir a la migraciÃ³n.
- La ruta de generaciÃ³n de informes de auditorÃ­a pasa por `EnsureNCAuditoriaGestionSelected` + `constructor.getNCAuditoria`, no por `constructor.getNCProyecto` (BR-DOC-3). El servicio de generaciÃ³n de informes de la web debe replicar esa decisiÃ³n, con un guard que rechace si la NC no es de auditorÃ­a.
- La cobertura de los formularios documentales de cada dominio: `Form_FormNCProyectoDocumentos`, `Form_FormNCAuditoriaDocumentos`, `Form_FormAuditoriaDocumentos`, `Form_FormARProyectoDocumentos`, `Form_FormARAuditoriaDocumentos`. La web debe mantener cinco rutas/parientes de UI separadas, no un Ãºnico punto comÃºn.
- El hecho de que el mÃ³dulo `InformeNCAuditorias` fue retirado como dead-code marker el 2026-06-15: la web no debe reintroducir una clase paralela de generaciÃ³n de informe de auditorÃ­a; la ruta canÃ³nica es `Informe.GenerarWordNoConformidades(p_EsDeProyecto:=No)`.

### Â§6.2 Transformar (mecanismo legacy que se reformula)
- Sustituir los formularios documentales por endpoints REST con upload directo a almacenamiento de objetos (S3/blob) + entrada de evidencia inmutable en base de datos, no por rutas de filesystem configuradas en `Entorno` o variables de entorno.
- Convertir el `DocumentoService` + `DocumentoProyecto`/`DocumentoAuditoria` en un Ãºnico servicio polimÃ³rfico de evidencias con un campo `parentType` que enrute al repositorio correcto, no cinco clases paralelas.
- Mover la generaciÃ³n de Word a un servicio server-side que use una plantilla controlada por configuraciÃ³n y devuelva una URL prefirmada de descarga, en lugar de `Informe.GenerarWordNoConformidades` ejecutÃ¡ndose en cliente.
- Reemplazar la BCC por defecto de `Correo.Registrar` y la polÃ­tica de privacidad embebida por una configuraciÃ³n auditada de privacidad y notificaciÃ³n, versionada y revisable.
- Sustituir el contrato de exportaciÃ³n Excel acoplado al formulario de seguimiento por un endpoint `GET /exportaciones/...` con un contrato explÃ­cito de columnas, filtros y formato, expuesto a travÃ©s de la API de documentos.

### Â§6.3 NO copiar (deuda legacy de Access que no debe portarse)
- No portar rutas de filesystem absolutas (`C:\...`, rutas UNC) como configuraciÃ³n de almacenamiento: la web debe usar un servicio de objetos o un volumen montado, no rutas locales de un usuario.
- No duplicar la lÃ³gica de "quÃ© plantilla Word aplica" en cada consumidor: la web debe tener un Ãºnico servicio de generaciÃ³n que centralice la decisiÃ³n `p_EsDeProyecto`.
- No usar el filename original sin sanear como nombre visible: la web debe aplicar nomenclatura versionada y revisiÃ³n de extensiones peligrosas antes de almacenar.
- No migrar la separaciÃ³n fÃ­sica de formularios documentales por dominio como regla de UI obligatoria: la web puede tener un solo formulario siempre que el discriminador `parentType` se mantenga.
- No usar `DoCmd.OpenForm` con `OpenArgs` para enlazar un documento a su padre: la API web debe recibir el `parentId` como path/query param, no como parÃ¡metro opaco.

### Â§6.4 Preguntas abiertas al product owner
- Â¿CuÃ¡l es la polÃ­tica de retenciÃ³n y borrado de documentos por tipo de padre? (BR-DOC-4) Â¿Se borran al cerrar la NC, se conservan N aÃ±os, o son inmutables?
- Â¿QuÃ© tipos de archivo, tamaÃ±o mÃ¡ximo y nomenclatura/versionado son obligatorios? (BR-DOC-4) Confirmar si Word, Excel, PDF, imÃ¡genes y zip estÃ¡n permitidos.
- Â¿QuÃ© permisos por rol aplican a aÃ±adir/eliminar/descargar documentos? (BR-DOC-4) Â¿La polÃ­tica es la misma para Proyecto y AuditorÃ­a?
- Â¿La generaciÃ³n de Word como evidencia requiere plantilla aprobada por Calidad o cada equipo puede subir la suya? (BR-DOC-3, BR-DOC-6)
- Â¿Las salidas Excel de seguimiento cuentan como "evidencia" para UAT/release o solo los adjuntos subidos? (BR-DOC-6)
- Â¿El `parentType` debe ser extensible (e.g. para futuras AC con subflujos) o se cierra al conjunto actual {Proyecto, AuditorÃ­a, AuditorÃ­a, AC, AR}?

## Â§7 Libro de confianza
| Hecho | Confianza | Evidencia | Fecha |
|---|---|---|---|
| BR-DOC-1 â€” Los enlaces de evidencia deben incluir el contexto correcto de dominio/padre: NC Proyecto, NC AuditorÃ­a, AuditorÃ­a, AC o AR. | Intended | FALTA â†’ crear mediante access-vba-tdd tras inspecciÃ³n de esquema | 2026-06-15 |
| BR-DOC-2 â€” Las rutas documentales de Proyecto y AuditorÃ­a no deben confundirse por servicios compartidos. | Intended | FALTA â†’ crear mediante access-vba-tdd; pruebas de enrutamiento de dominio | 2026-06-15 |
| BR-DOC-3 â€” La generaciÃ³n de informe de auditorÃ­a usa ruta de resolver/constructor de auditorÃ­a, no constructor de NC Proyecto. | Verified-static | `EnsureNCAuditoriaGestionSelected` / `constructor.getNCAuditoria`; referencia archivada `tests/tests.vba.audit-gestion-helper.json` / `Test_AuditGestionForm_ReportConstructorPath_Characterization`; FALTA â†’ reejecutar mediante access-vba-tdd para elevar a `Verified-runtime` | 2026-06-15 |
| BR-DOC-4 â€” Las reglas de permisos para aÃ±adir/eliminar documentos, obligatoriedad, nomenclatura/versionado y retenciÃ³n son explÃ­citas. | Intended | FALTA â†’ crear mediante access-vba-tdd tras confirmar reglas; aÃ±adir UAT cuando proceda | 2026-06-15 |
| BR-DOC-5 â€” La evidencia generada es trazable a UAT/release antes de afirmarla como pasada. | Intended | FALTA â†’ crear mediante access-vba-tdd donde sea automatizable; aÃ±adir filas de evidencia release/UAT | 2026-06-15 |
| BR-DOC-6 â€” Las salidas Word/Excel y correos se enlazan con documentos/evidencias sin perder contrato de dominio, filtros ni privacidad. | Intended | FALTA â†’ crear mediante access-vba-tdd para contrato Excel, send-vs-queue y BCC/privacidad | 2026-06-15 |
| La ruta de informe de auditorÃ­a debe usar resolver/constructor de auditorÃ­a. | Verified-static | Documento de funcionalidad de auditorÃ­a existente; sin reejecuciÃ³n | 2026-06-15 |
| Existen superficies UI documentales separadas para varios dominios. | Verified-static | Inventario de fuente de documentos existentes | 2026-06-15 |
| El comportamiento de aÃ±adir/eliminar/almacenar/retener documentos estÃ¡ protegido para release. | Intended | Faltan reglas/pruebas | 2026-06-15 |
| El contrato de Excel, correo send-vs-queue, BCC/privacidad y trazabilidad UAT/release estÃ¡ cerrado. | Intended | Faltan reglas/pruebas | 2026-06-15 |

**âš ï¸ Divergencias (intenciÃ³n SDD â‰  realidad del cÃ³digo)**
- Sospechada pero sin confirmar: el arreglo de ruta de informe de auditorÃ­a tiene trazabilidad de commit pendiente, por lo que la documentaciÃ³n puede describir comportamiento aÃºn no trazable a un ancestro de staging con commit.
