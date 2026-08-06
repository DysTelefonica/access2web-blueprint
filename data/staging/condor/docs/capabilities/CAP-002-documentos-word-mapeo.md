# Capacidad: Generación de documentos Word y validación de mapeos de plantilla

## §0 Identidad

- **ID de capacidad**: CAP-002
- **Tier**: critical
- **Estado**: active con evidencia runtime parcial y bloqueo de operabilidad E2E
- **Source**: hybrid
- **Responsable / autoridad de producto**: Pendiente de confirmación
- **Última verificación**: `dysflow.verify_binary(Test_DocumentTemplateMapping)` con `actionableOk=true` (solo `caseOnly`); `Test_DTM_StrictMissingFieldValidation` pasó en 6471 ms; `Test_DTM_AllMappedFields` agotó timeout aislado a 300 s y dejó `WINWORD.EXE /Automation -Embedding` huérfano, cerrado por PID diagnosticado. **Slice A4 (2026-06-15)**: dos refactors (split per-template + XML-directo sin Word COM) no salvaron el timeout — el runner de Dysflow embebido en Access se cuelga en operaciones PowerShell/Shell.Application desde VBA. Código WIP en el `.bas` y script manual en `scripts/run-dtm-tests.ps1`.
- **Confianza global**: mixta: `Verified-runtime` para el seam de validación estricta de campo Word ausente; `Verified-static` para el contrato completo plantilla↔`tbMapeoCampos` (4 refactors intentados en 2026-06-15 — split per-template, XML-directo con PowerShell/Shell.Application — todos se colgaron en el runner Dysflow con `timeoutMs=60000`; el XML corre en 2s en PowerShell standalone, lo que confirma que el cuello de botella es la invocación de subprocess desde VBA embebido, no el algoritmo).
- **Deuda TDD v2.4.2 transversal**: ver [Deuda crítica de pruebas Access/VBA TDD v2.4.2](../testing/access-vba-tdd-v2_4_2-debt.md). El E2E documental puede ser evidencia read-only de plantilla si se ejecuta en verde; no sustituye evidencia de mutación, persistencia ni generación con fixture de negocio.

**Contrato TDD vigente**: desde `access-vba-tdd` v2.4.2, ninguna prueba nueva puede reclamar `Verified-runtime` si no devuelve JSON canónico, no usa fixture propio cuando toca datos de negocio, no hace revisión schema-first, no inyecta `DAO.Database` explícito en rutas persistentes, no verifica cardinalidad de mutaciones o mezcla agregadores con tests atómicos. Para esta capacidad, las pruebas de contrato de plantilla pueden ser de solo lectura sobre `tbMapeoCampos`; esa configuración es el sujeto bajo prueba y no debe ser sembrada ni normalizada por el test.

## §1 Intención de negocio — POR QUÉ

- **Propósito**: generar documentos Word y PDF de solicitudes CONDOR a partir de plantillas corporativas, rellenando campos de formulario con datos técnicos de PC, CD/CA, CD/CA-SUB y PC-SUB.
- **Usuarios / perfiles**: técnicos, calidad, RAC, autoridad de decisión y personal que formaliza o intercambia documentación con terceros.
- **Problema que resuelve**: evita redactar documentos contractuales manualmente y centraliza el vínculo entre datos de solicitud y campos Word.
- **Valor de negocio**: reduce errores de transcripción, preserva trazabilidad de versiones de borrador y detecta deriva entre configuración `tbMapeoCampos` y plantillas reales.
- **No-objetivos**: no sustituye la revisión humana del contenido final ni valida firmas digitales; tampoco migra plantillas Word a otro formato.
- **Origen de la intención**: SDD `openspec/changes/document-template-mapping-e2e` y lectura estática de código.
- **Referencia de tracker de origen**: Pendiente de confirmación.

## §2 Contrato de comportamiento — QUÉ

### Escenarios

- **DADO** una solicitud con datos técnicos existentes **CUANDO** se invoca generación de documento **ENTONCES** se selecciona el generador por tipo (`PC`, `CD_CA`, `CD_CA_SUB`, `PC_SUB`) y se crea un DOCX a partir de la plantilla correspondiente. **Estado**: `Verified-static`.
- **DADO** un mapeo real en `tbMapeoCampos` **CUANDO** el generador intenta escribir un `nombreCampoWord` ausente **ENTONCES** la generación falla con contexto de plantilla, campo de tabla, campo Word y rutas. **Estado**: `Verified-runtime` para el seam de ausencia controlada (`Test_DTM_StrictMissingFieldValidation`); el barrido completo de campos reales sigue bloqueado por timeout.
- **DADO** un texto técnico de más de 255 caracteres **CUANDO** se exporta a Word **ENTONCES** el contenido se trocea en campo base, extensiones `_extN` y, si procede, bloque `Cont` sin sobrescribirlo después con vacío. **Estado**: `Verified-static`.
- **DADO** una solicitud que requiere PDF **CUANDO** se genera el Word intermedio y se exporta **ENTONCES** Word se abre en segundo plano, se exporta a PDF y se intenta cerrar Word y limpiar temporales. **Estado**: `Verified-static`; requiere E2E con cierre de Word garantizado.
- **DADO** un documento Word devuelto por RAC **CUANDO** se importan datos RAC **ENTONCES** se leen campos RAC mapeados, extensiones y comentarios Word hacia `observacionesRAC` según tipo de solicitud. **Estado**: `Verified-static`.

### Reglas de negocio

| ID regla | Enunciado | Autoridad | ¿Aplicada en código? | Prueba / evidencia | Confianza |
|---|---|---|---|---|---|
| BR-001 | Cada tipo de solicitud usa su plantilla y mapeo propios. | SDD + código | Sí: `GenerarDocumentoGenerico` selecciona `getMapeoPC`, `getMapeoCDCA`, `getMapeoCDCASUB` o `getMapeoPCSUB`. | `DocumentoServicio.cls`; `MapeoServicio.cls`. | Verified-static |
| BR-002 | La configuración real `tbMapeoCampos` no se sustituye por fixtures en los tests de contrato de plantilla. | SDD `document-template-mapping-e2e` | Pretendido y reflejado en manifest read-only. | `tests/tests.document-template-e2e.json`; spec SDD. | Intended / Verified-static documental |
| BR-003 | Un campo Word mapeado inexistente debe fallar de forma explícita, no omitirse silenciosamente. | SDD + código | Sí: `ValidarCampoWordMapeado` y `SetFormFieldValue`. | `Test_DTM_StrictMissingFieldValidation`: OK, 7870 ms. | Verified-runtime |
| BR-004 | Los textos largos se dividen en bloques de 255 caracteres por limitación de campos Word. | Código + PRD CDCA | Sí: `MAX_CHARS_POR_MARCADOR = 255`, `_extN` y `ProcesarContinuationBlock`. | `DocumentoServicio.cls`; PRD `06_Formulario_Datos_CDCA.md`. | Verified-static |
| BR-005 | La generación PDF debe eliminar el DOCX temporal y cerrar Word tanto en éxito como en error. | Código | Parcial: hay bloques de limpieza y cierre; se usan también terminaciones de WINWORD nuevos. | `DocumentoServicio.cls`. | Verified-static / riesgo |
| BR-006 | Para generar PDF por solicitud deben existir datos técnicos y, en PDF, aprobación de suministrador y dictamen RAC completos. | Código | Sí en `GenerarDocumentoPDFparaSolicitud`. | `DocumentoServicio.cls`; falta prueba fixture-first por tipo. | Verified-static |
| BR-007 | La importación RAC desde Word solo debe poblar campos RAC relevantes y comentarios. | Código + PRD | Sí en métodos `ImportarDatosRACDesdeWord*`. | `DocumentoServicio.cls`; PRD CDCA. | Verified-static |

### Validaciones

- Plantilla inexistente → error de validación con ruta de plantilla.
- Mapeo sin filas → error de validación para la plantilla.
- `nombreCampoWord` vacío → error con contexto de plantilla y campo de tabla.
- `FormField` ausente → error con plantilla, campo de tabla, campo Word, ruta de plantilla y ruta de destino.
- Solicitud inexistente o tipo no reconocido → error antes de generar documento.

### Señales de aceptación / presencia

- Existe `DocumentoServicio.GenerarDocumentoParaSolicitud` con ramas para `PC`, `CD_CA`, `CD_CA_SUB` y `PC_SUB`.
- Existe `DocumentoServicio.ValidarCampoWordMapeado`; la ausencia de un campo mapeado no se omite silenciosamente.
- Existe `tests/tests.document-template-e2e.json` con `Test_DTM_AllMappedFields` y `Test_DTM_StrictMissingFieldValidation`; el segundo está en verde, el primero requiere partición/instrumentación por timeout.
- `MapeoRepositorio.getMapeoParaPlantilla` lee `tbMapeoCampos` mediante `getdb()`.

## §3 Mapa de implementación — CÓMO

- **Puntos de entrada de código**:
  - `DocumentoServicio.GenerarDocumentoParaSolicitud(idSolicitud, numVersion, db)` genera DOCX por tipo.
  - `DocumentoServicio.GenerarDocumentoPDFparaSolicitud(idSolicitud)` genera PDF con validaciones adicionales.
  - `DocumentoServicio.GenerarDocumentoGenerico(objDatosTecnicos, objSolicitud, nombrePlantilla, rutaPlantilla)` copia plantilla, lee mapeos y escribe campos.
  - `DocumentoServicio.ImportarDatosRACDesdeWord*` importa datos RAC desde documentos devueltos.
- **Datos afectados**:
  - Lectura: `tbMapeoCampos`, `tbSolicitudes`, tablas específicas de datos técnicos.
  - Escritura: archivos DOCX/PDF en directorios configurados; importación RAC solo modifica objetos en memoria hasta que otro servicio persista.
- **Dependencias e integraciones**: Microsoft Word COM, `m_ObjEntorno` para rutas, `WorkflowServicio` para versión de borrador, servicios de datos por tipo de solicitud y `MapeoServicio`/`MapeoRepositorio`.
- **Valoración de diseño**: la generación genérica reduce duplicación entre tipos y la validación estricta de campos es un buen ancla contra regresiones. La deuda principal está en testabilidad: Word COM y archivos temporales requieren seam de servicio/helper para aislar cierre de procesos, y las pruebas que toquen solicitudes deben preparar fixture v2.4.2 completo antes de reclamar runtime.

## §4 Receta de reconstrucción — REPRODUCIBILIDAD

1. Restaurar `MapeoCampos.cls`, `MapeoServicio.cls` y `MapeoRepositorio.bas`.
2. Restaurar `DocumentoServicio.cls` con `GenerarDocumentoGenerico`, `SetFormFieldValue`, `ValidarCampoWordMapeado`, overflow `_extN` y bloque `Cont`.
3. Confirmar que `m_ObjEntorno` resuelve rutas de plantillas y directorios de salida para los cuatro tipos.
4. Mantener `tbMapeoCampos` como configuración real de contrato; no sembrarla en pruebas de mapeo.
5. Importar/compilar con Dysflow solo cuando haya cambios de código; esta tarea no ejecutó Access por restricción explícita.
6. Ejecutar `Test_DTM_StrictMissingFieldValidation` como prueba focal rápida. Para `Test_DTM_AllMappedFields`, dividir por plantilla o instrumentar duración antes de usarlo como gate, porque el barrido completo agotó 300 s en esta estación.

## §5 Evidencia y trazabilidad

- **Evidencia estática consultada**:
  - `src/classes/DocumentoServicio.cls`
  - `src/classes/MapeoServicio.cls`
  - `src/modules/MapeoRepositorio.bas`
  - `tests/tests.document-template-e2e.json`
  - `openspec/changes/document-template-mapping-e2e/proposal.md`
  - `openspec/changes/document-template-mapping-e2e/specs/document-template-mapping/spec.md`
- **Evidencia runtime incorporada**:
  - `dysflow.verify_binary(Test_DocumentTemplateMapping)`: `actionableOk=true`; diferencia `caseOnly` no funcional.
  - `Test_DTM_StrictMissingFieldValidation`: OK, 7870 ms; valida error 513 con `Plantilla`, `CampoTabla`, `CampoWord`, `RutaPlantilla` y `RutaDestino`.
  - `Test_DTM_AllMappedFields`: timeout aislado a 300 s; dejó `WINWORD.EXE /Automation -Embedding` huérfano y se cerró por PID diagnosticado. No reclamar como verde.
  - Refactor XML (lectura directa de `word/document.xml` del .docx via PowerShell `Expand-Archive`): corre en 2s en PowerShell standalone PERO cuelga al ejecutarse desde `dysflow.test_vba` con `timeoutMs=60000`. Cuello de botella es la invocación de subprocess desde VBA embebido en el runner de Dysflow.

### Diagnóstico de regresión

| Síntoma | Causa probable | Comprobación | Ancla |
|---|---|---|---|
| El documento se genera sin un campo esperado | mapeo apunta a campo Word inexistente o validación estricta no ejecutada | Ejecutar `Test_DTM_StrictMissingFieldValidation`; para barrido real, dividir `Test_DTM_AllMappedFields` por plantilla antes de usarlo como gate | §2 / §3 |
| Texto largo aparece truncado | faltan extensiones `_extN` o bloque `Cont` en plantilla/mapeo | test de plantilla + inspección de `numExtensiones` | §2 |
| Word queda abierto tras fallo | limpieza incompleta o proceso huérfano de Word | E2E con fallo controlado y comprobación de cierre | §5 |
| Prueba de generación pasa sin probar datos de negocio | test solo valida mapeo read-only, no fixture de solicitud | añadir E2E con fixture de solicitud legal | §5 |

### Deuda de pruebas v2.4.2

- El informe transversal [Deuda crítica de pruebas Access/VBA TDD v2.4.2](../testing/access-vba-tdd-v2_4_2-debt.md) centraliza el impacto general sobre la evidencia.
- `tests/tests.document-template-e2e.json` no puede tratarse como gate único todavía: `Test_DTM_StrictMissingFieldValidation` pasa, pero `Test_DTM_AllMappedFields` requiere partición/instrumentación por timeout y cierre de Word.
- Falta un E2E de generación con fixture de negocio propio que demuestre que valores sembrados llegan al DOCX sin escribir `tbMapeoCampos`.
- Cualquier E2E de PDF debe garantizar cierre de Word/documento en éxito y error, y reportar residuos de proceso como fallo o bloqueo diagnosticable.
- El runner probe no debe usarse como evidencia funcional; solo diagnostica que el runner puede invocar procedimientos.

## §6 Notas de migración web

- **Conservar**: contrato de mapeo explícito, error fuerte ante campo de plantilla ausente, separación por tipo de solicitud, versión de borrador y pruebas de compatibilidad plantilla↔mapeo.
- **Transformar**: Word COM debería reemplazarse por un generador documental desacoplado, con plantillas versionadas y validación de esquema en CI.
- **NO copiar**: terminación directa de procesos Word como mecanismo normal; dependencia de rutas globales Access; pruebas que pasan por disponibilidad ambiental de Word sin contrato claro.

## §7 Registro de confianza

| Hecho | Confianza | Evidencia | Fecha |
|---|---|---|---|
| La generación de documentos usa un generador genérico para PC, CD/CA, CD/CA-SUB y PC-SUB. | Verified-static | `DocumentoServicio.GenerarDocumentoGenerico` y wrappers por tipo. | 2026-06-15 |
| Los mapeos se leen desde `tbMapeoCampos` por nombre de plantilla. | Verified-static | `MapeoServicio` + `MapeoRepositorio.getMapeoParaPlantilla`. | 2026-06-15 |
| La ausencia de un campo Word mapeado falla con contexto. | Verified-runtime | `Test_DTM_StrictMissingFieldValidation`: OK, error 513 con contexto completo. | 2026-06-15 |
| El manifest documental existente es read-only y no fixturea `tbMapeoCampos`. | Verified-static documental | `tests/tests.document-template-e2e.json`. | 2026-06-15 |
| El manifest documental completo está en verde actualmente. | Divergent / bloqueado | `Test_DTM_AllMappedFields` timeout aislado a 300 s; no reclamar manifest completo como verde. | 2026-06-15 |

**Divergencias / riesgos pendientes**:

- La evidencia read-only de plantilla no sustituye pruebas de generación con fixture de negocio.
- La limpieza de Word debe validarse con E2E específico antes de tratarla como garantía runtime; el timeout de `Test_DTM_AllMappedFields` dejó Word huérfano y confirmó el riesgo.
