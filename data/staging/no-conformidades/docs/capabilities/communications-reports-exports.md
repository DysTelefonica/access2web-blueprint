# Capacidad: comunicaciones, informes y exportaciones

## Â§0 Identidad
- **ID de capacidad**: `CAP-COMMS-REPORTS`
- **Tier**: standard
- **Estado**: active / inventario documental inicial; contratos de salida pendientes de prueba runtime
- **Source**: hybrid
- **Responsable / autoridad de producto**: Pendiente de confirmaciÃ³n â€” Calidad / comunicaciones operativas
- **Ãšltima verificaciÃ³n**: 2026-06-15 mediante inspecciÃ³n estÃ¡tica; no se ejecutÃ³ Dysflow/Access
- **Confianza global**: mixta â€” rutas de cÃ³digo verificadas estÃ¡ticamente; comportamiento de envÃ­o/salida, contrato Excel y polÃ­tica BCC/privacidad sin prueba runtime reciente

## Â§1 IntenciÃ³n de negocio
- **PropÃ³sito**: Generar comunicaciones e informes de NC para evidenciar estado, acciones y tareas ante responsables y revisores.
- **Usuarios / perfiles**: Calidad, responsables de proyecto/auditorÃ­a, revisores UAT y soporte.
- **Problema que resuelve**: Sin salidas reproducibles, la informaciÃ³n de NC queda dentro de Access y no puede circular ni auditarse fuera de la aplicaciÃ³n.
- **Valor de negocio / por quÃ© existe**: Permite notificar, registrar correos pendientes de envÃ­o y generar documentos Word con informaciÃ³n de NC y acciones.
- **No-objetivos**: No define polÃ­tica corporativa de correo, BCC/privacidad ni almacenamiento/retenciÃ³n final de documentos.
- **Origen de la intenciÃ³n**: CÃ³digo exportado + docs de auditorÃ­a y capability de documentos.
- **Referencia de tracker de origen**: Issue #67; evidencia de regresiÃ³n de informe de auditorÃ­a 2026-06-14.

## Â§2 Contrato de comportamiento

### Escenarios (Dado / Cuando / Entonces)
- **DADO** una NC Proyecto activa **CUANDO** se abre `Form_FormCorreo` en modo proyecto **ENTONCES** se propone asunto de informe de NC de proyecto y destinatario del responsable TelefÃ³nica si existe.
- **DADO** una NC AuditorÃ­a activa **CUANDO** se abre `Form_FormCorreo` en modo auditorÃ­a **ENTONCES** se propone asunto de auditorÃ­a y destinatarios de calidad en pruebas.
- **DADO** que el usuario envÃ­a un correo **CUANDO** faltan asunto o destinatarios **ENTONCES** la operaciÃ³n se bloquea con mensaje de validaciÃ³n.
- **DADO** datos vÃ¡lidos **CUANDO** se registra el correo **ENTONCES** se crea una fila en `TbCorreosEnviados` con asunto, cuerpo HTML, originador y destinatarios.
- **DADO** una o varias NC **CUANDO** se genera Word **ENTONCES** se copia la plantilla de Proyecto/AuditorÃ­a al directorio local de informes y se inserta la informaciÃ³n de NC, AC y tareas.

### Reglas de negocio
| ID regla | Enunciado (pretendido) | Autoridad | Â¿Aplicada en cÃ³digo? | Prueba | Confianza |
|---|---|---|---|---|---|
| BR-COM-1 | No se puede ordenar correo sin asunto. | CÃ³digo exportado | SÃ­ â€” `Form_FormCorreo.ComandoEnviarCorreo_Click`, `Correo.Registrar` | FALTA â†’ crear mediante access-vba-tdd con objeto `Correo` y fixture/control de `TbCorreosEnviados` | Verified-static |
| BR-COM-2 | No se puede ordenar correo sin destinatario principal, copia ni copia oculta. | CÃ³digo exportado | SÃ­ â€” formulario y clase `Correo` | FALTA â†’ crear mediante access-vba-tdd | Verified-static |
| BR-COM-3 `#75` | El cuerpo del correo se genera en HTML desde la NC activa y debe incluir acciones. | CÃ³digo exportado | SÃ­ â€” `HTMLNCProyecto` / `HTMLNCAuditoria` con `p_ConAcciones:=EnumSino.SÃ­` | FALTA â†’ crear mediante access-vba-tdd sobre HTML | Verified-static |
| BR-COM-4 | El correo se registra como orden de envÃ­o en `TbCorreosEnviados`; no se afirma envÃ­o SMTP directo. | CÃ³digo exportado | SÃ­ â€” `Correo.Registrar` hace `AddNew` | FALTA â†’ crear mediante access-vba-tdd con cardinalidad | Verified-static |
| BR-COM-5 `#76` | Si no hay BCC, se aÃ±ade una copia oculta por defecto. | CÃ³digo exportado | SÃ­ â€” `Correo.Registrar` | FALTA â†’ crear mediante access-vba-tdd; confirmar si sigue siendo regla de negocio vÃ¡lida | Likely |
| BR-COM-6 `#77` | La generaciÃ³n Word exige saber si la NC es de Proyecto o AuditorÃ­a y usar la plantilla correspondiente. | CÃ³digo exportado | SÃ­ â€” `Informe.GenerarWordNoConformidades` / `PrepararPlantilla` | FALTA â†’ crear mediante access-vba-tdd o prueba de costura sin automatizar Word real | Verified-static |
| BR-COM-7 `#78` | Exportaciones Excel desde listados/seguimiento preservan filtros y columnas de negocio. | Nombres de eventos | Probable â€” eventos `ComandoExportarAExcel_Click` aparecen en formularios de seguimiento | FALTA â†’ crear mediante access-vba-tdd; mapear eventos exactos | Likely |
| BR-COM-8 | La ruta de informe de auditorÃ­a usa selecciÃ³n/constructor de auditorÃ­a, no constructor de NC Proyecto. | Docs de auditorÃ­a + capacidad documental | SÃ­ documentado estÃ¡ticamente â€” `EnsureNCAuditoriaGestionSelected` / `constructor.getNCAuditoria` | Referencia archivada: `Test_AuditGestionForm_ReportConstructorPath_Characterization`; FALTA â†’ reejecutar mediante access-vba-tdd | Verified-static |

### Validaciones
- Asunto obligatorio.
- Al menos un destinatario obligatorio.
- Cuerpo HTML obligatorio en `Correo.Registrar`.
- ParÃ¡metro `p_EsDeProyecto` obligatorio para informes Word.
- Plantilla de informe debe existir en ruta de entorno; si no existe, error bloqueante.
- Contrato de exportaciÃ³n Excel â€” columnas, filtros, orden y formato â€” pendiente de definiciÃ³n/prueba.
- PolÃ­tica de cola/envÃ­o real de correo y BCC/privacidad pendiente de confirmaciÃ³n.

### Transiciones de estado
- `Correo redactado` --(`Registrar`)--> `Correo ordenado para envÃ­o`.
- `Sin plantilla local` --(`PrepararPlantilla`)--> `Plantilla copiada a directorio local de informes`.
- `NC seleccionada` --(`GenerarWordNoConformidades`)--> `Documento Word generado`.

### Casos lÃ­mite y de error
- Formulario de correo abierto sin NC activa o sin `OpenArgs` suficiente se cierra con error.
- Word Automation puede dejar salidas parciales; las pruebas deberÃ­an usar costuras o doble de sistema de ficheros/Word.
- La BCC por defecto contiene un correo fijo y requiere revisiÃ³n de privacidad/operaciÃ³n antes de migraciÃ³n web.

### SeÃ±ales de aceptaciÃ³n / presencia
- El registro en `TbCorreosEnviados` se crea una sola vez con campos completos.
- Las salidas Word usan plantilla de Proyecto/AuditorÃ­a correcta.
- No se marca envÃ­o como `Verified-runtime` sin manifest que pruebe registro y contenido.
- No se marca exportaciÃ³n Excel como `Verified-runtime` sin contrato de columnas/filtros probado.

## Â§3 Mapa de implementaciÃ³n
- **Puntos de entrada de UI**: `Form_FormCorreo`; eventos de informe en `Form_FormNCAuditoriaGestion`; posibles `ComandoExportarAExcel_Click` en formularios de seguimiento.
- **Puntos de entrada de cÃ³digo**: `Correo`, `Informe` (incluye `GenerarWordNoConformidades(p_EsDeProyecto:=No)` para el caso auditorÃ­a), `HTML`, `MÃ³dulo1.EnviarCorreoReactivacionNC`, `CorreoAlAdministrador` como soporte de error. `InformeNCAuditorias` fue retirado como dead-code marker el 2026-06-15 (commit <SHA>) â€” el archivo `src/classes/InformeNCAuditorias.cls` estaba vacÃ­o desde el commit inicial `df3c17a` y ningÃºn path de runtime lo instanciaba; la generaciÃ³n de informe de auditorÃ­a se hace por `Informe.GenerarWordNoConformidades(p_EsDeProyecto:=No)`.
- **Datos afectados**: `TbCorreosEnviados`; rutas/plantillas de `Entorno`; NC Proyecto/AuditorÃ­a y acciones asociadas.
- **Salidas**: filas de correo/orden de envÃ­o, HTML, documentos Word, posibles exportaciones Excel y evidencia documental vinculada.
- **Dependencias e integraciones**: documentos/evidencia, NC Proyecto, NC AuditorÃ­a, acciones/seguimiento, soporte transversal.
- **SincronizaciÃ³n fuenteâ†”binario**: no comprobada; tarea solo documental.
- **ValoraciÃ³n de diseÃ±o**: separar registro de correo de envÃ­o real es razonable. Word Automation y rutas fijas deben encapsularse antes de web; el correo fijo en copia oculta es una deuda a revisar.

## Â§4 Receta de reconstrucciÃ³n
1. Confirmar con producto quÃ© comunicaciones son obligatorias, destinatarios, copias, privacidad y plantilla oficial.
2. Inspeccionar esquema de `TbCorreosEnviados` antes de crear fixtures.
3. Crear pruebas de `Correo.Registrar`: asunto vacÃ­o, destinatarios vacÃ­os, cuerpo vacÃ­o, registro correcto y BCC por defecto si se confirma.
4. Crear costuras para HTML e informe Word sin depender de Word real cuando sea posible.
5. Mapear y probar exportaciones Excel con filtros, columnas, orden y formato si se decide que son parte de la capacidad.
6. Confirmar polÃ­tica de envÃ­o directo frente a cola/orden de envÃ­o, BCC por defecto y privacidad antes de elevar la confianza.

## Â§5 Evidencia y trazabilidad
- **Tests**: no se localizÃ³ manifest dedicado a correo/informes/exportaciones. Hay evidencia adyacente de ruta de informe de auditorÃ­a en `tests/tests.vba.audit-gestion-helper.json` citada por docs de auditorÃ­a.
- **Candidatas para prueba runtime futura**:
  - `tests/tests.vba.audit-gestion-helper.json`: `Test_AuditGestionForm_ReportConstructorPath_Characterization`, `Test_AuditListadoHelper_RowAndReportContracts_RED`.
  - `tests/tests.vba.seguimiento-tareas-helper.json`: `Test_TareasHelper_DeterministicOrder_ExportInput`, `Test_TareasForm_Delegates_FilterPaths`.
  - `tests/tests.vba.cache-e2e.json` / `tests/tests.vba.cache-warmup.json`: `Test_E2E_Cache_PrecalentarSincronizar_LogEvidence_Atomic`.
- **Enlaces cruzados**: ver `docs/capabilities/documents-generated-evidence.md` para evidencia documental; ver `docs/capabilities/nc-auditoria-lifecycle.md` y `docs/features/audit/audit-backend-list-cache.md` para la ruta de informe de auditorÃ­a.

| Elemento | Ref. tracker | VersiÃ³n de staging (UAT) | Estado UAT | Release de producciÃ³n | Fecha en producciÃ³n | Nota |
|---|---|---|---|---|---|---|
| Ruta de informe de auditorÃ­a | Evidencia 2026-06-14 | Pendiente | pending | Pendiente | Pendiente | `Verified-static`: fuente documentada con `EnsureNCAuditoriaGestionSelected` / `constructor.getNCAuditoria`; runtime fresco pendiente. |
| Registro de correos | Pendiente | Pendiente | pending | Pendiente | Pendiente | Falta manifest dedicado. |
| GeneraciÃ³n Word/Excel | Pendiente | Pendiente | pending | Pendiente | Pendiente | Falta contrato de salida; para Excel faltan columnas/filtros/orden/formato. |
| PolÃ­tica send-vs-queue y BCC/privacidad | Pendiente | Pendiente | pending | Pendiente | Pendiente | No elevar por encima de `Likely`/`Intended` sin decisiÃ³n de producto y prueba. |

| SÃ­ntoma | Causa probable | ComprobaciÃ³n (Dysflow) | Ancla del documento |
|---|---|---|---|
| No se registra correo | ValidaciÃ³n o escritura en `TbCorreosEnviados` rota | Crear/rejecutar pruebas de `Correo.Registrar` | BR-COM-1..5 |
| Informe de auditorÃ­a abre dominio incorrecto | Ruta de constructor incorrecta | Reejecutar prueba de audit helper | BR-COM-8 |
| Word/Excel incompleto | Plantilla/ruta/columnas sin contrato | Crear prueba de costura de salida | BR-COM-6..7 |

## Â§6 Notas de migraciÃ³n web

### Â§6.1 Conservar (comportamiento de negocio que debe sobrevivir)
- La obligatoriedad de asunto, destinatario principal, copia y copia oculta antes de registrar un correo (BR-COM-1, BR-COM-2): la API de envÃ­o de la web debe rechazar la orden si falta cualquiera de esos campos, con el mismo orden de validaciÃ³n que `Form_FormCorreo.ComandoEnviarCorreo_Click` aplica hoy.
- La generaciÃ³n del cuerpo del correo en HTML desde la NC activa con `p_ConAcciones:=EnumSino.SÃ­` (BR-COM-3): la plantilla de correo de la web debe seguir incluyendo el bloque de acciones cuando `ConAcciones` venga verdadero, conservando las dos firmas `HTMLNCProyecto` y `HTMLNCAuditoria`.
- El modelo de **orden de envÃ­o** en `TbCorreosEnviados` (no envÃ­o SMTP directo) (BR-COM-4): el sistema de la web debe distinguir entre "orden registrada" y "envÃ­o efectivo", y nunca afirmar que un correo se enviÃ³ solo porque se persistiÃ³.
- La selecciÃ³n de plantilla Word por origen de la NC (Proyecto vs AuditorÃ­a) en `Informe.GenerarWordNoConformidades(p_EsDeProyecto:=)` y `PrepararPlantilla` (BR-COM-6): el generador documental de la web debe seguir exigiendo `p_EsDeProyecto` y rechazar la generaciÃ³n si no se resuelve, igual que el cÃ³digo VBA actual.
- La separaciÃ³n de constructores para informes: `EnsureNCAuditoriaGestionSelected` + `constructor.getNCAuditoria` para auditorÃ­a, jamÃ¡s `constructor.getNCProyecto` (BR-COM-8): la API REST de generaciÃ³n de informes debe mantener dos rutas explÃ­citas con guard de tipo de NC.

### Â§6.2 Transformar (mecanismo legacy que se reformula)
- Sustituir `Form_FormCorreo` y el botÃ³n `ComandoEnviarCorreo_Click` por una pantalla de redacciÃ³n con vista previa del HTML generado y un endpoint `POST /correos/ordenes` que delegue en una capa de aplicaciÃ³n en lugar de un formulario Access.
- Sustituir Word Automation local (`Informe.GenerarWordNoConformidades`) por un servicio server-side que use una plantilla controlada por configuraciÃ³n y devuelva una URL de descarga + entrada de auditorÃ­a inmutable, en lugar de un `.docx` generado en el cliente.
- Reemplazar la BCC por defecto embebida en `Correo.Registrar` (BR-COM-5) por una polÃ­tica de destinatarios por configuraciÃ³n auditada, versionada y revisable por producto antes de promover release.
- Convertir las exportaciones Excel de seguimiento en endpoints `GET /exportaciones/...` con un contrato explÃ­cito de columnas, filtros aplicados y formato, en lugar de eventos `ComandoExportarAExcel_Click` acoplados al formulario de seguimiento.

### Â§6.3 NO copiar (deuda legacy de Access que no debe portarse)
- No portar la direcciÃ³n fija de la copia oculta por defecto (BR-COM-5) como constante en cÃ³digo: moverla a configuraciÃ³n auditada o eliminarla si producto la considera obsoleta.
- No migrar la generaciÃ³n de Word como dependencia del lado cliente: el archivo `.docx` debe ser generado y firmado en servidor, no por la app web del usuario.
- No asumir que la fila en `TbCorreosEnviados` equivale a correo enviado: el estado de la orden y el estado del envÃ­o deben ser columnas separadas en el modelo de la web, con transiciones explÃ­citas.
- No duplicar la lÃ³gica de selecciÃ³n de plantilla de informe (`PrepararPlantilla` resuelve por `p_EsDeProyecto` y por plantilla de entorno) en cada consumidor; exponer un Ãºnico servicio de generaciÃ³n que centralice esa decisiÃ³n.
- No reutilizar `Me.OpenArgs` ni ribbon como contrato de selecciÃ³n de NC: la API web debe recibir un identificador de NC explÃ­cito en la URL o el body, sin parÃ¡metros opacos.

### Â§6.4 Preguntas abiertas al product owner
- Â¿La BCC por defecto (BR-COM-5) sigue siendo una regla vÃ¡lida o es deuda operativa que debe eliminarse antes de migrar? Confirmar con Calidad y privacidad.
- Â¿CuÃ¡l es la plantilla Word oficial de NC de Proyecto y de NC de AuditorÃ­a? Â¿Existe un repositorio versionado o hay que crearlo como parte de la migraciÃ³n? (BR-COM-6)
- Â¿QuÃ© polÃ­tica de retenciÃ³n aplica a los documentos Word generados y a los correos ordenados? Â¿Se conservan tras el cierre de la NC o se purgan? (BR-COM-4, BR-COM-6)
- Â¿El contrato de exportaciÃ³n Excel (BR-COM-7) debe cubrir los mismos seguimientos que hoy existen o se redefinen los reportes en la web? Confirmar lista de seguimientos, columnas y orden antes de la migraciÃ³n.
- Â¿QuiÃ©n es el originador por defecto de un correo cuando el usuario no lo proporciona? Â¿Se sigue derivando del usuario conectado como hace el cÃ³digo actual?
- Â¿El envÃ­o SMTP real lo hace un worker desacoplado, un servicio de la empresa o sigue siendo un Outlook local? Decidir y documentar antes de definir el modelo de estados de la orden.

## Â§7 Registro de confianza
| Hecho | Confianza | Evidencia | Fecha |
|---|---|---|---|
| BR-COM-1 â€” No se puede ordenar correo sin asunto. | Verified-static | `Form_FormCorreo.ComandoEnviarCorreo_Click` y `Correo.Registrar` en `src/forms/Form_FormCorreo.cls` + `src/classes/Correo.cls`; FALTA â†’ crear mediante access-vba-tdd con objeto `Correo` y fixture de `TbCorreosEnviados` | 2026-06-15 |
| BR-COM-2 â€” No se puede ordenar correo sin destinatario principal, copia ni copia oculta. | Verified-static | ValidaciÃ³n en `Form_FormCorreo.cls` y `Correo.Registrar`; FALTA â†’ crear mediante access-vba-tdd | 2026-06-15 |
| BR-COM-3 â€” El cuerpo del correo se genera en HTML desde la NC activa y debe incluir acciones. | Verified-static | `HTMLNCProyecto` / `HTMLNCAuditoria` con `p_ConAcciones:=EnumSino.SÃ­`; FALTA â†’ crear mediante access-vba-tdd sobre HTML | 2026-06-15 |
| BR-COM-4 â€” El correo se registra como orden de envÃ­o en `TbCorreosEnviados`; no se afirma envÃ­o SMTP directo. | Verified-static | `Correo.Registrar` hace `AddNew`; FALTA â†’ crear mediante access-vba-tdd con cardinalidad | 2026-06-15 |
| BR-COM-5 â€” Si no hay BCC, se aÃ±ade una copia oculta por defecto. | Likely | `Correo.Registrar`; FALTA â†’ crear mediante access-vba-tdd; confirmar si sigue siendo regla de negocio vÃ¡lida | 2026-06-15 |
| BR-COM-6 â€” La generaciÃ³n Word exige saber si la NC es de Proyecto o AuditorÃ­a y usa la plantilla correspondiente. | Verified-static | `Informe.GenerarWordNoConformidades` / `PrepararPlantilla`; FALTA â†’ crear mediante access-vba-tdd o prueba de costura sin automatizar Word real | 2026-06-15 |
| BR-COM-7 â€” Exportaciones Excel desde listados/seguimiento preservan filtros y columnas de negocio. | Likely | Eventos `ComandoExportarAExcel_Click` en formularios de seguimiento; FALTA â†’ crear mediante access-vba-tdd; mapear eventos exactos | 2026-06-15 |
| BR-COM-8 â€” La ruta de informe de auditorÃ­a usa selecciÃ³n/constructor de auditorÃ­a, no constructor de NC Proyecto. | Verified-static | `EnsureNCAuditoriaGestionSelected` / `constructor.getNCAuditoria`; referencia archivada `Test_AuditGestionForm_ReportConstructorPath_Characterization`; FALTA â†’ reejecutar mediante access-vba-tdd | 2026-06-15 |
| `Form_FormCorreo` valida asunto y destinatarios antes de registrar. | Verified-static | `src/forms/Form_FormCorreo.cls` | 2026-06-15 |
| `Correo.Registrar` escribe en `TbCorreosEnviados`. | Verified-static | `src/classes/Correo.cls` | 2026-06-15 |
| `Informe` genera documentos Word desde plantillas de Proyecto/AuditorÃ­a. | Verified-static | `src/classes/Informe.cls` | 2026-06-15 |
| Las exportaciones Excel preservan filtros/columnas. | Likely | Nombres de eventos de formularios; sin lectura exhaustiva de implementaciÃ³n | 2026-06-15 |
| La ruta de informe de auditorÃ­a usa `EnsureNCAuditoriaGestionSelected` / `constructor.getNCAuditoria`. | Verified-static | Docs de auditorÃ­a y referencia archivada `Test_AuditGestionForm_ReportConstructorPath_Characterization`; sin reejecuciÃ³n | 2026-06-15 |
| La polÃ­tica de envÃ­o real frente a cola, BCC/privacidad y contrato Excel estÃ¡ cerrada. | Intended | Faltan decisiÃ³n de producto y pruebas | 2026-06-15 |

**âš ï¸ Divergencias (intenciÃ³n SDD â‰  realidad del cÃ³digo)**
- Pendiente de revisiÃ³n: la copia oculta por defecto parece una regla operativa embebida en cÃ³digo, no una intenciÃ³n SDD documentada.
