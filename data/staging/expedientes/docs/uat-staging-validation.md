# UAT staging — EXPEDIENTES validation checklist

Este documento define qué validar en el Access publicado en `staging`. No cubre solo el hardening de backend: esta rama acumula cambios desde `main` en arranque, selección de backend, alta de expedientes, guardado automático, caché transaccional y protección anti-spam.

## Resultado esperado

La publicación es aceptable si:

1. La aplicación arranca en UAT con el backend configurado, sin caer silenciosamente a producción.
2. Las rutas críticas de alta, edición, búsqueda y navegación siguen funcionando.
3. Los cambios guardados se ven de forma consistente al reabrir el expediente y en las listas/cachés derivadas.
4. Las operaciones largas o repetibles no permiten doble ejecución por spam de clicks.
5. Los gaps conocidos quedan aceptados explícitamente como fuera de alcance o se corrigen antes de aprobar UAT.

## Alcance publicado en staging

| Área | Qué cambió | Estado de cobertura |
|---|---|---|
| Backend selection hardening | Resolución explícita de `BackendActivo`; rechazo de claves vacías, inválidas, rutas vacías o archivos inexistentes. | Tests automatizados presentes. UAT manual obligatoria. |
| Startup/anexos preflight | Validación de infraestructura/rutas de anexos antes de iniciar flujos que dependen de ellas. | Tests automatizados parciales. UAT manual obligatoria. |
| Alta y post-alta refresh | Ajustes en alta por tipo, ordinal, comportamiento modal y refresco de lista tras cerrar alta. | Sin cobertura automatizada suficiente. UAT manual obligatoria. |
| Guardado automático por pestañas | Separación de flujos alta/edición y persistencia parcial al cambiar pestañas/cerrar expediente. | Sin tests manifestados. UAT manual obligatoria. |
| frmBusy / anti-spam | Protección de botones y popup de progreso para operaciones lentas o repetibles. | Sin tests manifestados. UAT manual obligatoria. |
| Caché transaccional | Seam de fallo de caché y endurecimiento parcial de escrituras/caché bajo transacción. | Tests parciales. Hay gaps conocidos. |
| E2E integration | Base de `OrdinalE2E`: propiedad en dominio, persistencia en `Registrar`, lectura por `constructor.getExpediente`, utilidades de máximo y unicidad. | Tests automatizados presentes para el slice ordinal. |
| E2E exporter | Exportación automática/manual, JSON con `IDExportacion`/batch, histórico, trazabilidad de archivo y avance de hash exportado. | Tests targeted de servicio presentes. UAT manual obligatoria para validar UX/evidencia en staging. |

## Evidencia técnica ya ejecutada

| Evidencia | Resultado |
|---|---|
| Rama publicada | `staging` |
| Binario Access | `Expedientes.accdb` tomado desde la rama donante `develop` y sincronizado para staging |
| Config dysflow local | `.dysflow/project.json` con `Expedientes.accdb` + `Expedientes_datos.accdb` |
| Suite VBA | `tests/tests.vba.json` |
| Resultado suite | Tests targeted por área ejecutados por procedimiento individual; ejecución completa por manifest puede superar timeout MCP según carga Access/COM |
| Backend resolver | 6 tests OK |
| Anexos/startup preflight | 2 tests OK |
| Caché transaccional | 5 tests OK, cobertura parcial |
| E2E OrdinalE2E/domain/utils/mapper | 7 tests OK ejecutados por procedimiento individual tras compilación manual |
| E2E export traceability | 5/5 tests OK tras compilación manual; `verify_code` OK para `E2EExportService`, `Test_E2EExportService` y `TestFixtures` |

Procedimientos cubiertos por la suite actual:

- `Test_ValidarCarpetaEscribible_DetectaRutaNoAlcanzable`
- `Test_ValidarCarpetaEscribible_AceptaRutaTemporal`
- `Test_CacheFailureSeam_InactiveByDefault`
- `Test_CacheFailureSeam_ActivatedProducesControlledFailure`
- `Test_CacheFailureSeam_ResetClearsFailure`
- `Test_ActualizarNCsContratistas_RolledBackWhenCallerRollsBack`
- `Test_ActualizarNCs_UsesCallerTransactionAndRollsBack`
- `Test_BackendResolver_ResolveValidConfiguredBackend`
- `Test_BackendResolver_RejectInvalidBackendKey`
- `Test_BackendResolver_RejectEmptyBackendKey`
- `Test_BackendResolver_RejectEmptyBackendPath`
- `Test_BackendResolver_RejectMissingBackendFile`
- `Test_BackendResolver_NoFallbackToProdWhenKeyInvalid`
- `Test_Expediente_OrdinalE2E_DefaultIsNull`
- `Test_Expediente_OrdinalE2E_LetAcceptsNull`
- `Test_Expediente_OrdinalE2E_LetAcceptsNumericAndReturnsLong`
- `Test_Expediente_SetPropiedad_OrdinalE2E_MapsCorrectly`
- `Test_ExpedienteE2EUtils_GetMaxOrdinalE2E_ReturnsHighestIgnoringNull`
- `Test_ExpedienteE2EUtils_ValidateOrdinalE2EUniqueness_Cases`
- `Test_Constructor_getExpediente_OrdinalE2E_NullDbMapsToNull`
- `Test_E2EExport_AutomaticSelectsPendingOnly`
- `Test_E2EExport_ManualUsesRightSideSelectionOnly`
- `Test_E2EExport_JsonMetaIncludesExportId`
- `Test_E2EExport_SuccessRecordsHistoryAndAdvancesHash`
- `Test_E2EExport_FileFailureDoesNotMarkExported`

## Pruebas E2E manuales para UAT

> Hacer estas pruebas sobre el entorno UAT/staging, nunca sobre producción.

### 1. Arranque y backend activo

- [ ] Abrir el `Expedientes.accdb` publicado en UAT.
- [ ] Confirmar que la aplicación arranca sin errores de configuración.
- [ ] Confirmar que se abre el menú principal esperado para el usuario.
- [ ] Confirmar que el indicador/caption de backend, si aparece, muestra el backend esperado para UAT.
- [ ] Buscar y abrir un expediente existente sin errores de conexión.

### 2. No fallback silencioso a producción

Esta prueba solo debe hacerse si se puede modificar una copia/control de `TbConfiguracionBackends` en UAT y restaurarla después.

- [ ] Guardar evidencia de la configuración válida actual.
- [ ] Cambiar temporalmente `BackendActivo` a una clave inválida, vacía o no habilitada.
- [ ] Abrir la aplicación.
- [ ] Confirmar que la aplicación informa/bloquea el error de configuración.
- [ ] Confirmar que no continúa operando contra producción por defecto.
- [ ] Restaurar la configuración válida.
- [ ] Volver a abrir la aplicación y confirmar arranque normal.

### 3. Startup/anexos preflight

- [ ] Con configuración válida, abrir un expediente con anexos/documentación.
- [ ] Confirmar que no hay bloqueo de infraestructura ni errores de ruta.
- [ ] Si existe entorno seguro para probarlo, configurar temporalmente una ruta de anexos inexistente o no escribible.
- [ ] Confirmar que la aplicación bloquea o informa el problema antes de permitir operación dependiente de esa ruta.
- [ ] Restaurar la ruta válida.

### 4. Alta de expediente y refresco posterior

- [ ] Desde gestión de expedientes, lanzar alta de expediente individual.
- [ ] Lanzar alta AM/C, lote y derivado/basado si aplican al perfil de pruebas.
- [ ] Confirmar que el filtro de tipo muestra opciones correctas.
- [ ] Confirmar que el ordinal se muestra u oculta correctamente según tipo y padre.
- [ ] Guardar un alta de prueba.
- [ ] Confirmar que el formulario se cierra o navega según el flujo esperado.
- [ ] Confirmar que el nuevo expediente aparece en la lista sin tener que reiniciar Access ni cambiar manualmente filtros.
- [ ] Abrir el expediente creado y confirmar que aterriza en la pestaña esperada.

### 5. Guardado automático por pestañas

- [ ] Abrir un expediente existente en modo edición.
- [ ] Modificar un campo menor en pestaña General.
- [ ] Cambiar a Fechas, cerrar y reabrir el expediente.
- [ ] Confirmar que el cambio de General persiste.
- [ ] Modificar un campo menor en Fechas.
- [ ] Cambiar de pestaña o cerrar/reabrir.
- [ ] Confirmar que el cambio de Fechas persiste.
- [ ] Repetir con Hitos y Modificados si el expediente de prueba tiene datos adecuados.
- [ ] Abrir un flujo de Alta y confirmar que no aparece guardado automático no deseado antes del registro final.

### 6. frmBusy y anti-spam

- [ ] En botones protegidos, hacer doble click rápido en Alta/Guardar/Actualizar/Detalle/Reiniciar parámetros cuando aplique.
- [ ] Confirmar que solo se ejecuta una operación.
- [ ] Confirmar que los botones quedan bloqueados durante la operación y vuelven a habilitarse al terminar.
- [ ] Confirmar que `frmBusy` se abre y se cierra correctamente en operaciones largas.
- [ ] Provocar o simular un error controlado si el entorno lo permite y confirmar que no queda popup huérfano ni botón bloqueado.
- [ ] Confirmar que el caption/estado visual se restaura tras éxito y error.

### 7. Caché transaccional y consistencia de listas

Usar expedientes de prueba. La validación funcional debe comprobar que source tables y vistas/listas derivadas quedan coherentes.

- [ ] Editar entidades/contratistas de un expediente.
- [ ] Guardar, cerrar y reabrir.
- [ ] Confirmar que los datos persisten.
- [ ] Confirmar que la lista/búsqueda/cache muestra los mismos datos actualizados.
- [ ] Editar responsables, lugares de ejecución, PECAL, RAC y suministradores si el expediente de prueba los tiene.
- [ ] Confirmar que no quedan datos stale en `FormExpedientesGestion` ni en vistas derivadas.
- [ ] Si una operación falla, confirmar que no queda medio guardada: ni source actualizado sin cache, ni cache actualizada sin source.

### 8. Base E2E — OrdinalE2E

- [ ] Abrir un expediente de prueba sin `OrdinalE2E` asignado y confirmar que no muestra valor ordinal E2E accidental.
- [ ] Guardar/reabrir el expediente y confirmar que `OrdinalE2E` vacío se conserva como vacío/Null.
- [ ] Asignar un `OrdinalE2E` positivo en entorno controlado si existe UI/procedimiento disponible para ello.
- [ ] Confirmar que el valor positivo persiste al guardar y reabrir.
- [ ] Confirmar que no se permite duplicar el ordinal E2E cuando exista flujo UI de integración.

### 9. Plataforma E2E — estados de exportación

Preparar tres expedientes de prueba con `OrdinalE2E` asignado:

| Caso | Estado esperado | Cómo reconocerlo |
|---|---|---|
| Nunca exportado | Pendiente de exportación | `HashUltimaExportacion` vacío/Null. |
| Exportado sin cambios | No pendiente | `HashActual = HashUltimaExportacion`. |
| Cambiado desde última exportación | Pendiente/cambiado | `HashActual <> HashUltimaExportacion`. |

- [ ] Abrir la pantalla/lista de gestión E2E disponible en staging.
- [ ] Aplicar filtro de pendientes y confirmar que incluye nunca exportados y cambiados.
- [ ] Confirmar que el expediente exportado sin cambios no queda seleccionado automáticamente como pendiente.
- [ ] Guardar captura de la lista filtrada y los IDs/ordinales de prueba usados.

### 10. Batch manual — doble lista, filtros y multi-selección

- [ ] Abrir el flujo manual de selección E2E.
- [ ] Buscar/filtrar expedientes por estado, texto o criterio disponible.
- [ ] Mover al menos dos expedientes desde la lista izquierda a la derecha.
- [ ] Quitar uno de la lista derecha y confirmar que vuelve a la izquierda si todavía cumple el filtro.
- [ ] Confirmar que solo la lista derecha queda como selección manual del batch.
- [ ] Guardar captura de lista izquierda/derecha antes de exportar.

### 11. Export bajo demanda y JSON generado

- [ ] Ejecutar export bajo demanda desde el batch manual de prueba.
- [ ] Confirmar que se genera un archivo JSON en la carpeta esperada de UAT.
- [ ] Abrir el JSON y validar que contiene raíz `meta` y `data`.
- [ ] Confirmar que `meta` contiene `IDExportacion` o `IDBatch` y que coincide con el batch registrado.
- [ ] Confirmar que los expedientes exportados son exactamente los seleccionados en la lista derecha.
- [ ] Guardar como evidencia el archivo JSON generado.

### 12. Histórico, trazabilidad y rollback de fallo de archivo

- [ ] Revisar el histórico de exportación del batch.
- [ ] Confirmar una fila de cabecera de batch con usuario, sesión/fecha, estado y totales.
- [ ] Confirmar detalle por expediente exportado, incluyendo hash exportado y estado exitoso.
- [ ] Confirmar que, tras export OK, `HashUltimaExportacion` queda igual a `HashActual` para los expedientes exportados.
- [ ] Si el entorno permite simular carpeta inválida/no escribible, ejecutar un intento de export fallido.
- [ ] Confirmar que el fallo queda registrado como error y que ningún expediente se marca como exportado.
- [ ] Guardar capturas del histórico y, si aplica, del error controlado.

## Evidencia esperada para cerrar UAT E2E

| Evidencia | Mínimo aceptable |
|---|---|
| Capturas de estados | Pendiente nunca exportado, exportado sin cambios y cambiado desde última exportación. |
| Captura batch manual | Lista derecha con selección final y lista izquierda filtrada. |
| Archivo JSON | Fichero generado en UAT con `meta.IDExportacion` o `meta.IDBatch` y `data`. |
| Histórico | Captura o export de cabecera/detalle del batch. |
| Trazabilidad hash | Evidencia de `HashUltimaExportacion = HashActual` tras éxito. |
| Fallo de archivo | Si se ejecuta, evidencia de error sin marcar expedientes como exportados. |

## Riesgos, rollback y validaciones por PR de la cadena E2E

| PR / scope | Riesgo principal | Validación | Rollback |
|---|---|---|---|
| PR1 — base export JSON / hash | Cambiar contrato JSON o cálculo de hash canónico. | Tests de exporter/hash y revisión de JSON generado. | Revertir módulos de exporter/hash y binario Access asociado. |
| PR2 — gestión, selección y destino | Selección manual o destino por usuario incorrectos. | Tests de picker/session/destination y UAT de doble lista. | Revertir módulos de gestión/destino y limpiar filas de config de prueba. |
| PR3 — export traceability | Marcar como exportado sin archivo correcto o perder histórico. | 5 tests targeted de export, `verify_code`, UAT de JSON/histórico/hash. | Revertir `E2EExportService`, tests y binario; conservar backup del backend antes de tocar hash/histórico. |
| Transversal — UAT staging | Evidencia incompleta para aprobar publicación. | Este checklist completo con capturas y JSON adjunto. | No aprobar UAT, mantener staging sin promover a main. |

## Gaps técnicos conocidos antes de aprobar

Estos puntos deben decidirse explícitamente: corregir antes de UAT, aceptar como riesgo, o declararlos fuera de alcance.

| Gap | Impacto | Decisión UAT |
|---|---|---|
| Caché transaccional tiene cobertura parcial: seam + contratistas/NCs + `ActualizarNCs`. | Riesgo de inconsistencias en otros mutadores de cache. | [ ] Corregir / [ ] Aceptar riesgo / [ ] Fuera de alcance |
| Faltan rollback tests por slice para lugares, PECAL, RAC, responsables, suministradores, comerciales e hitos. | No hay prueba automática de atomicidad en todos los caminos. | [ ] Corregir / [ ] Aceptar riesgo / [ ] Fuera de alcance |
| Alta/post-alta, guardado automático y frmBusy/anti-spam dependen de validación manual. | Regresión UI posible sin test automatizado. | [ ] Corregir / [ ] Aceptar riesgo / [ ] Fuera de alcance |
| UX/botones finales de E2E pueden requerir validación manual adicional según pantalla disponible en staging. | Riesgo de regresión UI aunque el servicio de export tenga tests targeted. | [ ] Corregir / [ ] Aceptar riesgo / [ ] Fuera de alcance |
| Tareas OpenSpec de backend/cache pueden estar desactualizadas respecto al código. | Riesgo de trazabilidad incompleta. | [ ] Reconciliar / [ ] Aceptar |

## Tests automatizados recomendados que faltan

No bloquean por sí solos si UAT acepta validación manual, pero son los próximos tests que conviene agregar.

### Backend/startup

- `LeeConfiguracionLocal` limpia estado/caché si falla la resolución.
- `getdb()` devuelve `Nothing` o propaga error controlado cuando el resolver falla.
- Startup/splash no muestra backend PROD implícito después de error de configuración.

### Caché transaccional

- Forced cache failure rollback para `EliminarComercial` y `EliminarHito`.
- Forced cache failure rollback para `EditarEnvioCorreoResponsable` y `EliminarResponsable`.
- Forced cache failure rollback para `EliminarLugarEjecucion`, `EliminarPECAL`, `EliminarRAC` y `EliminarSuministrador`.
- Success-path equivalence: source tables y `TbExpedientesConEntidades` quedan coherentes tras guardar.

### Alta/guardado/anti-spam

- Parsing de OpenArgs de `FormExpedienteAltaTipo`: semicolon, single tipo, tipo vacío.
- Ordinal visible/oculto según tipo y padre.
- Alta refresca `FormExpedientesGestion` tras cerrar.
- Helpers de guardado automático: dirty detection, snapshot y no auto-save en ALTA.
- Helpers anti-spam: reentrancy, caption restore y desbloqueo en error.

## Criterio de aceptación UAT

Marcar la UAT como aprobada solo cuando:

- [ ] Arranque/backend activo validado.
- [ ] No fallback silencioso a producción validado o justificado como no ejecutable en entorno seguro.
- [ ] Anexos/preflight validado.
- [ ] Alta y refresco post-alta validados.
- [ ] Guardado automático por pestañas validado.
- [ ] frmBusy/anti-spam validado.
- [ ] Caché/listas consistentes tras edición de entidades relacionadas.
- [ ] Estados E2E, batch manual, export bajo demanda, JSON, histórico y trazabilidad validados.
- [ ] Evidencia UAT E2E adjunta: capturas, JSON generado y registros de histórico.
- [ ] Todos los gaps técnicos tienen decisión explícita.
- [ ] No se detectan regresiones bloqueantes en navegación, guardado o búsqueda.

## Fuera de alcance salvo decisión contraria

- Rediseño de pantallas.
- Cobertura automatizada completa de todos los formularios.
- Pruebas destructivas sobre producción.
