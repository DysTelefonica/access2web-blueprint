# Capacidad: expedientes, riesgos y responsables

## Â§0 Identidad
- **ID de capacidad**: `CAP-EXP-RISK-RESP`
- **Tier**: standard
- **Estado**: active / inventario documental inicial
- **Source**: reverse-engineered
- **Responsable / autoridad de producto**: Pendiente de confirmaciÃ³n â€” Calidad / gestiÃ³n de proyectos
- **Ãšltima verificaciÃ³n**: 2026-06-15 mediante inspecciÃ³n estÃ¡tica; no se ejecutÃ³ Dysflow/Access
- **Confianza global**: mixta â€” mayoritariamente `Verified-static` y `Likely`

## Â§1 IntenciÃ³n de negocio
- **PropÃ³sito**: Permitir que las no conformidades se vinculen a expedientes, responsables de calidad/tÃ©cnicos/jefes de proyecto y riesgos asociados.
- **Usuarios / perfiles**: Calidad, responsables de proyecto, tÃ©cnicos y usuarios que filtran o asignan trabajo por expediente/responsable.
- **Problema que resuelve**: Sin esta capacidad, las NC pierden contexto de contrato/proyecto, propietario y riesgo, y los listados/seguimientos pueden mostrar trabajo no trazable.
- **Valor de negocio / por quÃ© existe**: Aporta contexto contractual, responsabilidad operativa y visibilidad de riesgo para priorizar acciones y cierre.
- **No-objetivos**: No documenta el ciclo de vida completo de NC Proyecto/AuditorÃ­a ni todos los indicadores.
- **Origen de la intenciÃ³n**: Inferido desde cÃ³digo exportado y specs de seguimiento/listado; intenciÃ³n de producto pendiente.
- **Referencia de tracker de origen**: Issue #67; OpenSpec `seguimiento-tareas-helper` para filtros de `IDExpediente` y responsables.

## Â§2 Contrato de comportamiento

### Escenarios (Dado / Cuando / Entonces)
- **DADO** que un usuario busca expedientes **CUANDO** filtra por palabra clave o responsable de calidad **ENTONCES** se muestra una lista con `IDExp`, `Cod_Exp`, `NemotÃ©cnico` y `TÃ­tulo`, y la selecciÃ³n emite un `Expediente`.
- **DADO** que una NC o tarea estÃ¡ vinculada a un expediente **CUANDO** se filtra seguimiento por `IDExpediente` **ENTONCES** solo aparecen los elementos del expediente seleccionado.
- **DADO** que un expediente tiene responsables/jurÃ­dicas/riesgos asociados **CUANDO** se carga su objeto de dominio **ENTONCES** las relaciones se resuelven desde `constructor` y no se inventan valores por defecto.
- **DADO** que una NC Proyecto tiene riesgos relacionados **CUANDO** se usa la ruta cache-first **ENTONCES** un resultado cargado-vacÃ­o es vÃ¡lido y no debe forzar fallback a backend.

### Reglas de negocio
| ID regla | Enunciado (pretendido) | Autoridad | Â¿Aplicada en cÃ³digo? | Prueba | Confianza |
|---|---|---|---|---|---|
| BR-EXP-1 | La bÃºsqueda de expedientes filtra por palabra clave y responsable de calidad antes de seleccionar. | CÃ³digo exportado | SÃ­ â€” `Form_FormExpedientesBusqueda.Filtrar`, `constructor.getExpedientesBusqueda` | FALTA â†’ author via access-vba-tdd con fixtures de expedientes y responsables | Verified-static |
| BR-EXP-2 | La selecciÃ³n de expediente solo emite evento si hay un expediente cargado; si no hay selecciÃ³n, no fuerza un objeto invÃ¡lido. | CÃ³digo exportado | SÃ­ â€” `ComandoElegir_Click`, `ListaFiltrados_Click` | FALTA â†’ author via access-vba-tdd | Verified-static |
| BR-EXP-3 | Los responsables de calidad del combo proceden de `m_ObjEntorno.ColUsuariosCalidad`. | CÃ³digo exportado | SÃ­ â€” `EstablecerComboResponsablesCalidad` | FALTA â†’ author via access-vba-tdd | Verified-static |
| BR-EXP-4 | `Expediente.TextoExpediente` prioriza `Nemotecnico (CodExp)` y cae a `CodExp` si falta nemotÃ©cnico. | Tests `tests/tests.vba.cap-exp.json` | SÃ­ â€” `Test_EXP_TextoExpediente_NemotecnicoYCodExp_FormateaConParentesis_Atomic`, `..._SoloNemotecnico_FormateaSinParentesis_Atomic`, `..._SoloCodExp_FormateaSinNemotecnico_Atomic` 3/3 PASS contra staging HEAD | Cubierto. Cache memoization cubierto por `..._CacheMemoization_ReutilizaCacheEnSegundaLectura_Atomic` (BR-EXP-5) | Verified-runtime |
| BR-EXP-5 | Un expediente puede exponer jurÃ­dicas, responsables, responsable de calidad, jefe de proyecto y riesgos asociados. | Tests `tests/tests.vba.cap-exp.json` | SÃ­ â€” `Test_EXP_Expediente_ExponePropiedades_PropertiesRoundTrip_Atomic` valida 13/13 propiedades round-trip (`IDExpediente`, `Nemotecnico`, `CodExp`, `CodExpLargo`, `Titulo`, `Estado`, `CodProyecto`, `IDResponsableCalidad`, `IDUsuarioCreacion`, `IDUsuarioUltimoCambio`, `Ambito`, `Tipo`, `NPedido`) | Cubierto. Cache memoization cubierta por `..._CacheMemoization_ReutilizaCacheEnSegundaLectura_Atomic` | Verified-runtime |
| BR-EXP-6 | Los riesgos asociados a NC Proyecto deben leerse con semÃ¡ntica cache-first cuando la cachÃ© estÃ¡ cargada. | Feature `trust-ncproyecto-cache-hits` | SÃ­ segÃºn docs de feature | Evidencia runtime mÃ¡s cercana: `tests/tests.vba.cache-e2e.json` 7/7 PASS (2026-06-14, staging `20b71f64`), registrada en `docs/features/cache-management/trust-ncproyecto-cache-hits.md` â€” no reejecutado en esta sesiÃ³n documental; no afirmar runtime para 2026-06-15 | Verified-static |
| BR-EXP-7 `#81` | El ciclo de vida propio de riesgos â€” aceptaciÃ³n, mitigaciÃ³n, contingencia, materializaciÃ³n, retirada, cierre y retipificaciÃ³n â€” estÃ¡ definido y probado. | Producto pendiente | Desconocido | FALTA â†’ author via access-vba-tdd tras confirmar estados | Intended |

### Validaciones
- No seleccionar expediente sin `IDExpediente`.
- No afirmar responsable/jefe/proveedor/jurÃ­dica si el constructor no devuelve objeto.
- Tratar los valores de riesgo como dominio pendiente: muchas propiedades existen, pero el flujo de aprobaciÃ³n/rechazo no estÃ¡ probado.

### Transiciones de estado
- `Sin filtro` --(`ComandoActualizar`)--> `Lista de expedientes recargada`.
- `Lista con fila seleccionada` --(`ComandoElegir`)--> `Expediente seleccionado emitido`.
- `Riesgo abierto` --(`Aceptar/mitigar/materializar/retirar/cerrar/retipificar`)--> `Estado de riesgo actualizado` â€” pendiente de contrato y pruebas.

### Casos lÃ­mite y de error
- `constructor.getExpediente` o `constructor.getExpedientesBusqueda` con error debe bloquear selecciÃ³n.
- Si `ColUsuariosCalidad` falla, el combo de responsables no debe mostrar datos engaÃ±osos.
- Riesgo cache-first cargado-vacÃ­o es distinto de fallo de cachÃ©.

### SeÃ±ales de aceptaciÃ³n / presencia
- La UI de bÃºsqueda permite filtrar por palabra clave y responsable, seleccionar y cerrar el formulario.
- Los manifests de expediente/responsable/riesgo pasan con fixtures sandbox, no con datos existentes.
- El comportamiento de riesgo no se marca `Verified-runtime` hasta tener pruebas dedicadas.

## Â§3 Mapa de implementaciÃ³n
- **Puntos de entrada de UI**: `Form_FormExpedientesBusqueda`; consumidores en seguimiento/listados de Proyecto; `Form_formRiesgosSeleccion` como superficie de selecciÃ³n probable.
- **Puntos de entrada de cÃ³digo**: `Expediente`, `ExpedienteResponsable`, `Riesgo`, `RiesgoServicio`, `RiesgoRepositorio`, `NCProyectoSeguimientoTareasListadoHelper`, `constructor.getExpediente*`, `constructor.getRiesgosDeExpediente`.
- **Datos afectados**: `TbExpedientes`, `TbExpedientesResponsables`, `TbUsuariosAplicaciones`, `TbRiesgosNC` y tablas de vÃ­nculo exactas pendientes de esquema.
- **Salidas**: filtros de listas, contexto de NC, asignaciÃ³n de responsables, columnas de informe/listado.
- **Dependencias e integraciones**: NC Proyecto lifecycle, acciones/seguimiento, indicadores, soporte transversal de cachÃ©.
- **SincronizaciÃ³n fuenteâ†”binario**: no comprobada; tarea solo documental.
- **ValoraciÃ³n de diseÃ±o**: los objetos de dominio existen y son Ãºtiles para migraciÃ³n, pero el comportamiento de riesgo y responsables sigue demasiado inferido sin pruebas de negocio.

## Â§4 Receta de reconstrucciÃ³n
1. Confirmar con producto quÃ© campos de expediente, responsable, jurÃ­dica y riesgo son obligatorios para NC.
2. Inspeccionar esquema real antes de sembrar fixtures (`TbExpedientes`, usuarios, responsables, riesgos y vÃ­nculos).
3. Crear pruebas de bÃºsqueda/selecciÃ³n de expedientes, carga de responsables y propiedades de dominio.
4. Crear pruebas de ciclo de vida de riesgos solo despuÃ©s de definir estados y permisos.
5. Si se modifica VBA: importar con Dysflow; el usuario compila manualmente; despuÃ©s ejecutar `dysflow.test_vba`.

## Â§5 Evidencia y trazabilidad
- **Tests**: evidencia adyacente en `tests/tests.vba.seguimiento-tareas-helper.json` (9/9 PASS, 2026-06-15) â€” cobertura adyacente a BR-EXP-1: `IDExpediente` filter parity â€” y `tests/tests.vba.cache-e2e.json`; no hay manifest dedicado de expediente/riesgo.

| Elemento | Ref. tracker | VersiÃ³n de staging (UAT) | Estado UAT | Release de producciÃ³n | Fecha en producciÃ³n | Nota |
|---|---|---|---|---|---|---|
| BÃºsqueda/selecciÃ³n de expedientes | Pendiente | Pendiente | pending | Pendiente | Pendiente | Falta prueba dedicada. |
| Riesgos asociados cache-first | Issue #39 / #67 | Pendiente | pending | Pendiente | Pendiente | Evidencia de feature; falta reejecuciÃ³n. |
| Ciclo de vida propio de riesgos | Pendiente | Pendiente | pending | Pendiente | Pendiente | Contrato de producto pendiente. |

| SÃ­ntoma | Causa probable | ComprobaciÃ³n (Dysflow) | Ancla del documento |
|---|---|---|---|
| No se puede seleccionar expediente | RegresiÃ³n de filtro/constructor/evento | Crear manifest de bÃºsqueda de expedientes | BR-EXP-1..3 |
| Seguimiento muestra tareas de otro expediente | Filtro `IDExpediente` roto | Reejecutar/crear prueba de seguimiento por expediente | BR-EXP-1 |
| Riesgos ausentes u obsoletos | RegresiÃ³n cache-first o falta de vÃ­nculo | Reejecutar cache-e2e + prueba dedicada de riesgo | BR-EXP-6..7 |

## Â§6 Notas de migraciÃ³n web

### Â§6.1 Conservar (comportamiento de negocio que debe sobrevivir)
- El filtrado de la bÃºsqueda de expedientes por palabra clave y responsable de calidad antes de seleccionar (BR-EXP-1): la web debe seguir exigiendo al menos uno de los dos filtros, o el `IDExpediente` directo, antes de devolver candidatos.
- La selecciÃ³n de expediente solo se completa cuando hay un expediente cargado con `IDExpediente` vÃ¡lido (BR-EXP-2): la API de selecciÃ³n debe rechazar intentos de "elegir" cuando no hay fila seleccionada, replicando el `ComandoElegir_Click` y `ListaFiltrados_Click` del VBA.
- Los responsables de calidad del combo de bÃºsqueda proceden siempre de `m_ObjEntorno.ColUsuariosCalidad` (BR-EXP-3): la web no debe aceptar responsables arbitrarios; debe consultar un endpoint de "responsables de calidad" y limitar la bÃºsqueda a ese dominio.
- `Expediente.TextoExpediente` prioriza `Nemotecnico (CodExp)` y cae a `CodExp` si falta nemotÃ©cnico (BR-EXP-4): la API REST debe mantener esa regla de presentaciÃ³n y nunca devolver un literal vacÃ­o.
- El expediente expone sus `Juridicas`, `Responsables`, `RESPONSABLECALIDAD`, `JefeProyecto` y `Riesgos` asociados (BR-EXP-5): la web debe serializar el expediente como agregado con sus vÃ­nculos, no como entidad aislada.
- La lectura de riesgos de NC Proyecto con semÃ¡ntica cache-first cuando la cachÃ© estÃ¡ cargada (BR-EXP-6): si la cachÃ© tiene el expediente/riesgo, se responde desde ella; cargado-vacÃ­o es vÃ¡lido, no es fallback a backend.

### Â§6.2 Transformar (mecanismo legacy que se reformula)
- Sustituir `Form_FormExpedientesBusqueda` por un endpoint REST `GET /expedientes?keyword=...&responsableCalidad=...` con paginaciÃ³n, retornando `{IDExp, Cod_Exp, NemotÃ©cnico, TÃ­tulo}` y dejando que la UI pinte.
- Convertir la combinaciÃ³n `Expediente` + `ExpedienteResponsable` en un agregado de dominio, con `constructor.getExpediente(IDExp)` retornando el expediente hidratado con todas sus relaciones.
- Reemplazar el acceso a `m_ObjEntorno.ColUsuariosCalidad` por una API de catÃ¡logo de responsables de calidad, versionada y cacheable, no por un global mutado en runtime.
- Mover el ciclo de vida de riesgos a una mÃ¡quina de estados explÃ­cita con comandos (`Aceptar`, `Mitigar`, `Materializar`, `Retirar`, `Cerrar`, `Retipificar`) y eventos versionados, en lugar de un objeto `Riesgo` con propiedades dispersas.
- Sustituir `RiesgoServicio` y `RiesgoRepositorio` por un servicio de dominio que reciba un comando, valide el estado origen, y devuelva el nuevo estado, no por un objeto con setters mÃºltiples.

### Â§6.3 NO copiar (deuda legacy de Access que no debe portarse)
- No portar la dependencia de `m_ObjEntorno` y `m_ObjUsuarioConectado` como estado compartido implÃ­cito: la web debe inyectar el contexto del usuario al servicio, no leer de globals.
- No duplicar la lÃ³gica de "quÃ© es un responsable de calidad" en cada formulario: la web debe tener un Ãºnico servicio de "responsables" y un Ãºnico catÃ¡logo de roles.
- No usar la selecciÃ³n de expediente por `OpenArgs` o un string opaco: la API web debe recibir `IDExpediente` (entero) en la URL.
- No migrar el patrÃ³n "consulta viva a backend si la cachÃ© estÃ¡ vacÃ­a" como ruta normal: en la web, cargado-vacÃ­o se responde vacÃ­o y se permite reintento explÃ­cito, no fallback silencioso.
- No usar `Form_formRiesgosSeleccion` como Ãºnica puerta de entrada al riesgo: la API web debe permitir consultar riesgos por `IDExpediente` sin necesidad de un formulario de selecciÃ³n intermedio.

### Â§6.4 Preguntas abiertas al product owner
- Â¿CuÃ¡les son los estados canÃ³nicos del ciclo de vida de riesgos? (BR-EXP-7) Hoy se mencionan aceptaciÃ³n, mitigaciÃ³n, contingencia, materializaciÃ³n, retirada, cierre y retipificaciÃ³n â€” Â¿estÃ¡n todos o hay otros?
- Â¿Los riesgos asociados a NC Proyecto son siempre los mismos que los del expediente, o pueden existir riesgos solo de NC sin expediente? Confirmar cardinalidad.
- Â¿Un expediente puede tener mÃ¡s de un responsable de calidad o es Ãºnico? Hoy `ColUsuariosCalidad` parece ser colecciÃ³n; Â¿la bÃºsqueda debe permitir varios? (BR-EXP-3)
- Â¿La relaciÃ³n `Riesgo â†’ NC` es muchos-a-muchos o un riesgo estÃ¡ atado a una sola NC? (BR-EXP-5, BR-EXP-6)
- Â¿`NemotÃ©cnico` es obligatorio para todos los expedientes o puede estar vacÃ­o? Si estÃ¡ vacÃ­o, Â¿se debe permitir crear la NC asociada? (BR-EXP-4)
- Â¿La bÃºsqueda por palabra clave aplica a `NemotÃ©cnico`, `Cod_Exp` y `TÃ­tulo`, o solo a uno de ellos? Confirmar antes de definir el endpoint.

## Â§7 Registro de confianza
| Hecho | Confianza | Evidencia | Fecha |
|---|---|---|---|
| BR-EXP-1 â€” La bÃºsqueda de expedientes filtra por palabra clave y responsable de calidad antes de seleccionar. | Verified-static | `Form_FormExpedientesBusqueda.Filtrar`, `constructor.getExpedientesBusqueda`; FALTA â†’ author via access-vba-tdd con fixtures de expedientes y responsables | 2026-06-15 |
| BR-EXP-2 â€” La selecciÃ³n de expediente solo emite evento si hay un expediente cargado; si no hay selecciÃ³n, no fuerza un objeto invÃ¡lido. | Verified-static | `ComandoElegir_Click`, `ListaFiltrados_Click`; FALTA â†’ author via access-vba-tdd | 2026-06-15 |
| BR-EXP-3 â€” Los responsables de calidad del combo proceden de `m_ObjEntorno.ColUsuariosCalidad`. | Verified-static | `EstablecerComboResponsablesCalidad`; FALTA â†’ author via access-vba-tdd | 2026-06-15 |
| BR-EXP-4 â€” `Expediente.TextoExpediente` prioriza `Nemotecnico (CodExp)` y cae a `CodExp` si falta nemotÃ©cnico. | Verified-static | `Expediente.TextoExpediente`; FALTA â†’ author via access-vba-tdd unitario | 2026-06-15 |
| BR-EXP-5 â€” Un expediente puede exponer jurÃ­dicas, responsables, responsable de calidad, jefe de proyecto y riesgos asociados. | Verified-static | Propiedades `Juridicas`, `Responsables`, `RESPONSABLECALIDAD`, `JefeProyecto`, `Riesgos`; FALTA â†’ author via access-vba-tdd con esquema primero | 2026-06-15 |
| BR-EXP-6 â€” Los riesgos asociados a NC Proyecto deben leerse con semÃ¡ntica cache-first cuando la cachÃ© estÃ¡ cargada. | Verified-static | `tests/tests.vba.cache-e2e.json` 7/7 PASS (2026-06-14, staging `20b71f64`); referencia archivada en `docs/features/cache-management/trust-ncproyecto-cache-hits.md`; FALTA â†’ reejecutar | 2026-06-15 |
| BR-EXP-7 â€” El ciclo de vida propio de riesgos â€” aceptaciÃ³n, mitigaciÃ³n, contingencia, materializaciÃ³n, retirada, cierre y retipificaciÃ³n â€” estÃ¡ definido y probado. | Intended | FALTA â†’ author via access-vba-tdd tras confirmar estados | 2026-06-15 |
| Existe una UI de bÃºsqueda/selecciÃ³n de expedientes con filtros por palabra clave y responsable. | Verified-static | `src/forms/Form_FormExpedientesBusqueda.cls` | 2026-06-15 |
| El objeto `Expediente` expone responsables, jurÃ­dicas, responsable de calidad, jefe de proyecto y riesgos. | Verified-static | `src/classes/Expediente.cls` | 2026-06-15 |
| El objeto `Riesgo` contiene campos de aceptaciÃ³n, retirada, mitigaciÃ³n, cierre y retipificaciÃ³n. | Verified-static | `src/classes/Riesgo.cls` | 2026-06-15 |
| El flujo de negocio de riesgos estÃ¡ probado. | Intended | No hay manifest dedicado localizado | 2026-06-15 |

**âš ï¸ Divergencias (intenciÃ³n SDD â‰  realidad del cÃ³digo)**
- Sin divergencia confirmada. Hueco: la app tiene modelo de riesgos rico, pero el contrato de negocio y las pruebas son insuficientes.
