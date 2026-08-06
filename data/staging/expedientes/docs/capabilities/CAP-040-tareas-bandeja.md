<!--
Documento de capacidad CAP-040 — Bandeja de tareas pendientes.
Linaje: dysflow_get_schema (verificó que `TbTareas` no existe en backend) + lectura de Form_FormTareas.cls
Estado honesto: 2026-06-15 — la capacidad completa es TARGET post-PRUEBA-002 PR-F + PRUEBA-003 REFAC-4c. Hoy NO existe tabla backend, no existe helper, no existe test.
Idioma: castellano de España. Identificadores de código/test se mantienen tal cual.
-->

# Capacidad: Bandeja de tareas pendientes

## §0 Identidad
- **ID de capacidad**: CAP-040
- **Tier**: minimal (estado aspiracional; **sin implementación backend**)
- **Estado**: **deferred** (target post-PRUEBA-002 PR-F)
- **Source**: aspirational (PRUEBA-002 spec + PRUEBA-003 design; sin código actual que cumpla el contrato)
- **Responsable / autoridad de producto**: Pendiente de confirmación
- **Última verificación**: 2026-06-15 mediante `dysflow_get_schema({ tableName: "TbTareas" })` → `ACCESS_QUERY_FAILED: No se encontró el elemento en esta colección`. **La tabla `TbTareas` no existe en el backend actual.**
- **Confianza global**: **Intended / Divergent** — la spec describe una capacidad que el backend no soporta todavía. Hasta que se cree `TbTareas` y se implemente `Helper_BandejaTareas`, esta capacidad es deuda de producto, no funcionalidad activa.

## §1 Intención de negocio (≈ proposal SDD) — POR QUÉ
- **Propósito pretendido**: mostrar al usuario conectado la lista de tareas pendientes asignadas a él (o a su rol), permitirle marcarlas como completadas, reasignarlas, o ver el detalle del expediente origen.
- **Usuarios / perfiles pretendidos**: gestor de área, técnico, administrador; cada uno ve solo las tareas de su ámbito.
- **Problema que resolvería**: la bandeja actual (`Form_FormTareas.cls`, 436 líneas con 7 handlers) mezcla lógica de filtrado, renderizado y eventos sin separación testeable. Imposible de probar unitariamente sin abrir la UI; imposible de auditar "quién completó qué" sin log manual.
- **Valor de negocio pretendido**: bandeja determinista por usuario, trazabilidad de completitud, base para notificaciones.
- **No-objetivos pretendidos**: bandeja de consultas (CAP-022 usa otra estrategia); bandeja de informes (CAP-025); reasignación masiva.
- **Origen de la intención**: spec PRUEBA-002 §3.2 (BR-40-01..02) + design PRUEBA-003 §6.4 (`Helper_BandejaTareas` como clase stateful).
- **Referencia de tracker de origen**: PRUEBA-002 PR-F, PRUEBA-003 REFAC-4c.

## §2 Contrato de comportamiento (≈ spec SDD) — QUÉ ⟵ ANCLA DE REGRESIÓN

### Estado actual
**Esta capacidad NO está implementada.** No existe:
- Tabla `TbTareas` en el backend (verificado con `dysflow_get_schema` → `ACCESS_QUERY_FAILED`)
- Clase `Helper_BandejaTareas.cls` en `src/classes/`
- Tests atómicos en `tests/tests.vba.json` con tag `bandeja-tareas`
- Entry points en `ExpedienteOperaciones.cls` o similar

Lo que **SÍ existe**:
- `Form_FormTareas.cls` (form con UI y handlers; estado interno: `m_ColTareasFiltradas`, `Filtrar`, `ComandoActualizar_Click`, etc.)
- `Form_FormTareas.form.txt` (diseño de controles)

### Escenarios futuros (target post-PR-F)
- **DADO** `g_UsuarioConectado = "DOMINIO\jgarcia"` y `TbTareas` con 3 filas (1 pendiente suya, 1 pendiente de otro, 1 completada suya) **CUANDO** se invoca `Helper_BandejaTareas.CargarTareas(p_FiltroTipo, p_UsuarioConectadoId, p_Db, p_Error)` **ENTONCES** la colección resultante contiene solo la tarea 1.
- **DADO** una tarea pendiente con `IDTarea = 42` y `g_UsuarioConectado = "DOMINIO\jgarcia"` **CUANDO** se invoca `Helper_BandejaTareas.MarcarCompletada(p_IDTarea, p_Db, p_Error)` **ENTONCES** `TbTareas WHERE IDTarea = 42` tiene `Estado = 'Completada'`, `FechaCompletado = Now()`, `UsuarioCompletado = "DOMINIO\jgarcia"`.
- **DADO** un usuario sin tareas pendientes **CUANDO** se invoca `CargarTareas` **ENTONCES** retorna colección vacía; la UI muestra mensaje "No hay tareas pendientes" (manejo del form).
- **DADO** una tarea ya completada que el usuario intenta re-marcar **CUANDO** se invoca `MarcarCompletada` sobre un `IDTarea` ya completado **ENTONCES** retorna `False` con `MSG-TAR-001: La tarea ya está completada` (idempotencia).

### Reglas de negocio (target post-PR-F)

| ID regla | Enunciado pretendido | Autoridad | ¿Aplicada en código HOY? | Prueba | Confianza |
|---|---|---|---|---|---|
| BR-40-01 | La bandeja solo muestra tareas del usuario conectado en estado "Pendiente" | spec PRUEBA-002 §3.2 | **NO** — backend no tiene `TbTareas` | `Test_Helper_BandejaTareas_CargarTareas_FiltraPorUsuario` (target post-PR-F, no implementable hoy) | Intended (deuda de producto) |
| BR-40-02 | `MarcarCompletada` actualiza `Estado = 'Completada'`, `FechaCompletado`, `UsuarioCompletado` | spec | **NO** | `Test_Helper_BandejaTareas_MarcarCompletada_ActualizaEstadoYAudita` (target post-PR-F) | Intended |
| BR-40-03 | `MarcarCompletada` es idempotente (`MSG-TAR-001` si ya completada) | spec | **NO** | (no testeable hoy) | Intended |
| BR-40-04 | El form muestra mensaje "No hay tareas pendientes" cuando la colección está vacía | spec UI | parcial — el form tiene la lógica, pero el helper no existe | cubierto por test de integración post-PR-F | Intended |

### Validaciones (target)
- `p_UsuarioConectado = ""` → `MSG-TAR-002: Usuario no identificado`
- `p_IDTarea` ya completado → `MSG-TAR-001: La tarea ya está completada`
- `p_IDTarea` no existe → `MSG-TAR-003: Tarea no encontrada`

### Transiciones de estado (target)
- `Pendiente` --(MarcarCompletada)--> `Completada` con `FechaCompletado = Now()`, `UsuarioCompletado = g_UsuarioConectado`

### Casos límite y de error (target)
- `p_Db Is Nothing` → fallback a `getdb()`
- Concurrencia: si dos usuarios marcan la misma tarea, el segundo recibe `MSG-TAR-001` (idempotencia)

### Señales de aceptación / presencia (target)
- `Helper_BandejaTareas.cls` existe en el binario
- `Form_FormTareas.Filtrar` es una llamada delgada a `Helper_BandejaTareas.CargarTareas`
- `TbTareas` existe en el backend con PK `IDTarea` y FK a `TbExpedientes`
- `dysflow.test_vba` con filtro `bandeja-tareas` retorna `ok: true` con 2/2 PASS

## §3 Mapa de implementación (≈ design SDD) — CÓMO
- **Puntos de entrada de UI (HOY)**: `Form_FormTareas.cls` (436 líneas) — `Filtrar`, `ComandoActualizar_Click`, `ComandoDetalle_Click`, `ComandoVerInforme_Click`, `ListaFiltrados_Click/DblClick`, `ListaTipoTareas_Click`, `Form_Open`. Estado interno: `m_ColTareasFiltradas As Collection`. **No testeable** porque las tareas no se persisten en una tabla backend.
- **Puntos de entrada de código (HOY)**: ninguno. `Helper_BandejaTareas.cls` no existe.
- **Puntos de entrada de código (target post-PRUEBA-003 REFAC-4c)**:
  - `Helper_BandejaTareas.CargarTareas(ByVal p_FiltroTipo As String, ByVal p_UsuarioConectadoId As Long, Optional ByVal p_db As DAO.Database = Nothing, Optional ByRef p_Error As String) As Long` — retorna count
  - `Helper_BandejaTareas.ObtenerTareas(ByVal p_IdxDesde As Long, ByVal p_IdxHasta As Long, Optional ByRef p_Error As String) As Collection` — slice paginado
  - `Helper_BandejaTareas.MarcarCompletada(ByVal p_IDTarea As Long, Optional ByVal p_db As DAO.Database = Nothing, Optional ByRef p_Error As String) As Boolean`
  - `Helper_BandejaTareas.Reasignar(ByVal p_IDTarea As Long, ByVal p_NuevoUsuario As String, Optional ByVal p_db As DAO.Database = Nothing, Optional ByRef p_Error As String) As Boolean`
  - Estado interno: `m_ColTareasFiltradas As Collection`, `m_FiltroTipo As String`, `m_UsuarioConectado As String`
- **Datos afectados (target)**: `TbTareas` (lectura para filtrar; UPDATE para `MarcarCompletada` y `Reasignar`)
- **Dependencias (target)**: `TbTareas` (PK: `IDTarea`; FK a `TbExpedientes`); `g_UsuarioConectado`; `m_TestingMode` y `getdb()`; `Helper_Autorizacion` (CAP-041) para validar que el usuario puede reasignar
- **Sincronización fuente↔binario**: **no aplica** — el helper no existe
- **Valoración de diseño (tal-como-está vs ideal)**: **NO IMPLEMENTADO**. La capacidad completa (backend + helper + tests) es target de PRUEBA-002 PR-F + PRUEBA-003 REFAC-4c. **No se puede testear hoy.**

## §4 Receta de reconstrucción (≈ tasks SDD) — REPRODUCIBILIDAD
> Esta receta es target. Reconstruir esta capacidad requiere:
1. **Crear la tabla `TbTareas` en el backend** con schema:
   - `IDTarea` (Long, required, PK)
   - `IDExpediente` (Long, FK a `TbExpedientes`)
   - `Descripcion` (Text 255)
   - `TipoTarea` (Text 50)
   - `Estado` (Text 20, ∈ {`Pendiente`, `EnProgreso`, `Completada`, `Cancelada`})
   - `UsuarioAsignado` (Text 100, default = `g_UsuarioConectado`)
   - `FechaAsignacion` (Date, default = `Now()`)
   - `FechaCompletado` (Date, opcional)
   - `UsuarioCompletado` (Text 100, opcional)
   - `Prioridad` (Text 20, ∈ {`Baja`, `Media`, `Alta`, `Critica`})
   - **Pendiente de confirmación con el responsable de producto**
2. **Crear `src/classes/Helper_BandejaTareas.cls`** con las firmas de §3 (clase stateful con `m_ColTareasFiltradas`)
3. **Refactor `Form_FormTareas.cls`** para que `Filtrar` llame al helper con `New Helper_BandejaTareas`
4. **Crear `src/modules/Test_HelperBandejaTareas.bas`** con los 2 tests atómicos de §2 (BR-40-01 con `BeginTestSession`/`EndTestSession` + 3 filas fixture, BR-40-02 con cardinalidad antes/después)
5. **Registrar en `tests/tests.vba.json`** con `tags: ["verified-runtime", "pr-f", "cap-040", "br-40-NN", "bandeja-tareas", "fixture-first"]`
6. **Importar + compilar + testear** vía Dysflow MCP
7. **Demostrar** → `dysflow.test_vba({ projectId: "expedientes", filter: "bandeja-tareas" })` → 2/2 PASS

## §5 Evidencia y trazabilidad (≈ verify SDD)
- **Tests**: `src/modules/Test_HelperBandejaTareas.bas` (no existe). Pre-PRUEBA-002: 0 tests específicos de bandeja de tareas.
- **Trazabilidad de release** (solo-añadir):

| Elemento | Ref. tracker | Versión de staging | Estado UAT | Release de producción | Fecha en producción | Nota |
|---|---|---|---|---|---|---|
| Bandeja de tareas pendientes (backend + helper + tests) | PRUEBA-002 PR-F + PRUEBA-003 REFAC-4c | Pendiente | pending | Pendiente | — | **deuda de producto**; tabla `TbTareas` no existe, helper no existe, tests no existen |

- **Tabla de diagnóstico de regresión**: N/A (la capacidad no existe)

## §6 Notas de migración web
- **Conservar (target)**: el filtrado por usuario, la completitud con auditoría, la idempotencia.
- **Transformar (target)**: la cache `m_ColTareasFiltradas` → query + paginación server-side. La completitud → `POST /api/tareas/{id}/completar`.
- **NO copiar (legado)**: el estado en memoria de la cache — la web no mantiene estado entre requests.
- **Preguntas abiertas**:
  - ¿La bandeja debe refrescarse automáticamente (polling) o solo al abrir el form? Hoy parece manual (`ComandoActualizar_Click`).
  - ¿Cuál es el universo de `TipoTarea`? ¿`["TareaS4H", "VerificacionCalidad", "Aprobacion", "Reasignacion", ...]`?
  - ¿Cómo se crea una tarea? ¿Manual desde form, automática desde evento de expediente, o ambas?
  - ¿La reasignación masiva debe quedar en este helper o moverse a `Helper_ExpedienteOperacionesMasivas`?

## §7 Registro de confianza

| Hecho | Confianza | Evidencia | Fecha |
|---|---|---|---|
| Tabla `TbTareas` NO existe en el backend actual | Verified-runtime | `dysflow_get_schema` → `ACCESS_QUERY_FAILED` | 2026-06-15 |
| `Helper_BandejaTareas.cls` NO existe | Verified-static | `Get-ChildItem src/classes` no lo lista | 2026-06-15 |
| `Form_FormTareas.cls` EXISTE con 436 líneas y 7 handlers (UI) | Verified-runtime | `Get-ChildItem src/forms` lo lista; compila OK | 2026-06-15 |
| Tests `Test_Form_FormTareas_*` NO están en `tests/tests.vba.json` | Verified-static | `grep` en `tests.vba.json` retorna 0 matches | 2026-06-15 |
| La spec PRUEBA-002 §3.2 (CAP-040) menciona 2 reglas con tests target | Verified-static | lectura del spec | 2026-06-15 |

**⚠️ Divergencia crítica (intención SDD vs realidad del código)**:
- La spec PRUEBA-002 describe la capacidad CAP-040 como si existiera, con 2 reglas y tests target. La realidad: **no hay tabla, no hay helper, no hay test, no hay entry point**. La capacidad es 100% aspiracional.
- La spec también dice que `Form_FormTareas` tiene "filter en `Verified-runtime` por tests existentes" — **no es cierto**. No hay tests para `Form_FormTareas` en `tests.vba.json` (verificado con grep). El form existe pero su comportamiento no está cubierto por tests.
- **Esta capacidad es deuda de producto**: antes de que CAP-040 pueda moverse a `Verified-runtime`, hay que:
  1. Crear la tabla `TbTareas` en el backend
  2. Diseñar el `Helper_BandejaTareas` (PRUEBA-003 REFAC-4c)
  3. Extraer el form a shells thin
  4. Autor de los 2 tests atómicos con fixture-first
  5. Verificar verde con `dysflow.test_vba`
- **No se debe incluir CAP-040 en el rollup de "reglas que migran a `Verified-runtime`"** de PRUEBA-002 hasta que estos 5 pasos se ejecuten. El coverage-matrix §PR-F los lista como target, pero su consecución depende de infraestructura de backend que NO existe.
