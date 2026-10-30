# UltimoCambio

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/classes/UltimoCambio.cls` (231 líneas) · **PRD:** [README](../README.md)

Registro de quién modificó un expediente y cuándo, con su texto de pantalla. Lo escribe `UltimoCambioOperaciones.Registrar` (`src/classes/UltimoCambioOperaciones.cls`, 15-90) y lo dispara `RegistrarUltimoCambio` (`FUNCIONES UTILES.bas:189-224`), con 25 llamadores en cuatro ficheros de `src` (`FUNCIONES UTILES.bas`, `ExpedienteOperaciones.cls`, `ExpedienteModificadoOperaciones.cls` y `ExpedienteSuministradorServicio.cls`): casi toda operación de escritura (p. ej. `ExpedienteOperaciones.cls:554, 1821, 1862, 1935`) y `ExtenderEntidades` (`FUNCIONES UTILES.bas:3952, 3960`). Las citas son `UltimoCambio.cls:línea` salvo indicación.

## Miembros

| Miembro | Significado | Evidencia |
|---|---|---|
| `id`, `IDExpediente`, `FechaCambio`, `IDUsuarioCambio` | Las cuatro columnas de `TbUltimoCambio` | `:13-16`; `ERD/schema.sql:756-761` |
| `IDCalculado` | Siguiente identificador libre (`DameID("TbUltimoCambio", "ID")`) | `:26-45` |
| `texto` | «Última modificación por <usuario> para el Exp <CodExp> el <fecha>»; cada parte vale «Desconocido» (o «Desconocida» la fecha) si no se puede resolver | `:48-90` |

## Reglas (`UltimoCambioOperaciones.Registrar`)

1. **Una fila por usuario, no por expediente:** busca en `TbUltimoCambio` la fila de `IDUsuarioCambio` del usuario conectado; si no existe la inserta con un identificador nuevo, y si existe la sobrescribe con el expediente y la fecha actuales (`:32-50`).
2. **Auditoría en el propio expediente:** si `TbExpedientes.FechaCreacion` no es una fecha, la fija a `Now()` junto con `IDUsuarioCreacion`; si ya lo es, actualiza `FechaUltimoCambio` e `IDUsuarioUltimoCambio` (`:64-75`).
3. Guarda el resultado en `m_ObjEntorno.UltimoCambio` y, si `FormExpedientesGestion` está abierto, actualiza `lblUltimaModificacion` (`FUNCIONES UTILES.bas:213-215`); el rótulo solo es visible si el formulario está abierto para editar y el usuario es administrador; en otro caso se oculta (`Form_FormExpedientesGestion.cls:55-66`).

## Posibles defectos (hipótesis)

- `TbUltimoCambio` guarda solo la última modificación **de cada usuario**; no es un histórico de cambios de un expediente. Quien espere trazabilidad por expediente no la tiene.
- La primera escritura de un expediente se anota como creación y no como cambio (regla 2): si `FechaCreacion` estaba vacía por datos antiguos, un cambio posterior se registra como creación.
- `IDUsuarioCreacion` e `IDUsuarioUltimoCambio` son texto (`TEXT(255)`, `ERD/schema.sql:30-32`) y `IDUsuarioCambio` es numérico: dos tipos para el mismo concepto.

## Comportamiento que la migración debe conservar

La auditoría de creación y última modificación de cada expediente (fecha y usuario) y el rótulo de última modificación del administrador. La tabla `TbUltimoCambio` por usuario es un detalle de implementación; en web conviene un histórico de cambios por expediente (propuesta pendiente de decisión).

## Pendientes

- `TbUltimoCambio` figura como «no detallada» en el [modelo de datos](../../data-model.md); falta su fila en la matriz de campos.
