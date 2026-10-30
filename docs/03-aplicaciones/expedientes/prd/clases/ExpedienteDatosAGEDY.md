# ExpedienteDatosAGEDY

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/classes/ExpedienteDatosAGEDY.cls` (169 líneas) · **PRD:** [README](../README.md)

Contenedor de diez datos AGEDYS de un expediente, sin lógica propia: campos públicos, la lista `ColCampos` y acceso por nombre. Igual que en el snapshot. Las citas son `ExpedienteDatosAGEDY.cls:línea` salvo indicación.

**Sin uso en RC1:** ningún fichero de `src` instancia la clase ni cita su nombre fuera de ella. Tampoco hay constructor que la cargue ni código que la persista.

## Miembros públicos

Los diez campos son públicos y sin lógica; `getPropiedad` y `SetPropiedad` los leen y escriben por nombre (`:61-113, 116-168`). Un nombre desconocido deja «Propiedad no reconocida en ExpedienteDatosAGEDY.getPropiedad» (o `.SetPropiedad`) y lanza el error interno 1000. `ColCampos` devuelve los diez nombres en caché (`:28-58`) y `Error` guarda el último error (`:25`).

La columna homónima es «hipótesis»: la clase no cita tabla, pero los diez nombres coinciden con columnas de `TbExpAgedys` en [`ERD/schema.sql`](../../ERD/schema.sql) (`IdExpediente` solo difiere en la caja).

| Campo | Tipo en la clase | Columna homónima en `TbExpAgedys` | Evidencia |
|---|---|---|---|
| `IDExpediente` | `String` | `IdExpediente` `LONG(4)` | `:14`; `schema.sql:268` |
| `CarpetaArchivo` | `String` | `TEXT(11)` | `:15`; `schema.sql:272` |
| `FirmadelContrato` | `String` | `SHORT_DATE_TIME(8)` | `:16`; `schema.sql:273` |
| `FechaRecepcion` | `String` | `SHORT_DATE_TIME(8)` | `:17`; `schema.sql:281` |
| `Adjudicado` | `Boolean` | `BOOLEAN(1)` | `:18`; `schema.sql:283` |
| `Informe` | `Boolean` | `BOOLEAN(1)` | `:19`; `schema.sql:287` |
| `Pcal` | `String` | `TEXT(2)` | `:20`; `schema.sql:289` |
| `Generico` | `String` | `TEXT(2)` | `:21`; `schema.sql:291` |
| `AñosGarantia` | `Double` | `DOUBLE(8)` | `:22`; `schema.sql:292` |
| `AcuerdoMarco` | `String` | `TEXT(2)` | `:23`; `schema.sql:311` |

No hay propiedades calculadas ni reglas de validación.

## Persistencia

Ninguna en `src`. La clase no lee ni escribe ninguna tabla. La única lectura de `TbExpAgedys` en RC1 es `Instalador.RellenarOficinasProgramas`, que no usa la clase (ver [ExpedienteAntiguo](ExpedienteAntiguo.md#persistencia)).

## Comportamiento que la migración debe conservar

Ninguno: la clase no participa en ningún flujo. Los diez datos que representa forman parte de `TbExpAgedys`, cuya disposición está pendiente de negocio ([P-EXP-12](../issues-propuestas.md)).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- El manejador de `ColCampos` sobrescribe el mensaje con `m_Error = Me.Error` (`:54-57`): el error se pierde.
- `getPropiedad` no limpia `p_Error` al entrar (`:65`), a diferencia de `SetPropiedad` (`:120`).
- «Hipótesis»: un cargador con el patrón de `constructor.bas` (`SetPropiedad` con `Nz(..., "")`, p. ej. `constructor.bas:5059`) fallaría por tipo con `Adjudicado`, `Informe` o `AñosGarantia` nulos, porque asignaría `""` a un `Boolean` o un `Double`. No existe tal cargador en RC1.
- `Adjudicado` es `Boolean` aquí y `TEXT(2)` con `Sí`/`No` en `TbExpedientes` ([cabecera](../matriz-campos/cabecera.md#tbexpedientes)): mismo nombre con dos tipos.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: no portar la clase.

## Criterios de aceptación

Propuestos, no aprobados:

1. Ningún flujo web depende de esta clase.
2. Si negocio decide migrar `TbExpAgedys`, cada uno de los diez campos tiene destino o motivo de descarte en la [matriz](../matriz-campos.md).

## Pendientes

- Confirmar que la clase no se usa desde el binario (código no exportado, macros o referencias por cadena), fuera del alcance de la lectura estática de `src`.
