# ExpedienteAntiguo

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/classes/ExpedienteAntiguo.cls` (569 líneas) · **PRD:** [README](../README.md)

Objeto de datos con la forma de una fila del esquema anterior de expedientes de AGEDYS: 50 campos públicos, la lista `ColCampos` y acceso por nombre, sin lógica propia. Igual que en el snapshot. Las citas son `ExpedienteAntiguo.cls:línea` salvo indicación.

**Sin uso en RC1:** ningún fichero de `src` instancia la clase ni cita su nombre fuera de ella.

## Miembros públicos

Accesores triviales, agrupados: `getPropiedad` y `SetPropiedad` leen y escriben por nombre los 50 campos (`:141-354, 356-569`). Un nombre desconocido deja «Propiedad no reconocida en ExpedienteAntiguo.getPropiedad» (o `.SetPropiedad`) y lanza el error interno 1000. `ColCampos` devuelve los 50 nombres en caché (`:68-138`) y `Error` guarda el último error (`:65`).

| Campos | Tipo en la clase | Evidencia |
|---|---|---|
| `Adjudicado`, `EnPeriodoDeAdjudicacion`, `Informe` | `Boolean` | `:29, 30, 33` |
| Los 47 restantes, de `IDExpediente` a `Observaciones`, en el orden de las columnas de `TbExpAgedys` | `String` | `:14-28, 31-32, 34-63` |

Los 49 primeros campos son, sin distinguir mayúsculas, las 49 columnas de `TbExpAgedys` y de `TbExpAGEDYS1`, que el volcado registra con las mismas columnas ([`ERD/schema.sql`](../../ERD/schema.sql), `:267-316` y `:320-369`). `Observaciones` (`:63`) no existe en ninguna de las dos. Que la clase represente esas tablas es «hipótesis» por los nombres: no cita ninguna tabla.

Cuatro nombres cambian de caja entre la declaración y `ColCampos`/`getPropiedad`: `IDExpediente` y `IdExpediente`, `AÑO` y `año`, `Clasificacion` y `CLASIFICACION`, `ESTADO` y `Estado` (`:14, 80, 146`; `:32, 98, 218`; `:46, 112, 274`; `:56, 122, 314`). La comparación de nombres funciona con `Option Compare Database` (`:10`) si la base no distingue mayúsculas («hipótesis», como en la ficha de CPV).

No hay propiedades calculadas ni reglas de validación.

## Persistencia

La clase no lee ni escribe ninguna tabla. La única lectura de `TbExpAgedys` en RC1 está en `Instalador.RellenarOficinasProgramas` (`Instalador.bas:791-830`), que no tiene llamadas en `src` ni usa la clase:

1. Borra **todo** `TbExpedientes` con `DELETE * FROM TbExpedientes` (`Instalador.bas:799-800`).
2. Recorre `TbExpAgedys` y por cada fila hace `AddNew` en `TbExpedientes` (`Instalador.bas:802-823`) con: `IDExpediente`; `Expediente` → `CodExp`; `CodExpedienteLargo` → `CodExpLargo`; `TITULOEXP` → `Titulo`; `FirmadelContrato` → `FechaInicioContrato`; `FechaFinExp` → `FechaFinContrato`; `AcuerdoMarco` → `EsAM`; y, si `AcuerdoMarco` no es «Sí», `EsExpediente` = «Sí» y `EsLote` = «No» (`Instalador.bas:810-820`).

## Comportamiento que la migración debe conservar

Ninguno en ejecución. Si negocio decide migrar `TbExpAgedys` ([P-EXP-12](../issues-propuestas.md)), la correspondencia de columnas de `RellenarOficinasProgramas` es la única que `src` documenta entre esa tabla y `TbExpedientes`.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- El manejador de `ColCampos` sobrescribe el mensaje con `m_Error = Me.Error` (`:134-137`): el error se pierde.
- Tipos que no siguen al volcado: `IMPORTEEXP`, `AÑO` y las fechas (p. ej. `:19, 24-27`) son `String` (`:21, 32`) frente a `DOUBLE`, `LONG` y `SHORT_DATE_TIME` (`schema.sql:275, 286, 273, 278-281`); `AñosGarantia` es `String` aquí (`:38`) y `Double` en [ExpedienteDatosAGEDY](ExpedienteDatosAGEDY.md).
- `RellenarOficinasProgramas` vacía `TbExpedientes` entera antes de copiar (`Instalador.bas:799-800`), no cierra sus recordsets (`Instalador.bas:821-825`) y su mensaje de error cita «RellenoDeDatosDeAGEDYS» (`Instalador.bas:828`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: no portar la clase; tratar `TbExpAgedys` y `TbExpAGEDYS1` con las tablas históricas.

## Criterios de aceptación

Propuestos, no aprobados:

1. Ningún flujo web depende de esta clase.
2. Ninguna carga de datos web borra `TbExpedientes` entera para rellenarla desde `TbExpAgedys`.

## Pendientes

- Origen de `Observaciones` (`:63`): no está en el volcado del snapshot; confirmar con el ERD de RC1 ([P-EXP-1](../issues-propuestas.md)).
- Confirmar que la clase no se usa desde el binario, fuera del alcance de la lectura estática de `src`.
