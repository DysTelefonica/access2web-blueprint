# ExpedienteAGEDYS

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/classes/ExpedienteAGEDYS.cls` (152 líneas) · **PRD:** [README](../README.md)

Objeto de datos de un registro de `TbExpedientesDatosAGEDYS`: tres campos, acceso perezoso al `Expediente` y cálculo del siguiente identificador. Es la única clase que cita esa tabla. Frente al snapshot solo cambia la caja de un identificador privado (`m_objExpediente`, `:19`). Las citas son `ExpedienteAGEDYS.cls:línea` salvo indicación.

**Ningún flujo de RC1 llega a la clase:** `constructor.getExpedienteAGEDYS` no tiene llamadas en `src`, `IDExpedienteAGEDYSCalculado` tampoco, y `Expediente.cls:141` declara `m_ObjExpedienteAGEDYS` sin usarlo en ninguna otra línea.

## Miembros públicos

| Miembro | Tipo | Significado | Evidencia |
|---|---|---|---|
| `IDExpedienteAGEDYS` | `String` | Identificador del registro; columna homónima | `:14, 85` |
| `IDExpediente` | `String` | Expediente al que pertenece el registro | `:15, 86` |
| `Generico` | `String` | Marca de expediente genérico de AGEDYS; «hipótesis»: texto `Sí`/`No`, como `TbExpAgedys.Generico` (`TEXT(2)`, [`schema.sql:291`](../../ERD/schema.sql)) | `:16, 87` |
| `Expediente` | `Property Get` → `Expediente` | Carga perezosa con `constructor.getExpediente`; `Nothing` si `IDExpediente` está vacío; se guarda tras la primera lectura | `:24-50` |
| `IDExpedienteAGEDYSCalculado` | `Property Get` → `String` | Siguiente identificador libre de la tabla | `:52-71` |
| `ColCampos` | `Property Get` → `Collection` | Los tres nombres de columna anteriores, en caché | `:73-96` |
| `getPropiedad`, `SetPropiedad` | funciones | Lectura y escritura por nombre de los tres campos; un nombre desconocido deja «Propiedad no reconocida en ExpedienteAGEDYS.getPropiedad» (o `.SetPropiedad`) y lanza el error interno 1000 | `:99-123, 126-150` |
| `Error` | `String` | Último error de las propiedades | `:21` |

## Reglas y cálculos

- **Siguiente identificador:** `IDExpedienteAGEDYSCalculado` llama en cada lectura a `DameID("TbExpedientesDatosAGEDYS", "IDExpedienteAGEDYS", getdb())` (`:59`), que devuelve `Max + 1`, o `1` con la tabla vacía (`FUNCIONES UTILES.bas:2851-2864`). No guarda el valor entre lecturas.
- **Carga desde la tabla** (`constructor.getExpedienteAGEDYS`, `constructor.bas:5024-5073`):
  1. Sin `p_IDExpedienteAGEDYS` ni `p_IDExpediente`, sale sin consultar y devuelve `Nothing` (`constructor.bas:5037-5039`).
  2. Si llega `p_IDExpedienteAGEDYS`, filtra por él; si no, por `IDExpediente` (`constructor.bas:5040-5048`).
  3. Sin filas devuelve `Nothing` (`constructor.bas:5052-5056`). Con filas, crea la clase con la **primera** y copia cada columna de `ColCampos` con `Nz(..., "")` (`constructor.bas:5057-5063`).

## Persistencia

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lectura por ID o por expediente | `TbExpedientesDatosAGEDYS` | `SELECT *` concatenado, ID sin comillas | — | `constructor.bas:5041-5050` |
| Siguiente ID | `TbExpedientesDatosAGEDYS` | `SELECT Max(...)` | no | `:59`; `FUNCIONES UTILES.bas:2851-2853` |
| Alta, edición o baja | — | no existe en `src` | — | sin más citas de la tabla que las tres anteriores |

**Tabla `TbExpedientesDatosAGEDYS`: ausente del volcado ERD (anterior al commit fijado).** [`ERD/schema.sql`](../../ERD/schema.sql) no la contiene, y no es `TbExpAgedys`: esta no tiene columna `IDExpedienteAGEDYS` (`schema.sql:267-316`). Las columnas que el código exige están en la [matriz](../matriz-campos/definidas-en-src.md#tbexpedientesdatosagedys).

## Comportamiento que la migración debe conservar

Ninguno en ejecución: ningún formulario ni módulo de producción llega a la clase. Lo que se conserve depende de los datos: si la tabla existe en el backend de RC1 y tiene filas, su disposición la decide negocio, como en las tablas históricas ([P-EXP-12](../issues-propuestas.md)).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- El mensaje de error de `Expediente` cita «ExpedienteResponsable.Expediente» (`:47`).
- El manejador de `ColCampos` sobrescribe el mensaje con `m_Error = Me.Error` (`:92-95`): el error se pierde y la propiedad devuelve `Nothing`.
- La lectura por `IDExpediente` no ordena y toma la primera fila (`constructor.bas:5045-5048, 5057`): con varios registros por expediente, el elegido no es determinista.
- `getPropiedad` no limpia `p_Error` al entrar (`:103`), a diferencia de `SetPropiedad` (`:130`).
- `Max + 1` fuera de transacción ([coexistencia](../coexistencia.md#3-identidad-y-generación-de-claves)).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: no portar la clase y decidir la tabla con el resto de las históricas.

## Criterios de aceptación

Propuestos, no aprobados; aplican solo si se decide migrar la tabla:

1. Pedir los datos AGEDYS sin identificador de registro ni de expediente no consulta y devuelve vacío.
2. Con ambos identificadores, prevalece el del registro.
3. Un expediente sin registro devuelve vacío sin error.

## Pendientes

- Confirmar en el backend de RC1 si `TbExpedientesDatosAGEDYS` existe, con sus columnas, tipos, clave y filas ([P-EXP-1](../issues-propuestas.md)).
- Si `Generico` duplica `TbExpedientes.AGEDYSGenerico` ([cabecera](../matriz-campos/cabecera.md#tbexpedientes)) («hipótesis» por el nombre).
