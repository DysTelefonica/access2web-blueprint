# Form_FormExpedienteCambioTipo

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/forms/Form_FormExpedienteCambioTipo.cls` (336 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Diálogo «CAMBIO DE TIPO» (`.form.txt:857`): cambia la tipología de un expediente ya registrado (Acuerdo Marco, Lote, Basado o Expediente Independiente) y, según el destino, su expediente padre. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): trabaja sobre la global `m_ObjExpedienteActivo` y persiste con `ExpedienteOperaciones.RegistrarCambioTipo`. Las citas son `Form_FormExpedienteCambioTipo.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormExpedientesGestion` (botones `ComandoCambiarTipoAMoC`, `ComandoCambiarTipoLote`, `ComandoCambiarTipoBasado`) | `CambiarTipo` fija `m_ObjExpedienteActivo` con el expediente de la lista, cierra el diálogo si estaba abierto, lo abre y escucha `CambioRealizado` con `WithEvents` | `Form_FormExpedientesGestion.cls:6, 953-962, 982-992, 2584-2594, 2615-2677` |

Antes de abrir, Gestión rechaza un AM o un Lote con derivados: «El AM tiene derivados…» / «El Lote tiene derivados…» (`Form_FormExpedientesGestion.cls:2656-2667`). Los basados no tienen esa comprobación.

Abre `FormExpedientesParaCambioTipo` (búsqueda del padre) y escucha su evento `Seleccionado` (`:5, 40-53, 263-270`). Salidas: `cmdSalir` cierra sin guardar (`:10-14`); `ComandoRegistrar` cierra tras guardar (`:166`). Evento hacia quien lo abrió: `CambioRealizado` (`:6`, lanzado en `:165`); Gestión, al recibirlo, vacía `m_ObjEntorno.ColExpedientesCompletos`, `m_ExpBusquedaUltimo` y `m_ExpLoteSeleccionadoC` y vuelve a filtrar (`Form_FormExpedientesGestion.cls:3223-3233`).

**Permisos:** el formulario no comprueba rol ni `EsAdministrador`; quién puede pulsar los botones de Gestión no se ha trazado (ver Pendientes).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `TIPOACTUAL` | texto deshabilitado y bloqueado | — | Muestra `TipoCalculadoTexto` del expediente | `:231`, `.form.txt:956-969` |
| `Tipo` | combo de lista de valores (`LimitToList`) | `AfterUpdate` → `Tipo_AfterUpdate` | Habilita o bloquea `IDExpedientePadre` según el destino elegido | `:285-324`, `.form.txt:986-1008` |
| `IDExpedientePadre` | texto (deshabilitado en diseño) | — | ID del padre; se escribe a mano cuando está habilitado o lo rellena la búsqueda | `:225-226, 270`, `.form.txt:1028-1041` |
| `ComandoBuscar` (icono de lupa; ayuda «Formulario de Búsqueda del padre») | botón | `OnClick` → `ComandoBuscar_Click` | Cierra y abre `FormExpedientesParaCambioTipo` y se suscribe a `Seleccionado` | `:40-53`, `.form.txt:1145-1157` |
| `ComandoRegistrar` (rótulo «Editar») | botón | `OnClick` → `ComandoRegistrar_Click` | Valida, aplica los indicadores de tipo y guarda | `:68-170`, `.form.txt:1212-1214` |
| `ComandoAyuda` (rótulo «Salir») | botón | `OnClick` | `AbrirAyuda` | `:16-25`, `.form.txt:876-878` |
| `cmdSalir` | botón | `OnClick` | Cierra sin guardar | `:10-14`, `.form.txt:1299-1301` |

## Flujos

### Apertura (`Form_Open`, `:181-214`)

1. `Ajustar Me` y `Caption = m_TituloFormulario` (`:189-192`).
2. Sin `m_ObjExpedienteActivo`: «El formulario se ha abierto con parámetros insuficientes» (`:193-196`).
3. `EstablecerDatos` bloquea `IDExpedientePadre`, vacía el combo y lo rellena según `TipoCalculado` (`:225-252`):

| Tipo actual | Opciones ofrecidas | Evidencia |
|---|---|---|
| AM | Lote · Basado · Expediente Independiente | `:232-235` |
| Lote | Lote de Acuerdo Marco · Acuerdo Marco · Basado · Expediente Independiente | `:236-240` |
| Basado de AM o de Lote | Acuerdo Marco · Lote · Expediente Independiente · Basado | `:241-246` |
| Expediente individual | Acuerdo Marco · Lote · Basado | `:248-251` |

### Elección del destino (`Tipo_AfterUpdate`, `:285-324`)

`IDExpedientePadre` queda bloqueado para «Convertir a Acuerdo Marco» y «Convertir a Expediente Independiente» (`:304-307, 316-319`) y habilitado para Lote, Basado y Lote de Acuerdo Marco (`:308-315, 320-323`). No vacía el valor ya escrito.

### Búsqueda del padre

`FormExpedientesParaCambioTipo` busca por palabra clave (`Form_FormExpedientesParaCambioTipo.cls:47`), solo acepta un AM o un Lote («Se ha de seleccionar un AM o un Lote», `Form_FormExpedientesParaCambioTipo.cls:177-180`) y lanza `Seleccionado`; este formulario copia su `IDExpediente` en `IDExpedientePadre` (`:263-270`).

### Guardado (`ComandoRegistrar_Click`, `:68-170`)

1. Sin destino elegido: «Seleccione el cambio de tipo» (`:79-83`).
2. Según el destino, valida el padre y fija en `m_ObjExpedienteActivo` el padre y los indicadores `EsAM`/`EsLote`/`EsExpediente`/`EsBasado` (`:84-156`; detalle en Reglas).
3. `ExpedienteOperaciones.RegistrarCambioTipo` persiste (`:157-164`).
4. Éxito: lanza `CambioRealizado` y cierra (`:165-166`).

### Errores

Cada handler acumula el texto en `m_Error`, envía `CorreoAlAdministrador` salvo el error interno 1000 y muestra `MsgBox` crítico (`:171-178`, `:205-212`). `Form_Open` no cancela la apertura ante error (`:205-213`): el diálogo queda abierto. `EstablecerDatos` solo devuelve el texto en `p_Error` (`:256-260`).

## Reglas, validaciones y cálculos

| Destino | Padre | Indicadores resultantes | Evidencia |
|---|---|---|---|
| Acuerdo Marco | se vacía | `EsAM = Sí`, resto `No` | `:86-91` |
| Lote | opcional; si se indica, debe existir y ser AM («Un Lote puede tener un padre sólo de tipo Acuerdo Marco») | `EsLote = Sí` **solo si hay padre** | `:92-111` |
| Basado | obligatorio, debe existir y ser AM o Lote | `EsBasado = Sí` | `:112-133` |
| Expediente Independiente | debe estar vacío; si no, error | `EsExpediente = Sí` | `:134-143` |
| Lote de Acuerdo Marco | obligatorio; **no se comprueba existencia ni tipo** | `EsLote = Sí` | `:144-153` |

- Padre inexistente: «El IDExpediente indicado no se encuentra registrado» (`:98-101, 121-124`).
- El tipo se deriva de los indicadores en orden AM → Lote → Expediente → Basado; un basado es de AM o de Lote según el padre (`FUNCIONES UTILES.bas:5417-5450`). El texto sale de `m_ObjEntorno.ColTiposExpedientes` (`classes/Expediente.cls:1618-1640`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lectura del padre | vía `constructor.getExpediente` | DAO (no trazado aquí) | — | `:94, 117` |
| Cambio de tipo | `TbExpedientes`: `IDExpedientePadre`, `EsAM`, `EsLote`, `EsBasado`, `EsExpediente`, `TIpo` (vacío ⇒ `Null`) | SQL dinámico concatenado + `Recordset.Edit/Update` | ninguna | `classes/ExpedienteOperaciones.cls:2006-2032` |
| Caché de entidades | `RegistrarExpEntidades` (ámbito forzado a `Todo`) | DAO | ninguna | `classes/ExpedienteOperaciones.cls:2038, 607-621` |
| Último cambio | `RegistrarUltimoCambio` → `UltimoCambioOperaciones.Registrar` | no trazado aquí | ninguna | `classes/ExpedienteOperaciones.cls:2041`, `FUNCIONES UTILES.bas:189-215` |

`RegistrarCambioTipo` exige `IDExpediente` informado («El expediente no está registrado», `classes/ExpedienteOperaciones.cls:2000-2003`) y recalcula `TIpo` con `TipoCalculadoTexto` antes de escribir (`:2006`). Si la fila no existe, no escribe y sigue sin error (`:2014-2033`).

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: `WHERE IDExpediente=` concatenado sin parámetros (`classes/ExpedienteOperaciones.cls:2008-2010`); indicadores de tipo como texto «Sí»/«No» con tilde; tres escrituras (cabecera, caché, último cambio) sin transacción común.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- «Convertir a Lote» sin padre no cambia ningún indicador (`:93-111`) pero llama igualmente a `RegistrarCambioTipo` y cierra como éxito (`:157-166`): el usuario cree haber convertido el expediente.
- «Convertir a Lote de Acuerdo Marco» no comprueba que el padre exista ni que sea AM (`:144-153`), a diferencia de «Convertir a Lote» (`:102-105`).
- Ninguna rama impide que el padre sea el propio expediente (`:93-153`) si se escribe el ID a mano.
- Los indicadores se escriben en la global `m_ObjExpedienteActivo` antes de guardar (`:84-156`); si la validación o el guardado fallan, la global queda modificada.
- `ExpedientePadre` se memoriza en `m_ObjExpedientePadre` y nunca se vacía (`classes/Expediente.cls:1981-1983`; sin `Set m_ObjExpedientePadre = Nothing`): al cambiar el padre de un basado, `TIpo` (`classes/ExpedienteOperaciones.cls:2006`) podría calcularse con el padre anterior.
- Un tipo calculado fuera de los cuatro casos (p. ej. `EXPHPS`, `Variables Globales.bas:181`, o un basado sin padre, `FUNCIONES UTILES.bas:5439-5447`) deja el combo vacío (`:232-252`).
- Cabecera, caché y último cambio sin transacción: un fallo en `RegistrarExpEntidades` deja `TbExpedientes` ya actualizado (`classes/ExpedienteOperaciones.cls:2032-2039`).
- Rótulos de diseño incoherentes: `ComandoRegistrar` «Editar» y `ComandoAyuda` «Salir» (`.form.txt:877, 1213`); el mensaje de error cita `ComandoCambiarTipo_Click` (`:173`).
- `Form_Open` no cancela ante error (`:205-213`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: una sola operación transaccional de cambio de tipo con validación de padre en servidor; impedir el padre igual al propio expediente; ofrecer solo destinos válidos y mostrar error en vez de éxito silencioso.

## Criterios de aceptación

Propuestos, no aprobados:

1. Guardar sin destino muestra «Seleccione el cambio de tipo» y no escribe.
2. Convertir a Basado sin padre, o con un padre que no es AM ni Lote, se rechaza con el mensaje de la tabla de reglas.
3. Convertir a Expediente Independiente con padre informado se rechaza.
4. Convertir a Lote con padre que no es AM se rechaza.
5. Un AM o un Lote con derivados no puede abrir el cambio de tipo.
6. Tras un cambio correcto, la lista de Gestión se refresca y muestra el tipo nuevo.
7. Las opciones del combo dependen del tipo actual según la tabla de apertura.

## Pendientes

- Quién puede pulsar los botones de cambio de tipo en Gestión (habilitación por rol) y asignación de `m_TituloFormulario`.
- Contenido de `getColBusquedaParaCambio` (`FUNCIONES UTILES.bas:1214`) y de `UltimoCambioOperaciones.Registrar`.
- Diferencia funcional entre «Convertir a Lote» y «Convertir a Lote de Acuerdo Marco» y si un Lote sin padre es un estado válido.
- Si los derivados de un basado o de un AM convertido requieren reasignación (no se trata aquí).
