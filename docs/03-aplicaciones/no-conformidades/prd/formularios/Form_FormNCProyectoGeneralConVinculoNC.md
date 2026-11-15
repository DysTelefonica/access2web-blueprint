# Form_FormNCProyectoGeneralConVinculoNC

> **Commit fijado:** `1d7a2f826a1e3c582874e4e29886a6183efb9067` · **Fuente:** `src/forms/Form_FormNCProyectoGeneralConVinculoNC.cls` (454 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Ventana «ALTA DE NO CONFORMIDAD DE PROYECTO» (`.form.txt:594`) para crear una NC de proyecto nueva **vinculada** a otra NC ya existente: parte de los datos de la NC a vincular (expediente, descripción, entidad responsable, detectado por) y permite ajustarlos antes de registrar. Es un formulario **no enlazado** (sin `RecordSource`) que usa los objetos globales `m_ObjNCProyectoActivaPVinculada` (NC de origen, `Variables Globales.bas:118`) y `m_ObjNCProyectoActiva` (la nueva). Las citas son `Form_FormNCProyectoGeneralConVinculoNC.cls:línea` (abreviado `:línea`) salvo indicación; `Ops` abrevia `classes/NCProyectoOperaciones.cls`. La pestaña de edición de una NC ya vinculada es `FormNCProyectoGeneral` (ficha del mismo lote).

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `FormNCProyectoGestion` | Asigna la NC seleccionada a `m_ObjNCProyectoActivaPVinculada`, cierra una instancia previa y abre este formulario | `Form_FormNCProyectoGestion.cls:44-48` |

Abre `FormExpedientesBusqueda` (`:31`, `:85`, `:119`) y escucha su evento `Seleccionado` (`:5`, `:248-270`). Se cierra con «Salir» (`:8-11`) o tras registrar (`:197`), con `acSaveNo`; antes actualiza `FormNCProyectoGestion` si está abierto (`:194-196`).

## Controles y eventos

`.form.txt` declara 37 bloques `Begin`: el formulario (`:5`), 8 de valores por defecto, 3 secciones (`FormHeader` `:573`, `Section` `:614`, `FormFooter` `:1412`) y 25 controles.

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `lblTitulo` | etiqueta | — | Se reescribe con el código de la NC a vincular | `:231`, `.form.txt:583` |
| `PalabraClave` | texto | `OnKeyDown` | Intro/Tab lanza la búsqueda | `:315-333`, `.form.txt:641` |
| `ComandoBuscarExpediente` | botón | `OnClick` | Busca por palabra clave; vacía ⇒ abre la búsqueda | `:13-104`, `.form.txt:729` |
| `ComandoIrABusquedaExpedientes` | botón | `OnClick` | Abre la búsqueda filtrada | `:107-142`, `.form.txt:807` |
| `ComandoActualizararExpedientes` | botón, deshabilitado en diseño | sin `OnClick` | Se habilita o no tras buscar sin resultado | `.form.txt:856-857`, `:53-57` |
| `IDExpediente`, `Juridica`, `Titulo` | texto bloqueado (`IDExpediente` con `Tag="DATO"`) | — | Datos del expediente | `.form.txt:901-963`, `:373-375` |
| `RESPONSABLECALIDAD`, `ResponsableTelefonica` | texto bloqueado | — | Responsables que vienen del expediente | `.form.txt:1087, 1319`, `:376-381` |
| `Descripcion` | texto, `DATO;OBLIGATORIO` | `BeforeUpdate` | Repinta bordes | `.form.txt:1024-1040`, `:335-339` |
| `EntidadResponsable`, `DetectadoPor` | combos de lista de valores, `DATO;OBLIGATORIO` | `BeforeUpdate` | Selección cerrada | `.form.txt:1136-1163, 1243-1270`, `:342-351` |
| `FechaApertura` | texto, `DATO;OBLIGATORIO` | — | Se fija a hoy | `.form.txt:1349-1363`, `:303` |
| `ComandoRegistrar` | botón | `OnClick` | Crea la NC vinculada | `.form.txt:1495-1508`, `:144-212` |
| `cmdSalir` | botón | `OnClick` | Cierra sin guardar | `.form.txt:1422-1431`, `:8-11` |
| `Cuadro53` y 7 etiquetas (`Etiqueta43`, `44`, `49`, `50`, `59`, `249`, `250`, `252`) | rectángulo y rótulos | — | Decorativos | `.form.txt:658-1379` |

## Flujos

### Apertura (`Form_Load`, `:214-246`)

1. Sin NC a vincular ⇒ «No se ha indicado primero la NC de proyecto a vincular» (`:221-223`).
2. Vacía `m_ObjNCProyectoActiva`, ajusta tamaño (`Ajustar Me`) y llama a `EstablecerDatos` (`:225-229`), que bloquea la edición a técnicos (`:279-283`), copia expediente, descripción, entidad y detectado por de la NC de origen (`:284-302`), pone la fecha de apertura a hoy y pinta obligatorios (`:303-306`). Título: «ALTA DE NO CONFORMIDAD DE PROYECTO VINCULADA A » + código (`:231`).

### Registrar (`ComandoRegistrar_Click`, `:144-212`)

1. Técnico ⇒ «No tiene autorización para esa acción» (`:155-157`).
2. Crea un `NCProyecto` nuevo y copia expediente (jurídicas, código, nemotécnico, responsable de Calidad y jefe de proyecto como responsable de Telefónica), descripción, detectado por, identificador y código de la NC asociada y fecha de apertura (`:159-180`). La entidad responsable **no** se copia en este bloque (a verificar).
3. `RegistrarAltaDatosUnicosConVinculoNC` (`:181-188`; `Ops:594-714`).
4. Mensaje «Registro correcto de la NC de Proyecto … asociada con la NC …» (`:192-193`), actualiza la lista de gestión y cierra (`:194-197`).

### Búsqueda de expediente (`:13-104`)

Palabra clave vacía ⇒ abre la búsqueda (`:27-37`); sin resultados ⇒ limpia el expediente y abre la búsqueda (`:48-60`); un resultado ⇒ lo asigna y vacía la palabra clave (`:66-75`); varios ⇒ abre la búsqueda con la palabra clave (`:82-90`). `EstablecerDatosExpediente` (`:354-394`) rellena identificador, título, jurídicas, responsable de Calidad y jefe de proyecto, y da foco a `Descripcion`.

### Errores

Error 1000 ⇒ advertencia con `m_Error`; otro ⇒ `CorreoAlAdministrador` y aviso crítico (`:95-103`, `:203-211`, `:237-245`).

## Reglas, validaciones y cálculos

- Rol: técnico sin edición y sin registro (`:155`, `:279-283`).
- Obligatorios: solo se señalan con borde (`RellenarBordeCamposObligatorios`, `:396-454`: expediente, descripción, entidad, detectado por, fecha); `ComandoRegistrar_Click` no los valida (`:155-188`); la validación efectiva queda en `RegistrarAltaDatosUnicosConVinculoNC` (no leída por completo, ver Pendientes).
- La NC nueva toma el identificador de `IDNoConformidadCalculado` (`Ops:622`).

## Lecturas y escrituras

| Operación | Tabla o fuente | Tipo | Evidencia |
|---|---|---|---|
| Resultados de búsqueda | `constructor.getExpedientesBusqueda` | SQL en `constructor` (no leído aquí) | `:42-44` |
| Alta de la NC vinculada | `TbNoConformidades` | DAO `OpenRecordset/AddNew/Update` | `Ops:632-695` |
| Registro de cambios | `RegistrarLogProyecto` («Alta de NC de Proyecto Vinculada a …») | ver clase | `Ops:701` |

Sin transacción visible en el formulario.

## Coexistencia

El alta calcula el identificador en cliente (`IDNoConformidadCalculado`, `Ops:622`) y escribe por recordset DAO: adaptar a secuencia o `INSERT … RETURNING` en PostgreSQL. La búsqueda de expedientes recibe texto libre: parametrizar. Las listas de entidad y detectado por están en el diseño del formulario (`.form.txt:1158-1161, 1265-1268`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `EntidadResponsable` se muestra y puede editarse, pero `ComandoRegistrar_Click` no la traslada a la NC nueva (`:159-180`).
- `ComandoRegistrar_Click` sustituye `m_ObjNCProyectoActiva` por una NC nueva y no la restaura si el alta falla (`:159`, `:185-187`).
- `PalabraClave_KeyDown` no comprueba permiso de edición (comprobación comentada, `:320-322`).
- `ComandoActualizararExpedientes` sin `OnClick` (`.form.txt:856`).
- `Form_Load` deja el formulario abierto y utilizable tras un error (`:237-245`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: unificar con el formulario general mediante un modo «alta vinculada»; validar obligatorios antes de registrar.

## Criterios de aceptación

Propuestos, no aprobados:

1. Solo se abre desde la gestión con una NC seleccionada; sin ella avisa (`:221-223`).
2. Muestra expediente, descripción, entidad y detectado por de la NC de origen y la fecha de hoy.
3. Un técnico no puede registrar.
4. Buscar con palabra clave de un resultado único lo asigna; con varios o ninguno abre la búsqueda.
5. Registrar crea una NC nueva asociada a la de origen con código propio y lo comunica.
6. Tras registrar la ventana se cierra y la lista de gestión refleja la NC nueva.
7. «Salir» cierra sin crear nada.

## Pendientes

- `RegistrarAltaDatosUnicosConVinculoNC` (`Ops:594-714`): campos obligatorios y escrituras exactas no leídas por completo.
- `constructor.getExpedientesBusqueda` y `Ajustar` (`FG.bas:615`): comportamiento no leído aquí.
