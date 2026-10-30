# Form_FormOriginadorGestion

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormOriginadorGestion.cls` (218 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Mantenimiento del catálogo de originadores de evento (`TbOriginador`, un único campo `ORIGINADOR`): lista, alta por `InputBox` y baja con confirmación. Es un formulario sin `RecordSource`; la lista usa un `RowSource` propio (`Form_FormOriginadorGestion.form.txt:328`). Es gemelo de `Form_FormUbicacionGestion` (misma estructura, otra tabla). Las citas son `:línea` del `.cls` del formulario salvo indicación; `FG.bas` = `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_Form0BDOpcionesOtras` | Clic en la etiqueta `lblGestionOriginadores`: `DoCmd.OpenForm "FormOriginadorGestion"` | `Form_Form0BDOpcionesOtras.cls:95` |

No lanza eventos ni devuelve valores. `cmdSalir` cierra sin guardar (`:94-97`).

## Controles y eventos

Controles: 17 (17 `Begin` en el `.form.txt`): 1 formulario, 5 plantillas de estilo por defecto (Label, Rectangle, CommandButton, OptionGroup, ListBox; `Form_FormOriginadorGestion.form.txt:144-210`), 1 sección, 3 rectángulos, 1 lista, 3 botones, 1 grupo de opciones y 2 etiquetas.

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| Formulario | formulario | `OnOpen` → `Form_Open` | Ajusta tamaño, copia `lblTitulo1` al título de ventana | CAP-12 | `:81-93`; `Form_FormOriginadorGestion.form.txt:42` |
| `ListaOriginadores` | lista | `OnClick` → `ListaOriginadores_Click` | Guarda el originador seleccionado (columna 0) en `strOriginadorSeleccionado` y habilita o deshabilita `ComandoEliminar` | CAP-12 | `:98-106`; `Form_FormOriginadorGestion.form.txt:326-331` |
| `ComandoAlta` («Alta») | botón | `OnClick` → `ComandoAlta_Click` | Pide el nombre con `InputBox`, llama a `Alta` y refresca la lista | CAP-12 | `:5-37`; `Form_FormOriginadorGestion.form.txt:357-359` |
| `ComandoEliminar` («Eliminar») | botón | `OnClick` → `ComandoEliminar_Click` | Confirma, llama a `Eliminar` y refresca | CAP-12 | `:38-80`; `Form_FormOriginadorGestion.form.txt:434-436` |
| `cmdSalir` («Salir») | botón | `OnClick` → `cmdSalir_Click` | Cierra con `acSaveNo` | CAP-12 | `:94-97`; `Form_FormOriginadorGestion.form.txt:383-385` |
| `lblTitulo1`, `lblTitulo2` | etiquetas «ORIGINADORES» | — | Título | CAP-12 | `Form_FormOriginadorGestion.form.txt:464-486` |
| `Detalle`, `CuadroFormulario`, `CuadroSuperior`, `Cuadro2`, `Marco120` y 5 plantillas | sección, rectángulos, grupo de opciones, estilos | — | Solo diseño | sin asignar | `Form_FormOriginadorGestion.form.txt:233-410` |

## Flujos

### Apertura (`:81-93`)
Con `On Error Resume Next` (`:82`): cursor de espera, `AjustarTamaño` (`FG.bas:261-287`), título de ventana y repintado. Un fallo no se notifica.

### Alta (`:5-37`)
1. `InputBox("Nombre del Originador")`; vacío o cancelar ⇒ sale sin hacer nada (`:8-11`).
2. `Alta(nombre)` (`:107-155`): vacío ⇒ error (`:125-128`); busca el originador con SQL concatenado (`:129-132`); si existe, «Ya existe un Originador registrado con el nombre…» (`:134-137`); si no, `AddNew`/`Update` (`:138-141`).
3. Los errores vuelven como `#ERR|detalle` (`:154`); el llamador detecta `|` (`:21-25`) y muestra `MsgBox` crítico (`:31-36`).
4. Éxito: `ListaOriginadores.Requery` y fin del cursor de espera (`:26-28`). El texto de éxito (`:144`) no se muestra.

### Baja (`:38-80`)
1. Sin selección ⇒ «Se ha de seleccionar el Originador que desea eliminar» (`:44-47`).
2. Confirmación Sí/No con «No» por defecto (`:48-52`).
3. `Eliminar` (`:156-217`): comprueba con `Dame("TbEventos","Originador","Originador",valor)` (`:183`; `FG.bas:290-326`) si algún evento lo usa (`:189-192`); busca el registro (`:193-196`), si no existe error (`:198-201`) y lo borra con `.Delete` (`:202`).
4. Éxito: `Requery`, foco en `ComandoAlta` y deshabilita `ComandoEliminar` (`:67-69`). No vacía `strOriginadorSeleccionado`.

## Reglas, validaciones y cálculos

- **Obligatorio:** nombre no vacío (`:9, 125`).
- **Unicidad:** consulta de existencia por texto exacto en `TbOriginador` (`:129-137`); la comparación la resuelve el motor.
- **Integridad al borrar:** solo se comprueba `TbEventos.Originador` (`:165, 183-192`).
- **Confirmación** de la baja con botón por defecto «No» (`:48-49`).
- **Sin transacción**: cada operación es una única escritura.

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Rellenar la lista | `TbOriginador` | SQL guardada en `RowSource`, ordenada por `ORIGINADOR` | — | `Form_FormOriginadorGestion.form.txt:328` |
| Existencia en alta | `TbOriginador` | SQL dinámico concatenado con el nombre entre comillas | no | `:129-132` |
| Alta | `TbOriginador` | DAO `AddNew`/`Update` | no | `:138-141` |
| Uso en eventos | `TbEventos` | SQL dinámico concatenado, primero sin comillas y, si falla, entre comillas | no | `FG.bas:290-326` |
| Existencia en baja | `TbOriginador` | SQL dinámico concatenado | no | `:193-196` |
| Baja | `TbOriginador` | DAO `Delete` | no | `:202` |

## Coexistencia

Hay que adaptar: SQL concatenado con valor entre comillas simples (`:131, 195`); `Dame` genérico que prueba el valor sin comillas y recurre a comillas tras un error (`FG.bas:290-326`), que en PostgreSQL fallaría por columna inexistente en el primer intento; comparación de texto sin distinguir mayúsculas heredada de Jet/ACE (a verificar en el vinculado); `CurrentDb()` frente a la conexión vinculada; borrado físico sin restricción de clave foránea.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Un nombre con apóstrofo rompe la consulta de existencia (`:131`) y el alta falla con el mensaje genérico.
- Al responder «No» a la confirmación se sale con `Exit Sub` (`:50-52`) sin apagar el cursor de espera activado en `:42`.
- Mensajes copiados de otros formularios: el error genérico de baja habla de «ubicación» (`:76`) y el de nombre vacío en `Alta` habla de «nodo» (`:126`).
- `strOriginadorSeleccionado` no se vacía tras borrar (`:67-69`; en `Form_FormUbicacionGestion.cls:68` sí se hace).
- Solo se protege el borrado frente a `TbEventos`; otras tablas que guarden el originador podrían quedar huérfanas.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: alta en línea sin `InputBox`; restricción única e integridad referencial en base de datos; unificar este mantenimiento con el de ubicaciones en un componente de catálogo simple.

## Criterios de aceptación

Propuestos, no aprobados:

1. Dar de alta un originador nuevo lo deja en `ListaOriginadores` ordenado alfabéticamente (`:26`).
2. Dar de alta un nombre ya existente muestra «Ya existe un Originador registrado con el nombre…» y no escribe (`:134-137`).
3. Cancelar o dejar vacío el `InputBox` no hace nada (`:9-11`).
4. `ComandoEliminar` solo está habilitado con un originador seleccionado (`:101-105`).
5. Eliminar un originador usado en algún evento muestra «forma parte de algún evento» y no borra (`:189-192`).
6. Eliminar un originador libre pide confirmación y, al aceptar, lo quita de la lista (`:48-69`).
7. Cualquier error se muestra en un cuadro crítico y no deja el cursor de espera activo (`:78-79`).

## Pendientes

- Si otras tablas distintas de `TbEventos` referencian `ORIGINADOR` (ERD de origen).
- Longitud y colación de `ORIGINADOR` en el backend.
- Quién llama a `Alta` y `Eliminar` desde fuera del formulario (referencia dinámica no resuelta).
