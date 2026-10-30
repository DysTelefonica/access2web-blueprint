# Form_FormMaterialRepuestoEdicion

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormMaterialRepuestoEdicion.cls` (631 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Edición (o consulta, si está bloqueado) de un material usado como repuesto en una actividad, registro de `TbMaterial`. Es un formulario **no enlazado** (ningún control con `ControlSource`; el `.form.txt` solo declara `RowSource` en los combos): carga los datos por DAO en `Form_Open` y persiste con `EdicionMaterialEnActividad` (`FG.bas` = `modules/Funciones Generales.bas`, líneas 4350-4504). Las citas son `Form_FormMaterialRepuestoEdicion.cls:línea` salvo indicación; se abrevian a `:línea`.

Controles: 53 (62 `Begin` en el .form.txt: 1 formulario, 7 plantillas de estilo por defecto, 1 sección `Detalle` y 53 controles).

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormMaterialesGestion` (rama repuesto) | Cierra el formulario si estaba abierto y lo abre con `OpenArgs` = `IDMaterial\|franqueado\|facturado\|URL del anexo` | `Form_FormMaterialesGestion.cls:188-192` |

`OpenArgs` debe tener exactamente 4 partes (`:344-348`); `IDMaterial` no vacío (`:354`), «facturado» y «franqueado» solo `Sí`/`No` (`:358-365`). Cualquier fallo: «No se han cargado todos los datos» y cierre (`:341-347, 388-394`). Salidas: `Requery` de `ListaMateriales` en `FormEventoGestion` si está abierto (`:77-79, 128-130, 192-194`); abre `FormAnexos` con la global `m_ObjEntidadParaAnexoActiva` (`:145-148`); abre/enfoca `FormEventoGestion` (`:276-287`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `ComboMaterial` | combo | `OnChange` → `ComboMaterial_Change` | Material existente del equipo (3 columnas: Material, PN, NS); al elegir vacía `Material` y copia PN y NS (cols. 1 y 2) | CAP-7 Materiales, reparaciones y repuestos | `:299-328`; `.form.txt:423-451` |
| `Material` | texto | — | Material «no existente»; solo se lee en `HaHabidoCambios` | CAP-7 | `:25, 313`; `.form.txt:1558` |
| `PN`, `NS` | texto | — | Número de parte / serie; editables | CAP-7 | `:26-27, 502-507` |
| `COSTE` | texto | — | Coste con formato `#,##0.00 €` | CAP-7 | `:508-510`; `.form.txt:558` |
| `ComboReparadoPor` | combo | — | Suministrador: valores distintos de `TbMaterial.ReparadoPor` | CAP-7 | `.form.txt:593-597` |
| `ComboTipoAccion` | combo | — | `TbTipoAccion.ESTADOREPARABLE` | CAP-7 | `.form.txt:630-634` |
| `ComboGarantia` | combo | — | Lista de valores `Sí;No` | CAP-7 | `.form.txt:1043-1045` |
| `FechaIncio` | texto | — | Fecha («Fecha» en pantalla); obligatoria al guardar | CAP-7 | `:516, 526`; `FG.bas:4407-4410` |
| `OBSERVACIONES` | texto | — | Descripción libre | CAP-7 | `:493, 522` |
| `IDMATERIAL`, `IDEvento`, `IDActividad`, `BUI`, `SubSistema`, `Equipo` | texto `FIJOS` | — | Datos de contexto de solo lectura | CAP-7 | `:448-466, 538-544` |
| `ComandoRegistrar` | botón | `OnClick` | Guarda y cierra; se deshabilita en solo lectura | CAP-7 | `:160-206, 556, 569` |
| `cmdSalir` | botón | `OnClick` | Pregunta si guardar (solo en modo escritura) y cierra | CAP-7 | `:51-93` |
| `ComandoAdjuntar` | botón | `OnClick` | Guarda si hay cambios y abre `FormAnexos` del material | CAP-11 Anexos | `:94-159` |
| `ComandoVerEvento` | botón | `OnClick` | Abre `FormEventoGestion` en el evento origen; deshabilitado si no hay evento | CAP-3 Eventos | `:261-298, 460-462` |
| `ImagenAnexoMaterial` | imagen | `OnClick` | Visible solo con URL de anexo; con `;` delega en Adjuntar, si no copia a local y abre | CAP-11 | `:531-536, 592-626` |
| `lblFacturado`, `lblFranqueado` | etiqueta | — | Avisos según los argumentos | CAP-8 Facturación, gastos y subcontratación | `:420-429, 557, 570` |
| `lblTitulo1`, `lblTitulo2` | etiqueta | — | Título «EDICIÓN»/«DETALLE DE MATERIAL COMO REPUESTO» y `Caption` | CAP-7 | `:545-576` |
| 3 `OptionGroup`, 2 `Rectangle`, 23 `Label` rotulares (incluye 7 «*») | marcos y rótulos | — | Decoración y rótulos | sin asignar | `.form.txt:300-1534` |

Recuento: 4 combos + 12 textos + 4 botones + 1 imagen + 2 etiquetas de aviso + 2 títulos + 3 marcos + 2 rectángulos + 23 rótulos = 53 (27 etiquetas en total).

## Flujos

### Apertura (`Form_Open`, `:329-395`; `EstablecerDatos`, `:396-591`)

1. `strEscritura` = «No» si el material está facturado o franqueado, «Sí» en otro caso (`:366-370`).
2. Lee `IDActividad` e `IDEquipo` de `TbMaterial` (`:405-418`), luego evento, nodo, BUI, subsistema y equipo (`:433-447`) y rellena los controles `FIJOS` (`:448-466`).
3. Carga `ComboMaterial.RowSource` con los materiales distintos del mismo equipo (`:467-473`) y los valores del registro (`:477-530`), guardando copias «al inicio» (`:486-495`).
4. Modo lectura: título «DETALLE…», combos y textos bloqueados, `ComandoRegistrar` deshabilitado (`:545-557`). Modo edición: campos desbloqueados (`:558-571`).

### Guardado (`ComandoRegistrar_Click`, `:160-206`; `Registrar`, `:208-257`)

1. Sin cambios (`HaHabidoCambios`, `:7-50`): «No ha habido cambios» (`:172-175`).
2. `Registrar` envía a `EdicionMaterialEnActividad` los valores con `Nz` (vacío ⇒ `#` = poner `Null`, `:228-236`), `EsReparacion` = «No» (`:242-243`).
3. Tras éxito refresca `ListaMateriales` y cierra (`:192-195`). Errores: `MsgBox` crítico y el formulario sigue abierto (`:200-206`).

### Alternativos

- `cmdSalir`: con cambios y modo escritura pregunta «¿Desea guardar los cambios producidos?»; «No» cierra sin guardar (`:55-85`).
- `ComandoAdjuntar`: sin `strIDMaterial` ⇒ error (`:101-104`); con cambios registra primero (`:111-132`); `getMaterial` (`Constructor.bas:1815`) y apertura de `FormAnexos` (`:133-148`).
- Anexo: `AbrirEnLocal` comprueba que exista, copia al directorio local configurado y lo abre (`FG.bas:218-258`).

## Reglas, validaciones y cálculos

- Material obligatorio: vacío o `#` ⇒ «Se ha de indicar el material de que se trata» (`FG.bas:4390-4393`).
- Coste numérico si se indica (`FG.bas:4394-4399`); fecha inicial obligatoria y válida (`FG.bas:4407-4410`); garantía solo `Sí`/`No` (`FG.bas:4411-4414`).
- `PN` vacío no se modifica; `NS` vacío o `#` se pone a `Null` (`FG.bas:4432-4443`).
- `COSTE`, observaciones, suministrador y tipo de acción: `#` ⇒ `Null` (`FG.bas:4444-4478`).
- Bloqueo por facturado/franqueado: `:366-370, 545-557`.

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Actividad y equipo del material | `TbMaterial` | SQL dinámico concatenado (`IDMaterial` entre comillas) | — | `:405-408` |
| Evento, nodo, BUI, subsistema y equipo | `TbActividades`, `TbEventos`, `TbEquipos` | SQL dinámico concatenado | — | `:433-437` |
| Origen de `ComboMaterial` | `TbMaterial`, `TbActividades`, `TbEventos` | SQL dinámico con `DISTINCT` por equipo | — | `:467-472` |
| Lectura del registro | `TbMaterial` | SQL dinámico concatenado | — | `:477-480` |
| Edición | `TbMaterial` | DAO `Edit`/`Update` | no | `FG.bas:4415-4490` |
| Origen de otros combos | `TbMaterial`, `TbTipoAccion` | SQL guardada en `RowSource` | — | `.form.txt:595, 632` |

## Coexistencia

Pendiente de `coexistencia`. Ya se ve: `SELECT DISTINCT` con `ORDER BY`, comillas simples y paréntesis anidados de Jet en `JOIN` (`:467-471`); `Format` con `€` y fecha `dd/mm/yyyy` (`:509, 515`); `Nz`, `IsNumeric`/`IsDate` dependientes de la configuración regional; sin transacción en la edición.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `FechaEntrega` recibe el valor de `FechaIncio`: el último argumento de la llamada es `strFechaInicial` (`:242-243`, aplicado en `FG.bas:4451-4452`), y la lectura de `FechaEntrega` se sobrescribe en pantalla con `FechaIncio` (`:514-527`).
- El texto `Material` (no existente) se compara en `HaHabidoCambios` (`:25, 40`) pero `Registrar` solo envía `ComboMaterial` (`:228`): un material escrito a mano daría «Se ha de indicar el material» (`FG.bas:4390`).
- Mensaje erróneo «El método Eliminar ha devuelto un error» en `ComandoAdjuntar_Click` (`:120`).
- `Coste` se muestra con `€` (`:509`) y se reenvía como texto con ese símbolo (`:229`); depende de que `IsNumeric` lo acepte (`FG.bas:4394-4399`).
- Consulta concatenada con `IDMaterial` entre comillas (`:405-407, 477-479`) y `IDEquipo` sin comillas (`:470`).
- `ComandoAdjuntar` guarda sin comprobar `strEscritura` (`:111-117`); en modo lectura los campos están bloqueados (`:547-555`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: unificar `Material` y `ComboMaterial`; enviar `FechaEntrega` explícita; parametrizar consultas.

## Criterios de aceptación

Propuestos, no aprobados:

1. Un material facturado o franqueado se abre como «DETALLE…», con campos bloqueados y sin poder guardar (`:366-370, 545-557`).
2. Elegir un material en `ComboMaterial` copia su PN y NS (`:305-311`).
3. Guardar sin cambios muestra «No ha habido cambios» (`:172-175`).
4. Guardar sin fecha o con garantía distinta de `Sí`/`No` rechaza la operación (`FG.bas:4407-4414`).
5. Vaciar PN, NS o coste lo deja en blanco en `TbMaterial` (`:228-232`; `FG.bas:4432-4448`).
6. `Ver Evento` abre la gestión del evento origen; sin evento está deshabilitado (`:261-287, 460-462`).
7. Tras guardar, la lista de materiales del evento se refresca (`:192-194`).

## Pendientes

- Contenido de `AjustarTamaño` (`:335`) y de `Ejecutar` (`FG.bas:245`).
- Definición de `lngColorCampoNoEditable`/`lngColorCampoEditable` (`:540, 563`).
- Si `Form_FormEventoGestion.ComandoActualizarTodo_Click` es público y su efecto (`:283`).
- Índices, longitudes y nulabilidad de `TbMaterial` (ERD de origen).
