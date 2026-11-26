# Form_FormNCProyectoControlEficacia

> **Commit fijado:** `1d7a2f826a1e3c582874e4e29886a6183efb9067` · **Fuente:** `src/forms/Form_FormNCProyectoControlEficacia.cls` (448 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Pestaña «Control de eficacia» de la NC de proyecto: planificación del control (fecha prevista y descripción) y registro de su resultado (fecha, resultado y conformidad). Se incrusta en `FormNCProyecto` (usa `Me.Parent`, `:138,250`). Es **no enlazado** (sin `RecordSource`) y opera sobre `m_ObjNCProyectoActiva` y sobre la copia de partida `Form_FormNCProyecto.m_NCAlInicio`. Las citas son `Form_FormNCProyectoControlEficacia.cls:línea` salvo indicación; `NCProyectoOperaciones.cls` abrevia `classes/NCProyectoOperaciones.cls`. Todos los datos pertenecen a la propia NC (tabla `TbNoConformidades`), no hay entidad aparte (`NCProyectoOperaciones.cls:1072-1075`).

## Entrada y salida

Se carga como detalle de `FormNCProyecto` con `Form_Load` ⇒ `EstablecerDatos` (`:288-316`). Exige NC activa con identificador («… parámetros insuficentes», `:320-326`). Desde el formulario se puede saltar a la pestaña de acciones: `Me.Parent.NavAcciones.SetFocus` y `FrmDetalle.SourceObject = "FormNCProyectoAcciones"` (`:137-139`). Refresca la línea de `FormNCProyectoGestion` si cambia el estado (`:158-161`).

## Controles y eventos

`.form.txt` declara 23 bloques `Begin`: el formulario, 8 valores por defecto, la sección `Detalle` (`.form.txt:299`) y los 13 controles de la tabla.

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `RequiereControlEficacia` | texto bloqueado, `DATO` | `AfterUpdate` declarado; campo bloqueado | Se fija a «Sí» al cargar | `:347,377-383`, `.form.txt:706,719` |
| `FechaPrevistaControlEficacia` | texto, `DATO` | `AfterUpdate` | Fecha prevista; repinta bordes | `:209-217`, `.form.txt:322-326` |
| `ControlEficacia` | texto, `DATO` | `AfterUpdate` | Descripción del control; repinta bordes | `:191-198`, `.form.txt:417-420` |
| `FechaControlEficacia` | texto, `DATO` | `AfterUpdate` | Fecha del resultado; repinta bordes | `:200-207`, `.form.txt:479-483` |
| `ResultadoControlEficacia` | texto, `DATO` | `AfterUpdate` | Resultado; repinta bordes | `:442-448`, `.form.txt:580-583` |
| `ConformeControlEficacia` | combo «Sí;No», `DATO` | `AfterUpdate` | Conformidad; repinta bordes | `:182-189`, `.form.txt:650-656` |
| `ComandoGrabar` | botón | `OnClick` | Valida y graba | `:36-180`, `.form.txt:540-542` |
| `Etiqueta254`, `255`, `256`, `264`, `63`, `65` | etiquetas | — | Rótulos de los seis campos | `.form.txt:353-686` |

## Flujos

### Carga (`EstablecerDatos`, `:316-375`)
Con la NC sin NC vinculada y usuario no técnico habilita «Grabar»; en otro caso (técnico o NC vinculada) pone `AllowEdits = False` y lo deshabilita (`:329-340`). Vuelca los datos de la NC o de su vinculada (`:341-362`) y pinta bordes si no es técnico (`:364-365`).

### Grabar (`:36-180`)
1. Rechaza técnicos («No tiene autorización para esa acción») y, si no hay cambios en planificación ni en resultado, «No hay cambios que guardar» (`:49-65`).
2. Copia los campos `DATO` al objeto NC con `SetPropiedad` (`:10-27,66`).
3. Con fecha prevista informada y cambios en la planificación llama a `RegistrarControlEficacia` (`:75-77`); sin fecha prevista pero con fecha prevista previa llama a `EliminarResultadoControlEficacia` y salta al cierre (`:83-90`).
4. Con fecha de control informada y cambios en el resultado llama a `RegistrarResultadoControlEficacia` (`:95-97`); sin fecha y con fecha previa, `EliminarResultadoControlEficacia` (`:104-107`).
5. Si falla alguna operación, recarga la NC y muestra el motivo (`:78-81`).
6. Recarga `m_NCAlInicio`, recalcula el estado (`:118-131`). Con «No» conforme y control requerido pregunta «¿Desea planificar una nueva acción?» y, si se acepta, abre la pestaña de acciones (`:133-141`). Avisa del cierre o reapertura de la NC (`:146-157`).

### Errores
Error 1000 ⇒ aviso del motivo; otro ⇒ correo al administrador y mensaje crítico (`:173-178`); en `Form_Load` el patrón es el mismo (`:307-312`).

## Reglas, validaciones y cálculos

- Planificación: exige fecha prevista y texto del control (`NCProyectoOperaciones.cls:1057-1064`); marca «Requiere» = «Sí» (`:1065`).
- Resultado: exige requerir control, fecha prevista y control definidos, fecha de resultado, resultado y conformidad «Sí»/«No» (`NCProyectoOperaciones.cls:1202-1230`); rechaza si la NC marca que no requiere control (`NCProyectoOperaciones.cls:1206-1209`).
- Conforme «Sí» solo si hay acciones, todas con tareas y todas las tareas finalizadas (`NCProyectoOperaciones.cls:1066-1070,1231-1243`).
- Bordes obligatorios: prevista y descripción siempre; con fecha de resultado, resultado y conformidad; sin ella, esos dos se marcan si están rellenos (`:255-282`).
- Color de la pestaña según `EficaciaOK` (`:249-253`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Registrar planificación | `TbNoConformidades` (`RequiereControlEficacia`, `FechaPrevistaControlEficacia`, `ControlEficacia`, `Estado`) | DAO, SQL dinámico concatenado por `IDNoConformidad` + `Edit` | `NCProyectoOperaciones.cls:1072-1089` |
| Registrar resultado | `TbNoConformidades` (`FechaControlEficacia`, `ResultadoControlEficacia`, `ConformeControlEficacia`, `Estado`) | ídem | `NCProyectoOperaciones.cls:1245-1262` |
| Borrar resultado | `TbNoConformidades` (los tres campos anteriores a `Null`) | ídem | `NCProyectoOperaciones.cls:1302-1319` |
| Log y caché | `RegistrarLogProyecto`, `CacheNCProyecto.InvalidarCache` | no resuelto | `NCProyectoOperaciones.cls:1093-1095` |

Sin transacción: actualización, log y recálculo son pasos separados.

## Coexistencia

SQL con identificador concatenado (`NCProyectoOperaciones.cls:1072-1074`) ⇒ parámetros. Campos de fecha asignados desde texto y borrados con `Null` (`:1151-1155,1313-1315`): verificar tipos en PostgreSQL. Los cierres se sustentan en estado derivado (`EstadoCalculadoTexto`) que se persiste; mantener un único cálculo.

## Comportamiento actual · Posibles defectos · Propuestas

**Comportamiento actual.** «Requiere control de eficacia» no es editable: el formulario siempre lo muestra como «Sí» (`:347`, `.form.txt:706`). La baja del control completo (`EliminarControlEficacia`, `NCProyectoOperaciones.cls:1112`) existe pero este formulario no la invoca.

**Posibles defectos (a verificar).**
- Vaciar la fecha prevista ejecuta `EliminarResultadoControlEficacia` (`:85`), que solo borra el resultado (`NCProyectoOperaciones.cls:1313-1315`); fecha prevista y descripción permanecen en base.
- `HaHabidoCambios*` compara texto con fechas (`:395-401,424-430`); formatos distintos podrían dar cambios falsos.
- `RegistrarControlEficacia` comprueba conformidad usando valores que el formulario ya copió (`NCProyectoOperaciones.cls:1066`), por lo que un conforme se valida en la planificación.
- `Err.Raise 100` en lugar de 1000 (`:68`) muestra el mensaje técnico y avisa al administrador.

**Propuestas web pendientes de decisión.** Exponer la baja del control; separar planificación y resultado en dos acciones.

## Criterios de aceptación

1. Sin NC activa o sin identificador, aviso y no carga (`:320-326`).
2. Un técnico o una NC vinculada ven el formulario en solo lectura (`:329-340`).
3. Grabar sin cambios muestra «No hay cambios que guardar» (`:61-64`).
4. Sin fecha prevista o sin descripción no se registra la planificación (`NCProyectoOperaciones.cls:1057-1064`).
5. El resultado exige fecha, resultado y conformidad (`NCProyectoOperaciones.cls:1219-1230`).
6. Conforme «Sí» se rechaza si hay tareas sin finalizar (`NCProyectoOperaciones.cls:1231-1243`).
7. Conforme «No» ofrece planificar una acción nueva (`:133-141`).

## Pendientes

- Cuerpos de `SetPropiedad`, `EficaciaOK`, `TodasLasArsFinalizadas`, `AlgunaACSinAR`, `RequiereControlEficaciaCalculado` y `RegistrarLogProyecto`.
- Cuándo y cómo `FormNCProyecto` incrusta este formulario (`:138-139`).
