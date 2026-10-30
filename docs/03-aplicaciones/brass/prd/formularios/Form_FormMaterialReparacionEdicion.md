# Form_FormMaterialReparacionEdicion

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormMaterialReparacionEdicion.cls` (667 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Edición (o consulta, si está bloqueado) de un material marcado como reparación (`TbMaterial`). Es un formulario **no enlazado**: sin `RecordSource` en el `.form.txt`; los datos se cargan por DAO en `EstablecerDatos` (`Form_FormMaterialReparacionEdicion.cls:285-575`) y se guardan con `EdicionMaterialEnActividad` o `EdicionMaterialSinActividad` de `FG.bas` (`FG.bas` = `modules/Funciones Generales.bas`). Capacidad: CAP-7 Materiales, reparaciones y repuestos. Las citas son `Form_FormMaterialReparacionEdicion.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormMaterialesGestion` (`ComandoEdicionMaterialSinActividad_Click`) | Si la columna 6 de la lista es «Sí» (reparación): cierra el formulario si estaba abierto y lo abre con `OpenArgs = IDMaterial` solo | `Form_FormMaterialesGestion.cls:140-153` |

`Form_Open` admite dos formatos de `OpenArgs`: solo `IDMaterial`, o cuatro partes separadas por `|` (`IDMaterial|Franqueado|EstaFacturado|URLAnexo`); con otro número de partes o vacío, error de parámetros insuficientes y cierre (`Form_FormMaterialReparacionEdicion.cls:585-610, 628-634`). El único llamador localizado pasa solo el ID (`Form_FormMaterialesGestion.cls:153`; el formato de cuatro partes figura solo como comentario en `:583`).

Salidas: `cmdSalir` cierra (con pregunta de guardado si procede, `:58-97`); `ComandoRegistrar` guarda y cierra (`:207-244`); `ComandoAdjuntar` abre `FormAnexos` (`:98-132`); `ComandoVerEvento` abre `FormEventoGestion` (`:245-282`); `ImagenAnexoMaterial` abre el anexo local (`:637-660`).

## Controles y eventos

Controles: 51 (60 `Begin` en el `.form.txt`: 51 controles, `Form`, `Section` y 7 plantillas de estilo por tipo en `Form_FormMaterialReparacionEdicion.form.txt:147-263`).

| Control | Tipo | Evento y handler | Qué hace / origen de datos | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `ComandoRegistrar` («Registrar») | botón | `OnClick` → `ComandoRegistrar_Click` | Exige cambios y guarda; se deshabilita si franqueado o facturado | CAP-7 | `:207-244, 542, 556`; `Form_FormMaterialReparacionEdicion.form.txt:342-352` |
| `cmdSalir` («Salir») | botón | `OnClick` → `cmdSalir_Click` | Pregunta si guardar cuando hay cambios y escritura permitida; cierra sin guardar el formulario | CAP-7 | `:58-97`; `Form_FormMaterialReparacionEdicion.form.txt:367-378` |
| `ComandoAdjuntar` («Adjuntar») | botón | `OnClick` → `ComandoAdjuntar_Click` | Carga el `Material`, fija `m_ObjEntidadParaAnexoActiva` y abre `FormAnexos` pasando el nombre del formulario | CAP-11 | `:98-132`; `Form_FormMaterialReparacionEdicion.form.txt:1831-1842` |
| `ComandoVerEvento` («Ver Evento») | botón | `OnClick` → `ComandoVerEvento_Click` | Abre `FormEventoGestion` posicionado en `IDEVENTO`; deshabilitado en el diseño y sin actividad | CAP-3 | `:245-282, 403-411`; `Form_FormMaterialReparacionEdicion.form.txt:1566-1578` |
| `ImagenAnexoMaterial` | imagen | `OnClick` → `ImagenAnexoMaterial_Click` | Abre el anexo con `AbrirEnLocal`; visible solo si hay `strURLAnexo` | CAP-11 | `:499-505, 637-660`; `Form_FormMaterialReparacionEdicion.form.txt:1672-1683` |
| `ComboMaterial` | combo | — | Nombre del material; `RowSource` dinámico: materiales distintos de actividades del mismo equipo (o `'-1'` sin equipo) | CAP-7 | `:506-523`; `Form_FormMaterialReparacionEdicion.form.txt:394-417` |
| `PN`, `NS` | texto | — | Part number y número de serie; `#` si vacíos al guardar | CAP-7 | `:161-162, 469-474` |
| `COSTE` | texto | — | Coste; se muestra con `#,##0.00 €`; `#` si vacío | CAP-7 | `:163, 475-478`; `Form_FormMaterialReparacionEdicion.form.txt:662-676` |
| `FechaIncio` («Fecha Alta») | texto fecha | — | Fecha inicial (nombre con la errata del campo); obligatoria al guardar | CAP-7 | `:164, 492-495`; `FG.bas:6304-6307` |
| `FechaEntrega` («F.ENTREG») | texto fecha | — | Fecha de entrega o de reparación; opcional | CAP-7 | `:170, 482-485`; `Form_FormMaterialReparacionEdicion.form.txt:756-770` |
| `ComboReparadoPor` | combo | — | Valores distintos de `TbMaterial.ReparadoPor` | CAP-7 | `Form_FormMaterialReparacionEdicion.form.txt:692-712`; `:486-488` |
| `ComboTipoAccion` | combo | — | Valores de `TbTipoAccion.ESTADOREPARABLE` | CAP-7 | `Form_FormMaterialReparacionEdicion.form.txt:786-807`; `:479-481` |
| `ComboGarantia`, `ComboEsReparacion` | combo lista de valores `Sí;No` | — | Garantía y «Es Reparación»; este se normaliza a «No» si el valor guardado no es Sí/No | CAP-7 | `Form_FormMaterialReparacionEdicion.form.txt:1361-1382, 1455-1476`; `:362-365, 496-498` |
| `DESCRIPCION` | texto | — | Descripción libre | CAP-7 | `:169, 489-491` |
| `NODO`, `BUI`, `SUBSISTEMA`, `EQUIPO`, `IDEVENTO`, `IDActividad` | texto bloqueado (`Tag = FIJOS`) | — | Contexto de solo lectura obtenido de la actividad/evento o del equipo | CAP-12 | `:369-409, 428-439`; `Form_FormMaterialReparacionEdicion.form.txt:453-592, 1246-1261` |
| `lblFacturado`, `lblFranqueado` | etiqueta | — | Avisos, ocultos salvo facturado / franqueado | CAP-8 | `:352-356, 543, 557`; `Form_FormMaterialReparacionEdicion.form.txt:1397-1398, 1491-1492` |
| `lblTitulo1`, `lblTitulo2` | etiqueta | — | Título «EDICIÓN…» o «DETALLE DE MATERIAL.REPARACIÓN»; también la barra de título | sin asignar | `:531, 545, 559-560` |
| 17 etiquetas rótulo (`Etiqueta*`) | etiqueta | — | Rótulos de campo | sin asignar | `Form_FormMaterialReparacionEdicion.form.txt:823-1346` |
| 3 `OptionGroup` (`MarcoEvento`, `Marco240`, `Marco237`), 2 `Rectangle` (`CuadroSuperior`, `CuadroPie`) | marco / rectángulo | — | Decoración | sin asignar | `Form_FormMaterialReparacionEdicion.form.txt:296-340, 431, 931, 1170` |

## Flujos

**Apertura** (`Form_Open`, `:576-635`): `AjustarTamaño Me` (`:582`); interpreta `OpenArgs` (`:585-610`); `EstablecerDatos` (`:616`).

`EstablecerDatos`:
1. Lee `IDEquipo`, `IDActividad`, `EsReparacion`, `FechaEntrega` del material; si no existe, error (`:296-309`).
2. Con `OpenArgs` de un solo ID calcula franqueado (solo si hay actividad; si no, «No»), facturado y el anexo con `getMaterial(...).URLAnexos` (`:312-347`).
3. `strEsDeActividad = "Sí"` si hay `IDActividad` (`:357-361`). Con actividad lee evento, nodo, BUI, subsistema y equipo por `TbActividades`–`TbEventos`–`TbEquipos` (`:373-409`); sin actividad, por `TbEquipos.IDEquipo` (error si no existe, `:412-420`).
4. Carga el material completo, guarda los valores iniciales para comparar y rellena controles (`:444-498`); fija el `RowSource` de `ComboMaterial` (`:506-523`).
5. Franqueado o facturado ⇒ `strEscritura = "No"`: título «DETALLE…», controles combo/texto con `Tag` distinto de `FIJOS` bloqueados y con color de no editable, `ComandoRegistrar` deshabilitado y `lblFranqueado` visible; en otro caso, editable (`:525-558`).

**Guardado** (`ComandoRegistrar_Click`, `:207-244`): sin cambios ⇒ «No se ha registrado ningún cambio» (`:218-221`); `Registrar` delega en `EdicionMaterialEnActividad` si hay actividad y en `EdicionMaterialSinActividad` si no (`:171-197`); éxito ⇒ cierra (`:236`).

**Salida con cambios** (`cmdSalir_Click`, `:67-89`): si hay cambios y `strEscritura = "Sí"`, pregunta «¿Desea guardar los cambios producidos?»; con Sí ejecuta `Registrar`.

**Errores:** todos los handlers usan el patrón `Err.Raise 1000` y `MsgBox` crítico (`:91-96, 238-243`); `Form_Open` además cierra el formulario (`:634`).

## Reglas, validaciones y cálculos

- **Campos de `Registrar`:** los vacíos viajan como `#` (que borra el campo) salvo `ComboMaterial`, que viaja vacío (`:160-170`).
- **Material obligatorio** (`FG.bas:4390-4393` en la rama con actividad; la rama sin actividad tiene sus propias comprobaciones desde `FG.bas:6294`).
- **Coste:** si se informa debe ser numérico (`FG.bas:4394-4399`). **Fecha de entrega:** si se informa debe ser fecha (`FG.bas:4401-4406`). **Fecha inicial:** obligatoria (`FG.bas:4407-4410`). **Garantía** y **Es Reparación:** deben ser «Sí» o «No» (`FG.bas:4411-4418`).
- **Facturado:** la rama sin actividad rechaza con «Material ya facturado» si existe fila en `TbFacturaMaterialesInvolucrados` (`FG.bas:6316-6331`); la rama con actividad no lo comprueba (a verificar abajo).
- **Detección de cambios:** compara texto de los 10 campos con el valor inicial; fechas en `dd/mm/yyyy`, coste como `CDbl`; ante cualquier error devuelve «hay cambios» (`:25-56`).
- **Bloqueo:** `strEscritura = "No"` si franqueado o facturado (`:525-529`). Franqueado = `TbEventos.Franqueado` del evento de la actividad del material (`FG.bas:6111-6130`); facturado = existencia en `TbFacturaMaterialesInvolucrados` (`FG.bas:5818-5823`).
- **Campos de `EdicionMaterialEnActividad`:** `NS` vacío se guarda `Null` aunque no llegue `#` (`FG.bas:4439-4442`); `FechaEntrega` se escribe una segunda vez si `EsReparacion = "No"` (`FG.bas:4479-4487`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lectura inicial del material | `TbMaterial` | SQL dinámico concatenado (`IDMaterial` entre comillas) | — | `:296-299` |
| Datos de evento (con actividad) | `TbActividades`, `TbEventos`, `TbEquipos` | SQL dinámico concatenado | — | `:373-377` |
| Datos de equipo (sin actividad) | `TbEquipos` | SQL dinámico concatenado (`IDEquipo` numérico sin comillas) | — | `:412-415` |
| Material completo | `TbMaterial` | SQL dinámico concatenado | — | `:444-447` |
| `RowSource` de `ComboMaterial` | `TbMaterial`, `TbActividades`, `TbEventos`, `TbEquipos` | SQL dinámico concatenado (nombre de equipo entre comillas) | — | `:506-523` |
| Lectura por `Constructor.getMaterial` | `TbMaterial` | SQL dinámico concatenado | — | `Constructor.bas:1815-1840` |
| Franqueado | `TbEventos`, `TbActividades`, `TbMaterial` | SQL dinámico concatenado | — | `FG.bas:6112-6116` |
| Facturado | `TbFacturaMaterialesInvolucrados` | SQL dinámico construido por prefijo | — | `FG.bas:5758-5823` |
| Edición con actividad | `TbMaterial` | DAO `Edit`/`Update` | no | `FG.bas:4428-4489` |
| Comprobación de facturación y edición sin actividad | `TbFacturaMaterialesInvolucrados`, `TbMaterial` | SQL concatenado + DAO `Edit`/`Update` | no | `FG.bas:6316-6330, 6341-6399` |

## Coexistencia

Pendiente de `coexistencia.md`. Ya se ve: `Format` con `€` y `dd/mm/yyyy`; booleano de `Franqueado` leído con `If .Fields("Franqueado")` (`FG.bas:6121`); SQL con `IDMaterial` como texto entre comillas simples; `DISTINCT … ORDER BY` en `RowSource`; `Edit`/`Update` sobre `SELECT *` sin transacción; textos «Sí»/«No» como valores de columna en lugar de booleanos.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `EdicionMaterialEnActividad` no comprueba si el material está facturado (solo el bloqueo del formulario por `EstaFacturado`, `:525-529`), a diferencia de la rama sin actividad (`FG.bas:6316-6331`). A verificar si hay otra defensa.
- `HaHabidoCambios` devuelve cierto ante cualquier error (`:55-56`) y modifica `strCosteAlInicio` en cada llamada (`:44-46`).
- El nombre del control `FechaIncio` coincide con el campo (`:33, 164, 461`): la errata se propaga; no es un fallo pero condiciona la migración.
- Tras un guardado fallido en `cmdSalir_Click`, el handler de error no cierra el formulario ni restaura el cursor de espera de forma distinta a los demás (`:91-96`): queda abierto con el mensaje. A verificar si es el comportamiento buscado.
- `ComandoRegistrar_Click` indica «No se ha registrado ningún cambio» como error con `MsgBox` crítico (`:218-221, 243`), no como aviso.
- SQL concatenado con el nombre de equipo entre comillas (`:511`): un nombre con apóstrofo rompería el `RowSource`.
- La rama sin actividad escribe `FechaEntrega` sin condición `EsReparacion`, la de actividad la repite (`FG.bas:4479-4487`); revisar equivalencia entre ramas.
- Mensaje de error de `EstaFranqueado` etiquetado en el llamador al calcular `EstaFacturado` (`Form_FormMaterialesGestion.cls:161`): texto no coincide con la función invocada.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: una sola operación de guardado transaccional con la comprobación de facturado dentro; estado de bloqueo (franqueado/facturado) devuelto por el servidor y aplicado también al guardar.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir con un `IDMaterial` inexistente muestra «El formulario no se ha abierto con los datos suficientes» y cierra (`:301-304, 628-634`).
2. Material de actividad franqueada o facturado: título «DETALLE DE MATERIAL.REPARACIÓN», campos bloqueados y `ComandoRegistrar` deshabilitado (`:525-542`).
3. `lblFacturado` solo visible si está facturado; `lblFranqueado` solo si está bloqueado (`:352-356, 543, 557`).
4. Guardar sin cambios muestra «No se ha registrado ningún cambio» y no escribe (`:218-221`).
5. Guardar sin fecha inicial, con coste no numérico o garantía fuera de Sí/No se rechaza con el mensaje de la regla (`FG.bas:4394-4418`).
6. Vaciar `PN`, coste, descripción o fecha de entrega los deja nulos en el registro (`:160-170`; `FG.bas:4428-4489`).
7. Salir con cambios y escritura permitida pregunta si guardar; con «No» cierra sin escribir (`:67-89`).
8. `ComandoVerEvento` solo está activo con actividad y abre el evento correspondiente (`:401-411, 260-271`).
9. `ComandoAdjuntar` abre los anexos del material y el icono de anexo solo aparece con fichero asociado (`:98-132, 499-505`).

## Pendientes

- Contenido de `AjustarTamaño` (`:582`), `Dame` (`FG.bas:290`), `Ejecutar` (declaración API en `VG.bas:15`) y `FicheroAbierto` (`FG.bas:195`), propiedades y `ColCampos` de la clase `Material` (`:110, 346`).
- Quién cumplimenta el formato de `OpenArgs` de cuatro partes (`:583`): no hay llamador localizado.
- Procedimiento «Ver Evento»: en el `.form.txt` el `OnClick` aparece como `[Procedimiento de evento]` (`Form_FormMaterialReparacionEdicion.form.txt:1578`); comprobar que el binario lo enlaza igual.
- Referencias dinámicas: `Forms("FormEventoGestion").Controls("ComboIDEventoBusqueda")` y llamadas directas a `Form_FormEventoGestion.ComboIDEventoBusqueda_AfterUpdate` / `ComandoActualizarTodo_Click` (`:262-269`).
- Si `TbMaterial` tiene restricciones sobre `PN`/`NS`/`IDMaterial` (ERD de origen).
