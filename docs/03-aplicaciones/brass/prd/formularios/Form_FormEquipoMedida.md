# Form_FormEquipoMedida

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormEquipoMedida.cls` (206 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Alta y edición de un equipo de medida (`TbEquiposMedida`: nombre, número de serie, número de parte, fechas de servicio, marca, modelo y descripción). Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): lee el equipo de la global `m_ObjEquipoMedidaActivo` (`Variables Globales.bas:253`) y persiste con `EquipoMedida.Registrar` (`classes/EquipoMedida.cls:294-467`). Los controles con `Tag` «DATOS» se vuelcan al objeto y desde él por nombre (`getPropiedad`/`SetPropiedad`). Las citas son `Form_FormEquipoMedida.cls:línea` salvo indicación. Capacidad de todos los controles: CAP-5 Equipos, equipos de medida y calibraciones.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormEquiposMedidaGestion` (alta) | `m_ObjEquipoMedidaActivo = Nothing`, cierra el formulario si estaba abierto y lo abre; guarda la referencia en `m_FormEquipo` (`WithEvents`) | `Form_FormEquiposMedidaGestion.cls:66-80, 6` |
| `Form_FormEquiposMedidaGestion` (edición) | `m_ObjEquipoMedidaActivo = m_EquipoSeleccionado` y lo abre | `Form_FormEquiposMedidaGestion.cls:132-158` |

Eventos hacia quien lo abrió: `EquipoAlta` (`:8`, lanzado en `:150`) y `EquipoModificado` (`:7`, lanzado en `:152`); la gestión refiltra su lista en ambos (`Form_FormEquiposMedidaGestion.cls:398-445`). Tras guardar se cierra solo (`:157`); `cmdSalir` cierra sin guardar (`:168-171`).

## Controles y eventos

Controles: 27 en la sección Detalle (36 `Begin` en el `.form.txt`: 1 `Form`, 7 plantillas de estilo por defecto en `Form_FormEquipoMedida.form.txt:140-243`, 1 `Section` y los 27).

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `Nombre`, `NS`, `PN` | texto, `Tag` «DATOS» | — | Obligatorios al guardar (nombre, número de serie, número de parte) | `Form_FormEquipoMedida.form.txt:475-476, 623-624, 701-702`; `classes/EquipoMedida.cls:323-334` |
| `FechaInicioServicio` | texto, formato «Short Date», `Tag` «DATOS» | — | Obligatoria al guardar | `Form_FormEquipoMedida.form.txt:780-781`; `classes/EquipoMedida.cls:335-338` |
| `FechaFinServicio` | texto, formato «Short Date», `Tag` «DATOS» | — | Opcional; si existe, no puede ser anterior al inicio | `Form_FormEquipoMedida.form.txt:859-860`; `classes/EquipoMedida.cls:352-357` |
| `Marca`, `Modelo`, `Descripcion` | texto, `Tag` «DATOS» | — | Opcionales | `Form_FormEquipoMedida.form.txt:288, 916, 972` |
| `Numero_Actividad` | texto con `ControlSource` «IDActividad», `Enabled`/`Locked` distintos del valor por defecto | — | Sin `Tag`: el código no lo lee ni lo escribe | `Form_FormEquipoMedida.form.txt:305-306, 317-318` |
| `ComandoRegistrar` | botón «Registrar» | `OnClick` → `ComandoRegistrar_Click` | Guarda y cierra | `:120-166`; `Form_FormEquipoMedida.form.txt:395-397` |
| `cmdSalir` | botón «Salir» | `OnClick` → `cmdSalir_Click` | Cierra sin guardar | `:168-171`; `Form_FormEquipoMedida.form.txt:433-435` |
| `lblTitulo1`, `lblTitulo2` | etiquetas | — | Título «ALTA DE EQUIPO DE MEDIDA» o «EDICIÓN DE EQUIPO DE MEDIDA» | `:20-23, 34-37` |
| `Etiqueta196`, `197`, `5`, `8`, `11`, `14`, `16`, `18` | 8 etiquetas | — | Rótulos Nombre, DESCR., N/S, P/N, fechas, Marca, Modelo | `Form_FormEquipoMedida.form.txt:503-1000` |
| `Etiqueta2`, `6`, `9`, `12` | 4 etiquetas «*» | — | Marca de obligatorio (coincide con los 4 campos exigidos al guardar) | `Form_FormEquipoMedida.form.txt:597-598, 675-676, 753-754, 832-833` |
| `CuadroSuperior`, `Cuadro2` | 2 rectángulos | — | Decorativos | `Form_FormEquipoMedida.form.txt:348, 372` |

## Flujos

### Apertura (`Form_Load`, `:172-198`)

1. `AjustarTamaño Me` y título (`:180-182`); `EstablecerDatos` (`:11-57`).
2. Sin equipo activo, solo rótulos de alta (`:19-25`). Con equipo activo, **relee** el registro con `Constructor.getEquipoMedida` (`Constructor.bas:4878-4896`) en `m_ObjEquipoAlInicio`; si no existe, «No se encuentra el equipo de medida registrado» (`:26-33`).
3. Recorre los controles con `Tag` «DATOS» y copia al control el valor del objeto si no está vacío (`:38-48`).
4. Cualquier error: `MsgBox` crítico y cierre (`:191-197`).

### Guardado (`ComandoRegistrar_Click`, `:120-166`)

1. `HaHabidoCambios`: en alta siempre cierto; en edición compara cada control con el valor inicial (`:95-111`). Sin cambios: «No ha habido cambios» (`:129-135`).
2. `RellenarObjeto` crea `m_ObjEquipoMedidaActivo` si no existía y le copia los controles con `Nz` (`:65-76`).
3. `ParaAlta = No` si hubo registro inicial, `Sí` si no (`:140-144`); `Registrar` (`:145`).
4. Lanza `EquipoAlta` o `EquipoModificado` (`:149-153`) y cierra (`:157`).
5. Error: `MsgBox` crítico; el formulario queda abierto (`:160-165`).

## Reglas, validaciones y cálculos

- **Obligatorios:** nombre, NS, PN y fecha de inicio de servicio; mensajes «Falta el nombre», «Falta el NS», etc. (`classes/EquipoMedida.cls:323-338`).
- **Unicidad del NS:** solo en edición y si el NS cambia; «El Número de Serie ya está siendo usado por otro equipo de medida» (`classes/EquipoMedida.cls:339-349, 503-534`).
- **Coherencia de fechas:** fin de servicio anterior al inicio ⇒ «La fecha fin de servicio no puede ser anterior a la de comienzo» (`classes/EquipoMedida.cls:352-357`).
- **Identificador:** en alta `IDEquipoMedida = DameID("TbEquiposMedida", "IDEquipoMedida")`, máximo + 1 (`classes/EquipoMedida.cls:188-196, 358-367`; `FG.bas:371-420`, con `FG.bas` = `modules/Funciones Generales.bas`).
- **Edición:** un campo vacío se guarda como `Null` (`classes/EquipoMedida.cls:405-446`); en alta solo se escriben los no vacíos (`classes/EquipoMedida.cls:376-401`).
- **Caché:** si cambia `FechaFinServicio` en edición, recarga `m_ObjEntorno.ColEquiposMedidaActivos` (`classes/EquipoMedida.cls:452-458`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lectura por ID | `TbEquiposMedida` | SQL dinámico concatenado | — | `Constructor.bas:4892-4896` |
| Siguiente ID | `TbEquiposMedida` | recorrido del campo ID en VBA (`DameID`) | no | `FG.bas:409-411` |
| Comprobar NS repetido | `TbEquiposMedida` | SQL dinámico con el NS entre comillas | no | `classes/EquipoMedida.cls:515-519` |
| Localizar registro | `TbEquiposMedida` | SQL dinámico concatenado | no | `classes/EquipoMedida.cls:370-373` |
| Alta | `TbEquiposMedida` | DAO `AddNew`/`Update` | no | `classes/EquipoMedida.cls:376-403` |
| Edición | `TbEquiposMedida` | DAO `Edit`/`Update` | no | `classes/EquipoMedida.cls:405-447` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: identificador `Max + 1` fuera de transacción (`FG.bas:371-420`); fechas manejadas como texto y convertidas con `CDate` (`classes/EquipoMedida.cls:353`); `Null` explícito en edición; SQL concatenado con el NS entre comillas simples; uso de `Format = "Short Date"` dependiente de la configuración regional.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- El control de NS repetido solo se ejecuta cuando `ParaAlta = No` (`classes/EquipoMedida.cls:339-349`): un alta con un NS ya existente no se rechazaría.
- `Max + 1` sin transacción: dos altas simultáneas podrían obtener el mismo ID (`FG.bas:371-420`; `classes/EquipoMedida.cls:358-373`).
- La consulta de NS repetido concatena el NS entre comillas (`classes/EquipoMedida.cls:517`): un NS con apóstrofo rompería la consulta.
- `RellenarObjeto` modifica el objeto global antes de guardar (`:65-76`); en edición es la misma instancia seleccionada en la gestión (`Form_FormEquiposMedidaGestion.cls:153`), de modo que un guardado fallido dejaría valores sin guardar en la lista.
- `HaHabidoCambios` compara texto del objeto con `Nz(ctl.Value, "")` (`:105-106`); en las fechas, el formato del texto cargado y el del valor del control podrían diferir y dar un cambio falso.
- Una fecha no válida en `CDate` (`classes/EquipoMedida.cls:353`) produciría el error genérico de `Registrar`, no un mensaje específico.
- `Numero_Actividad` mantiene un `ControlSource` «IDActividad» en un formulario sin origen de registros (`Form_FormEquipoMedida.form.txt:317-318`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: restricción única sobre el NS en alta y edición; identificador generado por la base de datos; validación de fechas en el campo.

## Criterios de aceptación

Propuestos, no aprobados:

1. Guardar sin nombre, NS, PN o fecha de inicio muestra el mensaje «Falta el …» correspondiente y no escribe (`Nombre`, `NS`, `PN`, `FechaInicioServicio`).
2. Una fecha de fin anterior a la de inicio se rechaza con su mensaje (`FechaFinServicio`).
3. En edición, cambiar el NS a uno ya usado por otro equipo se rechaza.
4. En edición sin cambios aparece «No ha habido cambios» y no se escribe.
5. Vaciar un campo opcional en edición lo deja vacío en el registro.
6. Tras guardar, la lista de equipos de medida se refresca (`ComandoRegistrar`, `EquipoAlta`, `EquipoModificado`).
7. `cmdSalir` cierra sin escribir.

## Pendientes

- Esquema de `TbEquiposMedida`: tipos, longitudes e índice único sobre `NS` (ERD de origen).
- Valor devuelto por `getdb()` (`Variables Globales.bas:548`): origen de datos configurado, sin resolver aquí.
- Si `Numero_Actividad` muestra `#Nombre?` al ser un control con `ControlSource` en un formulario sin origen (`Form_FormEquipoMedida.form.txt:318`).
