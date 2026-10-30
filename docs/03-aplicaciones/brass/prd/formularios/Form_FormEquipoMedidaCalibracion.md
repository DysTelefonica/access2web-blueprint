# Form_FormEquipoMedidaCalibracion

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormEquipoMedidaCalibracion.cls` (202 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Alta y edición de una calibración de un equipo de medida (`TbEquiposMedidaCalibraciones`: fecha de calibración y fecha de fin de vigencia). Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): lee el equipo de la global `m_ObjEquipoMedidaActivo` (`Variables Globales.bas:253`) y la calibración de `m_ObjCalibracionEquipoMedidaActiva` (`Variables Globales.bas:254`), y persiste con `EquipoMedidaCalibracion.Registrar` (`classes/EquipoMedidaCalibracion.cls:78-167`). Las citas son `Form_FormEquipoMedidaCalibracion.cls:línea` salvo indicación. Capacidad de todos los controles: CAP-5 Equipos, equipos de medida y calibraciones.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormEquipoMedidaCalibracionesGestion` (alta) | `m_ObjCalibracionEquipoMedidaActiva = Nothing`, cierra el formulario si estaba abierto y lo abre; guarda la referencia en `m_FormCalibracion` (`WithEvents`) | `Form_FormEquipoMedidaCalibracionesGestion.cls:66-81, 6` |
| `Form_FormEquipoMedidaCalibracionesGestion` (edición) | Toma la calibración seleccionada (o la lee con `Constructor.getEquipoMedidaCalibracion` desde la columna 0 de `ListaFiltrados`), la asigna a la global y abre el formulario | `Form_FormEquipoMedidaCalibracionesGestion.cls:100-120` |

Eventos hacia quien lo abrió: `Alta` (`:7`, lanzado en `:44`) y `Edicion` (`:8`, lanzado en `:46`); la gestión los atiende en `m_FormCalibracion_Alta` y `m_FormCalibracion_Edicion` (`Form_FormEquipoMedidaCalibracionesGestion.cls:283, 308`). Tras guardar se cierra solo (`:51`); `cmdSalir` cierra sin guardar (`:10-12`).

## Controles y eventos

Controles: 22 (33 `Begin` en el `.form.txt`: 1 `Form`, 7 plantillas de estilo por defecto en `Form_FormEquipoMedidaCalibracion.form.txt:140-243`, 1 `FormHeader`, 1 `Section`, 1 `FormFooter`, y 22 controles: 2 en la cabecera, 18 en el detalle y 2 en el pie).

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `FechaCalibracion` | texto, formato «Short Date» | — | Fecha de calibración; obligatoria y válida al guardar | `Form_FormEquipoMedidaCalibracion.form.txt:365-366`; `classes/EquipoMedidaCalibracion.cls:111-114` |
| `FechaFinCalibracion` | texto, formato «Short Date» | — | Fecha de fin de calibración; obligatoria, válida y no anterior a la primera | `Form_FormEquipoMedidaCalibracion.form.txt:703-704`; `classes/EquipoMedidaCalibracion.cls:115-122` |
| `Nombre`, `Modelo`, `NS`, `PN` | 4 textos con `Enabled`/`Locked` distintos del valor por defecto | — | Informativos del equipo; se rellenan en `EstablecerDatos` (`Nombre` = nombre y marca entre paréntesis) | `:107-112`; `Form_FormEquipoMedidaCalibracion.form.txt:541-542, 555, 586, 617, 675` |
| `ComandoRegistrar` | botón «Registrar» (pie) | `OnClick` → `ComandoRegistrar_Click` | Guarda y cierra | `:14-60`; `Form_FormEquipoMedidaCalibracion.form.txt:862-864` |
| `CmdSalir` | botón «Salir» (pie) | `OnClick` → `cmdSalir_Click` | Cierra sin guardar | `:10-12`; `Form_FormEquipoMedidaCalibracion.form.txt:837-839` |
| `lblTitulo1`, `lblTitulo2` | etiquetas (cabecera) | — | «ALTA DE CALIBRACIÓN» o «EDICIÓN DE CALIBRACIÓN» | `:113-123` |
| `Etiqueta538`, `190`, `191`, `192`, `193`, `197`, `243`, `245` | 8 etiquetas | — | Rótulos de fechas, equipo, modelo, NS, PN y sección | `Form_FormEquipoMedidaCalibracion.form.txt:393-760` |
| `Etiqueta99`, `Etiqueta0` | 2 etiquetas «*» | — | Marca de obligatorio (una por fecha) | `Form_FormEquipoMedidaCalibracion.form.txt:782-783, 804-805` |
| `Marco175`, `Marco189` | 2 grupos de opciones | — | Decorativos | `Form_FormEquipoMedidaCalibracion.form.txt:341, 419` |

## Flujos

### Apertura (`Form_Load`, `:62-93`)

1. `AjustarTamaño Me` y título (`:70-72`). Sin equipo de medida activo: «No se ha indicado el equipo de medida» (`:73-76`).
2. `EstablecerDatos` (`:95-149`): rellena los cuatro datos del equipo (`:107-112`); sin calibración activa, rótulos de alta y nada más (`:113-119`).
3. Con calibración activa, la **relee** con `Constructor.getEquipoMedidaCalibracion` en `m_CalibracionAlInicio`; si no existe, «No parece estar registrado en la base de datos» (`:124-131`). Copia las fechas solo si son válidas (`:132-139`).
4. Cualquier error: `MsgBox` crítico y cierre (`:86-92`).

### Guardado (`ComandoRegistrar_Click`, `:14-60`)

1. `HaHabidoCambios`: en alta siempre cierto; en edición compara las dos fechas con el valor inicial (`:177-194`). Sin cambios: «No ha habido cambios» (`:23-29`).
2. `RellenarObjeto` crea la global si no existía y le copia las dos fechas con `Nz` y el `IDEquipoMedida` del equipo activo (`:158-165`).
3. `ParaAlta = No` si hubo registro inicial, `Sí` si no (`:34-38`); `Registrar` (`:39`).
4. Lanza `Alta` o `Edicion` (`:43-47`) y cierra (`:51`).
5. Error: `MsgBox` crítico; el formulario queda abierto (`:54-59`).

## Reglas, validaciones y cálculos

- **Fechas:** ambas deben ser fechas válidas, «La fecha inicial es obligatoria» / «La fecha Final es obligatoria», y la final no puede ser anterior a la inicial (`classes/EquipoMedidaCalibracion.cls:111-122`).
- **Identificador:** en alta `IDCalibracion = DameID("TbEquiposMedidaCalibraciones", "IDCalibracion")`, máximo + 1 (`classes/EquipoMedidaCalibracion.cls:28-46, 104-110`; `FG.bas:371-420`, con `FG.bas` = `modules/Funciones Generales.bas`).
- **Empresa calibradora:** `Registrar` la fija siempre a un valor literal del código, sin control en el formulario (`classes/EquipoMedidaCalibracion.cls:123`).
- **Observaciones:** el formulario no las muestra ni las modifica; en edición `Registrar` las deja `Null` si el objeto las trae vacías (`classes/EquipoMedidaCalibracion.cls:152-156`).
- **Estado del equipo:** `EquipoMedida.EstadoCalibracion` deriva de la última calibración (`classes/EquipoMedida.cls:68-95`); este formulario no recalcula ese estado.

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lectura inicial en edición | `TbEquiposMedidaCalibraciones` | `Constructor.getEquipoMedidaCalibracion` (SQL no leído aquí) | — | `:124` |
| Siguiente ID | `TbEquiposMedidaCalibraciones` | recorrido del campo ID en VBA (`DameID`) | no | `FG.bas:409-411` |
| Localizar registro | `TbEquiposMedidaCalibraciones` | SQL dinámico concatenado | no | `classes/EquipoMedidaCalibracion.cls:124-127` |
| Alta | `TbEquiposMedidaCalibraciones` | DAO `AddNew`/`Update` | no | `classes/EquipoMedidaCalibracion.cls:136-145` |
| Edición | `TbEquiposMedidaCalibraciones` | DAO `Edit`/`Update` | no | `classes/EquipoMedidaCalibracion.cls:147-157` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: identificador `Max + 1` fuera de transacción (`FG.bas:371-420`); fechas manejadas como texto con `IsDate`/`CDate` (`classes/EquipoMedidaCalibracion.cls:111-119`), sensibles a la configuración regional; `Null` explícito en edición; campo `Observaciones` opcional.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `Max + 1` sin transacción: dos altas simultáneas podrían obtener el mismo ID (`classes/EquipoMedidaCalibracion.cls:104-110, 127-136`).
- `RellenarObjeto` modifica la global antes de guardar (`:158-165`); en edición es la misma instancia seleccionada en la gestión (`Form_FormEquipoMedidaCalibracionesGestion.cls:100-114`), así que un guardado fallido dejaría fechas sin guardar en la lista.
- `HaHabidoCambios` compara el texto leído de la base con `Nz(Me.FechaCalibracion, "")` (`:186-190`); si el formato del texto cargado y el del control difieren, podría detectar un cambio inexistente.
- `Registrar` reescribe `EmpresaCalibradora` con el valor literal en cada guardado, también en edición (`classes/EquipoMedidaCalibracion.cls:123, 151`): cualquier valor distinto previo se perdería.
- `cmdSalir_Click` no tiene manejo de errores, a diferencia de los otros handlers (`:10-12`).
- No hay comprobación de solapamiento entre calibraciones del mismo equipo en `Registrar` (`classes/EquipoMedidaCalibracion.cls:78-167`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: identificador generado por la base de datos; empresa calibradora como dato editable; validación de periodos solapados.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir sin equipo de medida activo muestra «No se ha indicado el equipo de medida» y cierra el formulario.
2. Los datos del equipo (`Nombre`, `Modelo`, `NS`, `PN`) se muestran y no son editables.
3. Guardar sin fecha de calibración o sin fecha final muestra el mensaje correspondiente y no escribe.
4. Guardar con fecha final anterior a la inicial muestra «La fecha Final ha de ser posterior a la inicial».
5. En edición sin cambios aparece «No ha habido cambios» y no se escribe.
6. Un alta correcta crea la calibración asociada al equipo activo y la lista de la gestión se refresca (`Alta`, `Edicion`).
7. `CmdSalir` cierra sin escribir.

## Pendientes

- Esquema de `TbEquiposMedidaCalibraciones` (tipos, claves; ERD de origen).
- SQL de `Constructor.getEquipoMedidaCalibracion` (`:124`), no leído.
- Origen de la asignación de `m_ObjEquipoMedidaActivo` antes de abrir el formulario desde la gestión de calibraciones (no verificado).
- Valor devuelto por `getdb()` (`Variables Globales.bas:548`): origen de datos configurado.
