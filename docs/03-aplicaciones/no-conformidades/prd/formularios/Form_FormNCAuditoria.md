# Form_FormNCAuditoria

> **Commit fijado:** `1d7a2f826a1e3c582874e4e29886a6183efb9067` · **Fuente:** `src/forms/Form_FormNCAuditoria.cls` (219 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Contenedor de la ficha de una NC (o observación) de auditoría: muestra un título de alta o edición, una barra de pestañas (`NavigationControl`) y, debajo, el subformulario `FrmDetalle`, que en la carga apunta a `FormNCAuditoriaGeneral`. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); no lee ni escribe tablas por sí mismo y trabaja sobre los objetos globales `m_ObjNCAuditoriaActiva` y `m_ObjAuditoriaActiva` (`VG.bas:126-127`). Las citas son `Form_FormNCAuditoria.cls:línea` salvo indicación; `FG.bas` abrevia `modules/Funciones Generales.bas` y `VG.bas`, `modules/Variables Globales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormNCAuditoriaGestion` | `DoCmd.OpenForm "FormNCAuditoria"` | `Form_FormNCAuditoriaGestion.cls:254, 341, 587` |
| `Form_FormAuditoriasGestion`, `Form_FormAuditoriaSeleccion` | Ídem | `Form_FormAuditoriasGestion.cls:86`, `Form_FormAuditoriaSeleccion.cls:60` |
| `Form_FormNCAuditoriaSeguimientoNC`, `Form_FormNCAuditoriaSeguimientoTareas` | Ídem | `Form_FormNCAuditoriaSeguimientoNC.cls:86`, `Form_FormNCAuditoriaSeguimientoTareas.cls:56` |

El formulario no recibe `OpenArgs`: el contexto viaja en las globales (`:139-149, 162, 179`). Las pestañas navegan a `FormNCAuditoriaAcciones`, `FormNCAuditoriaControlEficacia`, `FormNCAuditoriaNota`, `FormNCAuditoriaDocumentos`, `FormNCAuditoriaReplanificaciones` y `FormNCProyectoAuditoriaEliminado` (`.form.txt:1033, 1089, 1143, 1197, 1251, 1306`). Se cierra con `ComandoSalir` (`:44-46`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| Formulario | form | `OnLoad`, `OnClose` | Prepara título, permisos y pestañas; al cerrar puede enviar el correo de alta | `.form.txt:59, 394`, `:86-117, 49-84` |
| `lblTitulo` | etiqueta | — | «ALTA DE REGISTRO DE AUDITORÍA <nombre>» o «EDICIÓN DE <partícula> DE AUDITORÍA <nº 2 dígitos>» | `:151, 168, 191`, `.form.txt:894` |
| `ControlDeNavegación0` | control de navegación | — | Contenedor de las pestañas | `.form.txt:935` |
| `NavGeneral` («General») | pestaña | — | Destino fijado en código a `FormNCAuditoriaGeneral` | `:207`, `.form.txt:963-964` |
| `NavAcciones`, `NavControlEficacia`, `NavNotas`, `NavDocumentos`, `NavReplanificaciones`, `NavMotivoBorrado` | pestañas | — | Se deshabilitan en el alta; el resto de reglas en «Reglas» | `:154-160, 171-177`, `.form.txt:1016-1306` |
| `BotónDeNavegación1` («[Agregar nuevo]») | pestaña | — | Pestaña de alta de la plantilla de Access; sin destino ni código | `.form.txt:1343-1344` |
| `FrmDetalle` | subformulario | — | `SourceObject = "FormNCAuditoriaGeneral"` al cargar | `:208`, `.form.txt:1394` |
| `ComandoInforme` («Informe») | botón | `OnClick` | Genera y abre el informe Word de la NC; oculto en el alta | `:16-42, 163, 166`, `.form.txt:1490-1501` |
| `ComandoSalir` (rótulo «Comando47») | botón | `OnClick` | Cierra el formulario sin guardar | `:44-46`, `.form.txt:1424-1432` |
| `EncabezadoDelFormulario`, `Detalle`, `PieDelFormulario` | secciones | — | Cabecera, detalle y pie | `.form.txt:876, 918, 1414` |

El `.form.txt` contiene además bloques `Begin` de valores por defecto (imagen, botón, subformulario, celda vacía, control de navegación, botón de navegación) sin control real asociado (`.form.txt:758-860`).

## Flujos

### Apertura (`Form_Load`, `:86-117`; `EstablecerDatos`, `:118-219`)

1. `Ajustar Me`, `blnRequiereCorreoAlta = False` y `EstablecerDatos` (`:97-99`). Si devuelve error, lo eleva como aviso (`:100-102`).
2. Si `EsTecnico = Sí`, el formulario queda en solo lectura (`AllowEdits = False`); si no, editable (`:130-137`).
3. Si no hay auditoría activa, la toma de la NC activa; sin ninguna de las dos avisa «No se conoce la auditoría de la que proviente» (`:139-149`).
4. **Alta** (NC activa vacía o sin `id`): título de alta, foco en `NavGeneral`, deshabilita seis pestañas, crea un `NCAuditoria` nuevo y activa `blnRequiereCorreoAlta` (`:150-164, 167-180`).
5. **Edición**: recarga la NC con `constructor.getNCAuditoria` (`:182`); si no existe, avisa «Parece que no se ha podido encontrar…» (`:186-189`); si `Borrado`, bloquea la edición (`:192-195`); llama a `EstablecerColorPestañasNCAuditoria` (`:197`).
6. Fija `NavGeneral` y `FrmDetalle` en `FormNCAuditoriaGeneral` (`:207-208`).

### Cierre (`Form_Close`, `:49-84`)

Si la NC ya tiene `id` y era un alta, comprueba: sin acciones y con acción correctiva «Sí», llama a `PintarIndicadores` con reinicio y avisa «No se va a mandar el correo a Calidad al no haber acciones registradas» (`:59-65`); en otro caso, `EnviarCorreoAltaNCAuditoria` (`:66`).

### Informe (`ComandoInforme_Click`, `:16-42`)

`GenerarWordNoConformidades` con `p_EsDeProyecto = No` (`:25`); genera el informe y lo abre (`FG.bas:3978-4009`).

### Errores

Error interno 1000 ⇒ `MsgBox` de advertencia; otros ⇒ correo al administrador y `MsgBox` crítico (`:33-41, 75-83, 108-116`). El cursor de espera se restablece en todos los caminos.

## Reglas, validaciones y cálculos

- **Pestañas habilitadas** (`FG.bas:2393-2432`): sin NC o sin `id` ⇒ quedan deshabilitadas todas salvo `NavGeneral`; con `id`, `NavMotivoBorrado` solo si `Borrado`; `NavControlEficacia` se deshabilita si `RequiereControlEficacia = "No"`; `NavAcciones` y `NavReplanificaciones`, si `RequiereAccionCorrectiva = "No"`.
- **Color de pestañas** (`FG.bas:2434-2461`): salvo NC cerrada o borrada, `NavGeneral`, `NavAcciones` (solo si requiere acción correctiva) y `NavControlEficacia` toman el color «OK sí/no» según `DatosGeneralesOK`, `AccionesOK` y `EficaciaOK`.
- **Correo de alta** (`FG.bas:3392-3435`): para Calidad (o los correos de Calidad en pruebas) con copia al usuario conectado; en pruebas con el usuario `adm` se vuelca a texto en lugar de enviar (`FG.bas:3406-3412`).
- **Solo lectura:** técnico o NC borrada (`:130-133, 192-195`).

## Lecturas y escrituras

| Operación | Objeto | Tipo | Evidencia |
|---|---|---|---|
| Recarga de la NC | `constructor.getNCAuditoria` (objeto `NCAuditoria`) | Pendiente: SQL no revisado en esta ficha | `:182` |
| Correo de alta | `Correo.Registrar` | Registro del correo; destino no revisado | `FG.bas:3425` |
| Informe Word | `Informe.GenerarWordNoConformidades` | Genera fichero y lo abre | `FG.bas:3998-4009` |

El formulario no ejecuta SQL propio ni transacciones.

## Coexistencia

Sin SQL propio: la adaptación recae en `constructor.getNCAuditoria`, `Correo.Registrar` y la generación del informe (automatización Word, no dependiente del motor). No hay dialecto Jet/ACE, autonuméricos ni transacciones en este formulario.

## Comportamiento actual · Posibles defectos · Propuestas web

**Comportamiento actual:** el correo de alta se dispara en `Form_Close`, no al guardar (`:49-66`); el estado de las pestañas depende de globales y de una función externa que localiza el formulario por nombre (`FG.bas:2380-2383`).

**Posibles defectos (a verificar):**
- La pestaña `NavMotivoBorrado` apunta a `FormNCProyectoAuditoriaEliminado` (`.form.txt:1306`), objeto que no existe en `src/forms` en el commit fijado.
- El mensaje «proviente» (errata) y el rótulo «Comando47» de `ComandoSalir` (`:141`, `.form.txt:1431`).
- `Form_Close` con error interno 1000 deja el cierre en curso y solo avisa (`:77-83`).

**Propuestas web pendientes de decisión:**
- Enviar el correo de alta al guardar la NC y no al cerrar.
- Sustituir las globales por un identificador de NC en la ruta.

## Criterios de aceptación

1. Con una NC sin `id`, el título empieza por «ALTA DE REGISTRO DE AUDITORÍA», solo `NavGeneral` queda habilitada y `ComandoInforme` está oculto (`:150-164`).
2. Con una NC con `id`, el título es «EDICIÓN DE <partícula> DE AUDITORÍA» con el número en dos dígitos y `ComandoInforme` está visible (`:166, 191`).
3. Un usuario técnico o una NC borrada no pueden editar (`:130-133, 192-195`).
4. `NavControlEficacia` se deshabilita si no requiere control de eficacia; `NavAcciones` y `NavReplanificaciones`, si no requiere acción correctiva (`FG.bas:2424-2430`).
5. Cerrar un alta con acción correctiva «Sí» y sin acciones no envía el correo y avisa (`:59-65`); con acciones, lo envía (`:66`).
6. `ComandoInforme` abre el informe de la NC activa (`:25`).

## Pendientes

- Contenido de `constructor.getNCAuditoria` y de las propiedades `DatosGeneralesOK`, `AccionesOK`, `EficaciaOK` (`:182`, `FG.bas:2441-2459`).
- Objeto `FormNCProyectoAuditoriaEliminado` ausente en main (`.form.txt:1306`).
- Origen del valor `EsTecnico` (`VG.bas:393`) para las reglas de solo lectura.
