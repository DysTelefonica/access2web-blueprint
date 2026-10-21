# Form_Form0BDOpcionesParteProyectos

> **Commit fijado:** `1d7a2f826a1e3c582874e4e29886a6183efb9067` · **Fuente:** `src/forms/Form_Form0BDOpcionesParteProyectos.cls` (369 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Submenú «PARTE DE PROYECTOS»: da acceso al alta y la búsqueda de NC de proyecto, al seguimiento, a la gestión de tipologías, al informe de indicadores y al vídeo de ayuda, y muestra el indicador «Seguimiento usuario / total» de NC de proyecto. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); los datos del indicador llegan por `PintarIndicadores`. Las citas son `Form_Form0BDOpcionesParteProyectos.cls:línea` salvo indicación; `FG.bas` abrevia `modules/Funciones Generales.bas` y `VG.bas`, `modules/Variables Globales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| [Form_Form0BDOpciones](Form_Form0BDOpciones.md) | Botón o etiqueta «NC de proyecto» | `Form_Form0BDOpciones.cls:16, 199-204` |
| [Form_Form0BDOpcionesAuditorias](Form_Form0BDOpcionesAuditorias.md) | Botón «Proyectos»; el de auditorías se cierra | `Form_Form0BDOpcionesAuditorias.cls:197-198` |

Abre `FormNCProyecto` (`:45`), `FormNCProyectoGestion` (`:130`), `FormNCProyectoSeguimiento` (`:226`), `FormNCProyectoTipologiaGestion` (`:102`), `FormIndicadores` (`:170`) y `Form0BDOpcionesAuditorias` cerrándose a sí mismo (`:198-199`). «Salir» solo cierra el formulario (`:342`).

## Controles y eventos

`.form.txt` declara 31 bloques `Begin`: el formulario (`.form.txt:4`), 8 de valores por defecto (`.form.txt:410-465`), 3 secciones (`EncabezadoDelFormulario` `.form.txt:480`, `Detalle` `.form.txt:832`, `PieDelFormulario` `.form.txt:3441`) y los 19 controles de la tabla.

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `ComandoAltaNCProyecto` | botón | `OnClick` | Rechaza técnicos; vacía `m_ObjNCProyectoActiva` y reabre `FormNCProyecto` en alta | `:35-56`, `.form.txt:2772-2774` |
| `ComandoGestionNCProyecto` | botón | `OnClick` | Reabre `FormNCProyectoGestion` con filtro inicial (ver Flujos) | `:119-152`, `.form.txt:849-851` |
| `lblBuscarNC` («Buscar no conformidad») | etiqueta con hipervínculo `#` | `OnClick` | Si el botón está habilitado, ejecuta su handler | `:354-359`, `.form.txt:2053-2057` |
| `ComandoConfiguracion` | botón | `OnClick` | Rechaza técnicos; reabre `FormNCProyectoTipologiaGestion` | `:87-117`, `.form.txt:1390-1392` |
| `lblTipologiaGestion` («Tipos») | etiqueta con hipervínculo `#` | `OnClick` | Si el botón está habilitado, ejecuta su handler | `:364-369`, `.form.txt:2087-2091` |
| `ComandoSeguimiento` | botón | `OnClick` | Reabre `FormNCProyectoSeguimiento` | `:216-241`, `.form.txt:2310-2312` |
| `ComandoIndicadores` | botón | `OnClick` | Abre `FormIndicadores` sin control de rol | `:156-189`, `.form.txt:3232-3234` |
| `ComandoActualizar` («Contador») | botón | `OnClick` | Recalcula el indicador vaciando la caché | `:7-33`, `.form.txt:3157-3159` |
| `ComandoAyuda` | botón (texto «Salir» en el `.form.txt`) | `OnClick` | Abre el vídeo de ayuda de NC de proyecto | `:60-85`, `.form.txt:497-499` |
| `ComandoMenuAuditorias` («Auditorías») | botón | `OnClick` | Abre el submenú de auditorías y se cierra | `:190-214`, `.form.txt:3587-3589` |
| `cmdSalir` | botón | `OnClick` | `DoCmd.Close` del formulario | `:340-351`, `.form.txt:3457-3459` |
| `lblNuevaNoConformidad`, `lblInformeIndicador`, `lblSeguimientos` | etiquetas con hipervínculo `#` | — | Sin `OnClick`; `lblSeguimientos` recibe «Seguimiento N / M» | `.form.txt:3126-3129, 3416-3419, 2740-2743`; `FG.bas:317-319` |
| `lblTitulo1`, `lblVersion` | etiquetas | — | Título según modo pruebas y versión | `:290-305`, `.form.txt:590, 2121` |
| `lblEstado` | etiqueta | — | Destino de `Avance`/`AvanceCerrar` | `:288, 307`, `FG.bas:1426, 1456`, `.form.txt:3563` |
| `Imagen376`, `Imagen70` | imágenes | — | Decorativas | `.form.txt:612, 2145` |
| Sección `Detalle` | sección | `OnDblClick` | Recarga el entorno con `EVE` (gesto oculto) | `:243-272`, `.form.txt:832-833` |

## Flujos

### Apertura (`Form_Open`, `:274-338`)

1. Bloquea la edición, `Ajustar Me` y deshabilita `ComandoConfiguracion` mientras carga (`:281-287`).
2. Título «PARTE DE PROYECTOS» (con «VERSIÓN DE PRUEBAS» y cinta solo para administradores si `EnPruebas = "Sí"`), versión y usuario conectado (`:290-306`).
3. `PintarIndicadores` sin reiniciar (`:307-311`); rehabilita `ComandoConfiguracion` y la edición (`:319, 325`).

### Indicador de seguimiento (`PintarIndicadores`, `FG.bas:75-345`)

- **Total de proyecto** = tareas pendientes de replanificar + NC registradas + NC con acciones sin tareas + NC pendientes de control de eficacia + control de eficacia caducado + control de eficacia no conforme (`FG.bas:148-177`).
- **Del usuario**: las mismas colecciones filtradas por el responsable de Calidad (`RespCalidad` en tareas, `NombreCalidad` en NC) igual al nombre del usuario conectado (`FG.bas:216-262, 306-308`; filtro en `FG.bas:398-409`).
- Escribe «Seguimiento usuario / total» en `lblSeguimientos` (`FG.bas:317-319`) y refresca `FormNCProyectoSeguimiento` si está abierto (`FG.bas:325-330`). «Contador» pone «Seguimiento #/#» y reinicia la caché (`:14-16`; `FG.bas:139-140, 353-363`).

### Búsqueda de NC de proyecto (`:119-152`)

Reabre `FormNCProyectoGestion`, fija `EstadoNC = "Abiertas"`; si el usuario **es** técnico fija `ResponsableTelefonica` a su nombre; después invoca `Form_FormNCProyectoGestion.EstadoNC_AfterUpdate` (`:131-136`; `Form_FormNCProyectoGestion.cls:755`).

### Ayuda (`:60-85`)

Si el vídeo de la ruta de red configurada (`Entorno.URLVideoAyudaNCProyecto`, `Entorno.cls:651-673`) no existe, avisa «El video de ayuda no se encuentra disponible en este momento»; si existe, lo abre con `ShellExecute` (`:68-72`).

### Errores

Error 1000 ⇒ aviso con `m_Error` («No tiene autorización para esa acción», `:38, 96`); otro ⇒ `CorreoAlAdministrador` y `MsgBox` crítico. El cierre de Access en `Form_Open` está comentado (`:337`): el formulario queda abierto tras el error.

## Reglas, validaciones y cálculos

- **Técnicos:** no pueden dar de alta NC de proyecto ni gestionar tipologías (`:37-40, 95-98`); sí buscar, ver seguimiento e indicadores.
- **Pendiente de replanificar:** tarea iniciada, `FechaFinPrevista <= Date()` y sin `FechaFinReal` (`constructor.bas:3410-3413`).
- **NC registrada:** sin acción correctiva y `Borrado = False` (`constructor.bas:3756`).
- **Control de eficacia caducado:** acción con `FechaFinReal` y `FechaPrevistaControlEficacia <= Date()` (`constructor.bas:4022-4024`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Colecciones del indicador (una por categoría) | `TbNoConformidades`, `TbExpedientes`, `TbNCAccionCorrectivas`, `TbNCAccionesRealizadas`, `TbUsuariosAplicaciones`, `TbTiposNCProyectos` | SQL dinámico concatenado de constantes, sin datos del usuario | `constructor.bas:3393-3414, 3731-3758, 3861-3890, 3929-3957, 3995-4026, 4065-4096`; base en `VG.bas:445-457` |
| Caché de colecciones | propiedades `ColSegs*` de `Entorno` (carga perezosa) | memoria | `Entorno.cls:1729-1909` |
| Registro de errores no controlados | `TbCorreosEnviados` | DAO `AddNew/Update` | `FG.bas:1291-1305` |

Sin transacciones.

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: `Date()` de Jet (`constructor.bas:3412, 4024`); booleano `Borrado=False` (`constructor.bas:3756`); concatenación de texto con `&` en el `SELECT` base (`VG.bas:448`); identificador de correo `Max + 1` (`constructor.bas:1165`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `lblNuevaNoConformidad` e `lblInformeIndicador` parecen enlaces pero no tienen `OnClick` (`.form.txt:3126-3129, 3416-3419`).
- El filtro por responsable se aplica al técnico (`:133-135`), al revés que en el submenú de auditorías (`Form_Form0BDOpcionesAuditorias.cls:131-133`).
- `m_ColSegsTareasProyectoIrregulares` nunca se asigna y su resultado se descarta (`FG.bas:223-226`).
- Si una propiedad `ColSegs*` falla, devuelve `Nothing` y el indicador cuenta 0 sin avisar (`Entorno.cls:1729-1748`; `FG.bas:148-151`).
- Si `Form_Open` falla tras `:287`, `ComandoConfiguracion` queda deshabilitado y el formulario abierto (`:327-337`).
- El mensaje de error de «Contador» cita `ComandoConfiguracion_Click` (`:27`); `ComandoAltaNCProyecto` no muestra reloj de arena (`:35-46`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: indicador calculado en servidor por consulta agregada; menú único de proyectos y auditorías con pestañas; criterio de filtro por responsable común a ambos submenús.

## Criterios de aceptación

Propuestos, no aprobados:

1. Un técnico que pulsa «Nueva no conformidad» o «Tipos» recibe «No tiene autorización para esa acción».
2. «Contador» recalcula y muestra «Seguimiento usuario / total» con las seis categorías.
3. «Buscar no conformidad» (botón y etiqueta) abre la búsqueda de NC abiertas; para un técnico, filtrada por él como responsable.
4. «Informe indicador» abre el informe de indicadores.
5. «Auditorías» cambia al submenú de auditorías y cierra este.

## Pendientes

- Contenido y filtros de `FormIndicadores` (`:170`): fuera de esta ficha.
- Efecto en Access de una etiqueta con hipervínculo `#` sin `OnClick` (`.form.txt:3129`).
