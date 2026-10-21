# Form_Form0BDOpcionesAuditorias

> **Commit fijado:** `1d7a2f826a1e3c582874e4e29886a6183efb9067` · **Fuente:** `src/forms/Form_Form0BDOpcionesAuditorias.cls` (290 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Submenú «PARTE DE AUDITORÍAS»: da acceso al alta y la búsqueda de auditorías, al alta y la búsqueda de NC de auditoría, al seguimiento y al vídeo de ayuda, y muestra el indicador «Seguimiento usuario / total» de NC de auditoría. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); los datos del indicador llegan por `PintarIndicadores`. Las citas son `Form_Form0BDOpcionesAuditorias.cls:línea` salvo indicación; `FG.bas` abrevia `modules/Funciones Generales.bas` y `VG.bas`, `modules/Variables Globales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| [Form_Form0BDOpciones](Form_Form0BDOpciones.md) | Botón «NC de auditoría» | `Form_Form0BDOpciones.cls:72` |
| [Form_Form0BDOpcionesParteProyectos](Form_Form0BDOpcionesParteProyectos.md) | Botón «Auditorías»; el de proyectos se cierra | `Form_Form0BDOpcionesParteProyectos.cls:198-199` |

Abre `FormAuditoria` (`:49`), `FormAuditoriasGestion` (`:103`), `FormAuditoriaSeleccion` (`:167`), `FormNCAuditoriaGestion` (`:128`), `FormNCAuditoriaSeguimiento` (`:226`) y `Form0BDOpcionesParteProyectos` cerrándose a sí mismo (`:197-198`). «Salir» solo cierra el formulario (`:186`).

## Controles y eventos

`.form.txt` declara 31 bloques `Begin`: el formulario (`.form.txt:4`), 8 de valores por defecto (`.form.txt:136-191`), 3 secciones (`EncabezadoDelFormulario` `.form.txt:206`, `Detalle` `.form.txt:558`, `PieDelFormulario` `.form.txt:11291`) y los 19 controles de la tabla.

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `ComandoAltaAuditoría` | botón | `OnClick` | Rechaza técnicos; cierra `FormAuditoria` si está abierto, vacía `m_ObjAuditoriaActiva` y lo abre en alta | `:34-63`, `.form.txt:8386-8388` |
| `ComandoGestionAuditorias` | botón | `OnClick` | Abre `FormAuditoriasGestion` | `:93-118`, `.form.txt:574-576` |
| `ComandoAltaNCAuditoría` | botón | `OnClick` | Rechaza técnicos; reabre `FormAuditoriaSeleccion` | `:152-183`, `.form.txt:6784-6786` |
| `ComandoGestionNCAuditorias` | botón | `OnClick` | Abre `FormNCAuditoriaGestion` con filtro inicial (ver Flujos) | `:120-150`, `.form.txt:8869-8871` |
| `ComandoSeguimiento` | botón | `OnClick` | Reabre `FormNCAuditoriaSeguimiento` | `:215-241`, `.form.txt:10762-10764` |
| `ComandoActualizar` («Contador») | botón | `OnClick` | Recalcula el indicador vaciando la caché | `:6-32`, `.form.txt:11223-11225` |
| `ComandoAyuda` | botón (texto «Salir» en el `.form.txt`) | `OnClick` | Abre el vídeo de ayuda de NC de auditoría | `:65-90`, `.form.txt:223-225` |
| `ComandoMenuParteProyecto` («Proyectos») | botón | `OnClick` | Abre el submenú de proyectos y se cierra | `:189-213`, `.form.txt:11437-11439` |
| `cmdSalir` | botón | `OnClick` | `DoCmd.Close` del formulario | `:185-187`, `.form.txt:11307-11309` |
| `lblNuevaAuditoría`, `lblBuscarAuditoria`, `lblNuevaNCAuditoria`, `lblBuscarNCAuditorias` | etiquetas con hipervínculo `#` | `OnClick` declarado | **Sin handler** en el `.cls` | `.form.txt:8836-8840, 10471-10475, 10505-10509, 10539-10543` |
| `lblSeguimientos` | etiqueta con hipervínculo `#` | — | Texto «Seguimiento N / M» escrito por `PintarIndicadores` | `:13`, `FG.bas:321-323`, `.form.txt:11192` |
| `lblTitulo1`, `lblVersion` | etiquetas | — | Título según modo pruebas y versión | `:253-268`, `.form.txt:316, 10573` |
| `lblEstado` | etiqueta | — | Destino de `Avance`/`AvanceCerrar` | `FG.bas:1426, 1456`, `.form.txt:11413` |
| `Imagen376`, `Imagen70` | imágenes | — | Decorativas | `.form.txt:338, 10597` |

## Flujos

### Carga (`Form_Load`, `:244-290`)

1. `Ajustar Me`; título «PARTE DE AUDITORÍAS» (con «VERSIÓN DE PRUEBAS» y cinta solo para administradores si `EnPruebas = "Sí"`) (`:250-267`).
2. Versión y usuario conectado en `lblVersion` y la barra de título (`:268-269`).
3. `PintarIndicadores` sin reiniciar: usa las colecciones ya cacheadas en `m_ObjEntorno` (`:270`; `FG.bas:139-144`).

### Indicador de seguimiento (`PintarIndicadores`, `FG.bas:75-345`)

- **Total de auditoría** = tareas pendientes de replanificar + NC registradas + NC con acciones sin tareas + NC pendientes de control de eficacia + control de eficacia caducado + control de eficacia no conforme (`FG.bas:184-213`).
- **Del usuario**: las mismas colecciones filtradas por `Responsable = Nombre` del usuario conectado (`FG.bas:263-305, 311-313`; filtro en `FG.bas:410-421`).
- Escribe «Seguimiento usuario / total» en `lblSeguimientos` si este formulario está abierto (`FG.bas:321-323`) y refresca `FormNCAuditoriaSeguimiento` si lo está (`FG.bas:332-337`).
- «Contador» primero pone «Seguimiento #/#» y llama con `p_Reiniciando = Sí`, que vacía la caché con `ResetearColTareas` (`:13-15`; `FG.bas:139-140, 346-372`).

### Búsqueda de NC de auditoría (`:120-150`)

Abre `FormNCAuditoriaGestion`, fija `Estado = "Abiertas"`; si el usuario **no** es técnico fija `RESPONSABLEIMPLANTACION` a su nombre; después invoca `Form_FormNCAuditoriaGestion.ESTADO_AfterUpdate` (`:129-134`; `Form_FormNCAuditoriaGestion.cls:666`).

### Ayuda (`:65-90`)

Si el vídeo de la ruta de red configurada (`Entorno.URLVideoAyudaNCAuditoria`, `Entorno.cls:714-736`) no existe, avisa «El video de ayuda no se encuentra disponible en este momento»; si existe, lo abre con `ShellExecute` (`:73-77`; `VG.bas:17`).

### Errores

Error 1000 ⇒ aviso con `m_Error` («No tiene autorización para esa acción» en las altas, `:42, 160`); otro ⇒ `CorreoAlAdministrador` y `MsgBox` crítico. `Form_Load` no cierra Access ante error (`:281-289`).

## Reglas, validaciones y cálculos

- **Técnicos:** no pueden dar de alta auditorías ni NC de auditoría (`:41-44, 159-162`); sí pueden buscar y ver el seguimiento.
- **Pendiente de replanificar:** tarea iniciada, con `FechaFinPrevista <= Date()` y sin `FechaFinReal` (`constructor.bas:3595-3598`).
- **Control de eficacia caducado:** `RequiereControlEficacia = 'Sí'` y `FechaPrevistaControlEficacia <= Date()` (`constructor.bas:4424-4427`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Colecciones del indicador (una por categoría) | `TbAuditorias`, `TbNoConformidadesAuditoria`, `TbNCAuditoriaAccionCorrectivas`, `TbNCAuditoriaAccionesRealizadas`, `TbUsuariosAplicaciones` | SQL dinámico concatenado de constantes, sin datos del usuario | `constructor.bas:3577-3599, 4134-4161, 4263-4292, 4330-4360, 4398-4429, 4467-4498`; base en `VG.bas:459` |
| Caché de colecciones | propiedades `ColSegs*` de `Entorno` (carga perezosa) | memoria | `Entorno.cls:1966-1990, 2018-2169` |
| Registro de errores no controlados | `TbCorreosEnviados` | DAO `AddNew/Update` | `FG.bas:1291-1305` |

Sin transacciones.

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: `Date()` de Jet en los filtros (`constructor.bas:3597, 4427`); comparaciones con el literal `'Sí'` en campos de texto (`constructor.bas:4289, 4356`); identificador de correo `Max + 1` (`constructor.bas:1165`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Las cuatro etiquetas-enlace declaran `OnClick` sin handler: pulsar el texto no hace nada (`.form.txt:8839, 10474, 10508, 10542`).
- `ResetearColTareas` no vacía `ColSegsNCAuditoriaCECaducada` ni `ColSegsNCAuditoriaCENoConforme` (`FG.bas:353-372`): «Contador» mantendría esas dos cifras sin refrescar.
- Si una propiedad `ColSegs*` falla, deja el error en `Entorno.Error` y devuelve `Nothing` (`Entorno.cls:2018-2038`); `PintarIndicadores` no lo consulta y cuenta 0 (`FG.bas:189-192`).
- El filtro por responsable se aplica a quien **no** es técnico (`:131-133`), al revés que en el submenú de proyectos (`Form_Form0BDOpcionesParteProyectos.cls:133-135`).
- El mensaje de error de «Contador» cita `ComandoConfiguracion_Click` (`:26`).
- En `ComandoAyuda` el error 1000 se muestra como crítico (`:89`) y el texto del botón en el `.form.txt` es «Salir» (`.form.txt:224`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: indicador calculado en servidor por consulta agregada en vez de colecciones en memoria; menú único de proyectos y auditorías con pestañas; ayuda como enlace a un recurso publicado.

## Criterios de aceptación

Propuestos, no aprobados:

1. Un técnico que pulsa «Nueva auditoría» o «Nueva NC de auditoría» recibe «No tiene autorización para esa acción».
2. «Contador» recalcula y muestra «Seguimiento usuario / total» con las seis categorías.
3. «Buscar NC auditorías» abre la búsqueda filtrada por NC abiertas.
4. Sin vídeo disponible, «Ayuda» avisa de que no está disponible.
5. «Proyectos» cambia al submenú de proyectos y cierra este.

## Pendientes

- Efecto en Access de `OnClick = [Event Procedure]` sin procedimiento (`.form.txt:8839`).
- Si el texto «Salir» de `ComandoAyuda` se ve o queda tapado por una imagen (`.form.txt:224`).
