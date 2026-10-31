# Form_FormNCProyectoSeguimiento

> **Commit fijado:** `1d7a2f826a1e3c582874e4e29886a6183efb9067` · **Fuente:** `src/forms/Form_FormNCProyectoSeguimiento.cls` (75 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Contenedor del seguimiento de NC de proyecto: título, botón «Contador» y un control de navegación con dos pestañas («Tareas» y «Rev. Estados») que cargan en el subformulario `FrmDetalle` los formularios `FormNCProyectoSeguimientoTareas` (`Form_FormNCProyectoSeguimientoTareas`) y `FormNCProyectoSeguimientoNC` ([Form_FormNCProyectoSeguimientoNC](Form_FormNCProyectoSeguimientoNC.md)) (`.form.txt:1011, 1063`). No tiene `RecordSource` en el `.form.txt` (no enlazado) y no lee ni escribe datos propios. Las citas son `Form_FormNCProyectoSeguimiento.cls:línea` salvo indicación; `FG.bas` abrevia `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Form_Form0BDOpcionesParteProyectos | Botón de seguimiento: si el formulario ya está abierto lo cierra sin guardar y lo vuelve a abrir | `Form_Form0BDOpcionesParteProyectos.cls:223-226` |
| Form_Form0BDTecnicos | Mismo patrón de cierre y reapertura | `Form_Form0BDTecnicos.cls:67-70` |

Se cierra con «Salir» (`:13-16`, `DoCmd.Close acForm … acSaveNo`) o, si falla la carga, desde el manejador de errores de `Form_Load` (`:71`, inalcanzable, véase Posibles defectos).

## Controles y eventos

`.form.txt` declara 26 bloques `Begin`: el formulario (`.form.txt:4`), un bloque de valores por defecto con 12 tipos (`.form.txt:718-806`), 3 secciones (`EncabezadoDelFormulario` `:840`, `Detalle` `:944`, `PieDelFormulario` `:1169`) y los 10 controles de la tabla.

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `lblTitulo` | etiqueta | — | Texto «SEGUIMIENTO DE TAREAS Y REVISIÓN DE ESTADOS DE NC DE PROYECTO»; `Form_Load` lo copia a `Me.Caption` | `:55`, `.form.txt:859-860` |
| `ComandoActualizar` («Contador») | botón | `OnClick` | Recalcula los indicadores con `PintarIndicadores` reiniciando las colecciones | `:18-43`, `.form.txt:879-881` |
| `Etiqueta237` | etiqueta | — | Decorativa, sin evento | `.form.txt:952-954` |
| `ControlDeNavegación0` | control de navegación | — | Contenedor de las pestañas | `.form.txt:961-968` |
| `NavTareas` («Tareas») | botón de navegación | destino sin handler | Carga `FormNCProyectoSeguimientoTareas` en `FrmDetalle` | `.form.txt:990-1011` |
| `NavNC` («Rev. Estados») | botón de navegación | destino sin handler | Carga `FormNCProyectoSeguimientoNC` en `FrmDetalle` | `.form.txt:1039-1063` |
| `BotónDeNavegación1` («[Agregar nuevo]») | botón de navegación | — | Marcador de posición del asistente de Access, sin destino | `.form.txt:1091-1099` |
| `FrmDetalle` | subformulario | — | Hospeda el formulario elegido; sin `LinkChildFields` ni `LinkMasterFields` | `.form.txt:1141-1149` |
| `cmdSalir` («Salir») | botón | `OnClick` | Cierra el formulario | `:13-16`, `.form.txt:1179-1190` |
| `lblEstado` | etiqueta | — | Rótulo de estado del pie («Obteniendo variables de Entorno…»); el `.cls` no lo usa | `.form.txt:1282-1294` |

## Flujos

### Carga (`Form_Load`, `:45-72`)

1. Cursor de espera, `m_Error = ""`, `Ajustar Me` (`:48-51`; `FG.bas:615`).
2. `Me.Caption = Me.lblTitulo.Caption` y cursor normal (`:55-60`).
3. El código no fija `SourceObject` de `FrmDetalle` (`:45-61`); el destino inicial queda pendiente de verificar en el binario.

### Contador (`ComandoActualizar_Click`, `:18-43`)

Llama a `PintarIndicadores p_Reiniciando:=EnumSino.Sí` (`:26`; `FG.bas:75`), que reinicia las colecciones del entorno, recuenta los indicadores, rotula `lblSeguimientos` en los menús abiertos y refresca el detalle abierto: ejecuta `Filtrar` del formulario de NC o de tareas de proyecto según el `SourceObject` de `FrmDetalle` (`FG.bas:325-331`).

### Errores

Si `PintarIndicadores` devuelve texto en `m_Error` se lanza el error 1000 (`:27-29`). Error 1000 ⇒ `MsgBox` de advertencia con `m_Error`; otro ⇒ mensaje con número y descripción, `CorreoAlAdministrador` y `MsgBox` crítico (`:36-42`).

## Reglas, validaciones y cálculos

- El formulario no valida entradas ni comprueba el rol.
- Los indicadores de proyecto (usuario y total) se calculan en `PintarIndicadores` (`FG.bas:75-139`); el detalle queda en la ficha del menú de proyectos.
- Variable pública `m_ARProyectoSeleccionado` declarada sin uso en el `.cls` (`:7`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Ninguna directa | — | — | `:1-75` |
| Indirecta: recuentos del entorno | Colecciones del objeto de entorno recalculadas por `PintarIndicadores` | DAO/SQL ejecutado en `Entorno`, fuera de esta ficha | `FG.bas:75-139` |

## Coexistencia

Sin SQL propio, no hay dialecto Jet/ACE que adaptar en el formulario. Dependen de PostgreSQL vinculado las consultas que alimentan las colecciones del entorno (fuera de alcance de esta ficha). `Ajustar` y el cursor de espera son de interfaz Access, sin equivalente directo en web.

## Comportamiento actual · Posibles defectos · Propuestas web

**Comportamiento actual**

- Contenedor sin datos propios con dos pestañas y un botón de recuento (`.form.txt:952-1149`).
- «Contador» actualiza también los menús y los formularios de auditoría abiertos, no solo el de proyecto (`FG.bas:318-338`).

**Posibles defectos (a verificar)**

- `Form_Load` tiene la etiqueta `errores:` pero ningún `On Error GoTo errores`; el manejador y el cierre de `:62-71` no se ejecutan (`:45-72`).
- El mensaje de error de `ComandoActualizar_Click` dice «ComandoConfiguracion_Click» (`:37`) y el de `Form_Load` dice «Form_Open» (`:65`).
- `BotónDeNavegación1` es un marcador sin destino (`.form.txt:1091-1099`).

**Propuestas web pendientes de decisión**

- Sustituir el control de navegación por pestañas de la aplicación y cargar el detalle bajo demanda.
- Recalcular el contador al cambiar de pestaña en lugar de por botón.

## Criterios de aceptación

1. Desde el menú de proyectos, «Seguimiento» abre el formulario y, si ya estaba abierto, lo reabre (`Form_Form0BDOpcionesParteProyectos.cls:223-226`).
2. «Tareas» y «Rev. Estados» muestran en `FrmDetalle` el formulario correspondiente (`.form.txt:1011, 1063`).
3. «Contador» recalcula los indicadores y refresca el detalle visible (`:26`; `FG.bas:325-331`).
4. Un error 1000 muestra advertencia; cualquier otro, mensaje crítico y aviso al administrador (`:36-42`).
5. «Salir» cierra sin guardar (`:15`).

## Pendientes

- Destino inicial de `FrmDetalle` al abrir: el código no lo fija (`:45-61`); falta confirmarlo en el binario.
- Origen de datos de las colecciones del entorno (`FG.bas:75-139`), a detallar en la ficha del módulo.
