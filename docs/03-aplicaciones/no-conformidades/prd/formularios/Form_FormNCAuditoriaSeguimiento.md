# Form_FormNCAuditoriaSeguimiento

> **Commit fijado:** `1d7a2f826a1e3c582874e4e29886a6183efb9067` · **Fuente:** `src/forms/Form_FormNCAuditoriaSeguimiento.cls` (69 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Contenedor del seguimiento de NC de auditoría: título, botón «Contador» y un control de navegación con dos pestañas («Tareas» y «Rev. Estados») que cargan en el subformulario `FrmDetalle` los formularios `FormNCAuditoriaSeguimientoTareas` (`Form_FormNCAuditoriaSeguimientoTareas`) y `FormNCAuditoriaSeguimientoNC` ([Form_FormNCAuditoriaSeguimientoNC](Form_FormNCAuditoriaSeguimientoNC.md)) (`.form.txt:1009, 1061`). No tiene `RecordSource` en el `.form.txt` (no enlazado) y no lee ni escribe datos propios. Las citas son `Form_FormNCAuditoriaSeguimiento.cls:línea` salvo indicación; `FG.bas` abrevia `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Form_Form0BDOpcionesAuditorias | Botón `ComandoSeguimiento`: si el formulario ya está abierto lo cierra sin guardar y lo vuelve a abrir | `Form_Form0BDOpcionesAuditorias.cls:223-226` |

Se cierra con «Salir» (`:8-11`, `DoCmd.Close acForm … acSaveNo`) o, si falla la carga, desde el manejador de errores de `Form_Load` (`:66`, inalcanzable, véase Posibles defectos).

## Controles y eventos

`.form.txt` declara 25 bloques `Begin`: el formulario (`.form.txt:4`), un bloque de valores por defecto con 11 tipos (`.form.txt:718-806`), 3 secciones (`EncabezadoDelFormulario` `:840`, `Detalle` `:942`, `PieDelFormulario` `:1167`) y los 9 controles de la tabla.

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `lblTitulo` | etiqueta | — | Texto «SEGUIMIENTO DE TAREAS Y REVISIÓN DE ESTADOS DE NC DE AUDITORÍA»; `Form_Load` lo copia a `Me.Caption` | `:50`, `.form.txt:858-859` |
| `ComandoActualizar` («Contador») | botón | `OnClick` | Recalcula los indicadores con `PintarIndicadores` reiniciando las colecciones | `:13-38`, `.form.txt:877-879` |
| `Etiqueta237` | etiqueta | — | Decorativa, sin evento | `.form.txt:952` |
| `ControlDeNavegación0` | control de navegación | — | Contenedor de las pestañas | `.form.txt:966` |
| `NavTareas` («Tareas») | botón de navegación | destino sin handler | Carga `FormNCAuditoriaSeguimientoTareas` en `FrmDetalle` | `.form.txt:994-1010` |
| `NavNC` («Rev. Estados») | botón de navegación | destino sin handler | Carga `FormNCAuditoriaSeguimientoNC` en `FrmDetalle` | `.form.txt:1044-1062` |
| `BotónDeNavegación1` («[Agregar nuevo]») | botón de navegación | — | Marcador de posición del asistente de Access, sin destino | `.form.txt:1089-1096` |
| `FrmDetalle` | subformulario | — | Hospeda el formulario elegido; sin `LinkChildFields` ni `LinkMasterFields` | `.form.txt:1139-1147` |
| `cmdSalir` («Salir») | botón | `OnClick` | Cierra el formulario | `:8-11`, `.form.txt:1186-1188` |

## Flujos

### Carga (`Form_Load`, `:40-67`)

1. Cursor de espera, `m_Error = ""`, `Ajustar Me` (`:43-46`; `FG.bas:615`).
2. `Me.Caption = Me.lblTitulo.Caption` y cursor normal (`:50-55`).
3. El subformulario muestra la pestaña de navegación que Access seleccione; el código no fija `SourceObject` (`:40-56`). El destino inicial queda pendiente de verificar en el binario.

### Contador (`ComandoActualizar_Click`, `:13-38`)

Llama a `PintarIndicadores p_Reiniciando:=EnumSino.Sí` (`:21`; `FG.bas:75`), que reinicia las colecciones del entorno, recuenta los indicadores de proyecto y auditoría, rotula `lblSeguimientos` en los menús abiertos (`FG.bas:320-324`) y refresca el detalle abierto: `Filtrar` del formulario de NC o de tareas según su `SourceObject` (`FG.bas:332-338`).

### Errores

Error 1000 ⇒ `MsgBox` de advertencia con `m_Error`; otro ⇒ mensaje con número y descripción, `CorreoAlAdministrador` y `MsgBox` crítico (`:29-37`).

## Reglas, validaciones y cálculos

- El formulario no valida entradas ni comprueba el rol.
- El indicador total de auditoría y el del usuario se calculan en `PintarIndicadores` sumando recuentos de colecciones del entorno (`FG.bas:75-139`); el detalle de la suma queda en la ficha del menú de auditorías.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Ninguna directa | — | — | `:1-69` |
| Indirecta: recuentos del entorno | Colecciones `ColSegs…` de `m_ObjEntorno` | DAO/SQL ejecutado en `Entorno`, fuera de esta ficha | `FG.bas:147-181` |

## Coexistencia

Sin SQL propio, no hay dialecto Jet/ACE que adaptar en el formulario. Dependen de PostgreSQL vinculado las consultas que alimentan las colecciones del entorno (fuera de alcance de esta ficha). `Ajustar` y el cursor de espera son de interfaz Access, sin equivalente directo en web.

## Comportamiento actual · Posibles defectos · Propuestas web

**Comportamiento actual**

- Contenedor sin datos propios con dos pestañas y un botón de recuento (`.form.txt:959-1147`).
- «Contador» actualiza también los menús y los formularios de proyecto abiertos, no solo el de auditoría (`FG.bas:320-338`).

**Posibles defectos (a verificar)**

- `Form_Load` tiene la etiqueta `errores:` pero ningún `On Error GoTo errores`; el manejador y el cierre de `:57-66` no se ejecutan (`:40-67`).
- El mensaje de error de `ComandoActualizar_Click` dice «ComandoConfiguracion_Click» (`:32`) y el de `Form_Load` dice «Form_Open» (`:60`).
- `BotónDeNavegación1` es un marcador sin destino (`.form.txt:1089-1105`).

**Propuestas web pendientes de decisión**

- Sustituir el control de navegación por pestañas de la aplicación y cargar el detalle bajo demanda.
- Recalcular el contador al cambiar de pestaña en lugar de por botón.

## Criterios de aceptación

1. Desde el menú de auditorías, «Seguimiento» abre el formulario y, si ya estaba abierto, lo reabre (`Form_Form0BDOpcionesAuditorias.cls:223-226`).
2. «Tareas» y «Rev. Estados» muestran en `FrmDetalle` el formulario correspondiente (`.form.txt:1009, 1061`).
3. «Contador» recalcula los indicadores y refresca el detalle visible (`:21`; `FG.bas:332-338`).
4. Un error 1000 muestra advertencia; cualquier otro, mensaje crítico y aviso al administrador (`:29-37`).
5. «Salir» cierra sin guardar (`:10`).

## Pendientes

- Destino inicial de `FrmDetalle` al abrir: el código no lo fija (`:40-56`); falta confirmarlo en el binario.
- Origen de datos de `ColSegs…` en `Entorno` (`FG.bas:147-181`), a detallar en la ficha del módulo.
