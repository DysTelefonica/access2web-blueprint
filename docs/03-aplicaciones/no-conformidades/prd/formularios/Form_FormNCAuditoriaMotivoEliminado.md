# Form_FormNCAuditoriaMotivoEliminado

> **Commit fijado:** `1d7a2f826a1e3c582874e4e29886a6183efb9067` · **Fuente:** `src/forms/Form_FormNCAuditoriaMotivoEliminado.cls` (210 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Formulario de un solo campo que muestra y permite corregir el motivo con el que una no conformidad (NC) de auditoría figura como borrada. Es **no enlazado** (sin `RecordSource`) y trabaja sobre la NC activa global `m_ObjNCAuditoriaActiva` (`VG.bas:127`) y la copia inicial `Form_FormNCAuditoria.m_NCAlInicio`. Las citas son `Form_FormNCAuditoriaMotivoEliminado.cls:línea` salvo indicación; `VG.bas` abrevia `modules/Variables Globales.bas` y `Ops`, `classes/NCaUDITORIAOperaciones.cls`.

## Entrada y salida

No hay en `src` ninguna llamada ni destino de navegación a este nombre fuera de sus propios ficheros. `Form_FormNCAuditoria.form.txt:1289-1306` tiene un control `NavMotivoBorrado` con destino `FormNCProyectoAuditoriaEliminado`, objeto que no existe en `src/forms` (ver Pendientes). No abre otros formularios; sin botón de cierre.

## Controles y eventos

`.form.txt` declara 13 bloques `Begin`: el formulario (`.form.txt:5`), 8 de valores por defecto (`:141-278`), la sección `Detalle` (`:299`) y 3 controles.

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `MotivoBorrado` | cuadro de texto, `Tag = "DATO"` | `AfterUpdate` | Recolorea el borde según esté relleno o no | `:107-111`; `.form.txt:324-327` |
| `Etiqueta50` («Motivo») | etiqueta | — | Rótulo | `.form.txt:354-355` |
| `ComandoGrabar` («Grabar») | botón, inicialmente deshabilitado | `OnClick` | Guarda el motivo | `:158-211`; `.form.txt:374-386` |
| Formulario | — | `Form_Load` | Carga datos de la NC activa | `:9-35` |

## Flujos

### Carga (`Form_Load` → `EstablecerDatos`, `:9-79`)
1. Exige `m_ObjNCAuditoriaActiva` con `id` no vacío y `m_NCAlInicio` informado; si no, «El formulario se ha abierto con parámetros insuficientes» (`:44-56`).
2. Si la NC inicial figura como borrada, habilita «Grabar» (`:59-61`) y vuelca su `MotivoBorrado` al campo si no está vacío (`:63-70`).
3. Pinta el borde (`EstableceColorBordes`, `:80-105`).

### Grabar (`ComandoGrabar_Click`, `:158-211`)
1. Rechaza a técnicos: «No tiene autorización para esa acción» (`:169-172`).
2. Compara el campo con `m_NCAlInicio.MotivoBorrado`; sin diferencias: «No hay cambios que guardar» (`:173-180`, `:112-133`).
3. Copia los controles con `Tag = "DATO"` a la NC activa con `SetPropiedad` (`:134-157, 181`).
4. Llama a `Ops.ModificarMotivoBorrado` (`:185-193`) y refleja el nuevo motivo en `m_NCAlInicio` (`:195`), tras lo que recarga (`:196`).

### Errores
Error 1000 ⇒ aviso con `m_Error`; otro error ⇒ número y descripción, `CorreoAlAdministrador` y `MsgBox` crítico (`:26-34, 202-210`). `EstablecerDatos` asigna a `m_Error` y no a `p_Error` en un caso (`:50`).

## Reglas, validaciones y cálculos

- Solo no técnicos pueden grabar (`:169-172`).
- Motivo obligatorio y NC borrada: `Ops.ModificarMotivoBorrado` rechaza NC sin `id` («No parece estar registrada la No Conformidad»), motivo vacío («Se ha de indicar el motivo») y NC no borrada («No aparece como borrado») (`Ops:808-818`).
- Borde: color de campo relleno o no relleno según haya motivo y la NC activa esté o no borrada (`:84-100`; constantes en `VG.bas:84-85`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Evidencia |
|---|---|---|---|
| Actualiza `MotivoBorrado` | `TbNoConformidadesAuditoria` por `ID` | SQL dinámico concatenado con el texto del usuario (`Execute`) | `Ops:820-822` |
| Registro en log | `RegistrarLogAuditoria "Registrar Motivo Borrado"` | utilidad | `Ops:823` |

Sin transacciones. El formulario no lee datos propios.

## Coexistencia

El motivo se concatena en la sentencia sin escapar comillas (`Ops:820-821`): con comillas simples en el texto falla en Jet y sería inyectable en PostgreSQL; debe pasar a parámetro. Fuera de eso, `UPDATE … WHERE ID=` es compatible. Pendiente de la coexistencia (documento aún no redactado).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `HaHabidoCambios` devuelve falso si `m_NCAlInicio.MotivoBorrado` es nulo frente a un campo vacío (`:121`).
- «Grabar» solo se habilita si la NC *inicial* está borrada (`:59-61`), pero `Ops` valida la NC activa (`Ops:816`); si divergen no hay aviso.
- La rama `p_Error` de `RellenarDatosObjeto` lanza el error 100 (`:182-184`), que cae en el bloque de errores no controlados (`:204-207`).
- Con `On Error Resume Next` en el recoloreado (`:82, 108`) un fallo pasa inadvertido.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: fusionar motivo y marcado como borrada en un único diálogo (hoy el motivo se pide también por `InputBox` en la pantalla de gestión).

## Criterios de aceptación

Propuestos, no aprobados:

1. Abierto sin NC activa, informa de parámetros insuficientes (`:44-48`).
2. Con NC borrada, muestra su motivo y habilita «Grabar» (`:59-70`).
3. Un técnico no puede grabar (`:169-172`).
4. Sin cambios no escribe y avisa (`:177-180`).
5. Con motivo vacío no escribe (`Ops:812-815`).
6. Tras grabar, el motivo guardado se muestra de nuevo (`:195-196`).

## Pendientes

- Quién abre este formulario: `src` no lo referencia; el control `NavMotivoBorrado` apunta a `FormNCProyectoAuditoriaEliminado` (`Form_FormNCAuditoria.form.txt:1306`), objeto inexistente en main.
