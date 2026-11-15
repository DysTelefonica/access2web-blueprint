# Form_FormMotivoHPS

> **Commit fijado:** `cfa1937ab05a9e40150ea9236243898de96f84da` · **Fuente:** `src/forms/Form_FormMotivoHPS.cls` (124 líneas) y `.form.txt` (594 líneas) · **PRD:** [README](../README.md)
>
> **Abreviaturas:** `FU.bas` = `modules/FUNCIONES UTILES.bas`, `VG.bas` = `modules/Variables Globales.bas`; las clases de `src/classes/` se citan por el nombre del fichero. Las citas son `Form_FormMotivoHPS.cls:línea` del commit fijado.

Diálogo modal de alta o edición de un motivo del catálogo `TbMotivoHPS`. Cuatro procedimientos en total. El doble comportamiento (alta o edición) se decide por la presencia de `m_ObjMotivoActivo` (`VG.bas:199`); si está vacío se da de alta, si no se renombra el motivo. La persistencia delega en `MotivoHPS.AltaMotivo` y `MotivoHPS.EditarMotivo`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormMotivosHPS` (catálogo) | `ComandoAlta_Click` deja `m_ObjMotivoActivo = Nothing` y abre el diálogo (`:57-63`) | `Form_FormMotivosHPS.cls:55-78` |
| `Form_FormMotivosHPS` (catálogo) | `ComandoEditar_Click` asigna `Set m_ObjMotivoActivo = m_MotivoSeleccionado` y abre el diálogo (`:94-99`) | `Form_FormMotivosHPS.cls:82-115` |

Eventos hacia el exterior: `Alta` (`:5`, lanzado en `:64`) y `Editado` (`:6`, lanzado en `:81`). Estado compartido: la global `m_ObjMotivoActivo` (`VG.bas:199`) y `m_ObjEntorno` (`VG.bas:190`, cuya `ColMotivos` se anula en el alta, `:59`). El formulario no declara `OpenArgs`; la decisión alta/edición se infiere de la instancia.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `lblTitulo` | etiqueta «MOTIVO HPS» | — | Título que `Form_Open` copia a la `Caption` del formulario (`:104`) | `:225-227`, `:104` |
| `MotivoHPS` | texto | — | Texto del motivo; validado en `ComandoGrabar_Click` (`:44-45`) | `:258-263` |
| `cmdSalir` | botón «Salir» | `OnClick` | Cierra con `acSaveNo` sin capturador de error (`:9-12`) | `:320-322`, `:9-12` |
| `ComandoGrabar` | botón «Registrar» | `OnClick` | Valida texto y oficina; crea o edita el motivo; lanza `Alta` o `Editado` y cierra (`:38-92`) | `:426-428`, `:38-92` |

## Flujos

### Apertura (`Form_Open`, `:94-124`)

1. `Ajustar Me` (`:103`; `FU.bas:735-759`).
2. `Me.Caption = Me.lblTitulo.Caption` (`:104`) para unificar la barra de título con la etiqueta interna.
3. `EstablecerDatos m_Error` (`:105`): si hay `m_ObjMotivoActivo`, copia su `MotivoHPS` a `Me.MotivoHPS` (`:25-31`).
4. `AvanceCerrar` (`:109`; `FU.bas:555-582`).
5. Reactiva `Me.AllowEdits` (`:112`).

### Alta (`ComandoGrabar_Click` con `m_ObjMotivoActivo Is Nothing`, `:52-67`)

1. Tras validar texto y oficina (`:44-50`), crea `New MotivoHPS`, le asigna `MotivoHPS = Me.MotivoHPS` y llama a `m_ObjMotivoActivo.AltaMotivo m_Error` (`:53-57`; clase `MotivoHPS.cls:60-95`).
2. Anula `m_ObjEntorno.ColMotivos` (`:59`) para forzar recarga del caché.
3. Muestra el aviso «Registrado correctamente» (`:63`), lanza `RaiseEvent Alta` (`:64`) y cierra con `acSaveNo` (`:65`).

### Edición (`ComandoGrabar_Click` con `m_ObjMotivoActivo` existente, `:69-83`)

1. Exige cambio real frente al valor en memoria (`:69-71`); aborta con «Sin cambios que guardar» si no lo hay.
2. Llama a `m_ObjMotivoActivo.EditarMotivo m_ObjMotivoActivo.MotivoHPS, Me.MotivoHPS, m_Error` (`:73-76`; `MotivoHPS.cls:22-60`); ese método propaga el cambio a `TbSolicitudes`, `TbUsuarios` y `TbUsuariosHistoricos` (`MotivoHPS.cls:180-182, 208-213`).
3. Muestra el aviso, lanza `RaiseEvent Editado` (`:81`) y cierra.

### Errores

Los cuatro capturadores diferencian el error interno 1000 (aviso previsto, se reporta sin `CorreoAlAdministrador`) del resto, donde se envía el correo (`FU.bas:585`) y se muestra `MsgBox` crítico (`:86-91, 116-123`). `cmdSalir_Click` (`:9-12`) no captura nada.

## Reglas, validaciones y cálculos

- **Texto obligatorio:** `Nz(Me.MotivoHPS, "") = ""` aborta con «Introduzca el motivo» (`:44-46`).
- **Oficina obligatoria:** `m_EnOficina = EnumSiNo.no` aborta con «Sólo puede hacerse en la oficina» (`:48-50`). El mismo chequeo se duplica dentro de `MotivoHPS.EditarMotivo` (`MotivoHPS.cls:32-35`).
- **Cambio real exigido:** la rama de edición solo guarda si el texto en memoria difiere del control (`:69-71`).
- **Renombrado en cascada:** `EditarMotivo` reescribe `Motivo_HPS` en `TbSolicitudes`, `TbUsuarios` y `TbUsuariosHistoricos` (clase `MotivoHPS.cls:180-182, 208-213`). No aparecen trazas de `TbSolicitudesHistoricos`; queda como punto a revisar.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Alta | `TbMotivoHPS` (Recordset: AddNew) | DAO Recordset | `MotivoHPS.cls:82-84` |
| Edición | `TbMotivoHPS` (UPDATE), `TbSolicitudes` (UPDATE), `TbUsuarios` (UPDATE), `TbUsuariosHistoricos` (UPDATE) | `getdb().Execute` | `MotivoHPS.cls:41-44, 180-182, 208-213` |
| Validación previa | `TbMotivoHPS` (existencia por `MotivoHPS`) | DAO Recordset | `MotivoHPS.cls:74-77` |

## Coexistencia

Pendiente de `../coexistencia.md`. Lo que ya se ve: la unicidad del motivo se valida por SQL abierta en Jet/ACE con `WHERE MotivoHPS='<texto>'` (`MotivoHPS.cls:74-77`); la validación web debe ser `EXISTS` sobre una clave única y no por comparación de cadenas escapadas. La propagación en cascada a tres tablas sin transacción es una decisión de «todo o nada» a revisar en la versión web.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-HPS-S-28.** Los `UPDATE` en cascada dentro de `EditarMotivo` (`MotivoHPS.cls:41-44, 180-182, 208-213`) no están envueltos en transacción: una caída entre el `UPDATE` del catálogo y los `UPDATE` de las tablas satelite deja el motivo renombrado en algunos sitios y sin renombrar en otros.
- `cmdSalir_Click` cierra con `acSaveNo` sin capturador de error (`:9-12`); el patrón del resto de formularios del proyecto avisa antes de descartar.
- `EstablecerDatos` no rellena `Me.MotivoHPS` cuando `m_ObjMotivoActivo` existe pero su `MotivoHPS` viene vacío (`:25-31`): el control arranca en blanco y el operador no detecta que está editando.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: exponer el catálogo como endpoint REST con validación de clave única; mover la propagación en cascada a una transacción servidor; registrar la edición en una tabla de auditoría.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir el diálogo sin `m_ObjMotivoActivo` da de alta; con uno cargado edita.
2. Sin texto ni estando en oficina, el botón «Registrar» aborta con el aviso correspondiente.
3. Tras un alta, `m_ObjEntorno.ColMotivos` queda `Nothing`.
4. Tras una edición, el `MotivoHPS` del motivo en memoria refleja el nuevo texto y `TbSolicitudes`, `TbUsuarios` y `TbUsuariosHistoricos` han propagado el cambio.
5. «Salir» cierra el diálogo sin preguntar.
6. El alta lanza `Alta` y la edición lanza `Editado` exactamente una vez por pulsación.

## Pendientes

- Política sobre `TbSolicitudesHistoricos`: ¿debería propagarse también el renombrado?
- Si conviene consolidar el chequeo de oficina, hoy duplicado en formulario y en clase.
