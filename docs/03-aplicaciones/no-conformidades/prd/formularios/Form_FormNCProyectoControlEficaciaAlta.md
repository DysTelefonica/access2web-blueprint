# Form_FormNCProyectoControlEficaciaAlta

> **Commit fijado:** `1d7a2f826a1e3c582874e4e29886a6183efb9067` · **Fuente:** `src/forms/Form_FormNCProyectoControlEficaciaAlta.cls` (243 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Ventana auxiliar «ALTA DE CONTROL DE EFICACIA» (`.form.txt:40`) con dos datos de la NC de proyecto activa: fecha prevista del control de eficacia y descripción del control. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): trabaja sobre el objeto global `m_ObjNCProyectoActiva` (`Variables Globales.bas:117`) y no lee ni escribe tablas por sí mismo. Las citas son `Form_FormNCProyectoControlEficaciaAlta.cls:línea` (abreviado `:línea`) salvo indicación; `General` abrevia `Form_FormNCProyectoGeneral.cls`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `FormNCProyectoGeneral` | Botón de datos de control de eficacia: exige datos generales completos, copia `RequiereControlEficacia` al objeto, cierra una instancia previa de esta ventana y la abre | `General:255-263` |

Se cierra con «Salir» (`:36-39`) o tras grabar (`:124`), siempre con `acSaveNo`. Al grabar puede ejecutar `ComandoGrabar_Click` de `FormNCProyectoGeneral` (`:114-122`; `General:277`).

## Controles y eventos

`.form.txt` declara 23 bloques `Begin`: el formulario (`:4`), 10 de valores por defecto (`:132-194`), 3 secciones (`EncabezadoDelFormulario` `:204`, `Detalle` `:242`, `PieDelFormulario` `:472`) y los 9 controles de la tabla.

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `lblTitulo` | etiqueta | — | Título «Alta control de eficacia»; se copia a `Caption` | `.form.txt:221-222`, `:226` |
| `RequiereControlEficacia` | cuadro de texto bloqueado, `Tag="DATO"` | — | Fijado a «Sí» al cargar; solo informativo | `.form.txt:408-427`, `:228` |
| `FechaPrevistaControlEficacia` | cuadro de texto, formato fecha corta, `Tag="DATO"` | `AfterUpdate` | Recalcula bordes de campos rellenos/vacíos | `.form.txt:261-265`, `:147-151` |
| `ControlEficacia` | cuadro de texto, `Tag="DATO"` | `AfterUpdate` | Ídem | `.form.txt:340-343`, `:140-143` |
| `ComandoGrabar` | botón «Grabar» | `OnClick` | Valida y traslada los datos al objeto activo | `.form.txt:370-372`, `:70-138` |
| `cmdSalir` | botón «Salir» | `OnClick` | Cierra sin guardar | `.form.txt:487-489`, `:36-39` |
| `Etiqueta254`, `Etiqueta255`, `Etiqueta256` | etiquetas | — | Rótulos de los tres campos | `.form.txt:292, 313, 456` |

## Flujos

### Apertura (`Form_Load`, `:185-211`)

1. Reloj de arena; `EstablecerDatos` (`:213-238`). Si `m_ObjNCProyectoActiva` es `Nothing`, error «El formulario se ha abierto con parámetros insuficentes» (`:217-220`).
2. Pone `Caption`, habilita «Grabar», fija `RequiereControlEficacia = "Sí"` y pinta bordes (`:226-229`). Los campos no se precargan desde el objeto: se muestran vacíos aunque la NC ya tenga valores (a verificar).

### Grabar (`ComandoGrabar_Click`, `:70-138`)

1. Si `EsTecnico = EnumSino.Sí`: «No tiene autorización para esa acción» (`:82-85`).
2. Exige fecha prevista y control de eficacia no vacíos (`:86-93`).
3. Vuelca ambos valores en `m_ObjNCProyectoActiva` y comprueba `HaHabidoCambios` contra `Form_FormNCProyecto.m_NCAlInicio` (`:94-105`, `:40-68`); sin cambios, «No hay cambios que guardar».
4. Llama a `General.RellenarDatosObjeto` (`:106`), exige `DatosGeneralesOK` (`:110-113`; `NCProyecto.cls:993-1050`).
5. Si `FormNCProyecto` está abierto y su subformulario `FrmDetalle` muestra `FormNCProyectoGeneral`, ejecuta `General.ComandoGrabar_Click` (`:114-122`), que realiza el guardado real. Cierra la ventana (`:124`).

### Errores

Error 1000 ⇒ `MsgBox` de advertencia con `m_Error`; otro error ⇒ mensaje con número y descripción, `CorreoAlAdministrador` y `MsgBox` crítico (`:129-137`, `:202-210`). `RellenarDatosObjeto` y `EstablecerDatos` devuelven el error por `p_Error` (`:30-33`, `:234-237`).

## Reglas, validaciones y cálculos

- **Rol:** los técnicos no pueden grabar (`:82-85`).
- **Obligatorios:** fecha prevista y control de eficacia (`:86-93`); además, datos generales de la NC completos (`:110-113`).
- **Cambios:** `HaHabidoCambios` compara fecha y control con la copia inicial; si `m_NCAlInicio` es `Nothing`, devuelve `True` (`:44-47`).
- **Aviso visual:** borde con color/grosor de «relleno» u «obligatorio vacío» según `Nz(campo,"")` (`:153-179`; `Variables Globales.bas:83-86`).
- **Volcado genérico:** `RellenarDatosObjeto` recorre los controles con `Tag` que contiene «DATO» y llama a `NCProyecto.SetPropiedad` (`:18-26`; `NCProyecto.cls:1971`).

## Lecturas y escrituras

| Operación | Objeto | Tipo | Evidencia |
|---|---|---|---|
| Fecha y control de eficacia | propiedades de `NCProyecto` en memoria | sin acceso a datos | `:94-97` |
| Persistencia | delegada en `General.ComandoGrabar_Click` | ver ficha de `FormNCProyectoGeneral` | `:119` |
| Registro de errores | `CorreoAlAdministrador` | ver ficha de `Form0BDOpciones` | `:133` |

Sin transacciones propias.

## Coexistencia

Sin SQL propio. La fecha se compara como texto y se guarda en una propiedad `String` (`NCProyecto.cls:37`): al pasar a PostgreSQL conviene fijar el formato de fecha (`date`) en el mapeo, no en el formulario.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Tras `RellenarDatosObjeto` se lanza `Err.Raise 100` (`:108`), no 1000: si hay error se mostrará el mensaje crítico y se enviará correo al administrador en vez de la advertencia.
- `DatosGeneralesOK` se invoca sin `p_MenosCef` (`:110`), mientras que `General` lo invoca con `EnumSino.Sí` (`General:255`); con el valor por defecto exige también los datos de control de eficacia.
- Si `FormNCProyecto` no está abierto o el detalle no es `FormNCProyectoGeneral`, se cierra la ventana sin persistir (`:114-124`).
- Los campos no se precargan con el valor existente (`:213-238`).
- `Form_FormNCProyecto.m_NCAlInicio` se lee aunque `FormNCProyecto` no esté abierto (`:44`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: integrar estos dos campos en el formulario general en lugar de ventana auxiliar.

## Criterios de aceptación

Propuestos, no aprobados:

1. Un técnico no puede grabar (`:82-85`).
2. Sin fecha prevista o sin control de eficacia, se muestra el mensaje correspondiente y no se guarda (`:86-93`).
3. Sin cambios respecto al estado inicial, «No hay cambios que guardar» (`:102-105`).
4. Con datos generales incompletos no se permite grabar (`:110-113`).
5. Tras grabar, la ventana se cierra y la NC conserva ambos valores.
6. Abrir sin NC activa muestra el aviso de parámetros insuficientes (`:217-220`).

## Pendientes

- Definición de `CorreoAlAdministrador` y `EsTecnico` fuera del formulario: `EsTecnico` se declara en `Variables Globales.bas:98`; su asignación no se ha rastreado aquí.
- Persistencia efectiva: depende de la ficha de `FormNCProyectoGeneral`.
