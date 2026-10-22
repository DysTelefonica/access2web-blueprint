# Form_FormResponsable

> **Commit fijado:** `cfa1937a` · **Fuente:** `src/forms/Form_FormResponsable.cls` (213 líneas) y `.form.txt` (931 líneas) · **PRD:** [README](../README.md)

Detalle de un responsable para alta o edición. Tres campos (Nombre, Correo, Descripción) más el botón Registrar. Es invocado por `Form_FormResponsablesGestion` (`ComandoAlta_Click`, `ComandoEditar_Click`) o por `formularioMantenimientoResponsables` de la ficha madre. Al guardar emite el evento `Alta` o `Editado`, que el contenedor consume para refrescar su lista. Las citas son `Form_FormResponsable.cls:línea` salvo que se indique el `.form.txt`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormResponsablesGestion` | `DoCmd.OpenForm "FormResponsable"` (alta) o tras asignar `m_ObjResposableActivo` (edición) | global `m_ObjResposableActivo` (`Variables Globales.bas:204`) |
| Receptor del evento `Alta` | `ComandoRegistrar_Click` cuando `m_ResponsableAlInicio Is Nothing` | `:5, 99` |
| Receptor del evento `Editado` | `ComandoRegistrar_Click` cuando `m_ResponsableAlInicio` no es `Nothing` | `:6, 100-101` |

Estado compartido: la global `m_ObjResposableActivo` (misma referencia que la ficha madre); `m_ResponsableAlInicio` es interna (`:8`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `lblTitulo` | etiqueta | — | Texto dinámico «ALTA DE RESPONSABLE» o «EDICIÓN DE RESPONSABLE» | `:148-153`, `.form.txt:225-226` |
| `Nombre` | texto | — | Nombre del responsable | `.form.txt:339-349` |
| `Correo` | texto | — | Correo; se valida con `EsCorreoValido` antes de grabar | `.form.txt:394-405`, `ResponsableOperaciones.cls:46` |
| `Descripcion` | texto (memo) | — | Descripción libre | `.form.txt:450-461` |
| `ComandoRegistrar` | botón | `OnClick` | Valida, graba y emite evento | `:55-103`, `.form.txt:629-633` |
| `ComandoAyuda` | botón | `OnClick` | Llama a `AbrirAyuda` | `:21-47`, `.form.txt:244-248` |
| `cmdSalir` | botón | `OnClick` | Cierra el formulario (`acSaveNo`) | `:15-18`, `.form.txt:523-525` |

## Flujos

### Apertura (`Form_Open`, `:196-227`)

1. `Ajustar Me`, fija `Caption = m_TituloFormulario` y deja `ComandoRegistrar.Enabled = False` (`:201-204`).
2. Si `EsTecnico <> Sí`, habilita el botón y permite edición (`:207-211`).
3. Llama a `EstablecerDatos` (`:213`) y, tras éxito, desbloquea y reinicia el cursor.

### Datos iniciales (`EstablecerDatos`, `:146-189`)

1. Si `m_ObjResposableActivo Is Nothing`, pinta «ALTA DE RESPONSABLE» y sale sin poblar (`:148-153`).
2. Para edición, pinta «EDICIÓN DE RESPONSABLE», recarga con `constructor.getResponsable(p_IDResponsable:=…)` (`constructor.bas:3431-3474`) y, si falla, aborta con error (`:155-164`).
3. Vuelca `Nombre`, `Correo` y `Descripcion` del responsable inicial (`:170-180`).

### Guardar (`ComandoRegistrar_Click`, `:55-103`)

1. `HaHabidoCambios` (`:106-141`) compara contra `m_ResponsableAlInicio`; si no hay cambios, aborta con «No hay cambios que guardar» (`:62-69`).
2. Si `m_ObjResposableActivo Is Nothing`, instancia un `Responsable` nuevo (`:71-73`) — sólo es válido el alta que pasa por `ComandoAlta_Click`, pero la defensa cubre otras entradas.
3. Copia `Me.Nombre`, `Me.Correo`, `Me.Descripcion` al objeto (`:74-79`).
4. `ResponsableOperaciones.Registrar m_ResponsableAlInicio, m_Error` (`ResponsableOperaciones.cls:30-118`) valida, graba y limpia `m_ObjEntorno.Responsables = Nothing` (`:84-92`).
5. Emite `Alta` o `Editado` según haya `m_ResponsableAlInicio` (`:98-101`) y cierra el formulario (`:102`).

### Cambios pendientes (`HaHabidoCambios`, `:106-141`)

Compara tres campos uno a uno (`Nombre`, `Correo`, `Descripcion`) contra `m_ResponsableAlInicio`. Si es alta, devuelve `True` sin comparar (`:113-117`). Si la comparación coincide, devuelve `False`.

### Errores

Cada handler acumula `m_Error`, envía `CorreoAlAdministrador` cuando procede y muestra `MsgBox vbCritical`; `Form_Open` cierra el formulario tras el aviso (`:222-227`). `EstablecerDatos` y `HaHabidoCambios` sólo devuelven el texto en `p_Error`.

## Reglas, validaciones y cálculos

- **Permisos**: `EsTecnico = Sí` deshabilita `ComandoRegistrar` (`:212`); la validación vive en `Variables Globales.bas:192-193, 318-329`.
- **Validación de correo**: `EsCorreoValido(.Correo)` en `ResponsableOperaciones.Registrar` (`:46`) y revalidación de unicidad vía `constructor.getResponsable(p_Correo:=…)` (`:48-58`).
- **Concurrencia**: si la edición cambia `Correo`, vuelve a comprobar unicidad (`:50-58`).
- **Campos opcionales**: `Descripcion` se almacena `Null` si el alta la deja vacía y como `Null` en edición si se vacía (`ResponsableOperaciones.cls:99-108`).
- **Generación de ID**: `Responsable.IDResponsableCalculado` (`Responsable.cls:24-41`) usa `DameID("TbResponsables","IDResponsable", getdb())`; numeración delegada a la base.

## Lecturas y escrituras

| Operación | Tabla | Tipo | Evidencia |
|---|---|---|---|
| Carga por ID | `TbResponsables` | SQL dinámico concatenado | `constructor.bas:3431-3474` |
| Carga por correo | `TbResponsables` | SQL dinámico concatenado | `constructor.bas:3431-3474` (sobrecarga por `p_Correo`) |
| Insert / Update | `TbResponsables` | DAO `AddNew`/`Edit` | `ResponsableOperaciones.cls:73-110` |
| Reset de caché | `m_ObjEntorno.Responsables = Nothing` | en memoria | `ResponsableOperaciones.cls:114` |

Sin borrado propio; la baja está centralizada en la ficha madre (`Form_FormResponsablesGestion.cls:179-217`).

## Coexistencia

Pendiente de `../coexistencia.md`. Mismas observaciones que la ficha madre: identificadores numéricos concatenados en `WHERE` (`ResponsableOperaciones.cls:81, 171`) y `LIKE` para nombre en `constructor.bas:2597-2599`. La validación de unicidad por correo conviene trasladarla a una restricción UNIQUE en el esquema.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- La propiedad `Caption` del formulario quedó como «Edición de Suministrador» (`.form.txt:40`), rótulo heredado de un clon previo. `lblTitulo.Caption` lo corrige a «EDICIÓN RESPONSABLE» (`.form.txt:226`), pero el caption de la ventana Access queda inconsistente.
- `ComandoAyuda` tiene `Caption = "Salir"` en el diseño (`.form.txt:245`); al ser un botón con imagen (`.form.txt` con `Picture = "Ayda1_25x25.png"`) el efecto sólo se ve al pasar el puntero. Inocuo, pero defecto de rotulación.
- `ComandoRegistrar` tiene `Caption = "Editar"` (`.form.txt:630`) y `ControlTipText = "Registrar"` (`.form.txt:633`). El handler lo trata como «Registrar»; el tooltip lo aclara, pero el rótulo visible es engañoso.
- `EsTecnico` sólo deshabilita `ComandoRegistrar` pero deja `AllowEdits = True` en la misma condición negada (`:212-218`); no protege contra edición manual por teclado si el botón queda habilitado por otros flujos.
- `HaHabidoCambios` no compara el caso `Descripcion = ""` en el alta (`:117`): el alta siempre sale por la rama `True` antes de comparar; si se entra en edición con un valor inicial vacío y se introduce texto, se detecta correctamente (`:135-138`).
- `EsCorreoValido` no se aplica en `EstablecerDatos`; un correo mal formado vivo en BBDD se mostraría tal cual.
- En `ComandoRegistrar_Click`, `m_ObjResposableActivo = New Responsable` (`:71-73`) salta la validación de unicidad por correo cuando se abre el formulario sin pasar por la gestión (no es el flujo previsto, pero conviene blindarlo).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: edición desacoplada del binario Access, validación de correo reutilizable desde la API de catálogos y borrado lógico (`FechaBaja`) en lugar de `DELETE *` en `TbResponsables`.

## Criterios de aceptación

Propuestos, no aprobados:

1. Invocado sin `m_ObjResposableActivo`, el rótulo dice «ALTA DE RESPONSABLE», los campos están vacíos y `ComandoRegistrar` está habilitado salvo para técnico.
2. Invocado con un responsable, el rótulo dice «EDICIÓN DE RESPONSABLE» y los campos coinciden con los valores en base.
3. Sin cambios, pulsar Registrar muestra «No hay cambios que guardar».
4. Con cambios, valida el formato del correo; si es único, lo graba; si ya existe, aborta mostrando el aviso.
5. Tras grabar, el contenedor de la lista recibe `Alta` o `Editado` y el formulario se cierra.

## Pendientes

- Política actual de contraseñas para responsables: el modelo `Responsable.cls:14-17` no la tiene y `Usuario.cls:16` sí (`Password`). Si el alta de responsable llega a generar un usuario implícito, el puente no está documentado en este archivo.
- Limpieza del caché `m_ObjEntorno.Responsables` (`ResponsableOperaciones.cls:114`) cuando la edición se hace desde una sesión distinta a la del alta.
- Diferencia entre flujo `EsTecnico = Sí` que llega a `ComandoRegistrar.Enabled = False` y la persistencia de los datos en pantalla: ¿se pierden al pulsar Cerrar?
