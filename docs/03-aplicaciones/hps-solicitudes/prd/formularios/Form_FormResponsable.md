# Form_FormResponsable

> **Commit fijado:** `cfa1937ab05a9e40150ea9236243898de96f84da` · **Fuente:** `src/forms/Form_FormResponsable.cls` (214 líneas) y `.form.txt` (25 bloques `Begin`) · **PRD:** [README](../README.md)
>
> **Abreviaturas:** `FU.bas` = `modules/FUNCIONES UTILES.bas`, `VG.bas` = `modules/Variables Globales.bas`; las clases de `src/classes/` se citan por el nombre del fichero. Las citas son `Form_FormResponsable.cls:línea` salvo indicación.

Alta y edición de un responsable de solicitudes: nombre, correo y descripción. Es un formulario **no enlazado**; decide si está en alta o en edición a partir de la global `m_ObjResposableActivo` (`VG.bas:204`), carga la fotografía inicial con `constructor.getResponsable` y delega la escritura en `ResponsableOperaciones.Registrar`. Avisa al formulario padre con los eventos `Alta` o `Editado` (`:4-5`, lanzados en `:81-85`).

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormResponsablesGestion` | `ComandoAlta_Click` pone `m_ObjResposableActivo = Nothing` y abre el formulario; `ComandoEditar_Click` le asigna el responsable elegido | `Form_FormResponsablesGestion.cls:133-163, 165-203` |
| `Form_FormSolicitudRenovacionDatos` | `ComandoNuevoResponsable_Click` limpia la global y abre el formulario | `Form_FormSolicitudRenovacionDatos.cls:305-312` |
| `Form_FormSolicitudesAltaTraspasoDatos` | Igual, desde su botón de nuevo responsable | `Form_FormSolicitudesAltaTraspasoDatos.cls:299-306` |

No declara parámetros y no abre formularios. Al terminar el registro se cierra a sí mismo (`:89`). El título de la ventana lo fija `m_TituloFormulario` (`:189`), una global de sesión (`VG.bas:196`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| Formulario | — | `OnOpen` → `Form_Open` | Dimensiona, aplica la restricción de perfil, carga los datos y cierra la etapa de carga | `:178-213`, `.form.txt:40-41` |
| `lblTitulo` | etiqueta de encabezado | — | «ALTA DE RESPONSABLE» o «EDICIÓN DE RESPONSABLE», según haya responsable activo | `:139, 143`, `.form.txt:225-226` |
| `ComandoAyuda` | botón | `OnClick` → `ComandoAyuda_Click` | Abre el PDF de ayuda | `:14-39`, `.form.txt:244-246` |
| `Nombre` | cuadro de texto | — | Nombre del responsable | `:67`, `.form.txt:349` |
| `Etiqueta232` | etiqueta | — | Rótulo «NOMBRE» | `.form.txt:379-380` |
| `Correo` | cuadro de texto | — | Correo del responsable | `:68`, `.form.txt:405` |
| `Etiqueta245` | etiqueta | — | Rótulo «CORREO» | `.form.txt:435-436` |
| `Descripcion` | cuadro de texto | — | Descripción libre | `:69`, `.form.txt:461` |
| `Etiqueta256` | etiqueta | — | Rótulo «DESC.» | `.form.txt:491-492` |
| `ComandoRegistrar` | botón | `OnClick` → `ComandoRegistrar_Click` | Comprueba que haya cambios, copia los campos y registra | `:42-99`, `.form.txt:629-631` |
| `cmdSalir` | botón «Salir» | `OnClick` → `cmdSalir_Click` | Cierra el formulario | `:9-12`, `.form.txt:523-525` |

## Flujos

### Apertura (`Form_Open`, `:178-213`)

1. `Ajustar Me` dimensiona el formulario (`:186`; `FU.bas:735-759`).
2. Deja deshabilitados el botón de registro y la edición (`:187-188`).
3. Fija el título de la ventana con `m_TituloFormulario` (`:189`).
4. Si el usuario no es técnico, rehabilita el botón y la edición (`:190-193`); la condición usa la global `EsTecnico` (`VG.bas:193`, alimentada desde `VG.bas:324`).
5. `EstablecerDatos` decide el rótulo y carga la fotografía inicial (`:195-198`).
6. Cierra la etapa de carga con `AvanceCerrar` (`:199`; `FU.bas:555-582`).

### Carga de datos (`EstablecerDatos`, `:133-176`)

1. Sin responsable activo, rotula «ALTA DE RESPONSABLE» y sale (`:138-142`); no carga nada más.
2. Con responsable activo, rotula «EDICIÓN DE RESPONSABLE» (`:143`) y pide la fotografía inicial con `constructor.getResponsable` por `IDResponsable` (`:145`).
3. Si la lectura falla, propaga el error (`:146-148`); si no encuentra el responsable, avisa «No se ha podido encontrar el Responsable registrado» (`:149-152`).
4. Copia los tres campos a los controles, pero sólo si en el registro no están vacíos (`:155-166`).

### Registro (`ComandoRegistrar_Click`, `:42-99`)

1. `HaHabidoCambios` decide si hay algo que guardar (`:54-61`); si no, avisa «No hay cambios que guardar».
2. Si no hay responsable activo, crea uno nuevo (`:62-64`).
3. Copia `Nombre`, `Correo` y `Descripcion` con `Nz` (`:65-71`).
4. `ResponsableOperaciones.Registrar` valida y escribe (`:73-80`).
5. Lanza `Alta` o `Editado` según hubiera fotografía inicial (`:81-85`) y cierra la ventana (`:89`).

### Errores

- `cmdSalir_Click` usa `On Error Resume Next` sin aviso (`:9-12`).
- La rama `errores:` de `ComandoAyuda_Click` ignora el centinela 1000 y, para el resto, envía correo al administrador y muestra un `MsgBox` (`:31-38`).
- `ComandoRegistrar_Click` hace lo mismo (`:91-98`).
- `Form_Open` propaga el error, avisa y cierra el formulario (`:205-213`); en esa rama también reactiva la edición (`:210`).
- `HaHabidoCambios` y `EstablecerDatos` llevan manejador propio que sólo rellena `p_Error` (`:127-131`, `:170-173`); no envían correo.

## Reglas, validaciones y cálculos

- **Perfil técnico:** el usuario técnico no puede registrar; el botón y la edición quedan deshabilitados (`:187-193`). La condición es la global `EsTecnico`, no el permiso por entidad.
- **Nombre obligatorio:** lo exige la capa de datos (`ResponsableOperaciones.cls:27-30`).
- **Correo con formato válido:** lo exige la capa de datos mediante `EsCorreoValido` (`ResponsableOperaciones.cls:31-34`; `FU.bas:990`).
- **Correo único:** si el correo cambia o se trata de un alta, se busca un responsable con ese correo y se rechaza la duplicidad (`ResponsableOperaciones.cls:36-53`).
- **Detección de cambios:** la compara campo por campo contra la fotografía inicial (`:101-131`); en alta siempre hay cambios (`:104-107`).
- **Escritura:** `ResponsableOperaciones.Registrar` inserta o actualiza `TbResponsables` con un `Recordset` en modo `AddNew`/`Edit` (`ResponsableOperaciones.cls:56-90`) e invalida la caché de responsables del entorno (`:92`).
- **Identificador del alta:** se reserva con `Max(id)+1` en `Responsable.IDResponsableCalculado` (`Responsable.cls:32`; `FU.bas:760-793`, consulta `Max` en `:775`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Fotografía inicial | vía `constructor.getResponsable` | DAO, consulta por `IDResponsable` | `:146`, `constructor.bas:3431` |
| Validación de correo duplicado | vía `constructor.getResponsable` | DAO, consulta por `Correo` | `ResponsableOperaciones.cls:43` |
| Alta o edición | `TbResponsables` | DAO (`Recordset` en `AddNew`/`Edit`) | `ResponsableOperaciones.cls:56-87` |
| Invalidación de caché | caché de responsables del entorno | Derivado | `ResponsableOperaciones.cls:88` |

## Coexistencia

Pendiente de `../coexistencia.md`. Lo que ya se ve: la escritura es una operación de una sola fila con `Recordset`, sin transacción explícita; el identificador del alta se calcula con `Max+1` en cliente (`Responsable.cls:32`; `FU.bas:775`), patrón que en PostgreSQL quedaría mejor resuelto con una secuencia o un `INSERT … RETURNING`. El formulario tiene un único modo de alta y uno de edición, y la elección entre ambos depende de una variable global de sesión (`VG.bas:204`) y no de un parámetro explícito.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-HPS-S-69.** `Form_Open` reactiva `AllowEdits` incondicionalmente al terminar (`:203`), y también en la rama de error (`:210`): la restricción de perfil técnico aplicada en `:190-193` queda anulada. El botón «Registrar» sigue deshabilitado para el técnico, de modo que puede teclear y perder el trabajo sin aviso.
- **H-HPS-S-70.** Rótulos heredados de otra aplicación: el `Caption` de diseño es «Edición de Suministrador» (`.form.txt:40`), el botón de ayuda muestra «Salir» (`.form.txt:245`) y el de registro dice «Editar» frente a un texto emergente «Registrar» (`.form.txt:630-631`). El `Caption` de la ventana lo sustituye `Form_Open` (`:189`), pero el rótulo interno «EDICIÓN RESPONSABLE» (`.form.txt:226`) permanece fijo y contradice el rótulo que el código asigna en alta («ALTA DE RESPONSABLE», `:139`).
- **H-HPS-S-71.** `HaHabidoCambios` compara a mano los tres campos (`:110-121`): un campo nuevo en `Responsable` no se detectaría como cambio, y la validación de correo y de unicidad sólo se produce al registrar, no antes, de modo que el usuario escribe un correo inválido y lo descubre al pulsar el botón.
- **H-HPS-S-72.** El identificador del responsable en alta se reserva con `Max(id)+1` (`Responsable.cls:32`; `FU.bas:775-788`) y el `INSERT` se hace después (`ResponsableOperaciones.cls:55-73`): dos altas simultáneas pueden obtener el mismo identificador.
- `EstablecerDatos` sólo copia los campos que vienen informados en el registro (`:156-164`): un campo vacío en la base deja el valor de diseño en el control, de modo que el alta «hereda» un valor que no es del responsable.
- La comprobación de duplicidad de correo depende de que el correo haya cambiado (`ResponsableOperaciones.cls:36-42`): una edición que sólo toca el nombre no vuelve a comprobar el correo, lo cual es correcto, pero el criterio está repartido entre dos archivos y no es evidente a primera vista.
- `pregunta`, la variable que recoge el resultado de los `MsgBox`, es una global compartida (`VG.bas:182`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: formulario de alta y edición como recurso único con identificador de responsable en la URL; validación de correo y de unicidad en el propio formulario antes de enabling el botón de registro; y el perfil técnico como permiso explícito por acción en lugar de una comprobación global de identidad.

## Criterios de aceptación

Propuestos, no aprobados:

1. Sin solicitud ni responsable activo, el formulario se abre en modo alta y así lo indica el rótulo.
2. En modo edición, los tres campos muestran los valores registrados.
3. Sin cambios, no se escribe nada y se informa del motivo.
4. Con nombre vacío o correo con formato inválido, no se escribe y se informa del campo concreto.
5. Con un correo ya registrado en otro responsable, no se escribe y se informa.
6. El perfil técnico no puede registrar; si puede escribir, la interfaz avisa de que no se va a guardar.
7. Tras registrar, el formulario padre se repinta y la lista de responsables refleja el alta o la edición.

## Pendientes

- Si `EsTecnico` es la única comprobación de autorización o si el entorno de permisos por entidad (`UsuarioAplicacionPermisos`, `UsuarioEntidad`) debería gobernar también esta pantalla: en este `.cls` no hay ninguna referencia a permisos.
- Qué ocurre con el responsable cuando la solicitud que lo usa se elimina: `ResponsableOperaciones.ValidacionEliminar` bloquea el borrado si el correo está en uso (`ResponsableOperaciones.cls:98-135`), pero no consta quién avisa al usuario de esa dependencia.
- Recuento de controles: el inventario fija 28 y el `.form.txt` contiene 25 bloques `Begin`; la diferencia no se reproduce con una regla simple de recuento.
