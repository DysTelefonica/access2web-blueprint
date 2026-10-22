# Form_FormCambioTipo

> **Commit fijado:** `cfa1937ab05a9e40150ea9236243898de96f84da` · **Fuente:** `src/forms/Form_FormCambioTipo.cls` (107 líneas) y `.form.txt` (26 controles `Begin`) · **PRD:** [README](../README.md)

Diálogo modal de un solo campo: muestra el tipo vigente de la solicitud activa, deja elegir entre tres tipos y delega el cambio a `SolicitudOperaciones.CambiarTipo`. Es un formulario **no enlazado** (sin `RecordSource`): toma la solicitud de la global `m_ObjSolicitudActiva` y escribe a través de la capa de operaciones, no sobre controles. Las citas son `Form_FormCambioTipo.cls:línea` salvo indicación; `FU.bas` = `modules/FUNCIONES UTILES.bas`, `VG.bas` = `modules/Variables Globales.bas`, `classes/` = `src/classes/`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `FormSolicitudesAltaTraspasoDatos` (subformulario) | `DoCmd.OpenForm "FormCambioTipo"` desde `ComandoCambioTipo_Click`, sin argumentos; el padre lo declara `WithEvents m_FormCambioTipo` | `Form_FormSolicitudesAltaTraspasoDatos.cls:231-243`, este cls `:4` |
| `FormSolicitudRenovacionDatos` | Mismo patrón; además cierra su propio padre nada más abrir este diálogo | `Form_FormSolicitudRenovacionDatos.cls:914-920` |

| Hacia | Cómo | Evidencia |
|---|---|---|
| Padre suscriptor | `RaiseEvent TipoCambiado` tras el guardado; el padre cierra su formulario contenedor | `:61`; `Form_FormSolicitudesAltaTraspasoDatos.cls:834-840`, `Form_FormSolicitudRenovacionDatos.cls:936-941` |
| Menú e indicador de tareas | `PintarIndicadores` recalcula totales del entorno y repinta otros formularios abiertos | `:64`; `FU.bas:41-72` |
| Administrador (solo ante error inesperado) | `CorreoAlAdministrador m_Error` | `:33`, `:74`, `:101` |

No recibe parámetros: la única entrada es la global `m_ObjSolicitudActiva` (`VG.bas:197`), que el padre fija antes de abrir el diálogo. `Form_Open` solo la lee (`:88-90`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| Formulario | — | `OnOpen` → `Form_Open` | `Ajustar Me` y refleja el tipo actual en `TipoActual` | `:80-107`, `.form.txt:38` |
| `TipoActual` | texto deshabilitado | — | Tipo vigente, tal cual lo trae la solicitud | `:90`, `.form.txt:372-373` |
| `TipoFuturo` | combo de lista de valores fija (tres literales), `Tag = "DATO"` | — | Tipo a aplicar; sin `AfterUpdate`, el valor no se filtra en ningún momento | `.form.txt:434-439` |
| `ComandoRegistrar` | botón, rótulo «Editar», descripción «Registrar» | `OnClick` → `ComandoRegistrar_Click` | Valida el combo, llama a `CambiarTipo`, emite `TipoCambiado`, repinta y cierra | `:40-78`, `.form.txt:477-478` |
| `ComandoAyuda` | botón, rótulo «Salir», descripción «Ayuda» | `OnClick` → `ComandoAyuda_Click` | Abre el PDF de ayuda del formulario activo; **no cierra** | `:12-37`, `.form.txt:245-249` |
| `cmdSalir` | botón, rótulo «Salir» | `OnClick` → `cmdSalir_Click` | Cierra sin guardar (`acSaveNo`), con `On Error Resume Next` | `:7-10`, `.form.txt:~900` |

## Flujos

### Normal: cambiar el tipo (`ComandoRegistrar_Click`, `:40-78`)

1. Enciende el reloj de arena, limpia `m_Error` y exige `TipoFuturo` no vacío; si falta, fija el aviso «Seleccione un tipo» y salta a la rama de error (`:45-53`).
2. Instancia `SolicitudOperaciones`, le asigna `m_ObjSolicitudActiva` y llama a `CambiarTipo Me.TipoFuturo, m_Error` (`:54-59`; `SolicitudOperaciones.cls:1267-1308`). Cualquier error devuelto por omisión aborta antes de emitir el evento.
3. Emite `TipoCambiado` (`:61`), lo que hace que el padre cierre su contenedor.
4. Llama a `PintarIndicadores` (`:64`), apaga el reloj y se cierra a sí mismo con `acSaveNo` (`:68`).

### Alternativos

- **Abandonar o consultar la ayuda:** `cmdSalir_Click` cierra sin tocar la base de datos (`:7-10`) y `ComandoAyuda_Click` compone la ruta con el nombre del formulario activo y abre el PDF, avisa si el archivo no existe y deja el formulario abierto (`:12-37`; `FU.bas:475-503`).
- **Cambio rechazado por la capa de operaciones:** tipo vacío, tipo idéntico al actual o tipo actual «Traspaso» (`:1275-1288`). El aviso distingue los tres casos y no se emite `TipoCambiado`.

### Errores

Los cuatro manejadores siguen el mismo patrón: solo los errores cuyo número **no** es 1000 se consideran internos; en ese caso se notifica al administrador y el `MsgBox` muestra el mensaje original (`:30-37`, `:71-78`, `:99-106`). `pregunta` es una global `Long` compartida con el resto de la aplicación (`VG.bas:182`).

## Reglas, validaciones y cálculos

- **Lista cerrada de tipos:** el combo se alimenta de tres literales fijos en diseño (`.form.txt:436`), no de `TbSolicitudes` ni de `EnumTipo`. No hay validación de que el tipo exista en el dominio: depende de que la lista de diseño no se desincronice de `Entorno.cls:1835-1836`.
- **Precondiciones en la capa de operaciones** (`SolicitudOperaciones.cls:1275-1288`): tipo no vacío, tipo distinto del actual y tipo calculado distinto de «Traspaso».
- **Escritura:** una única sentencia `UPDATE` sobre `TbSolicitudes` que fija `TIPO` por identificador (`SolicitudOperaciones.cls:1291-1294`), seguida de `RegistrarActualizacion` (`FU.bas:83`) y del asiento de registro con la acción «Cambio del tipo … al …» (`FU.bas:895`).
- **Sin control de permisos ni transacción:** el formulario no consulta el rol ni el modo de oficina, y la escritura se confirma antes de registrar la auditoría; un fallo posterior deja la solicitud cambiada y la operación comunicada como fallida.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Tipo vigente | propiedad en memoria de `m_ObjSolicitudActiva` | sin acceso a base de datos | `:88-90` |
| Cambio de tipo | `TbSolicitudes` | `UPDATE` con valores concatenados | `SolicitudOperaciones.cls:1291-1294` |
| Auditoría del cambio | registro de la aplicación vía `RegistrarLogGeneral` | escritura indirecta | `SolicitudOperaciones.cls:1300`; `FU.bas:895` |
| Recuento de pendientes | totales del entorno (`CalcularTotalesPendientes`) | lectura en caché | `:64`; `FU.bas:52-53` |
| Ayuda | fichero PDF en el directorio de documentación | sistema de ficheros | `:20`; `FU.bas:492-499` |

## Coexistencia

Pendiente de `../coexistencia.md`. Puntos previsibles: el `UPDATE` concatena el valor del combo sin parámetros (`SolicitudOperaciones.cls:1291-1294`), patrón que en PostgreSQL exigiría consulta parametrizada; los tres literales de tipo llevan tilde y son los mismos que usa `EnumTipo` a través de `Entorno` (`Entorno.cls:1835-1836`), de modo que la columna `TIPO` es una cadena mágica y no una clave foránea; `PintarIndicadores` (`FU.bas:41-72`) depende de que estén abiertos dos formularios concretos de Access. La lógica vive **duplicada** en el `.cls` y en el bloque `CodeBehindForm` del `.form.txt` (`.form.txt:667-777`), que es una segunda copia del mismo código y debe mantenerse sincronizada con él.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **Rótulo de ventana heredado de otra pantalla:** la `Caption` de diseño es «Edición de Suministrador» (`.form.txt:40`) y el formulario no la sobrescribe al abrir, a diferencia de los formularios hermanos; el título de ventana no corresponde a su contenido.
- **Dos botones rotulados «Salir»:** `ComandoAyuda` muestra «Salir» pero su descripción es «Ayuda» y su handler abre la ayuda sin cerrar (`.form.txt:245-249`, `:12-33`). El segundo, `cmdSalir`, sí cierra.
- **Orden de emisión y cierre:** el evento `TipoCambiado` se emite antes de repintar y antes del cierre propio (`:61-68`) y el suscriptor cierra el formulario padre, de modo que la ejecución continúa sobre un formulario cuyo contenedor ya no existe. En la ruta de renovación el padre se cierra antes de abrir este diálogo (`Form_FormSolicitudRenovacionDatos.cls:920`): ambos caminos resuelven el ciclo de vida de forma distinta.
- **Objeto en memoria modificado antes de persistir y auditoría fuera de la transacción:** el tipo se asigna al objeto antes de ejecutar la sentencia (`SolicitudOperaciones.cls:1289-1294`) y el asiento de registro se escribe después (`:1300`); si falla cualquiera de los dos pasos, la base de datos, el objeto y el aviso al usuario quedan divergentes.
- **`Form_Open` sin protección en su manejador:** los handlers de botón intercalan `On Error Resume Next` antes de `CorreoAlAdministrador` (`:32-33`, `:73-74`), pero el de apertura no (`:101`): un fallo dentro del propio aviso puede propagarse fuera del manejador.
- **Concatenación de SQL y ausencia de control de acceso:** el tipo se inyecta sin validar contra el dominio, lo que hoy no es explotable porque el combo es una lista cerrada y sí lo sería si se volviera editable; y cambiar «Alta» a «Renovación» o «Elevación Grado» por aquí evita la comprobación de oficina que la vía de renovación sí impone (`Form_FormSolicitudAlta.cls:63-67`), única barrera entre un alta y una renovación.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: tipos cargados del dominio en lugar de una lista literal en diseño; el tipo como parámetro de ruta con el diálogo superpuesto en lugar de un formulario hermano abierto con `OpenForm`; comprobación explícita de rol y de modo de oficina en el propio caso de uso; transacción única que abarque cambio y asiento de auditoría; confirmación previa cuando el destino cambia la vía del trámite.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir el diálogo muestra el tipo vigente en `TipoActual`.
2. Sin elegir tipo, «Editar» avisa y no escribe nada.
3. Elegir el mismo tipo avisa que ya lo tiene y no escribe.
4. Elegir un tipo distinto persiste el cambio, queda anotado en el registro y el padre se cierra.
5. Una solicitud de traspaso no permite cambiar de tipo.
6. «Salir» cierra sin escribir.

## Pendientes

- Si el suscriptor `TipoCambiado` del padre de renovación se conecta a la instancia correcta o a la instancia por defecto, y qué hace `RegistrarActualizacion` (`FU.bas:83-141`) con la fila de último cambio que escribe aquí.
- Si el dominio de tipos de `Entorno` se mantiene sincronizado con la lista literal del combo (`.form.txt:436`; `Entorno.cls:1835-1836`) y cuál de las dos copias del código —`.cls` o bloque `CodeBehindForm` del `.form.txt` (`.form.txt:667`)— es la que se exporta.