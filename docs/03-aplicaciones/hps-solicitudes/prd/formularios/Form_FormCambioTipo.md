# Form_FormCambioTipo

> **Commit fijado:** `cfa1937ab05a9e40150ea9236243898de96f84da` · **Fuente:** `src/forms/Form_FormCambioTipo.cls` (107 líneas) y `.form.txt` (23 bloques `Begin`) · **PRD:** [README](../README.md)
>
> **Abreviaturas:** `FU.bas` = `modules/FUNCIONES UTILES.bas`, `VG.bas` = `modules/Variables Globales.bas`; las clases de `src/classes/` se citan por el nombre del fichero. Las citas son `Form_FormCambioTipo.cls:línea` salvo indicación.

Cambia el tipo de una solicitud ya registrada («Alta», «Renovación», «Elevación Grado»). Es un formulario **no enlazado** (sin `RecordSource`); muestra el tipo actual en un cuadro deshabilitado, ofrece el futuro en un combo de lista de valores y delega la escritura en `SolicitudOperaciones.CambiarTipo`. Avisa al formulario padre con el evento `TipoCambiado` (`:4`, lanzado en `:61`).

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormSolicitudRenovacionDatos` | `ComandoCambioTipo_Click` cierra la instancia previa, abre este formulario y se suscribe con `WithEvents`; acto seguido cierra el formulario padre | `Form_FormSolicitudRenovacionDatos.cls:907-922` |
| `Form_FormSolicitudesAltaTraspasoDatos` | `ComandoCambioTipo_Click`, igual, con `WithEvents` a `m_FormCambioTipo` | `Form_FormSolicitudesAltaTraspasoDatos.cls:231-245` |

No declara parámetros: lee el tipo de `m_ObjSolicitudActiva` (global de sesión, `VG.bas:197`), que los dos llamantes dejan preparado. No abre formularios. Al terminar el registro se cierra a sí mismo (`:68`). El tipo de trámite no se valida en este formulario: la comprobación de que el tipo actual no sea «Traspaso» la hace la capa de datos (`SolicitudOperaciones.cls:1287-1290`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| Formulario | — | `OnOpen` → `Form_Open` | Dimensiona el formulario y pinta el tipo actual | `:80-104`, `.form.txt:40-41` |
| `lblTitulo` | etiqueta de encabezado | — | Rótulo «CAMBIO DE TIPO» | `.form.txt:226-227` |
| `ComandoAyuda` | botón | `OnClick` → `ComandoAyuda_Click` | Abre el PDF de ayuda | `:12-37`, `.form.txt:245-247` |
| `Etiqueta109` | etiqueta | — | Rótulo «Tipo Actual» | `.form.txt:346-347` |
| `TipoActual` | cuadro de texto deshabilitado | — | Muestra `m_ObjSolicitudActiva.TIPO` | `:89-90`, `.form.txt:358, 372` |
| `Etiqueta250` | etiqueta | — | Rótulo «Tipo Futuro» | `.form.txt:402-403` |
| `TipoFuturo` | combo | — (sin handler) | Tipo a aplicar; lista de valores `Alta;Renovación;Elevación Grado` | `:50`, `.form.txt:434-436` |
| `ComandoRegistrar` | botón | `OnClick` → `ComandoRegistrar_Click` | Valida el tipo, llama a `CambiarTipo` y avisa al padre | `:40-78`, `.form.txt:477-479` |
| `cmdSalir` | botón «Salir» | `OnClick` → `cmdSalir_Click` | Cierra el formulario | `:7-10`, `.form.txt:568-570` |

## Flujos

### Apertura (`Form_Open`, `:80-104`)

1. `Ajustar Me` dimensiona el formulario (`:88`; `FU.bas:735-759`).
2. Si hay solicitud activa, copia su `TIPO` al cuadro `TipoActual` (`:89-90`); si no la hay, el cuadro conserva el valor de diseño.
3. El `Caption` de la ventana no se toca: se queda el de diseño (`.form.txt:40`).

### Registro (`ComandoRegistrar_Click`, `:40-78`)

1. Sin tipo futuro elegido, aborta con «Seleccione un tipo» (`:50-53`).
2. Crea `SolicitudOperaciones`, le asigna `m_ObjSolicitudActiva` y llama a `CambiarTipo` con el valor del combo (`:54-57`).
3. Si la capa de datos devuelve error, se propaga con el centinela 1000 y no se avisa al padre (`:58-60`).
4. Con éxito, lanza `TipoCambiado` (`:61`), llama a `PintarIndicadores` (`:64`; `FU.bas:41-83`) y cierra la ventana (`:68`).

### Errores

- `cmdSalir_Click` usa `On Error Resume Next` sin aviso (`:7-10`).
- Las dos ramas `errores:` de `ComandoAyuda_Click` y `ComandoRegistrar_Click` ignoran el centinela `Err.Number = 1000`, que en este formulario significa «error propio ya descrito en `m_Error`»; los errores ajenos rellenan el texto, envían `CorreoAlAdministrador` (`FU.bas:585`) y muestran un `MsgBox` crítico (`:29-36`, `:70-77`).
- `Form_Open` tiene el mismo patrón y, además, cierra el formulario tras el aviso (`:98-106`).

## Reglas, validaciones y cálculos

- **Tipo futuro obligatorio:** no se escribe sin valor en el combo (`:55-58`).
- **Tipo actual no traspaso:** regla de la capa de datos; un tipo «Traspaso» no admite cambio (`SolicitudOperaciones.cls:1287-1290`).
- **Tipo distinto del actual:** la capa de datos rechaza el cambio si el tipo nuevo es el ya registrado (`SolicitudOperaciones.cls:1282-1285`).
- **Lista cerrada:** el combo es de lista de valores con tres entradas fijas (`.form.txt:436`); no se admiten otros tipos aunque la base tenga más.
- **Escritura:** `CambiarTipo` actualiza `TbSolicitudes.TIPO` con una sentencia construida por concatenación (`SolicitudOperaciones.cls:1292-1295`), registra la actualización (`FU.bas:83`) y deja traza en el log general (`FU.bas:895`).
- **Sin datos propios:** este `.cls` no lee ni escribe tablas; todo el acceso a datos está en `SolicitudOperaciones` y en `constructor`.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Tipo actual | propiedad `TIPO` de la solicitud de sesión | Derivado, sin DAO en este `.cls` | `:88-90` |
| Cambio de tipo | `TbSolicitudes` | `UPDATE` con SQL concatenado | `SolicitudOperaciones.cls:1292-1295` |
| Marca de modificación y estado | `TbSolicitudes` | DAO (`RegistrarActualizacion`) | `FU.bas:83-124` |
| Registro de cambio de tipo | log general | DAO (`RegistrarLogGeneral`) | `FU.bas:895-937` |
| Contadores de pendientes | caché del entorno | Derivado | `FU.bas:41-83` |

## Coexistencia

Pendiente de `../coexistencia.md`. Lo que ya se ve: el cambio de tipo es una operación de una sola escritura sobre `TbSolicitudes.TIPO`, con la sentencia armada por concatenación (`SolicitudOperaciones.cls:1292-1295`); en PostgreSQL pasaría a consulta parametrizada. El formulario no tiene estado propio: todo depende de la solicitud de sesión, que es lo que un servicio de dominio puede recibir por identificador.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-HPS-S-66.** Rótulos heredados de otra aplicación: la ventana arranca como «Edición de Suministrador» (`.form.txt:40`) y `Form_Open` no la sobrescribe, el botón de ayuda muestra «Salir» (`.form.txt:246`) y el de registro dice «Editar» frente a un texto emergente «Registrar» (`.form.txt:478-479`).
- **H-HPS-S-67.** `Form_Open` no comprueba que exista solicitud activa (`:89`): si `m_ObjSolicitudActiva` es `Nothing`, `TipoActual` queda vacío y `ComandoRegistrar_Click` asigna esa referencia nula a `SolicitudOperaciones.solicitud` (`:56`); `CambiarTipo` opera con `With Me.solicitud` sobre un objeto inexistente (`SolicitudOperaciones.cls:1280-1281`) y el fallo sale como error no controlado, con correo al administrador. No hay mensaje de «no hay solicitud en curso» para el usuario.
- **H-HPS-S-68.** `CambiarTipo` interpola el tipo nuevo en la sentencia SQL (`SolicitudOperaciones.cls:1292-1295`); el valor viene de una lista cerrada en este formulario, pero el método es público y cualquier otro llamante podría pasar texto libre.
- El evento `TipoCambiado` se lanza dentro del bloque `With m_SolOp` (`:61`), antes de `PintarIndicadores` (`:64`) y del cierre (`:68`): el padre se repinta con el estado anterior al cierre de la ventana.
- Los dos llamantes cierran el formulario padre nada más abrir este (`Form_FormSolicitudRenovacionDatos.cls:922`, `Form_FormSolicitudesAltaTraspasoDatos.cls` no lo hace): el contenedor de la renovación se cierra con el cambio de tipo a medias, y el manejador `m_FormCambioTipo_TipoCambiado` vuelve a cerrarlo (`Form_FormSolicitudRenovacionDatos.cls:943`).
- `pregunta`, la variable que recoge el resultado de los `MsgBox`, es una global compartida (`VG.bas:182`): dos diálogos concurrentes en la misma sesión se pisan el resultado.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: selector de tipo en la propia ficha de la solicitud, sin ventana aparte; el tipo como un atributo del expediente con historial de cambios en lugar de una sobrescritura; y la lista de tipos como dato configurable en lugar de una constante en el formulario.

## Criterios de aceptación

Propuestos, no aprobados:

1. Sin solicitud activa, el formulario no se abre y se informa del motivo.
2. Sin tipo futuro elegido, no se escribe nada y se informa.
3. Con un tipo actual de «Traspaso», el cambio no se permite y se explica por qué.
4. Elegir el mismo tipo que ya tiene la solicitud no escribe y avisa.
5. Con un tipo válido, el tipo de la solicitud queda actualizado, se deja traza del cambio y el formulario padre se repinta.
6. Un error al escribir no deja el formulario padre con un estado que no corresponda.

## Pendientes

- Comportamiento de `TipoFuturo` cuando el combo se abre con un valor ya seleccionado por el usuario y se confirma sin cambios: la capa de datos lo rechaza, pero no está escrito en ninguna parte de la interfaz.
- Si el tipo «Elevación Grado» admite subtypes o condiciones que la lista de valores no expresa.
- Recuento de controles: el inventario fija 26 y el `.form.txt` contiene 23 bloques `Begin`; la diferencia no se reproduce con una regla simple de recuento.
