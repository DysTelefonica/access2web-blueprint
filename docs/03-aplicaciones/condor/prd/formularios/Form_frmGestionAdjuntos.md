# Form_frmGestionAdjuntos

> **Commit fijado:** `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · **Fuente:** `src/forms/Form_frmGestionAdjuntos.cls` (380 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Visor de los adjuntos de una solicitud: los lista, abre el seleccionado, añade ficheros y elimina los que no sean generados. Al añadir o borrar el «Documento Final Firmado» encadena cambios de decisión final y de estado. Es **no enlazado**: la lista es una `Value List` construida en código (`:39-44, 235-247`). Las citas son `Form_frmGestionAdjuntos.cls:línea` salvo indicación; `FU.bas` abrevia `modules/FUNCIONES UTILES.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Form_frmGestionSolicitud, «Gestionar adjuntos» | `OpenForm` en `acDialog` con el identificador de la solicitud | `Form_frmGestionSolicitud.cls:367-371` |
| Form_frmGestionSolicitud, acción `ADJUNTAR_PDF_FIRMADO` | Crea `TempVars("AutoUploadEtapa") = "Cierre"` y abre en `acDialog`; al volver, si existe adjunto «Cierre» o «Documento Final Firmado», ejecuta la aprobación final | `Form_frmGestionSolicitud.cls:746-768` |

Abre Form_frmElegirEtapaAdjunto (`:86`), el selector de ficheros de Office (`:125`) y el fichero con `AbrirEnLocal` (`:368`; `FU.bas:404-418`). Avisa a Form_frmGestionSolicitud con `NotificarCambioExterno` si está abierto (`:214-216, 341-343`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `OnLoad`, `OnTimer` | Carga la lista y, con `AutoUploadEtapa`, lanza la subida a los 300 ms | `:14-80`, `.form.txt:221-222` |
| `lstAdjuntos` | lista | `OnClick`, `OnDblClick` | Seis columnas (id y ruta ocultas; nombre, tipo, etapa, descripción); habilita botones; doble clic abre | `:39-44, 255-269, 371-373`, `.form.txt:647-656` |
| `cmdAnadirAdjunto` | botón | `OnClick` | Pide descripción y sube un fichero; deshabilitado para el rol técnico | `:46, 82-104`, `.form.txt:685-687` |
| `cmdAbrirAdjunto` | botón | `OnClick` | Abre la ruta de la columna 1 | `:365-369`, `.form.txt:734-736` |
| `cmdEliminarAdjunto` | botón | `OnClick` | Borra el adjunto y, si es el documento de cierre, revierte la aprobación | `:273-363`, `.form.txt:783-785` |
| `lblBloqueo` | etiqueta oculta | — | Aviso de proceso en curso | `:20, 141, 225`, `.form.txt:822-835` |
| `cmdCerrar` («Salir») | botón | `OnClick` | Cierra el formulario | `:375-377`, `.form.txt:866-868` |

## Flujos

### Apertura (`Form_Load`, `:14-63`)

1. Exige `OpenArgs` mayor que 0 (`:26-29`); titula con el código de la solicitud si la encuentra (`:31-37`).
2. Habilita «Añadir» salvo para el rol técnico (`:46`) y carga la lista (`:47`).
3. Si existe `AutoUploadEtapa`, arranca el temporizador (`:52-54`); `Form_Timer` lo para, borra la variable y llama a `EjecutarFlujoSubida(etapa, etapa)` **sin descripción** (`:65-80`).

### Subida (`EjecutarFlujoSubida`, `:108-233`)

1. Manual: descripción en Form_frmElegirEtapaAdjunto; etapa y nombre llegan como «Otro» (`:86-96`; `Form_frmElegirEtapaAdjunto.cls:6-12`).
2. Selector de un solo fichero; con etapa «Cierre» solo PDF, si no, todos, PDF o Word (`:125-137`).
3. Si la etapa es «Documento Final Firmado» y ya hay uno, pide confirmación y borra el anterior (`:145-170`).
4. `GuardarAdjuntoDesdeArchivo` copia el fichero a la carpeta de documentación (renombrándolo con fecha y hora si ya existe) e inserta en `tbAdjuntos`, en su propia transacción y borrando la copia si falla (`AdjuntosServicio.cls:32-38, 61-98, 114-123`). Sin descripción, guarda «Adjunto de la etapa …» (`AdjuntosServicio.cls:83-88`).
5. Solo con «Documento Final Firmado»: en una transacción nueva graba `decisionFinal = "APROBADO"` según el tipo (`PC`, `CD_CA`, `CD_CA_SUB`) y ejecuta la transición a aprobada (`:176-210`; `WorkflowServicio.cls:1211-1301`). Si no, muestra «Archivo adjuntado correctamente.» (`:218`).
6. Recarga la lista (`:221`).

### Eliminación (`cmdEliminarAdjunto_Click`, `:273-363`)

1. Lee id y etapa de la fila; detecta «documento de cierre» si la etapa es «Documento Final Firmado» y la solicitud está aprobada (`:287-297`).
2. Confirma con aviso reforzado o simple (`:300-313`).
3. `EliminarAdjunto` borra la fila y el fichero en su transacción (`:319`; `AdjuntosServicio.cls:484-520`).
4. Si era el documento de cierre: limpia la decisión final según tipo y llama a `RevertirAFaseAnterior` hacia formalización (`:322-336`).

- **Errores:** cada handler acaba en `ManejadorErroresFormulario` (`:61, 103, 232, 361`). Un error en `Form_Load` cierra el formulario (`:62`). `Form_Timer` silencia todo error (`:67`).

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| Técnico no añade ni elimina | `Form_Load`, `lstAdjuntos_Click` | `:46, 268` |
| No se eliminan adjuntos con tipo «Generado» | `lstAdjuntos_Click` | `:268` |
| «Documento Final Firmado» debe ser PDF | servicio | `AdjuntosServicio.cls:53-59` |
| Un solo «Documento Final Firmado» (sustituir con confirmación) | subida | `:145-169` |
| Revertir solo Calidad o Administrador | servicio | `WorkflowServicio.cls:2389-2391` |
| Ruta mostrada = carpeta configurada + nombre | `CargarLista` | `:243` |

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lista de adjuntos | `tbAdjuntos` | SQL parametrizado | no | `AdjuntoRepositorio.bas:114-128` |
| Alta de adjunto (id = máximo + 1) | `tbAdjuntos` | `INSERT` parametrizado | propia | `AdjuntoRepositorio.bas:12-27, 38-45` |
| Baja de adjunto y fichero | `tbAdjuntos` | `DELETE` parametrizado | propia | `AdjuntoRepositorio.bas:177-183`; `AdjuntosServicio.cls:513-516` |
| Decisión final de PC | `tbDatosPC` | DAO `Edit/Update` | la del formulario | `DatosPCRepositorio.bas:384-412` |
| Limpieza de decisión final de PC | `tbDatosPC` | SQL concatenado + `Edit/Update` | no | `DatosPCRepositorio.bas:515-533` |
| Transición y reversión de estado | solicitud y registro de estados | servicio | propia o compartida | `WorkflowServicio.cls:1282-1291, 2366-2554` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: identificadores por `MAX + 1` (`AdjuntoRepositorio.bas:45`), que no es seguro con dos escritores concurrentes; transacciones DAO abiertas con contraseña de Jet/ACE (`:181`); fichero en carpeta compartida coordinado con la fila sin transacción común. Debe preservarse el valor textual de `etapaWF` y `TipoAccion`, del que dependen la sustitución, la reversión y el permiso de borrado.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Las ramas de «Documento Final Firmado» (sustitución y aprobación automática, `:145, 176`) parecen inalcanzables desde este formulario: la subida manual llega con «Otro» y la automática con «Cierre» (`Form_frmGestionSolicitud.cls:762`).
- La subida automática pasa `etapa` como nombre de etapa sin descripción (`:78`), y la PDF solo se impone por filtro del selector, no por validación (`:130-131`; `AdjuntosServicio.cls:53`).
- En la aprobación automática, el borrado previo y el alta van en transacciones separadas de la decisión y la transición (`:168, 173, 182`); `db` no se cierra (`:224-228`); la notificación por correo de la transición se envía antes del `CommitTrans` del formulario (`WorkflowServicio.cls:1311-1375`; `:210`).
- La eliminación encadena tres operaciones no atómicas (`:319, 328-331, 336`); además `RevertirAFaseAnterior` desde aprobada ya borra esos adjuntos y la decisión final (`WorkflowServicio.cls:2469-2492`).
- El aviso promete borrar los datos de «Autoridad Competente», pero `LimpiarDecisionFinal` deja `obsAprobacionAutoridadDiseno` y `NombreAutoridadDiseno` (`:304`; `DatosPCRepositorio.bas:529-531`).
- Si `getSolicitudPorID` devuelve `Nothing`, `sol.idEstadoInterno` falla (`:294-297`).
- Una descripción con «;» desplaza las columnas de la `Value List` (`:244`); la ruta no añade separador si la carpeta configurada no termina en «\» (`:243`; compárese `AdjuntosServicio.cls:63`).
- Coexisten dos `.form.txt` con `RowSource` de diseño distinto (`.form.txt:649`; `frmGestionAdjuntos.form.txt:648`); el diseño incluye una ruta de red que el código sustituye en la carga.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: almacenamiento de ficheros con transacción compensada única; etapa elegida de una lista cerrada.

## Criterios de aceptación

Propuestos, no aprobados:

1. La lista muestra nombre, tipo, etapa y descripción de cada adjunto de la solicitud.
2. El rol técnico no puede añadir ni eliminar; nadie elimina un adjunto «Generado».
3. Una subida guarda fichero y fila, o ninguno de los dos.
4. Abierto desde «Adjuntar PDF firmado», el selector solo admite PDF y la subida empieza sin pulsar «Añadir».
5. Eliminar el documento de cierre de una solicitud aprobada la devuelve a formalización y limpia la decisión final, todo o nada.

## Pendientes

- Valores reales de `etapaWF` en datos: si existen filas «Documento Final Firmado» y quién las crea.
- Comportamiento de `AbrirEnLocal` tras la comprobación de acceso (`FU.bas:420` en adelante).
- Cuál de los dos `.form.txt` corresponde al binario en uso.
