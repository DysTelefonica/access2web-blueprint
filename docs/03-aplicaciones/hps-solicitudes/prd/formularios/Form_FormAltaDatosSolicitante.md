# Form_FormAltaDatosSolicitante

> **Commit fijado:** `cfa1937ab05a9e40150ea9236243898de96f84da` · **Fuente:** `src/forms/Form_FormAltaDatosSolicitante.cls` (452 líneas) y `.form.txt` (2 915 líneas) · **PRD:** [README](../README.md)
>
> **Abreviaturas:** `FU.bas` = `modules/FUNCIONES UTILES.bas`, `VG.bas` = `modules/Variables Globales.bas`; las clases de `src/classes/` se citan por el nombre del fichero. Las citas son `Form_FormAltaDatosSolicitante.cls:línea` del commit fijado.

Formulario **no enlazado** (sin `RecordSource`) de captura/edición de los datos del solicitante: tipo, nombre, apellidos, DNI/NIE, fecha y lugar de nacimiento, contacto, correos, motivo HPS, clasificación de la empresa del usuario y de la tramitadora, y expediente asociado. La persistencia se delega en `SolicitudOperaciones.RegistrarDatos` (`:345-388`). Tras registrar, abre `FormSolicitudAltaTraspaso` y se cierra (`:382-384`); ninguno de los dos pasos escribe ni lee directamente tablas desde el `.cls` del formulario.

> **Importante (hallazgo previo a esta ficha).** Una búsqueda global en `src/` no encuentra ningún `DoCmd.OpenForm "FormAltaDatosSolicitante"` ni un `SourceObject = "FormAltaDatosSolicitante"` en `.form.txt`. En `cfa1937a` el formulario está compilado pero **no está cableado** desde ningún otro formulario del proyecto; queda como módulo residente sin camino de apertura conocido.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| (nadie) | `DoCmd.OpenForm "FormAltaDatosSolicitante"` no aparece en `src/`; ver hallazgo previo. | — |
| Hacia `Form_FormSolicitudAltaTraspaso` | Tras `ComandoRegistrar`, cierra cualquier instancia previa con `acSaveNo` y abre el destino (`:372-377`) | `:372-377` |

Eventos hacia el exterior: `Alta` (`:4`, lanzado en `:366`). Estado compartido: la global `m_ObjSolicitudActiva` (`VG.bas:197`), que aquí se rellena con `RellenarObjeto` (`:307-343`) y se rehidrata con `constructor.getSolicitud` después de registrar (`:367-371`); también `m_ObjUsuarioConectado` (`VG.bas:189`) para `EstablecerDatos` (`:290`) y `m_ObjEntorno` (`VG.bas:190`) para combos (`:406, 422`) y para invalidar `Expedientes` en `ComandoActualizar` (`:131-133, 169-171`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `ComandoAyuda` | botón (caption `Comando46`, ImageData) | `OnClick` | `ComandoAyuda_Click` deshabilita edición, abre el PDF de ayuda del formulario activo y la rehabilita | `:190-216`, `:603` |
| `lblTitulo1` | etiqueta «ALTA DE DATOS DEL SOLICITANTE» | — | Encabezado | `:666-668` |
| `TIPO` | combo, `RowSource = "Alta;Traspaso"` | — | Tipo de la solicitud | `:1104-1106` |
| `IDEmpresaUsuario` | combo, `RowSource` con cinco valores de clasificación (lista enum cerrada) | — | Clasificación de la empresa del usuario | `:835-839` |
| `IDEmpresaTramitadora` | combo, mismo `RowSource` de clasificación | — | Clasificación de la tramitadora | `:1479-1483` |
| `CorreosAutomaticos` | combo, `RowSource = "Sí;No"` | — | Activa o desactiva el envío automático de correos | `:909-913` |
| `Gestor` | combo (sin `RowSource` fijo) | — | Rellenado en `EstablecerCombos` con `RellenarComboTramitadores` | `:315, 393-398`, `FU.bas:370-399` |
| `Motivo_HPS` | combo | — | Rellenado en `EstablecerCombos` con `RellenarComboMotivos` | `:427-430`, `FU.bas:402-428` |
| `ComandoMotivos` | botón «Editar» | `OnClick` | Pide un motivo por `InputBox`, llama a `MeterUnMotivoEnListaMOtivos` y reasigna `Me.Motivo_HPS` (`:222-244`) | `:222-244`, `:1741-1743` |
| `ComandoGenerarDNIValido` | botón (caption `Comando247`, ImageData) | `OnClick` | Asigna a `Me.DNI` un DNI/NIE sintético; solo visible cuando `TempVars("EnPruebas") = "Sí"` | `:218-220`, `:292-296`, `:1869-1872` |
| `DNI` / `Nombre` / `Apellido1` / `Apellido2` / `FNacimiento` / `LugarNacimiento` / `email` / `emailResponsable` / `Telefono` / `Observaciones` | textos libres | — | Volcados a `Solicitud` en `RellenarObjeto` | `:307-343` |
| `PalabraClave` / `IDExpediente` / `Titulo` | textos del bloque expediente | `KeyDown` (Enter) en `PalabraClave` | Enter en la palabra clave dispara `ComandoBuscarExpediente_Click` | `:172-189`, `:1962, 1993, 2232` |
| `ComandoBuscarExpediente` | botón (caption `Comando36`) | `OnClick` | Busca expedientes por palabra clave; con 0 resultados abre `FormExpedientesBusqueda`; con 1 resultado lo carga en `IDExpediente` | `:54-130`, `:2047` |
| `ComandoIrABusquedaExpedientes` | botón (caption `Comando36`) | `OnClick` | Abre `FormExpedientesBusqueda` con la palabra clave | `:17-52`, `:2125` |
| `ComandoActualizar` | botón (caption `Contador`) | `OnClick` | Anula `m_ObjEntorno.Expedientes` para forzar recarga | `:131-133, 169-171`, `:2183-2185` |
| `cmdSalir` | botón «Salir» | `OnClick` | Cierra el formulario sin parámetro | `:10-16`, `:2324-2326` |
| `ComandoRegistrar` | botón (caption `Editar`) | `OnClick` | Volca el formulario en `m_ObjSolicitudActiva` y llama a `SolicitudOperaciones.RegistrarDatos`; abre `FormSolicitudAltaTraspaso` y se cierra | `:345-388`, `:2401-2403` |

## Flujos

### Apertura (`Form_Load`, `:250-278`)

1. `Ajustar Me` (`:259`; `FU.bas:735-759`).
2. `EstablecerDatos m_Error` (`:260`): rellena combos por `EstablecerCombos` (`EstablecerCombos`, `:390-451`), asigna `Me.Gestor = m_ObjUsuarioConectado.Nombre` (`:290`), `Me.CorreosAutomaticos = m_ObjEntorno.CorreosAutomaticos` (`:291`) y muestra u oculta `ComandoGenerarDNIValido` por la `TempVar("EnPruebas")` (`:292-296`).
3. Reactiva `Me.AllowEdits` al final (`:268`).

### Búsqueda y selección de expediente (`ComandoBuscarExpediente_Click`, `:54-130`)

1. Ignora si la palabra clave viene vacía (`:68-74`).
2. `constructor.getExpedientesBusqueda` (`:76`; `constructor.bas:3855-3920`) devuelve un diccionario por `IDExpediente`.
3. Un único resultado se vuelca en `Me.IDExpediente` y `Me.Titulo` (`:95-106`); cero resultados abren la búsqueda global (`:80-91`); varios resultados la abren filtrada (`:107-119`).

### Alta (`ComandoRegistrar_Click`, `:345-388`)

1. `RellenarObjeto m_Error` (`:354`; cuerpo `:307-343`) crea `m_ObjSolicitudActiva` y mapea cada control uno a uno.
2. Crea un `SolicitudOperaciones`, le inyecta `m_ObjSolicitudActiva` y llama a `.RegistrarDatos , m_Error` (`:357-366`; `SolicitudOperaciones.cls:42-280`).
3. Lanza `RaiseEvent Alta` (`:366`).
4. Recarga `m_ObjSolicitudActiva` por `constructor.getSolicitud(p_IDSolicitud:=….)` (`:367-371`).
5. Cierra cualquier `FormSolicitudAltaTraspaso` previo y abre uno nuevo (`:372-377`) y se cierra (`DoCmd.Close acForm, Me.Name`, `:377`).

### Errores

Los handlers acumulan el texto en `m_Error`, llaman a `CorreoAlAdministrador m_Error` (`FU.bas:585`) salvo el error interno 1000 y muestran `MsgBox` crítico (`:46-51, 117-129, 154-167, 209-216, 269-278, 382-388`). El `MsgBox` del cierre (`cmdSalir_Click`, `:10-16`) no tiene capturador: usa `On Error Resume Next` implícito (`End Sub` directo, sin `Exit Sub`).

## Reglas, validaciones y cálculos

- **`ComandoMotivos` con texto vacío:** si el operador cancela el `InputBox`, el handler aborta sin error (`:228-230`).
- **`ComandoGenerarDNIValido` solo en pruebas:** la visibilidad depende de la `TempVar("EnPruebas")` consultada en `EstablecerDatos` (`:292-296`).
- **`Set m_ObjEntorno.Expedientes = Nothing` antes de reabrir:** fuerza la recarga de la colección desde el backend (`:131-133, 169-171`).
- **Perfil técnico:** este formulario no comprueba `EsTecnico` (`VG.bas:193`); el lector decide por `m_ObjUsuarioConectado.Nombre` sin restricción local.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Búsqueda de expedientes | `TbExpedientes` | DAO Recordset | `constructor.bas:3855-3920` |
| Inserción/edición de solicitud | `TbSolicitudes` | DAO Recordset (AddNew/Edit) | `SolicitudOperaciones.cls:42-280` |

El formulario no accede a tablas por su cuenta; la persistencia pasa por `SolicitudOperaciones.RegistrarDatos`.

## Coexistencia

Pendiente de `../coexistencia.md`. Lo que ya se ve: el `RellenarObjeto` mapea cada control uno a uno a `Solicitud` (`:307-343`) sin capa de validación; la validación de negocio vive en `SolicitudOperaciones.MotivoNoOK` (`SolicitudOperaciones.cls:42-280`), llamado antes de persistir. La clasificación de empresa (`IDEmpresaUsuario`, `IDEmpresaTramitadora`) se almacena como código numérico pequeño en `TbSolicitudes`; un esquema web puede mantener el catálogo como tabla de referencia (1 a 5).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-HPS-S-23.** Formulario **no cableado**: `DoCmd.OpenForm "FormAltaDatosSolicitante"` y `SourceObject = "FormAltaDatosSolicitante"` no aparecen en `src/`; en la versión analizada `cfa1937a` el formulario queda compilado pero sin camino de apertura. Hay que verificar con el área responsable si esto es residuo de un refactor previo o un módulo de respaldo; mientras tanto, la ficha describe lo que hace si llegase a abrirse desde otra rama o desde un complemento.
- **H-HPS-S-24.** `ComandoRegistrar_Click` cierra `FormSolicitudAltaTraspaso` con `acSaveNo` si está abierto y luego abre uno nuevo (`:372-377`): el trabajo en curso en esa instancia se descarta sin aviso. Patrón idéntico a `H-HPS-S-16` para `Form_FormSolicitudAlta`.
- **H-HPS-S-25.** Tipos a verificar sin decisión: `ComandoRegistrar_Click` referencia `Err.Numbr` (`:383`) — propiedad inexistente, mensaje de error mostrará «0» en lugar del número real. `cmdSalir_Click` (`:10-16`) cierra sin `Err` handler; el resto de formularios del proyecto capturan y reportan.
- **H-HPS-S-26.** `ComandoGenerarDNIValido` (`:218-220`, caption `Comando247`) asigna un DNI/NIE sintético al control `DNI`. La generación no escribe en base de datos por sí sola, pero el flujo combinado con `ComandoRegistrar` puede persistir un documento aleatorio sin validación de letra. Hay que verificar la política de uso en pruebas y si el botón llega a estar visible en producción.
- **H-HPS-S-27.** Rótulos y estética heredados de otra aplicación: la `Caption` del formulario es una cadena fija de tipo «Expedientes, versión X Administrador» (`:40`), `ComandoAyuda` (`:603`), `ComandoGenerarDNIValido` (`:1870`) y los dos botones `Comando36` (`:2047, 2125`) conservan sus nombres de diseño; `ComandoActualizar` (caption «Contador») y `ComandoRegistrar` (caption «Editar`) tienen `ControlTipText` correctos pero rótulos desalineados. Mismo patrón que `H-HPS-S-17`, `H-HPS-S-19` y `H-HPS-S-21`.
- `ComandoRegistrar_Click` no comprueba que la `Solicitud` haya cambiado: si el usuario pulsa dos veces con los mismos datos, el flujo vuelve a registrar (`:345-388`); aquí no aplica la guardia `If .Cambia` que sí se ve en otros formularios del proyecto.
- `EstablecerCombos` añade opciones con `cmb.AddItem` (`:411-420, 425-434`); el `RowSource` se vacía primero (`:406, 422`), pero `Me.Motivo_HPS.Requery` se hace solo en `ComandoMotivos_Click` (`:235`), no tras altas externas.
- `ComandoActualizar` (caption «Contador») y `ComandoActualizarar` (con errata, ambos asignan lo mismo; `:131-133, 169-171`) dejan dos puntos de entrada que reciclan `m_ObjEntorno.Expedientes`.

> Aviso de seguridad sensible: la generación de DNI/NIE sintético, la falta de validación de letra antes de persistir y la inyección del valor en `TbSolicitudes.IDEmpresaUsuario`/`IDEmpresaTramitadora` requieren revisión fuera de este lote. Solo se reportan como defectos a verificar; sin contraseñas, hashes, correos, URLs ni datos personales en esta ficha.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: modelar la pantalla como un formulario multipaso con esquema de validación por paso (DNI → datos personales → contacto → motivo → expediente) en lugar de un único formulario plano; cargar `Motivo_HPS` desde un endpoint paginado en lugar del combo; sustituir el botón «Editar motivo» por un modal; validar `DNI`/`NIE` por servidor antes de persistir.

## Criterios de aceptación

Propuestos, no aprobados:

1. Si se abriese con un `m_ObjSolicitudActiva` vacío, el alta crea una nueva fila en `TbSolicitudes` con `IDSolicitud` calculado y los demás campos vacíos o nulos.
2. Tras registrar, se cierra la instancia actual de `FormSolicitudAltaTraspaso` (si la hubiera) y se abre una nueva.
3. Las cuatro clasificaciones disponibles en `IDEmpresaUsuario` y `IDEmpresaTramitadora` siguen siendo las del `RowSource` del diseño.
4. Con `TempVars("EnPruebas") = "Sí"` aparece `ComandoGenerarDNIValido`; con cualquier otro valor, no.
5. Un expediente seleccionado desde `FormExpedientesBusqueda` actualiza `IDExpediente` y `Titulo` y deshabilita la sección hasta el siguiente «Actualizar».
6. El alta dispara `RaiseEvent Alta` exactamente una vez por pulsación.

## Pendientes

- Confirmar si la aplicación llega a abrir este formulario en algún flujo (módulo de administración, importación, atajo).
- Política interna sobre el botón `ComandoGenerarDNIValido` en producción.
- Si `ComandoActualizar` debería distinguir entre «recargar expedientes» y la errata `ComandoActualizarar` que duplica la acción.
- Cómo se rellenará `Me.Gestor` cuando el usuario no sea técnico (el `RellenarComboTramitadores` actúa con `EnumSiNo.no`, `:400`; `FU.bas:370`).
