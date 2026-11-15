# Form_FormSolicitudRenovacionDatos

> **Commit fijado:** `cfa1937ab05a9e40150ea9236243898de96f84da` · **Fuente:** `src/forms/Form_FormSolicitudRenovacionDatos.cls` (957 líneas) y `.form.txt` (63 bloques `Begin`) · **PRD:** [README](../README.md)
>
> **Abreviaturas:** `FU.bas` = `modules/FUNCIONES UTILES.bas`, `VG.bas` = `modules/Variables Globales.bas`, `constructor.bas` = `modules/constructor.bas`; las clases de `src/classes/` se citan por el nombre del fichero. Las citas son `Form_FormSolicitudRenovacionDatos.cls:línea` salvo indicación.

Formulario de datos de una solicitud de renovación o elevación de grado. Es el subformulario que `FormSolicitudRenovacion` monta en su pestaña «Datos» y concentra casi toda la lógica de la pantalla: carga la fotografía inicial, pinta los combos, decide si hay cambios, vuelca los controles a la entidad y la persiste. **Sí está enlazado por control** (`Me.TIPO`, `Me.Nombre`, `Me.DNI`, …), pero su fuente es la entidad en memoria, no un `RecordSource`: los `RowSource` son listas de valores en diseño o se rellenan por código.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormSolicitudRenovacion` | Lo monta como subformulario y le fija el contexto de sesión (`m_ObjSolicitudActiva`, `m_ObjUsuarioHPSActivo`) | `Form_FormSolicitudRenovacion.cls:82-83` |
| `Form_FormSolicitudesAltaTraspaso` | El mismo `.form.txt` sirve al contenedor de traspaso; el propio formulario distingue ambos por `Me.Parent.Name` | `:857-861` |

| Hacia | Cómo | Evidencia |
|---|---|---|
| `Form_FormResponsable` | `ComandoNuevoResponsable_Click` la abre y se enlaza con `WithEvents` | `:298-312` |
| `Form_FormCambioTipo` | `ComandoCambioTipo_Click` y el evento `TipoCambiado` | `:907, 936` |
| `Form_FormExpedientesBusqueda` | `ComandoIrABusquedaExpedientes_Click` la abre, opcionalmente prefiltrada, y escucha `Seleccionado` | `:16-51, 220` |
| Contenedor padre | Tras guardar, invoca `EstablecerDatos` del padre según su nombre | `:857-861` |
| `SolicitudOperaciones` | `RegistrarDatos`, `ReprogramarAvisosFuturos` y `AnularTodosLosProgramados` | `:700, 717, 851` |

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| Formulario | — | `OnLoad` → `Form_Load` | `EstablecerDatos` y reactiva la edición | `:327-355` |
| `TIPO`, `Gestor`, `Nombre`, `Apellido1`, `Apellido2`, `FNacimiento`, `LugarNacimiento`, `email`, `emailResponsable`, `Telefono`, `DNI`, `Observaciones`, `Motivo_HPS`, `IDEmpresaUsuario`, `IDEmpresaTramitadora`, `IDExpediente`, `IDUsuarioHPS`, `Titulo`, `CorreosAutomaticos` | controles enlazados por nombre a la entidad | — | Los 16 que `EstablecerDatos` pinta, `HaHabidoCambios` compara y `RellenarObjeto` vuelca | `:439-511, 534-603, 621-643`, `.form.txt:608-2640` |
| `ComandoRegistrar` | botón | `OnClick` → `ComandoRegistrar_Click` | Guarda y ejecuta la lógica de correos automáticos | `:656-746`, `.form.txt:1475` |
| `ComandoActualizar` | botón | `OnClick` → `ComandoActualizar_Click` | Repinta el rótulo de estado de la solicitud | `:278-294`, `.form.txt:2039` |
| `ComandoActualizarar` | botón | `OnClick` → `ComandoActualizarar_Click` | Invalida la caché de expedientes de `Entorno` | `:12-14`, `.form.txt:2414` |
| `ComandoNuevoResponsable` | botón | `OnClick` → `ComandoNuevoResponsable_Click` | Abre el editor de responsable | `:298-325`, `.form.txt:1917` |
| `ComandoNuevaEmpresaUsuario` | botón | `OnClick` → `ComandoBuscarExpediente_Click` | Comparte handler con la búsqueda de expediente | `:52`, `.form.txt:1386` |
| `ComandoMotivos`, `ComandoVerDetalleExpediente`, `ComandoCambioTipo` | botones | varios | Motivos, detalle de expediente y cambio de tipo | `:152, 183, 907`, `.form.txt:1656, 2545, 2098` |
| `ComandoBuscarExpediente`, `ComandoIrABusquedaExpedientes` | botones | varios | Búsqueda directa y apertura del buscador | `:52, 16`, `.form.txt:2276, 2354` |
| `PalabraClave` | texto | `OnKeyDown` → `PalabraClave_KeyDown` | Con Enter o Tab lanza la búsqueda de expediente | `:259-276`, `.form.txt:2192` |
| `lblEstadoSolicitud` | etiqueta | — | Estado calculado de la solicitud; sólo visible en edición | `:283, 434, 512-513`, `.form.txt:1783` |
| `Etiqueta36`, `Etiqueta51`, `Etiqueta62`, `Etiqueta72`, `Etiqueta80`, `Etiqueta94`, `Etiqueta96`, `Etiqueta98`, `Etiqueta100`, `Etiqueta102`, `Etiqueta104`, `Etiqueta225`, `Etiqueta281` | etiquetas | — | Rótulos de los campos; nombres internos de plantilla | `.form.txt:636-1281` |

## Flujos

### Normal — carga

1. `Form_Load` bloquea la edición, llama a `EstablecerDatos` y falla ruidosamente si ésta devuelve error (`:337-340`).
2. `EstablecerDatos` (`:359-523`) empieza por `EstablecerCombos` (`:364`), exige que haya solicitud activa (`:368-371`) y carga la fotografía inicial si la solicitud ya tiene identificador (`:372-377`).
3. **Alta** (`m_ObjSolicitudAlInicio Is Nothing`, `:378-437`): copia los datos del solicitante activo a los controles, habilita el detalle de expediente si el usuario tiene uno (`:401-410`), bloquea el cambio de tipo (`:432`) y oculta el rótulo de estado (`:434`).
4. **Edición** (`:439-511`): repinta los mismos campos desde `m_ObjSolicitudAlInicio`, decide si el cambio de tipo está permitido (`:440-449`), deja el rótulo de estado visible con `EstadoCalculadoTitulo` (`:512-513`) y recuerda el valor inicial de `CorreosAutomaticos` para detectar el cambio (`:514`).

### Normal — guardado

1. `ComandoRegistrar_Click` llama a `Registrar EnumSiNo.Sí` (`:666`), que aborta con «No hay cambios que guardar» si `HaHabidoCambios` es falso (`:839-842`), vuelca los controles con `RellenarObjeto` (`:844`) y persiste con `SolicitudOperaciones.RegistrarDatos` (`:851`).
2. Recarga la fotografía inicial (`:856`) y repinta el contenedor padre (`:857-861`).
3. Vuelve al handler, que evalúa el giro de `CorreosAutomaticos` de «No» a «Sí» (`:675`) —preguntando al usuario sólo si el proceso ya tiene algún disparador de fechas (`:680-695`) y reprogramando avisos con `ReprogramarAvisosFuturos` (`:700`)— y el giro de «Sí» a «No», que anula lo programado con `AnularTodosLosProgramados` (`:711-722`).
4. Cierra `FormSolicitudAlta` si estuviera abierta (`:729-731`) y confirma «Datos grabados correctamente» (`:734`).

### Alternativos

- **Sin cambios:** `Registrar` corta antes de escribir y el error se muestra como si fuera un fallo (`:839-842`, `:745`).
- **Reprogramación fallida:** el aviso distingue «Datos guardados, pero hubo un error reprogramando» y no aborta (`:702-705`).
- **Búsqueda con Enter o Tab:** `PalabraClave_KeyDown` da el foco al botón de búsqueda, lo devuelve al campo, dispara la búsqueda y anula la tecla (`:265-275`).
- **Fallo al limpiar la cola de correos:** se avisa pero el guardado se da por bueno y `m_Error` se limpia (`:719-722`).

### Errores

- Todos los handlers siguen el patrón `On Error GoTo errores` con el centinela `1000` para los errores propios ya descritos en `m_Error`; los errores reales se formatean, se envían por `CorreoAlAdministrador` (`FU.bas:585`) y se muestran en un `MsgBox` crítico.
- `ComandoRegistrar_Click` reactiva `AllowEdits` y quita el reloj de arena incluso en su rama de error (`:743-744`), y **muestra el mensaje aunque `m_Error` esté vacío**, que es lo que ocurre cuando el error es el centinela y el mensaje ya se había mostrado antes (`:745`).

## Reglas, validaciones y cálculos

- **Cambio de tipo:** deshabilitado en alta (`:432`) y en traspaso (`:443`), habilitado en el resto de los tipos con `TIPO` informado (`:445`).
- **Sin cambios no se escribe:** `Registrar` compara 16 campos con `Nz` y aborta si todos coinciden (`:534-603, 839-842`).
- **Reprogramación condicionada:** sólo se pregunta al usuario si ya existe `FechaEnvioExcel` o `FechaTramitacionAltaMarga` en `SolicitudFechas` (`:680-686`); sin proceso iniciado, activar los correos no genera ninguna cola.
- **Combos por código:** `EstablecerCombos` (`:747-825`) llena empresas usuarias, empresas tramitadoras y responsables desde los diccionarios de `Entorno` (`SumTramitadoras`, `Suministradores`, `Responsables`), más `RellenarComboTramitadores` (`FU.bas`) y `RellenarComboMotivos`.
- **Identidad del solicitante:** `RellenarObjeto` escribe `IDUsuarioHPS` dos veces seguidas, primero desde la sesión y después desde el control (`:640-643`).
- **Solo lectura efectiva:** `AllowEdits` se manipula en casi todos los handlers (`:335, 345, 352, 663, 728, 743`) sin que ningún control quede bloqueado por regla de negocio; el bloqueo real es la absence de `RecordSource`.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Fotografía inicial y recarga | `constructor.getSolicitud` | DAO, no acotado en este fichero | `:373, 856` |
| Persistencia de la solicitud | `SolicitudOperaciones.RegistrarDatos` | DAO, no acotado en este fichero | `:851` |
| Cola de correos automáticos | `ReprogramarAvisosFuturos` y `AnularTodosLosProgramados` | DAO sobre automatismos | `:700, 717` |
| Combos de tramitadoras y responsables | diccionarios de `Entorno`, con origen en `constructor` | DAO, cacheados | `:767, 783, 798` |
| Motivos | `RellenarComboMotivos` | DAO | `:813` |
| Invalidar caché de expedientes | `Set m_ObjEntorno.Expedientes = Nothing` | Derivado | `:13` |

## Coexistencia

Pendiente de `../coexistencia.md`. Lo relevante para la conversión: el formulario trabaja contra la entidad en memoria y delega toda la escritura en `SolicitudOperaciones`, lo que lo hace traducible a una operación de API sin más trámite. Lo que no se traslada es la detección de cambios por comparación explícita de 16 campos (`:534-603`), que en web es responsabilidad del servidor, y la carga de combos desde diccionarios cacheados en `Entorno`, que pasan a ser consultas pagitadas.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-HPS-S-50.** `Err.Numbr` en `ComandoActualizar_Click` (`:289`): la propiedad no existe, así que el aviso de error sale sin código, a diferencia de los otros cinco manejadores del mismo `.cls`, que usan `Err.Number`.
- **H-HPS-S-51.** `RellenarObjeto` asigna `IDUsuarioHPS` dos veces consecutivas: primero desde `m_ObjUsuarioHPSActivo.id` (`:640-642`) y acto seguido desde el control, con `Nz(…, "")` (`:643`). La segunda asignación anula la primera, de modo que un control vacío deja la solicitud sin solicitante aunque la sesión tenga uno válido, y ese `""` es lo que llega a `RegistrarDatos`.
- **H-HPS-S-52.** Al terminar el guardado, `ComandoRegistrar_Click` cierra `FormSolicitudAlta` (`:729-731`), un formulario que pertenece al flujo de alta y no al contenedor de renovación ni al de traspaso: en este flujo la orden no tiene efecto, y si el alta estuviera abierta desde otra ruta, se cierra sin guardar y sin aviso.
- **H-HPS-S-53.** El mismo grupo de campos se enumera tres veces de forma literal —`EstablecerDatos` (`:439-511`), `HaHabidoCambios` (`:534-603`) y `RellenarObjeto` (`:621-643`)— y las tres listas no coinciden: `RellenarObjeto` escribe `IDUsuarioHPS`, que `HaHabidoCambios` no compara, de modo que cambiar ese control no marca la solicitud como modificada y el guardado aborta con «No hay cambios que guardar».
- `Registrar` recarga la fotografía inicial con `constructor.getSolicitud` sin pasar `p_Error` (`:856`), a diferencia de las dos llamadas anteriores del mismo `.cls` (`:373`): un fallo de lectura deja `m_ObjSolicitudAlInicio` a `Nothing` y el siguiente guardado vuelve a tratar la solicitud como alta.
- `ComandoRegistrar_Click` lee `m_ObjEntorno.Configuracion.CorreosAutomaticos` (`:672`) sin comprobar que `Configuracion` esté cargada.
- `ComandoIrABusquedaExpedientes_Click` invoca `Form_FormExpedientesBusqueda.Filtrar` sobre el **nombre de clase** (`:32`) en lugar de la instancia `m_FormExpedientes` que acaba de enlazar (`:29`): funciona por la instancia abierta, pero depende de que la resolución de nombres de VBA acierte.
- Los combos se arman concatenando `id & ";" & nombre` (`:776, 791`) y `correo & ";" & nombre & ";" & descripción` (`:806`), con la misma fragilidad de separador que el resto de listas del legacy.
- El rótulo de estado se repinta en `ComandoActualizar_Click` sin recargar datos (`:283`): si otra sesión modificó la solicitud, el rótulo va obsoleto.
- **Defectos de interfaz heredados (a verificar):** el `Caption` de diseño es «EXPEDIENTES Versión 2024-001 Administrador» (`.form.txt:39`), y cinco combos llevan `RowSourceType ="Value List"` con datos reales de empresa (`1;Appian;2;AXPE;…`, `.form.txt:741`, `:1324`), de personas (`Martina Torralba Rodríguez;Esperanza del Álamo Arriba`, `.form.txt:1549`) y de motivos (`ARGO;BESKAR;…`, `.form.txt:1620`)versionados en el fichero. Doce etiquetas conservan nombres internos (`Etiqueta36` … `Etiqueta281`, `.form.txt:636-1281`) y el par `ComandoActualizarar` / `ComandoActualizarar_Click` arrastra un error tipográfico en el nombre (`:12`, `.form.txt:2414`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: el formulario como paso de un flujo guardado por el servidor, sin detección de cambios en cliente; combos con búsqueda remota y paginación en lugar de diccionarios cacheados; y la reprogramación de correos automáticos como efecto explícito de la operación de guardar, con su resultado devuelto en la respuesta.

## Criterios de aceptación

Propuestos, no aprobados:

1. Sin solicitud activa o sin solicitante, el formulario no se abre y el motivo se comunica.
2. Los datos de la solicitud se pintan completos tanto en alta como en edición, y el estado calculado sólo aparece en edición.
3. Sin cambios, no se escribe nada y se informa de que no hay nada que guardar.
4. Con cambios, se guarda una sola fila con el solicitante de la sesión, no el del control vacío.
5. Activar los correos automáticos cuando el proceso ya tiene fechas pide confirmación y reprograma los avisos; si no las tiene, no pregunta ni reprograma.
6. Desactivar los correos automáticos anula lo programado y avisa del resultado, sin invalidar el guardado.
7. Un fallo al reprogramar tras un guardado correcto no se presenta como fallo del guardado.
8. Tras guardar, el contenedor se repinta con la fotografía nueva.

## Pendientes

- Si `Constructor.getSolicitud` acepta `p_Error` en su firma y qué devuelve cuando el identificador no existe: el formulario lo llama de dos maneras distintas (`:373` y `:856`).
- Si `SolicitudOperaciones.RegistrarDatos` trata `IDUsuarioHPS` vacío como error: de ello depende la gravedad de H-HPS-S-51, y no es visible desde este `.cls`.
- Qué plantilla usa `AbrirAyuda` para este formulario: el `.cls` no declara un botón de ayuda propio, luego el control heredado no se identifica por nombre.
- Comportamiento de `m_FormResponsable_Alta` (`:877`) y `m_FormCambioTipo_TipoCambiado` (`:936`): los eventos se reciben con `WithEvents` pero su cuerpo no se cita en esta ficha.
- Recuento de controles: el inventario fija 64 y el `.form.txt` contiene 63 bloques `Begin <Tipo>`; la diferencia no se reproduce con una regla simple de recuento.