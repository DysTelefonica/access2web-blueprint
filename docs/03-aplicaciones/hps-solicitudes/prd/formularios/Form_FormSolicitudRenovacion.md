# Form_FormSolicitudRenovacion

> **Commit fijado:** `cfa1937ab05a9e40150ea9236243898de96f84da` · **Fuente:** `src/forms/Form_FormSolicitudRenovacion.cls` (151 líneas) y `.form.txt` (21 bloques `Begin`) · **PRD:** [README](../README.md)
>
> **Abreviaturas:** `FU.bas` = `modules/FUNCIONES UTILES.bas`, `VG.bas` = `modules/Variables Globales.bas`, `constructor.bas` = `modules/constructor.bas`; las clases de `src/classes/` se citan por el nombre del fichero. Las citas son `Form_FormSolicitudRenovacion.cls:línea` salvo indicación.

Contenedor de la renovación y la elevación de grado. No pinta datos: los monta en un control de navegación con dos destinos (`NavDatos` y `NavFechas`) cuyo `SourceObject` es `FormSolicitudRenovacionDatos`, y cuyo subformulario `FrmDetalle` recibe `FormSolicitudAltaTraspasoDatos` (`.form.txt:585`). Es **no enlazado** y su contrato de entrada es `OpenArgs`, que debe valer `"Renovación"` o `"Elevación Grado"` (`:69-73`). La lógica de datos vive en los formularios hijos.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormSolicitudesGestion` | `DoCmd.OpenForm "FormSolicitudRenovacion", OpenArgs:=.TIPO`, cerrando antes la instancia previa | `Form_FormSolicitudesGestion.cls:542-546` |
| `Form_FormUsuariosHPSBusqueda` | Igual, con `m_Tipo`; además borra la solicitud activa antes de abrir | `Form_FormUsuariosHPSBusqueda.cls:150-155` |
| `Form_FormTareasTramitadorPendientes` | Igual, con el `.TIPO` de la tarea | `Form_FormTareasTramitadorPendientes.cls:192-196` |

| Hacia | Cómo | Evidencia |
|---|---|---|
| `FormSolicitudRenovacionDatos` | `Me.FrmDetalle.SourceObject` y `Me.NavDatos.NavigationTargetName` se fijan al cargar | `:82-83` |
| `FormSolicitudRenovacionFechas` | Segundo destino del control de navegación, definido en diseño | `.form.txt:497` |
| `Form_FormSolicitudesAltaTraspasoDatos` | El subformulario invoca `EstablecerDatos` del padre por `Me.Parent` | `Form_FormSolicitudesAltaTraspasoDatos.cls:820-821` |
| `FormWeb` (vía `InformeDelRegistro`) | `ComandoVerLogs` | `:35`, `FU.bas:1721-1748` |

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| Formulario | — | `OnLoad` → `Form_Load` | Valida parámetros y usuario, prepara datos y monta los destinos del navegador | `:58-99` |
| `NavDatos` | control de navegación | — | Destino «Datos»; lo fija el código al cargar | `:82`, `.form.txt:427-428` |
| `NavFechas` | pestaña del control de navegación | — | Destino «Fechas»; se habilita o deshabilita según haya solicitud | `:110, 121, 135`, `.form.txt:479-497` |
| `FrmDetalle` | subformulario | — | Alberga `FormSolicitudAltaTraspasoDatos` | `:83`, `.form.txt:585` |
| `lblTitulo` | etiqueta de encabezado | — | Rotulada en `EstablecerDatos` | `:109, 137`, `.form.txt:286-287` |
| `lblModificacion` | etiqueta | — | «Última modificación por …» | `:108, 125-126`, `.form.txt:707-708` |
| `ComandoAyuda` | botón | `OnClick` → `ComandoAyuda_Click` | Abre el PDF de ayuda | `:7-23`, `.form.txt:313-315` |
| `ComandoVerLogs` | botón | `OnClick` → `ComandoVerLogs_Click` | Genera y abre el informe de acciones de la solicitud | `:25-55`, `.form.txt:736-738` |
| `cmdSalir` | botón | `OnClick` → `cmdSalir_Click` | Cierra sin guardar | `:149-151`, `.form.txt:623-625` |
| `BotónDeNavegación1` | botón de navegación | — | «[Agregar nuevo]»; no tiene handler en este `.cls` | `.form.txt:534-535` |

## Flujos

### Normal — alta

1. `Form_Load` valida que `OpenArgs` sea uno de los dos tipos admitidos (`:69-73`) y que haya un usuario HPS activo (`:74-77`); cualquiera de los dos fallos cierra el formulario.
2. `EstablecerDatos` (`:101-146`) ve que no hay solicitud activa, crea una `New solicitud` con `TIPO = m_Tipo` e `IDUsuarioHPS = m_ObjUsuarioHPSActivo.id` (`:130-132`) y deshabilita `NavFechas` (`:135`).
3. `Form_Load` monta los destinos (`:82-83`) y reactiva la edición (`:88`).

### Normal — edición

1. Con `m_ObjSolicitudActiva.IDSolicitud` informado, `EstablecerDatos` carga la fotografía inicial con `constructor.getSolicitud` (`:112`) y habilita `NavFechas` (`:121`).
2. El rótulo y la línea de última modificación se pintan con los datos de la sesión (`:125-126, 137`).
3. `ComandoVerLogs` delega en `InformeDelRegistro` (`:35`), que genera el HTML y abre `FormWeb` (`FU.bas:1734-1741`).

### Alternativos

- **`OpenArgs` vacío o distinto de los dos tipos:** `Form_Load` lo trata como parámetro insuficiente, lanza el centinela 1000 y cierra (`:70-73, 98`).
- **Sin usuario HPS seleccionado:** mismo camino con otro mensaje (`:74-77`); es la guarda que evita abrir la renovación sin identidad de solicitante.
- **`EstablecerDatos` público:** el subformulario de datos lo invoca para repintar el padre tras una edición suelta (`Form_FormSolicitudesAltaTraspasoDatos.cls:820-821`).

### Errores

- Las tres ramas `errores:` (`:17-22`, `:45-54`, `:90-98`) ignoran el centinela `Err.Number = 1000`, que en este formulario se usa como «error propio ya descrito en `m_Error`»; sólo los errores ajenos rellenan el texto y, en `ComandoVerLogs` y `Form_Load`, envían correo al administrador (`:48, 93`).
- `EstablecerDatos` tiene su propio manejador (`:140-145`) que sólo propaga `p_Error` y no cierra el formulario; la decisión de cerrar la toma su llamador (`Form_Load`, `:78-81, 98`).

## Reglas, validaciones y cálculos

- **Contrato de apertura:** `OpenArgs` ∈ {`Renovación`, `Elevación Grado`}; es el mismo valor que se copia a `Solicitud.TIPO` (`:131`).
- **Precondición de identidad:** `m_ObjUsuarioHPSActivo` tiene que estar informado (`:74`).
- **`IDSolicitud` es `String`** (`Solicitud.cls:14`), de modo que la comparación con `""` de `:107` distingue alta de edición sin conversiones.
- **Sin datos propios:** las únicas lecturas de este `.cls` son `constructor.getSolicitud` (`:112`) y las propiedades de sesión; toda la persistencia ocurre en los formularios hijos.
- **`AllowEdits`** se pone a `False` al entrar en cada handler y a `True` al salir, incluso en las ramas de error (`:32, 43, 51, 67, 88`); sobre un contenedor sin `RecordSource` la propiedad no protege nada.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Fotografía inicial de la solicitud | vía `constructor.getSolicitud` | DAO, no acotado en este fichero | `:112` |
| Última modificación | `Solicitud.UltimaModificacionTexto`, derivado de los datos de la solicitud | Derivado, sin DAO en este `.cls` | `:126`, `Solicitud.cls:244-279` |
| Informe de acciones | `HTMLLog` → «Informe de Acciones», abierto en `FormWeb` | DAO y E/S de fichero | `:35`, `FU.bas:1613-1641, 1734` |
| Identidad de solicitante | `m_ObjUsuarioHPSActivo` (sesión) y `UsuarioHPS.id` | Derivado | `:74, 132` |

## Coexistencia

Pendiente de `../coexistencia.md`. El patrón es el más replicable del legacy para la web: un contenedor que sólo compone destinos y delega el trabajo. Su equivalente es una ruta con pestañas y un formulario por pestaña, pero obliga a decidir qué pasa con la fotografía inicial (`constructor.getSolicitud`) que hoy se toma al abrir el contenedor y se vuelve a tomar cuando el subformulario pide repintar.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-HPS-S-42.** `Err.Numbr` en `ComandoAyuda_Click` (`:19`): la propiedad no existe, así que el mensaje que llega al usuario es «El método ComandoAyudaFormulario_Click ha devuelto el error: » sin código de error. Los otros dos manejadores de este mismo `.cls` sí usan `Err.Number` (`:47, 92`).
- **H-HPS-S-43.** Las dos decisiones de la rama de alta se sobrescriben después: `Me.lblTitulo.Caption = "ALTA DE SOLICITUD DE HPS"` (`:109`) queda inerte porque `:137` vuelve a escribir el rótulo con «RENOVACIÓN PARA …» para los dos casos, y `Me.lblModificacion.Visible = False` (`:108`) se deshace en `:125`, que la hace visible y le asigna `UltimaModificacionTexto` de una solicitud sin identificador.
- **H-HPS-S-44.** `ComandoVerLogs` no comprueba que exista una solicitud registrada y el diseño no lo deshabilita (`.form.txt:736-738`): en la rama de alta, donde `IDSolicitud` queda vacío (`:130-132`), el botón sigue disponible y llama a `InformeDelRegistro` con un identificador inexistente (`FU.bas:1730` → `HTMLLog`), con lo que se genera y abre un informe de acciones de una solicitud que todavía no existe.
- **H-HPS-S-45.** `EstablecerDatos` es `Public` (`:101`) sólo para que el subformulario de datos lo invoque sobre el padre: `Form_FormSolicitudesAltaTraspasoDatos.cls:820-821` resuelve el contenedor con `Me.Parent` y llama a su método de presentación, con lo que el hijo prepinta al padre y ninguno de los dos sabe quién inició el refresco.
- `BotónDeNavegación1`, rotulado «[Agregar nuevo]», no tiene handler en este `.cls` (`.form.txt:534-535`): o lo maneja el control de navegación, o es otro resto de plantilla (a verificar).
- Rótulos heredados: el `Caption` de diseño es «EXPEDIENTES Versión 2024-003 Martina Torralba Rodríguez EN PRUEBAS» (`.form.txt:41`) y `lblModificacion` guarda «Última modificación por Martina Torralba Rodríguez el 12/11/2024» (`.form.txt:708`), con datos personales reales en el fichero versionado. `Form_Load` sobrescribe el primero (`:68`) y `EstablecerDatos` el segundo (`:126`), así que sólo son visibles en diseño, pero permanecen en el repositorio. `ComandoAyuda` aparece rotulado «Salir», igual que `cmdSalir` (`.form.txt:314, 624`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: ruta de renovación con pestañas en lugar de ventana anidada; el tipo de trámite como parámetro de ruta en lugar de `OpenArgs`; y el informe de acciones como descarga o vista en el propio expediente, sin generar un fichero HTML temporal.

## Criterios de aceptación

Propuestos, no aprobados:

1. Sin `OpenArgs` válido el formulario no se abre y el motivo se comunica.
2. Sin usuario HPS seleccionado no se abre ningún formulario de datos.
3. En alta, la pestaña de fechas no está disponible; en edición, sí.
4. En alta, el rótulo distingue el alta de la edición y no se sobrescribe después.
5. Con una solicitud registrada, «Ver Logs» produce el informe de acciones de esa solicitud.
6. En alta, «Ver Logs» no está disponible.
7. Cerrar el contenedor no deja a medias el trabajo del subformulario de datos.

## Pendientes

- Comportamiento de `BotónDeNavegación1`: no hay handler en el `.cls` y no consta en `src` qué lo invoca.
- Si `Form_FormSolicitudesAltaTraspasoDatos` es realmente el subformulario de este contenedor en todos los tipos de trámite, o sólo en «Renovación»: el `SourceObject` se fija por código (`:83`) pero la otra rama comparte el mismo `.form.txt`.
- Qué ocurre con la solicitud en curso cuando `Form_FormUsuariosHPSBusqueda` cierra este formulario con `acSaveNo` antes de reabrirlo (`Form_FormUsuariosHPSBusqueda.cls:151`): el comportamiento depende del estado del subformulario de datos, que no se analiza aquí.
- Recuento de controles: el inventario fija 25 y el `.form.txt` contiene 21 bloques `Begin <Tipo>`; la diferencia no se reproduce con una regla simple de recuento.