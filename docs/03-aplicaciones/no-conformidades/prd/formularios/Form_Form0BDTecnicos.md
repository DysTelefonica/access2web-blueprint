# Form_Form0BDTecnicos

> **Commit fijado:** `1d7a2f826a1e3c582874e4e29886a6183efb9067` · **Fuente:** `src/forms/Form_Form0BDTecnicos.cls` (191 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Portada de arranque: carga el entorno (`EVE`), decide el rol del usuario y, si es técnico, se queda como menú reducido (NC de proyecto, seguimiento y vídeo de ayuda); si no, abre [Form_Form0BDOpciones](Form_Form0BDOpciones.md) y se cierra. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`). Las citas son `Form_Form0BDTecnicos.cls:línea` salvo indicación; `FG.bas` abrevia `modules/Funciones Generales.bas` y `VG.bas`, `modules/Variables Globales.bas`.

## Entrada y salida

Ningún objeto de `src` abre este formulario: es el candidato a formulario de inicio (ver Pendientes). Abre `FormNCProyectoGestion` (`:42`), `FormNCProyectoSeguimiento` (`:70`) y, para no técnicos, `Form0BDOpciones` cerrándose (`:133-134`). Cierra Access con `acQuitSaveAll` desde «Salir» (`:172`) y con `acQuitSaveNone` si falla la apertura (`:154`).

## Controles y eventos

`.form.txt` declara 23 bloques `Begin`: el formulario (`.form.txt:4`), 8 de valores por defecto (`.form.txt:410-465`), 3 secciones (`EncabezadoDelFormulario` `.form.txt:480`, `Detalle` `.form.txt:832`, `PieDelFormulario` `.form.txt:2149`) y los 11 controles de la tabla.

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `ComandoMenuParteProyecto` | botón | `OnClick` | Abre `FormNCProyectoGestion` sin filtro inicial | `:33-58`, `.form.txt:848-850` |
| `lblNoConformidadesDeProyecto` («NC de proyecto») | etiqueta con hipervínculo `#` | `OnClick` | Si el botón está habilitado, ejecuta su handler | `:185-190`, `.form.txt:1203-1207` |
| `ComandoSeguimiento` | botón | `OnClick` | Reabre `FormNCProyectoSeguimiento` | `:60-85`, `.form.txt:1695-1697` |
| `ComandoAyuda` | botón (texto «Salir» en el `.form.txt`) | `OnClick` | Abre el vídeo de ayuda para técnicos | `:7-31`, `.form.txt:746-748` |
| `cmdSalir` | botón | `OnClick` | Confirma y cierra Access | `:157-183`, `.form.txt:2196-2198` |
| `lblSeguimientos` | etiqueta con hipervínculo `#` | — | Texto fijo «Seguimiento»; ningún código lo actualiza | `.form.txt:2125-2128`; `FG.bas:317-324` |
| `lblTitulo1`, `lblVersion` | etiquetas | — | Título según modo pruebas (16 puntos) y versión | `:105-123`, `.form.txt:498, 1236` |
| `lblEstado` | etiqueta oculta | — | Destino de `Avance`/`AvanceCerrar` | `:99, 124`, `FG.bas:1426, 1456`, `.form.txt:2157-2169` |
| `Imagen376`, `Imagen70` | imágenes | — | Decorativas | `.form.txt:520, 1261` |

## Flujos

### Apertura (`Form_Open`, `:87-155`)

1. `Ajustar Me` y `Avance "Obteniendo Variables de estado"` (`:94-99`).
2. `EVE` (`:101-104`; `VG.bas:294`): reinicia `TempVars` y fija `EnPruebas = "No"` (`VG.bas:313, 322`); crea el `Entorno` (`VG.bas:336`); identifica al usuario por el argumento de línea de órdenes o, si no lo hay, por el usuario de Windows en `TbUsuariosAplicaciones` (`VG.bas:354-381`); sin usuario, error «No se ha podido determinar el usuario que está usando la herramienta» (`VG.bas:382-385`); calcula los roles (`VG.bas:390-404`) y precarga las propiedades del entorno, acumulando los fallos en «Ha habido errores al inicio» (`VG.bas:416-444`).
3. Título, cinta, usuario conectado y versión (`:105-123`).
4. Técnico ⇒ se queda abierto (`:125-132`); resto ⇒ abre `Form0BDOpciones` y se cierra (`:133-134`).

### Navegación del técnico

- «NC de proyecto» ⇒ `DoCmd.OpenForm "FormNCProyectoGestion"` (`:42`).
- «Seguimiento» ⇒ cierra y reabre `FormNCProyectoSeguimiento` (`:67-70`).
- «Ayuda» ⇒ si el vídeo de la ruta de red configurada (`Entorno.URLVideoAyudaNCProyectoParaTecnico`, `Entorno.cls:682-704`) no existe, «El video de ayuda no se encuentra disponible en este momento»; si existe, `ShellExecute` (`:14-18`).
- «Salir» ⇒ pregunta «¿Desea salir de la aplicación» con «No» por defecto; con «Sí», `DoCmd.Quit acQuitSaveAll` (`:159-172`).

### Errores

Error 1000 ⇒ aviso con `m_Error`; otro ⇒ `CorreoAlAdministrador` y `MsgBox` crítico. En `Form_Open` cualquier error, incluido no identificar al usuario, cierra Access sin guardar (`:144-154`).

## Reglas, validaciones y cálculos

- **Asignación de rol** (`VG.bas:390-404`): administrador si `EsAdministradorCalculado` (campo `EsAdministrador = "Sí"` o permisos, `Usuario.cls:59-80`); si no, Calidad si `EsUsuarioCalidadCalculado` (`Usuario.cls:116`); **cualquier otro usuario registrado es técnico** (`VG.bas:404`).
- **Alias de equipo:** los usuarios de Windows `Local1` y `adm1` se tratan como `adm` (`VG.bas:374-375`).
- **Usuario por argumento:** si Access se abre con un argumento de línea de órdenes, se toma como correo del usuario (`VG.bas:354, 366-367`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Usuario conectado | `TbUsuariosAplicaciones` vía `constructor.getUsuario` | SQL dinámico concatenado | `VG.bas:367, 376` |
| Propiedades del entorno | varias, una por propiedad de `Entorno.ColItems` | según propiedad | `VG.bas:416-437` |
| Versión de la aplicación | fichero de configuración | `LeerIni` | `Entorno.cls:1044` |
| Registro de errores no controlados | `TbCorreosEnviados` | DAO `AddNew/Update` | `FG.bas:1291-1305` |

Sin transacciones.

## Coexistencia

Pendiente de la coexistencia (documento aún no redactado). Ya se ve: la identidad depende del usuario de Windows (`VG.bas:372-376`) y de una ruta de red fija para aplicaciones y vídeos (`VG.bas:337`; `Entorno.cls:700`); identificador de correo `Max + 1` (`constructor.bas:1165`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Cualquier usuario registrado sin rol de administrador ni de Calidad entra como técnico, incluso sin permisos para esta aplicación (`VG.bas:404`; `Usuario.cls:70-73`).
- Un argumento de línea de órdenes suplanta la identidad sin más comprobación (`VG.bas:366-367`).
- `lblSeguimientos` nunca muestra cifras: `PintarIndicadores` solo actualiza los dos submenús (`FG.bas:317-324`).
- «NC de proyecto» abre la búsqueda sin el filtro de NC abiertas y de responsable que aplica el submenú de proyectos (`:42` frente a `Form_Form0BDOpcionesParteProyectos.cls:131-136`).
- La sección `Detalle` no tiene `OnDblClick` (`.form.txt:832`), a diferencia de los otros menús.
- «Salir» usa `acQuitSaveAll` (`:172`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: identidad por autenticación corporativa en vez de usuario de Windows o argumento; portada única con opciones por rol; denegar el acceso a usuarios sin permiso explícito.

## Criterios de aceptación

Propuestos, no aprobados:

1. Un usuario no identificable recibe «No se ha podido determinar el usuario que está usando la herramienta» y no accede.
2. Un técnico ve el menú reducido con su nombre y colectivo «Técnico».
3. Un administrador o usuario de Calidad pasa directamente a la portada completa.
4. «Seguimiento» abre el seguimiento de NC de proyecto.
5. «Salir» pide confirmación y no sale si se responde «No».

## Pendientes

- Confirmar que es el formulario de inicio (propiedad de arranque o `AutoExec`): no está en `src`.
- Lista de propiedades de `Entorno.ColItems` que precarga `EVE` y sus consultas (`VG.bas:416`).
- Formato del argumento de línea de órdenes que usa el lanzador (`VG.bas:354`).
