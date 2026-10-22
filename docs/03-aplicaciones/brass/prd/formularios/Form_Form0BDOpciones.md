# Form_Form0BDOpciones

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_Form0BDOpciones.cls` (175 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Menú principal de Brass («GESTIÓN BRASS (PARTE GESTORA)»): diez etiquetas que abren los formularios de gestión, un botón para salir de la aplicación y la comprobación del entorno al abrir. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`). Al abrirse llama a `EVE`, que inicializa globales, usuario, permisos y colores. Las citas son `Form_Form0BDOpciones.cls:línea` (abreviadas `:N`) salvo indicación; `VG.bas` = `modules/Variables Globales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Arranque de la aplicación | No hay ninguna llamada a `Form0BDOpciones` en `src`; se asume formulario de inicio definido en el binario | Pendiente (ver Pendientes) |

Salidas: cada etiqueta abre un formulario de gestión sin cerrar el menú (tabla siguiente). `cmdSalir` cierra la aplicación entera con `DoCmd.Quit acQuitSaveAll` (`:20`).

## Controles y eventos

Controles: 28 (37 `Begin` en el `.form.txt`: 1 `Form`, 7 plantillas de estilo por defecto —`Label`, `Rectangle`, `Line`, `Image`, `CommandButton`, `OptionGroup`, `UnboundObjectFrame`—, 1 `Section` y 28 controles con nombre).

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `lblEventos` («EVENTOS») | etiqueta | `Click`; `MouseMove` | Abre `FormEventoGestion` | CAP-3 | `:120-127`; `Form_Form0BDOpciones.form.txt:273-275` |
| `lblActividades` («ACTIVIDADES») | etiqueta | `Click`; `MouseMove` | Abre `FormActividadGestion` | CAP-4 | `:88-95`; `Form_Form0BDOpciones.form.txt:306-308` |
| `lblGestionFacturacion` («FACTURACIÓN») | etiqueta | `Click`; `MouseMove` | Abre `FormFacturacionGestion` | CAP-8 | `:144-151`; `Form_Form0BDOpciones.form.txt:340-342` |
| `lblGestionDeLosMateriales` («MATERIALES») | etiqueta | `Click`; `MouseMove` | Abre `FormMaterialesGestion` | CAP-7 | `:128-135`; `Form_Form0BDOpciones.form.txt:408-410` |
| `lblGestionEquipos` («EQUIPOS») | etiqueta | `Click`; `MouseMove` | Abre `FormEquipoGestion` | CAP-5 | `:136-143`; `Form_Form0BDOpciones.form.txt:442-444` |
| `lblGestionSubContrataciones` («SUBCONTRATACIONES») | etiqueta | `Click`; `MouseMove` | Abre `FormSubContratacionGestion` | CAP-8 | `:152-159`; `Form_Form0BDOpciones.form.txt:476-478` |
| `lblGastosYDietas` («GASTOS Y DIETAS») | etiqueta | `Click`; `MouseMove` | Abre `FormGastoGestion` | CAP-8 | `:96-103`; `Form_Form0BDOpciones.form.txt:510-512` |
| `lblGestionTecnicos` («TÉCNICOS») | etiqueta | `Click`; `MouseMove` | Abre `FormTecnicoGestion` | CAP-9 | `:112-119`; `Form_Form0BDOpciones.form.txt:1657-1659` |
| `lblInformes` («INFORMES») | etiqueta | `Click`; `MouseMove` | Abre `Form0BDInformes` | CAP-10 | `:104-111`; `Form_Form0BDOpciones.form.txt:544-546` |
| `lblOtrasOpciones` («OTRAS OPCIONES») | etiqueta | `Click`; `MouseMove` | Abre `Form0BDOpcionesOtras` | CAP-2 | `:160-167`; `Form_Form0BDOpciones.form.txt:374-376` |
| `cmdSalir` («Salir») | botón | `Click` → `cmdSalir_Click` | Confirma y cierra Access guardando todo | CAP-1 | `:5-28`; `Form_Form0BDOpciones.form.txt:649-651` |
| `lbVersion` | etiqueta | — | Recibe `m_ObjEntorno.VersionAplicacion` al abrir | CAP-2 | `:72`; `Form_Form0BDOpciones.form.txt:1580-1581` |
| `lbTitulo`, `lblEstado` | etiquetas | — | `lbTitulo` rotula y su `Caption` se copia a `Me.Caption` (luego se sobrescribe); `lblEstado` es el rótulo que `Avance` rellena | CAP-1 | `:53, :70`; `Form_Form0BDOpciones.form.txt:601-602, 693-694`; `FG.bas:7263-7283` |
| `ComandoAyudaFormulario` | botón | `OnClick` declarado, **sin handler** | Sin efecto | sin asignar | `Form_Form0BDOpciones.form.txt:1603-1605` |
| `CuadroSuperior`, `Cuadro1`, `Marco120` | 2 rectángulos y 1 grupo de opciones | — | Decoración | sin asignar | `Form_Form0BDOpciones.form.txt:576, 628, 725` |
| 10 imágenes `OLEIndependiente3`…`13` | imagen | — | Iconos de las opciones | sin asignar | `Form_Form0BDOpciones.form.txt:744-1491` |
| Formulario | — | `Form_Open` | Hourglass, `AjustarTamaño`, `EVE`, ribbon, títulos | CAP-1 | `:45-87` |
| Formulario | — | `Detalle_DblClick` | Muestra el ribbon sin comprobar rol | CAP-1 | `:35-40` |
| Formulario | — | `Detalle_MouseMove`, `CuadroFormulario_MouseMove` | `DesactivaLabel Me` | CAP-1 | `:30-33, :41-44` |

El `MouseMove` de cada etiqueta pone `ForeColor = longColorEtiquetaMenuPulsado` (`26367`, `VG.bas:357`); `DesactivaLabel` lo restaura (`FG.bas:157-169`).

## Flujos

### Apertura (`Form_Open`, `:45-87`)

1. Reloj de arena, `AjustarTamaño Me` y título de `lbTitulo` (`:49-53`).
2. `EVE strTextoError` (`:57`); si devuelve texto de error, `Err.Raise 1000` (`:58-60`). `EVE` (`VG.bas:264-457`) fija `IDAplicacion`, la bitness, `TempVars` `DatosEnLocal`/`EnDesarrollo` (valores «No» en código, `VG.bas:293-294`), rutas, usuario y permisos (`VG.bas:310-347`).
3. Ribbon visible solo si `EnDesarrollo = "Sí"` y el usuario es administrador (`:61-69`); en el resto de casos, oculto.
4. `Me.Caption = "Usuario conectado: " & Nombre & " " & m_TextoWin64`, versión en `lbVersion`, repintado (`:70-76`).
5. Error: `MsgBox` crítico y `DoCmd.Quit acQuitSaveNone` (`:80-86`). La aplicación se cierra si `EVE` falla o el usuario no tiene permiso (`VG.bas:338-347`).

### Salir (`cmdSalir_Click`, `:5-28`)

Pregunta «¿Desea salir de la aplicación» con «No» por defecto (`:8-11`); si acepta, `DoCmd.Quit acQuitSaveAll` (`:20`). Error: `MsgBox` con el texto (`:22-27`).

### Navegación

Cada `Click` ejecuta `DoCmd.OpenForm` bajo `On Error Resume Next` y no cierra el menú (`:89-91, :97-99, …`).

## Reglas, validaciones y cálculos

- **Acceso a la aplicación:** solo si el usuario existe y es administrador o técnico; si no, «El usuario no tiene permiso para abrir la aplicación» (`VG.bas:338-347`).
- **Usuario:** se obtiene de la línea de comandos (`VG.bas:310-317`) o del nombre de usuario de Windows, con dos alias fijos reasignados a `adm` (`VG.bas:318-327`).
- **Ribbon:** ver paso 3 del flujo de apertura (`:61-69`).
- **Colores:** `lngColorEtiquetaMenuSinPulsar = 8210719`, `longColorEtiquetaMenuPulsado = 26367` (`VG.bas:356-357`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Usuario conectado y permisos | tabla de usuarios (vía `Constructor.getUsuario`) | por determinar en la ficha del módulo | no | `VG.bas:313, 322` |
| Propiedades de entorno | vía `m_ObjEntorno` | por determinar | no | `VG.bas:377-379` |

El formulario no escribe datos. La versión mostrada es `m_ObjEntorno.VersionAplicacion` (`:72`).

## Coexistencia

Sin SQL en el formulario. A adaptar: `EVE` depende de `Wscript.Network` y de `Command` (`VG.bas:310-318`), no de la base. El enlace al backend lo fija el módulo de variables globales (contraseña literal en el código, `VG.bas:539, 573`, no se transcribe; `getdb` prueba antes la ruta de las tablas vinculadas, `VG.bas:556-562`); queda para la ficha de coexistencia global.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `Detalle_DblClick` muestra el ribbon a cualquier usuario, saltándose la regla de `Form_Open` (`:35-40` frente a `:61-69`).
- `cmdSalir` con `acQuitSaveAll` guarda sin preguntar cualquier objeto abierto (`:20`).
- En `errores` de `cmdSalir_Click` y `Form_Open`, si el error es 1000 sin texto previo, `strTextoError` queda vacío y el `MsgBox` sale en blanco (`:22-27`, `:80-86`).
- `Form_Open` cierra Access ante cualquier fallo de `EVE` (`:86`), sin reintento.
- `ComandoAyudaFormulario` tiene `OnClick` pero no handler (`Form_Form0BDOpciones.form.txt:1605`); `CuadroFormulario_MouseMove` (`:30-33`) apunta a un nombre que no está entre los controles del `.form.txt`.
- Mensaje sin cierre de interrogación («¿Desea salir de la aplicación», `:8`) y texto «Salir de Gestión de Proyectos» como título.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: menú por rol en lugar de diez enlaces fijos.

## Criterios de aceptación

Propuestos, no aprobados:

1. Un usuario sin perfil de administrador o técnico no accede (`VG.bas:343-347`).
2. Cada una de las diez entradas lleva a su pantalla de gestión (`:88-167`).
3. «Salir» pide confirmación con «No» por defecto (`:8-11`).
4. Se muestra «Usuario conectado: <nombre>» y la versión (`:70-72`).
5. El ribbon/herramientas de desarrollo solo para administradores y entorno de desarrollo (`:61-69`).

## Pendientes

- Quién abre este formulario (formulario de inicio del binario; no está en `src`).
- Contenido de `m_ObjEntorno.VersionAplicacion` y de `Constructor.getUsuario`.
