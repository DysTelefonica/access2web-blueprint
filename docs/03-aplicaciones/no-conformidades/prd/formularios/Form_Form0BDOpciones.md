# Form_Form0BDOpciones

> **Commit fijado:** `1d7a2f826a1e3c582874e4e29886a6183efb9067` · **Fuente:** `src/forms/Form_Form0BDOpciones.cls` (205 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Menú raíz de No Conformidades para administradores y usuarios de Calidad: muestra título, versión y usuario conectado, y da paso a los dos submenús, el de NC de proyecto ([Form_Form0BDOpcionesParteProyectos](Form_Form0BDOpcionesParteProyectos.md)) y el de NC de auditoría ([Form_Form0BDOpcionesAuditorias](Form_Form0BDOpcionesAuditorias.md)). Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`) y no lee ni escribe datos propios. Las citas son `Form_Form0BDOpciones.cls:línea` salvo indicación; `FG.bas` abrevia `modules/Funciones Generales.bas` y `VG.bas`, `modules/Variables Globales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| [Form_Form0BDTecnicos](Form_Form0BDTecnicos.md) | Tras cargar el entorno con `EVE`, si el usuario **no** es técnico lo abre y se cierra | `Form_Form0BDTecnicos.cls:125-134` |

Abre `Form0BDOpcionesParteProyectos` (`:16`) y `Form0BDOpcionesAuditorias` (`:72`) **sin cerrarse**; queda debajo como menú de vuelta. Cierra Access con `DoCmd.Quit acQuitSaveAll` desde «Salir» (`:185`) y con `acQuitSaveNone` si falla la apertura (`:167`).

## Controles y eventos

`.form.txt` declara 22 bloques `Begin`: el formulario (`.form.txt:4`), 8 bloques de valores por defecto (`.form.txt:718-773`), 3 secciones (`EncabezadoDelFormulario` `.form.txt:788`, `Detalle` `.form.txt:1047`, `PieDelFormulario` `.form.txt:3538`) y los 10 controles de la tabla.

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `ComandoMenuParteProyecto` | botón | `OnClick` | Abre el submenú de NC de proyecto | `:7-32`, `.form.txt:1064-1066` |
| `lblNoConformidadesDeProyecto` («NC de proyecto») | etiqueta con hipervínculo `#` | `OnClick` | Si el botón anterior está habilitado, ejecuta su handler | `:199-204`, `.form.txt:1419-1423` |
| `ComandoMenuAuditorias` | botón | `OnClick` | Abre el submenú de NC de auditoría | `:64-88`, `.form.txt:1451-1453` |
| `lblMenuAuditoria` («NC de auditoría») | etiqueta con hipervínculo `#` | — | Sin `OnClick` ni handler | `.form.txt:3054-3057` |
| `cmdSalir` | botón | `OnClick` | Confirma y cierra Access | `:170-196`, `.form.txt:3585-3587` |
| `lblTitulo1` | etiqueta | — | Título según modo pruebas | `:113-127`, `.form.txt:806` |
| `lblVersion` | etiqueta | — | «Versión: » + `VersionAplicacion` | `:130`, `.form.txt:3086` |
| `lblEstado` | etiqueta oculta | — | Destino de los mensajes de progreso de `Avance`/`AvanceCerrar` | `FG.bas:1422-1438, 1456`, `.form.txt:3546-3558` |
| `Imagen376`, `Imagen70` | imágenes | — | Decorativas | `.form.txt:828, 3111` |
| Sección `Detalle` | sección | `OnDblClick` | Recarga el entorno con `EVE` (gesto oculto) | `:34-62`, `.form.txt:1047-1048` |

## Flujos

### Apertura (`Form_Open`, `:95-168`)

1. Bloquea la edición, ajusta el tamaño con `Ajustar Me` (`:101-102`; `FG.bas:615`). La llamada a `EVE` está comentada (`:107-112`): el formulario depende de que el entorno ya esté cargado.
2. Si `TempVars("EnPruebas") = "Sí"`, título «… VERSIÓN DE PRUEBAS» a 14 puntos y cinta visible solo para administradores; si no, título normal a 25 puntos y cinta oculta (`:113-127`). `EVE` fija `EnPruebas = "No"` (`VG.bas:322`).
3. Pone en la barra de título `TituloUsuarioConectado` («Usuario conectado: nombre (colectivo)», `Entorno.cls:399-424`) y la versión leída del fichero de configuración (`:129-130`; `Entorno.cls:1034-1048`).
4. El bloque que deshabilitaba el menú de auditorías a los técnicos está comentado (`:133-147`).

### Navegación

- «NC de proyecto» (botón o etiqueta) ⇒ `DoCmd.OpenForm "Form0BDOpcionesParteProyectos"` (`:16, 199-204`).
- «NC de auditoría» (botón) ⇒ `DoCmd.OpenForm "Form0BDOpcionesAuditorias"` (`:72`).
- «Salir» ⇒ pregunta «¿Desea salir de la aplicación» con «No» por defecto; con «Sí», `DoCmd.Quit acQuitSaveAll` (`:172-185`).
- Doble clic en el fondo ⇒ `EVE` reinicia `TempVars`, crea un `Entorno` nuevo y recalcula usuario y roles (`:42`; `VG.bas:313, 336, 366-404`).

### Errores

Patrón común: error 1000 ⇒ aviso con `m_Error`; otro error ⇒ mensaje con número y descripción, `CorreoAlAdministrador` y `MsgBox` crítico (`:23-31, 79-87, 187-195`). En `Form_Open` cualquier error cierra Access sin guardar (`:157-167`).

## Reglas, validaciones y cálculos

- **Rol:** el formulario no comprueba el rol; el filtro lo hace `Form_Form0BDTecnicos` antes de abrirlo (`Form_Form0BDTecnicos.cls:125-133`). `EVE` asigna administrador, Calidad o, en cualquier otro caso, técnico (`VG.bas:390-404`).
- **Modo pruebas:** controla título, tamaño de letra y visibilidad de la cinta (`:113-126`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Usuario conectado y roles (solo con doble clic en el fondo) | `TbUsuariosAplicaciones` vía `constructor.getUsuario` | SQL dinámico concatenado | `VG.bas:367, 376` |
| Versión de la aplicación | fichero de configuración | `LeerIni` | `Entorno.cls:1044`, `VG.bas:491` |
| Registro de errores no controlados | `TbCorreosEnviados` (identificador `Max + 1`) | DAO `AddNew/Update` | `FG.bas:1291-1305`, `constructor.bas:1165` |

Sin transacciones.

## Coexistencia

Pendiente de `../coexistencia.md`. Este formulario no consulta datos; lo único sensible es `CorreoAlAdministrador`, que calcula el identificador con `SELECT Max(...)` y lo inserta a mano (`FG.bas:1291-1297`; `constructor.bas:1165`), incompatible con una secuencia de PostgreSQL sin adaptación.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Abierto directamente (sin pasar por `Form_Form0BDTecnicos`), `m_ObjEntorno` no existe y `:129` fallaría; el manejador cerraría Access (`:167`).
- `lblMenuAuditoria` parece un enlace (hipervínculo `#`) pero no tiene `OnClick`, a diferencia de `lblNoConformidadesDeProyecto` (`.form.txt:3054-3057` frente a `.form.txt:1419-1423`).
- Los submenús se abren sin cerrar este formulario y sin comprobar si ya están abiertos (`:16, 72`).
- «Salir» usa `acQuitSaveAll` (`:185`), que guarda cambios de diseño pendientes.
- El texto de confirmación carece del signo de cierre «?» (`:172`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: fusionar este menú con el de técnicos en una portada única que muestre opciones según rol; sustituir el doble clic oculto por una acción explícita de «recargar».

## Criterios de aceptación

Propuestos, no aprobados:

1. Un administrador o usuario de Calidad ve la portada con su nombre, colectivo y versión.
2. «NC de proyecto» (botón y etiqueta) lleva al submenú de proyectos.
3. «NC de auditoría» lleva al submenú de auditorías.
4. «Salir» pide confirmación y no sale si se responde «No».
5. Un técnico nunca llega a esta portada.

## Pendientes

- Forma en que se arranca la aplicación (formulario de inicio o `AutoExec`): no está en `src`.
- Comportamiento real de una etiqueta con hipervínculo `#` y sin `OnClick` en Access (`.form.txt:3057`).
