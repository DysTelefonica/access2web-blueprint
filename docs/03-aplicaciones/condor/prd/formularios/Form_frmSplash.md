# Form_frmSplash

> **Commit fijado:** `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · **Fuente:** `src/forms/Form_frmSplash.cls` (77 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Pantalla de arranque: inicializa el entorno con `EVE`, muestra una barra de progreso y abre el menú principal según el rol. Es **no enlazado** (sin `RecordSource`). Las citas son `Form_frmSplash.cls:línea` salvo indicación; `FU.bas` abrevia `modules/FUNCIONES UTILES.bas` y `VG.bas`, `modules/Variables Globales.bas`.

## Entrada y salida

- **Se abre desde:** no consta en `src` (propiedad de arranque de la base). `Avance` lo localiza por nombre (`FU.bas:363-364`).
- **Abre:** Form_frm0Ppal (rol Administrador o Calidad) o Form_frm0PpalTecnico (rol Técnico) y se cierra (`:57-66`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| Formulario | — | `OnOpen` | Oculta la cinta salvo con `/dev`, ajusta tamaño, reinicia el contador y vacía la barra | `:11-36`, `.form.txt:41` |
| Formulario | — | `OnTimer` (`TimerInterval = 50`) | Arranque completo (ver Flujos) | `:39-78`, `.form.txt:28, 376` |
| `lblEstado` | etiqueta | — | `Avance` escribe en ella el paso en curso | `FU.bas:356-359`, `.form.txt:1015-1016` |
| `lblProgresoFondo`, `lblProgresoBarra` | etiquetas | — | Ancho de la barra = pasos / elementos del `Entorno` | `FU.bas:369-386`, `.form.txt:26447, 26465` |

## Flujos

### Normal

1. `Form_Open`: si la línea de comandos no contiene `/dev`, `GestionarRibbon(False)` oculta la cinta (`:17-18`; `FU.bas:1024-1033`); `Ajustar Me` (`:21`; `FU.bas:325-339`); `ResetAvanceCounter` pone el contador a 0 (`:26`; `FU.bas:346-350`); barra a ancho 0 (`:30`).
2. `Form_Timer` desactiva el temporizador (`:45`) y llama a `EVE` (`:48`).
3. `EVE` (`VG.bas:56-152`): crea un `Workspace` DAO propio (`VG.bas:68`); reinicia `TempVars` y fija `EnDesarrollo = "Sí"` (`VG.bas:74-82`); valida la conexión a Lanzadera (`VG.bas:93-97`); identifica al usuario por la línea de comandos como correo o, si está vacía, por el usuario de Windows (`VG.bas:103-108`); asigna usuario activo, real y rol (`VG.bas:115-118`); recorre las propiedades del `Entorno` y falla si alguna da error (`VG.bas:122-144`).
4. Si `EVE` tarda menos de 2 s, espera hasta completarlos (`:40, 53-55`).
5. Abre el menú por rol y se cierra (`:57-66`).

### Alternativos

- Con `/dev` en la línea de comandos la cinta queda visible (`:17`).
- El rol sale de `DeterminarRol`: `EsAdministrador = "Sí"` ⇒ Administrador; sin permisos de la aplicación ⇒ Técnico; `EsUsuarioCalidad = "Sí"` ⇒ Calidad; resto ⇒ Técnico (`classes/UsuarioServicio.cls:86-100`).

### Errores

Cualquier error oculta el formulario, muestra «Se ha producido un error crítico durante el arranque:» con `FullDescription` y cierra las conexiones globales (`:69-76`; `FU.bas:264-274`). `Application.Quit` está comentado (`:77`). Mensajes de `EVE`: usuario no registrado (`VG.bas:112`), sin conexión a Lanzadera (`VG.bas:96`) y error de configuración (`VG.bas:142-143`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Usuario por correo o usuario de red | `TbUsuariosAplicaciones` (Lanzadera) | `QueryDef` parametrizada | `modules/UsuarioRepositorio.bas:34-42` |
| Permisos por aplicación | `TbUsuariosAplicacionesPermisos` (Lanzadera) | `QueryDef` parametrizada | `classes/UsuarioServicio.cls:29`; `modules/AplicacionRepositorio.bas:24-27` |
| Configuración del entorno | fichero INI y propiedades de `Entorno` | API de perfil | `VG.bas:128-131, 154-158` |

Sin escrituras en tablas ni transacciones.

## Coexistencia

`EVE` crea un `Workspace` Jet con el usuario `admin` (`VG.bas:68`): con tablas PostgreSQL vinculadas debe verificarse que las conexiones ODBC se abren dentro de ese espacio. `getdbLanzadera` es punto único de fallo del arranque (`VG.bas:94-97`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `/dev` en la línea de comandos se pasa también a `EVE` como correo (`:17`; `VG.bas:103-105`): no habría usuario con ese correo y el arranque fallaría en `VG.bas:112`.
- Un usuario de Lanzadera sin permisos de la aplicación recibe rol Técnico (`classes/UsuarioServicio.cls:91-93`) y entra.
- La rama «Rol de usuario desconocido» (`:62-63`) es inalcanzable: `DeterminarRol` siempre devuelve 1, 2 o 3.
- Tras un error crítico Access queda abierto sin formulario visible (`:70, 77`).
- `EnDesarrollo = "Sí"` fijo en cada arranque (`VG.bas:75`).
- `Avance` crea un `Entorno` nuevo en cada paso para contar elementos (`FU.bas:368-370`).
- `ProgresoActual` no se usa (`:9`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: sustituir el arranque por la sesión autenticada y la redirección por rol.

## Criterios de aceptación

Propuestos, no aprobados:

1. Un usuario Administrador o Calidad llega al menú general; un Técnico, al menú técnico.
2. Un usuario no registrado ve un aviso y no accede a ningún menú.
3. Si falla la conexión a usuarios o la configuración, el usuario ve un error y no queda una sesión a medias.

## Pendientes

- Propiedad de arranque de la base: no está en `src`.
- Propiedades de `Entorno` recorridas por `EVE` (`ColItems`) y cuáles fallan sin INI.
- `src/forms/frmSplash.form.txt` duplica el diseño y el código con otro nombre; falta decidir cuál rige.
