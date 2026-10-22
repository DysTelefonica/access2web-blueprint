# Form_FormExpediente

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/forms/Form_FormExpediente.cls` (597 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Contenedor de la ficha de un expediente: cabecera, barra de pestañas (control de navegación), subformulario `FrmDetalle` que aloja cada pestaña y botones de guardado. Las citas son `Form_FormExpediente.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormExpedientesGestion` | `DoCmd.OpenForm "FormExpediente"` con `OpenArgs`; escucha sus eventos | `Form_FormExpedientesGestion.cls:814-819` |
| `Form_FormExpedienteAltaTipo` | abre el formulario con un DTO nuevo o copiado del padre | `Form_FormExpedienteAltaTipo.cls:178-184, 201-207` |
| `Form_FormTareas` | cierra el formulario si estaba abierto y lo abre | `Form_FormTareas.cls:181-186` |

Salidas: `CmdSalir_Click` cierra sin guardar (`:107-109`); `ComandoRegistrar_Click` cierra tras guardar (`:98`). Eventos hacia quien lo abrió: `AltaExpediente` (`:5`, lanzado en `:87`) y `ExpedienteEditado` (`:6`, lanzado en `:89`).

Estado compartido (variables globales de `Variables Globales.bas`): `m_ObjExpedienteDTOActivo` (`:265`), `m_ObjExpedienteAlInicio` (`:266`), `m_ObjExpedienteActivo` (`:255`), `EsAdministrador` (`:244`) y `m_TituloFormulario` (`:273`). El formulario no recibe el expediente por parámetro: lo lee de esas globales.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `FrmDetalle` | subformulario | — | Aloja la pestaña activa; se inicializa con `FormExpedienteGeneral` | `.cls:146` |
| `NavGeneral` | botón de navegación | `OnEnter` → `NavGeneral_Enter` | Fija `NavigationTargetName = "FormExpedienteGeneral"`; no vuelca el DTO | `.cls:442-468`, `.form.txt:722-738` |
| `NavFechas` | botón de navegación | `OnClick` → `NavFechas_Click` | Llama a `RellenarDTO` antes de cambiar de pestaña | `.cls:418-440` |
| `NavEntidades` | botón de navegación | `OnEnter` → `NavEntidades_Enter` | `RellenarDTO` | `.cls:470-492` |
| `NavSuministradores` (rótulo «Empresas») | botón de navegación | `OnEnter` | `RellenarDTO` | `.cls:576-597` |
| `NavAnexos` (rótulo «Anexos»; destino `FormExpedienteDocumentacion`) | botón de navegación | `OnEnter` | `RellenarDTO` | `.cls:391-413`, `.form.txt:940, 958` |
| `NavHitos`, `BotónDeNavegación69` («Modificados») | botón de navegación | sin handler en el `.cls` | Navegan a `FormExpedienteHitos` y `FormExpedienteModificados` sin volcar el DTO | `.form.txt:995-1012, 1049-1066` |
| `BotónDeNavegación1` («[Agregar nuevo]») | botón de navegación | sin handler en el `.cls` | Sin comportamiento identificado en esta ficha | `.form.txt:1103` |
| `ComandoRegistrar` (rótulo «Editar») | botón | `OnClick` → `ComandoRegistrar_Click` | Guarda el expediente | `.cls:55-106` |
| `ComandoActualizarCompleto` (rótulo «Editar») | botón | `OnClick` | Refresca la caché `TbExpedientesConEntidades` del expediente | `.cls:10-33` |
| `ComandoAyuda` (rótulo «Salir» en `.form.txt`) | botón | `OnClick` | Abre la ayuda (`AbrirAyuda`, `FUNCIONES UTILES.bas:3076`) | `.cls:35-51`, `.form.txt:610-612` |
| `cmdSalir` | botón | `OnClick` | Cierra el formulario sin guardar | `.cls:107-109` |
| `lblTitulo`, `lblModificacion` | etiquetas | — | Título del formulario y «última modificación» | `.cls:184-190` |

## Flujos

### Apertura (`Form_Load`, `.cls:111-160`)

1. `OpenArgs = "Sololectura"` deja `m_AbiertoParaEditar = False`; cualquier otro valor lo deja en `True` (`:121-125`).
2. Parte con `ComandoRegistrar` deshabilitado y `AllowEdits = False`; solo si se abre para editar **y** `EsAdministrador = Sí` los habilita (`:129-140`).
3. `EstablecerDatos` exige un DTO con `Expediente`, `TIpo` y `Derivable` informados; si falta cualquiera, error «No se sabe para qué tipo de registro se ha abierto» (`:162-182`). Con `IDExpediente` vacío oculta `lblModificacion` y `ComandoActualizarCompleto`; con él informado los muestra (`:186-197`).
4. Carga `FormExpedienteGeneral` en `FrmDetalle` (`:146`).

### Guardado (`ComandoRegistrar_Click`, `.cls:55-106`)

1. `RellenarDTO` copia al DTO los controles de la pestaña visible (`:65`).
2. `HaHabidoCambios` compara el expediente con `m_ObjExpedienteAlInicio`; si no hay cambios **y** el árbol de suministradores está vacío, aborta con «Ningún cambio que guardar» (`:68-76`). Una alta (sin expediente inicial) siempre cuenta como cambio (`:212-215`).
3. `ExpedienteOperaciones.Registrar` ejecuta el guardado transaccional (`:79-83`; detalle en [Lecturas y escrituras](#lecturas-y-escrituras)).
4. Éxito: lanza `AltaExpediente` si no había expediente inicial y `ExpedienteEditado` si lo había (`:86-90`), limpia las tres globales y cierra el formulario (`:92-98`).

### Alternativos y de error

- Cualquier fallo se acumula en `m_Error` y se muestra con `MsgBox` crítico (`:100-105`). `Form_Load` además envía correo al administrador (`CorreoAlAdministrador`, `:155`) y cierra el formulario (`:159`); el resto de handlers de pestaña también envían correo (p. ej. `:408-409`).
- `ComandoActualizarCompleto_Click` solo invoca `RegistrarExpEntidades`; no valida ni guarda la cabecera (`:16-22`).

## Reglas, validaciones y cálculos

- Permiso de edición: `m_AbiertoParaEditar` y `EsAdministrador = Sí` (`:132-136`). Las pestañas hijas consultan `Form_FormExpediente.m_AbiertoParaEditar` (p. ej. `Form_FormExpedienteFechas.cls:138`, `Form_FormExpedienteHitos.cls:175`). `EsAdministrador` procede de `Usuario.EsAdministradorCalculado` (`classes/Usuario.cls:59`); su asignación no se ha trazado.
- `RellenarDTO` solo vuelca dos pestañas: **General** (25 campos copiados con `Nz`, algunos combos leídos con `.Column(0)` y `POSTAGEDO` leído del control `EnListas`; además el combo `AGEDYSAplica` se traduce al par `AGEDYSAplica` + `AGEDYSGenerico` según `EnumAPLICAAGEDYS`) y **Fechas** (12 campos) (`:502-563`). Entidades, Empresas, Anexos, Hitos y Modificados no se vuelcan aquí: cada pestaña mantiene su propio estado. General y Fechas, además, se vuelcan solas al descargarse (`Form_Unload`, ver Posibles defectos).
- Reglas que aplica `ExpedienteOperaciones.Registrar` antes de escribir (`classes/ExpedienteOperaciones.cls`):
  - `POSTAGEDO` distinto de «Sí»/«No» pasa a «Sí» (`:57-59`).
  - `MotivoNoOK` valida y devuelve el motivo de rechazo (`:61-64`); su contenido no se analiza en esta ficha.
  - `Ambito = "HPS"` fuerza `HPSAplica = "Sí"` (`:66`); `Ordinal` numérico se formatea a dos dígitos (`:67`).
  - `ESTADO`, `FechaFinGarantia`, `TIpo` y `TipoInforme` se recalculan (`:68-69, 82-83`); `IDResponsableSeguridad = "0"` pasa a vacío (`:72`).
  - En alta, `IDExpediente = IDExpedienteCalculado` (`:75-77`). Las cadenas vacías se guardan como `Null` en todos los campos de cabecera (`:99-381`), salvo `ContratistaPrincipal`, que se guarda tal cual, incluso vacío (`:383-387`).

## Lecturas y escrituras

`ExpedienteOperaciones.Registrar` (`:19-575`) abre una transacción de espacio de trabajo DAO (`ws.BeginTrans`, `:47`) y escribe en la base de datos devuelta por `getdb()` (`:45`):

| Paso | Efecto | Evidencia |
|---|---|---|
| Cabecera | `AddNew`/`Edit` + `Update` de `TbExpedientes` (49 columnas asignadas) | `:76-92, 390` |
| Alta: NC | `ActualizarNCs` | `:395-397` |
| Alta: hijos | Anualidades, comerciales, códigos de compra, CPV, PECAL, RAC, anexos, responsables, lugares de ejecución y árbol de suministradores, desde las colecciones del DTO | `:402-495` |
| Alta: cadenas | `ActualizarCadenaPECAL` | `:497-498` |
| Caché | `RegistrarExpEntidades` con ámbito `Cabecera` dentro de la transacción | `:514` |
| Commit | `ws.CommitTrans` | `:520` |
| Post-commit | Hook `modNotifyChangeE2E.NotifyChange` con `On Error Resume Next`: su fallo no revierte ni notifica error | `:523-534` |
| Post-proceso | `PintarTareas` según cambios de estado o necesidad de `CodS4H`; `RegistrarUltimoCambio` | `:540-554` |
| Error | `Rollback` total | `:568-575` |

En edición, los hijos **no** se escriben desde `Registrar`: se asume que cada pestaña los persiste por su cuenta (no verificado aquí; lo confirman las fichas de las pestañas).

`RegistrarExpEntidades` pasa siempre a ámbito `Todo` aunque se pida otro (`ExpedienteOperaciones.cls:607`), invalida los derivados del expediente y actualiza el objeto global con `ActualizarExpedienteCompleto` (`:621`).

## Coexistencia

Pendiente de `coexistencia.md` (tanda posterior). Puntos que ya se ven: `BeginTrans`/`CommitTrans` de DAO sobre `TbExpedientes` y tablas hijas; `Recordset` con `AddNew`/`Edit`; `IDExpedienteCalculado` genera el identificador en VBA; el hook E2E y `PintarTareas` ocurren fuera de la transacción.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Los tres manejadores de error de `ComandoActualizarCompleto_Click`, `ComandoAyuda_Click` y `ComandoRegistrar_Click` escriben `Err.Numbr` (`:29, 47, 102`), no `Err.Number`. Si compila, el mensaje no llevaría el número de error. Pendiente de comprobar en el binario.
- `HaHabidoCambios` no compara `ContratistaPrincipal`, `EsAM`, `EsLote`, `EsBasado` ni `EsExpediente`, que `Registrar` sí persiste (`:216-379`; `Registrar` `:161-181, 387`), y compara `ObjetoContrato` dos veces (`:259, 315`).
- `HaHabidoCambios` no detecta cambios en hijos (entidades, hitos, etc.) y el aviso «Ningún cambio que guardar» solo mira el árbol de suministradores (`:71-76`).
- `ComandoRegistrar` y `ComandoActualizarCompleto` comparten rótulo «Editar» y `ComandoAyuda` rotula «Salir» en el `.form.txt`; no se ha comprobado si el código los cambia en ejecución.
- Las pestañas General y Fechas duplican el volcado: cada una tiene su propio `RellenarDTO` (`Form_FormExpedienteGeneral.cls:846`, `Form_FormExpedienteFechas.cls:298`) y lo ejecuta en `Form_Unload` (`Form_FormExpedienteGeneral.cls:580-587`, `Form_FormExpedienteFechas.cls:238-245`), además del `RellenarDTO` del contenedor. El resultado depende de que Access dispare `Unload` al cambiar de pestaña, lo que está sin verificar. Hitos, Modificados y Anexos no vuelcan campos del expediente, así que `NavHitos` y `Modificados` sin handler no pierden datos de cabecera por sí mismos.
- El correo al administrador sale en cada error de navegación entre pestañas (p. ej. `:408-409`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: guardado explícito por pestaña frente a guardado único; sustituir las globales de sesión por el expediente como parámetro de la ruta.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir un expediente existente en solo lectura deja deshabilitado el guardado; abrirlo para editar sin ser administrador también.
2. Guardar sin cambios devuelve «Ningún cambio que guardar» y no escribe.
3. Guardar un alta crea cabecera e hijos en una sola transacción: si falla un hijo, no queda cabecera.
4. Guardar una edición lanza `ExpedienteEditado`; una alta, `AltaExpediente`.
5. Un fallo del hook E2E tras el commit no revierte ni oculta el guardado.

## Pendientes

- Contenido de `MotivoNoOK` (`ExpedienteOperaciones.cls:800`) y de `IDExpedienteCalculado`, `ESTADOCalculadoTexto`, `TipoCalculadoTexto`.
- Quién asigna `EsAdministrador` y `m_TituloFormulario`.
- Comportamiento de «[Agregar nuevo]» y de los rótulos en ejecución.
- Persistencia de hijos en edición (fichas de las pestañas).
- Referencias dinámicas: `Forms("FormExpediente").FrmDetalle.Form` desde `Form_FormExpedientesResponsables.cls:155`.
