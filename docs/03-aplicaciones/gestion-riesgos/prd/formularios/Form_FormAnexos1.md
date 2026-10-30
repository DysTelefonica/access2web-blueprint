# Form_FormAnexos1

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Fuente:** `src/forms/Form_FormAnexos1.cls` (614 líneas) y `src/forms/FormAnexos1.form.txt` · **PRD:** [README](../README.md)

Predecesor del gestor de anexos. Comparte casi todo con `Form_FormAnexos`, pero no lo abre nadie en el código fuente y su módulo **no compila tal como está exportado**. Se documenta aparte porque conserva decisiones que el sucesor descartó (validación de contexto, orden de la sustitución UTE, correo al administrador ante errores) y porque es la única forma de leer el diseño anterior. Las citas son `Form_FormAnexos1.cls:línea` salvo indicación.

## Entrada y salida

Acepta la misma palabra por `openArgs` (`:405`) y, a diferencia del sucesor, **valida que el contexto global exista** y que el argumento sea uno de los tres.

| Ámbito | `openArgs` | Objeto de contexto | Evidencia |
|---|---|---|---|
| Proyecto | `Proyecto` | `m_ObjProyectoActivo` | `:407-417` |
| Edición | `Edición` | `m_ObjEdicionActiva` | `:418-435` |
| Riesgo | `Riesgo` | `m_ObjRiesgoActivo` | `:436-448` |
| Cualquier otro | — | error «parámetros insuficientes» | `:449-451` |

| Quién lo abre | Evidencia |
|---|---|
| **Nadie en `src`** | No hay ninguna referencia a `FormAnexos1` fuera de su propia clase |

Solo una apertura manual o una macro no exportada lo pondría en marcha. Es el punto de partida de la ficha hermana, que reimplementa la misma idea con más controles.

## Diferencias con `Form_FormAnexos`

| Aspecto | `Form_FormAnexos1` | `Form_FormAnexos` |
|---|---|---|
| Botones de alta | Uno solo, `ComandoAnexar`: valida, elige fichero y persiste | Tres: `btnExaminar`, `btnAceptar`, `btnActualizarNombreArchivo` (`.form.txt:971, 972, 1340, 1439`) |
| Selección de fichero | `Seleccionar` (`Funciones Generales.bas:1101-1143`), que recuerda el último archivo | `Application.FileDialog` propio (`:73`) |
| Sustitución UTE | Borra la anterior **después** de registrar (`:217-226`) | Borra **antes** (`:162-166` frente a `:173`) |
| Duplicados y renombrado | Ni una cosa ni la otra | `ExisteNombreEnAmbito` (`:243-267`) y renombrado (`:199-231`) |
| Ayuda | `ComandoAyuda_Click` con `AbrirAyuda` (`:255-281`) | Sin manejador |
| Errores inesperados | `CorreoAlAdministrador` en todos los manejadores | Solo `MsgBox` |
| Título del formulario | `m_ObjEntorno.TituloUsuarioConectado` (`:479`) | Fijo en el diseño (`FormAnexos.form.txt:40`) |

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `Form_Open` | Hourglass, contexto con validación, permiso, listado y rótulo | `:384-494` |
| `ComandoAnexar` | botón | `Click` | Exige título, elige fichero, persiste y recarga | `:139-249`, `.form.txt:971` |
| `TituloDoc` | texto | — | Descripción del anexo; se bloquea tras anexar evidencia UTE (`:160`) | `.form.txt:820` |
| `EvidenciaUTE` | opción | — | Visible solo si el proyecto está en UTE (`:360-382`) | `.form.txt:915` |
| `ListaDocumentos` | lista | `Click` / `DblClick` | Selecciona y resuelve el anexo; doble clic abre el fichero | `:527-600`, `:606-613`, `.form.txt:783` |
| `ImagenAnexo` | imagen | `Click` | Abre con `AbrirEnLocal` y avisa si no hay selección | `:496-525`, `.form.txt:628` |
| `ComandoEliminarAnexo` | botón | `Click` | Confirma, borra y recarga el padre | `:283-359`, `.form.txt:1304` |
| `ComandoAyuda` | botón | `Click` | Abre la ayuda de la aplicación | `:255-281`, `.form.txt:700` |
| `cmdSalir` | botón | `Click` | Cierra sin guardar | `:250-253`, `.form.txt:1349` |

## Flujos

### Apertura (`Form_Open`, `:384-494`)

1. Hourglass, `AjustarTamaño Me`, `AllowEdits = True` y deshabilita las tres acciones (`:392-399`).
2. **Comprueba que el objeto global exista** en cada rama y, si falta, fija un mensaje concreto («No se conoce el proyecto activo», `:408-410`) (`:407-448`).
3. Rama `else` con «Se ha abierto el formulario con parámmetros insuficientes» (`:449-451`, con errata en el original).
4. Comprueba `EsActivo` y, en ambos caminos, **propaga `m_ObjDELAnexo.Error`** (`:453-467`); habilita `ComandoAnexar` solo si está activo.
5. `AdaptarTamañoFormulario` (`:468-471`) y `EstablecerLista` (`:472`), rótulo y título de ventana (`:477-479`).

El manejador distingue error propio (1000, solo `MsgBox`) de error inesperado (correo al administrador y `MsgBox` crítico) y **cierra siempre** (`:486-493`).

### Flujo normal: anexar (`ComandoAnexar_Click`, `:139-249`)

1. Sin edición conocida y con ámbito de edición → error «No se conoce la edición» (`:146-150`).
2. Título obligatorio (`:163`, `:169`); con evidencia UTE **se genera el título y `TituloDoc` se deshabilita** (`:156-160`).
3. `Seleccionar(True, …)` abre el diálogo y devuelve la ruta (`:174`); si se cancela, sale sin escribir (`:176-182`).
4. Asigna ámbito y, con evidencia UTE, **pregunta si se sustituye** pero no borra todavía (`:190-205`).
5. `Registrar` (`:217`) y **después** borra la evidencia UTE anterior (`:222-226`).
6. `RaiseEvent AnexoAñadido` (`:227`) y recarga (`:228`).

### Flujo de error: abrir el anexo (`ImagenAnexo_Click`, `:496-525`)

Si no hay selección, mensaje «Seleccione un elemento de la lista» (`:504-507`). `AbrirEnLocal` informa por `m_Error` (`:508`) y, si vuelve con mensaje, se lanza `Err.Raise 100` (`:510`).

### Baja (`ComandoEliminarAnexo_Click`, `:283-359`)

Si no hay selección en memoria, la resuelve por la lista (`:290-302`); pide confirmación (`:304-307`), llama a `EliminarAnexo` (`:312`), recarga el objeto padre según el ámbito (`:318-343`) y vuelve a listar (`:342`).

## Reglas, validaciones y cálculos

- **El alta exige título siempre** (`:161-170`), sin la excepción de la evidencia UTE que existe en el sucesor: aquí la evidencia UTE se salta la validación porque el título se genera solo.
- **El permiso sigue siendo `EsActivo`**, no un rol (`:453-467`), y la lista usa la misma `RowSource` de seis columnas que el sucesor (`:75`): `Tipo;IDAnexo;Título;Riesgo;Edic_Anexado;Fecha`. Con proyecto recorre `ColAnexosTotales`, con edición o riesgo `ColAnexos` (`:77-81`).
- **La columna de edición anexada tiene dos lecturas**: la edición del anexo (`:96-107`) y, para anexos de tipo riesgo, la edición en la que nació el riesgo (`:120`).

## Lecturas y escrituras

| Operación | Origen | Evidencia |
|---|---|---|
| Ruta de origen | Diálogo `Seleccionar`, que recuerda el último archivo del usuario | `:174`, `Funciones Generales.bas:1135-1137` |
| Ruta de destino | `URLDirectorioDocumentacion` | `Anexo.cls:330-332` |
| Alta de fila y copia | `Anexo.Registrar` | `:217`, `Anexo.cls:376-407` |
| Baja de fila y fichero | `Anexo.EliminarAnexo` | `:312`, `Anexo.cls:497-502` |
| Apertura | Copia local y `Shell open` | `:508`, `Funciones Generales.bas:874-887` |
| Evidencia UTE vigente | `Constructor.getAnexoEvidenciaUTE` | `Edicion.cls:1103-1122`, `Constructor.bas:2659-2685` |

## Coexistencia

Pendiente de `../coexistencia.md`. Es la misma dependencia que el sucesor: repositorio de ficheros compartido y comprobación con `fso.FileExists`. Su valor documental está en el orden de la sustitución UTE y en el envío de correo al administrador ante errores, dos decisiones que conviene entender antes de decidir qué se conserva.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-GR-71.** La variable `pregunta` se usa en seis procedimientos (`:199, 243, 245, 277, 279, 304, 355, 357, 489, 491, 521, 523, 596, 598`) y **no está declarada en ninguno**: el módulo tiene `Option Explicit` (`:11`). Tal como está exportado, no compila. Como ningún código de `src` abre el formulario, Access solo lo compilaría al abrirlo, y el hallazgo está pendiente de verificar en el binario.
- **H-GR-72.** Ningún punto de `src` abre `FormAnexos1`: el formulario es código muerto y su única vía de entrada posible es una macro o una apertura manual que no están en el código fuente. El sucesor (`Form_FormRiesgosGestion.cls:686`) ya cubre los mismos ámbitos.
- **H-GR-73.** La búsqueda de respaldo para borrar usa `Me.ListaDocumentos.Column(0)` (`:291`), que es la columna **Tipo** de la `RowSource` de `:75`, mientras la selección usa `Column(1)`, que es `IDAnexo` (`:541`). Con el valor «R» o «E» como identificador, la consulta de `Constructor.getAnexo` (`Constructor.bas:2634`) no devuelve nada.
- **H-GR-74.** `Me.ListaDocumentos.RemoveItem (Me.ListaDocumentos.Column(0))` (`:565`) pasa a `RemoveItem` el **valor** de la primera columna en lugar del índice del elemento: o falla o borra la fila equivocada.
- **H-GR-75.** `ImagenAnexo_Click` lanza `Err.Raise 100` (`:510`) cuando `AbrirEnLocal` devuelve un mensaje de negocio, pero el manejador solo trata como propio el error **1000** (`:521`). Un anexo inaccesible o ya abierto se trata como error inesperado: se manda correo al administrador y sale un `MsgBox` crítico. El sucesor sí distingue (`:480-482`).
- **H-GR-76.** Aquí la sustitución UTE borra la evidencia anterior **después** de registrar la nueva (`:217-226`): si el borrado falla quedan dos evidencias UTE en la misma edición, y `getAnexoEvidenciaUTE` (`Constructor.bas:2678`, `ORDER BY IDAnexo DESC`) elige una sin avisar. El orden contrario del sucesor tiene el defecto simétrico y peor (`Form_FormAnexos.cls:162-166`).
- **H-GR-77.** `EstablecerLista` lee `m_ObjAnexo.riesgo.EdicionEnLaQueNace.Edicion` (`:120`) sin comprobar nada, a diferencia de la rama de edición, que valida `.Error` (`:98`). Si el riesgo no resuelve, la lista se queda a medias.
- **H-GR-78.** `Err.Raise 10000` (`:470`) solo se usa para «`AdaptarTamañoFormulario` ha fallado», un problema de maquetación. Como no es 1000, el manejador lo trata como error inesperado y **manda correo al administrador** (`:486-489`).
- **H-GR-79.** `Me.TituloDoc.Enabled = False` (`:160`) no se vuelve a activar en ningún punto del módulo: tras anexar una evidencia UTE, el usuario ya no puede escribir otro título en esa sesión.
- **H-GR-80.** No hay control de títulos duplicados ni renombrado, y `Seleccionar` recuerda el último archivo del usuario en el entorno (`Funciones Generales.bas:1135-1137`): quedan dos rutas de un fichero local en el estado de sesión.
- **H-GR-81.** `EvidenciaSuministrador` no se fija en ninguno de los dos formularios; `Anexo.MotivoNoOK` lo normaliza a «No» (`Anexo.cls:322-323`). La evidencia de suministrador se gestiona en `Form_FormPublicacionUTE.cls:43` y solo llega a la lista a través de `EdicionSuministrador.Anexo` (`EdicionSuministrador.cls:78-97`, con su `URLAnexo` en `:163-193`): desde aquí no se puede renombrar ni borrar. `AnexoAntiguo` es una clase de solo lectura de 267 líneas sin `Registrar` ni `EliminarAnexo`, y solo sobrevive en una rama `TypeOf` (`Funciones Generales.bas:1273`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: decidir explícitamente si el formulario se retira del binario o se reconstruye; si se conserva, reutilizar sus validaciones de contexto y su política de correo, y descartar su orden de sustitución UTE; subida por HTTP en lugar de selección de fichero local.

## Criterios de aceptación

Propuestos, no aprobados:

1. El formulario solo se abre con un ámbito válido y con el objeto global disponible; si no, avisa del motivo.
2. Un contexto inactivo deja el alta deshabilitada.
3. Un anexo no alcanzable o ya abierto produce un mensaje al usuario sin correo al administrador.

## Pendientes

- Si el binario recompila este módulo y, por tanto, si `pregunta` está declarada en algún sitio que el exportador no recogió.
- Si alguna macro del binario lo abre, ya que en `src` no hay ninguna referencia.