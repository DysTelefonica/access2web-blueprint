# Form_FormExpedienteAltaTipo

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/forms/Form_FormExpedienteAltaTipo.cls` (363 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Selector del tipo de alta: muestra una lista de tipos de expediente, prepara un DTO vacío o copiado del expediente padre y abre el formulario de alta correspondiente. Las citas son `Form_FormExpedienteAltaTipo.cls:línea` salvo indicación.

## Entrada y salida

- **Apertura:** `Form_Open` (`:249-303`). El argumento `OpenArgs` admite tres formas: `"IDExpediente|TipoA;TipoB"` (expediente activo y tipos permitidos, `:271-274`), `"DesdeInicio"` (sin padre ni expediente activo, `:276-279`) o vacío, caso en el que el padre es la global `m_ObjExpedienteActivo` (`:269`).
- **Salida normal:** `ComandoAlta_Click` abre `FormExpedienteAltaParaHPS` o `FormExpediente` (`:99, 181, 204`). Cuando `FormExpediente` lanza `AltaExpediente`, este formulario relanza `Alta` hacia quien lo abrió y se cierra (`:343-351`).
- **Salida manual:** `CmdSalir_Click` cierra sin guardar (`:15-19`); `ComandoAyuda_Click` abre la ayuda (`:225-247`).
- **Quién lo abre:** pendiente de trazar (no se ha buscado en esta tanda).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `ListaTipoAlta` | lista de 2 columnas (valor «Value List») | `OnClick` | Habilita `ComandoAlta` si hay una opción seleccionada | `:305-319`, `.form.txt:631-651` |
| `ListaTipoAlta` | — | `OnDblClick` | Equivale a pulsar `ComandoAlta` si está habilitado | `:335-341` |
| `ComandoAlta` | botón («Alta») | `OnClick` | Construye el DTO y abre el formulario de alta | `:66-221`, `.form.txt:778-780` |
| `ComandoAyuda` | botón (rótulo «Salir» en `.form.txt`) | `OnClick` | `AbrirAyuda` | `:225-247`, `.form.txt:551-553` |
| `cmdSalir` | botón | `OnClick` | Cierra | `:15-19`, `.form.txt:689-691` |

## Flujos

### Carga de la lista (`Form_Open` → `EstablecerLista`, `:21-64`)

1. `ComandoAlta` y `AllowEdits` parten deshabilitados; solo `EsAdministrador = Sí` los habilita (`:261-268`).
2. Sin tipos en el argumento, la lista contiene los tipos no derivados de `m_ObjEntorno.ColTiposExpedientesNoDerivados`, con el texto «Alta de <tipo>» (`:35-43`).
3. Con tipos en el argumento, la lista contiene exactamente esos tipos (`:44-51`). Si queda un único elemento, lo selecciona, ejecuta el alta y cierra el formulario (`:52-56`).

### Alta (`ComandoAlta_Click`, `:66-221`)

1. Sin opción seleccionada: error «No se ha seleccionado ninguna opción» (`:81-84`). Limpia `m_ObjExpedienteActivo` y `m_ObjExpedienteAlInicio` (`:86-87`).
2. **Tipo HPS** (`EnumTipoExpediente.EXPHPS`): DTO nuevo con el tipo y `SetDatosDeExpedienteTipo`, y abre `FormExpedienteAltaParaHPS` (`:88-108`).
3. **Con expediente padre** (alta de lote o de basado): copia el DTO del padre con `EstablecerDTODeExpediente(..., p_ParaCopia:=Sí)` (`:110-112`) y aplica:
   - Vacía `IDExpediente`, `Observaciones`, fechas de creación y último cambio, usuarios, fechas de contrato y garantía, `CodExp`, `CodExpLargo`, `ImporteContratacion`, `ImporteLicitacion`, `ESTADO` (`:117-130, 164`).
   - Conserva `IDResponsableSeguridad`, `ObjetoContrato` y, si es numérico, `IDOrganoContratacion` del padre (`:132-133, 161-163`); `Nemotecnico` = `NemotecnicoCalculado` del padre (`:127`).
   - `Ordinal`: «01» si el padre no tiene ningún derivado del mismo tipo; si lo tiene y su ordinal es numérico, el siguiente con dos dígitos (`:134-159`).
   - `EsAM = "No"`; `EsLote = "Sí"`/`EsBasado = "No"` para lote y al contrario para basado; `EsExpediente = "No"`; `POSTAGEDO = "Sí"`; `TIpo = TipoCalculadoTexto` (`:165-175`).
   - Abre `FormExpediente` (`:178-184`).
4. **Sin padre**: DTO nuevo con el tipo y `SetDatosDeExpedienteTipo`, y abre `FormExpediente` (`:194-207`).

### Error (`:213-220`)

Mensaje en `MsgBox` crítico; si no es el error controlado 1000, envía correo al administrador (`CorreoAlAdministrador`).

## Reglas, validaciones y cálculos

- Solo un administrador puede dar de alta (`:265-268`).
- El tipo de alta disponible depende del contexto: sin padre, los no derivados; con padre, los que indique el llamador (`:35-51`). El catálogo de tipos viene de `m_ObjEntorno`, sin trazar aquí.
- La numeración `Ordinal` de lotes y basados la calcula este formulario, no `Registrar`; `Registrar` solo formatea el ordinal numérico a dos dígitos (`ExpedienteOperaciones.cls:67`).

## Lecturas y escrituras

No escribe en base de datos. Lee `m_ObjEntorno.ColTiposExpedientes*`, `constructor.getExpediente(IDExpediente)` y `UltimoLoteDerivado`/`UltimoBasadoDerivado` del expediente activo (`:135-147`). La escritura ocurre en `FormExpediente`/`FormExpedienteAltaParaHPS`.

## Coexistencia

Pendiente de `coexistencia.md` (tanda posterior). Este formulario no tiene SQL propio; sus dependencias son la carga de catálogos y del expediente padre.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Con un único tipo permitido, `EstablecerLista` invoca `ComandoAlta_Click` y después `DoCmd.Close` (`:52-56`) dentro de `Form_Open`; `ComandoAlta_Click` ya había abierto el formulario de alta, y el cierre del selector deja al alta sin quien escuche su evento `Alta` (el `WithEvents` vive en este formulario). Pendiente de comprobar en ejecución.
- `ComandoAlta_Click` no vuelve a comprobar `EsAdministrador`: depende de que el botón esté deshabilitado, pero `EstablecerLista` lo invoca directamente (`:52-54`) aunque el usuario no sea administrador.
- `m_FormExpediente_ParaHPS` se declara con `WithEvents` (`:11`) pero el único manejador del fichero es `m_FormExpediente_AltaExpediente` (`:343`); el alta de HPS no reenvía `Alta`.
- `ComandoAyuda` rotula «Salir» en el `.form.txt` (`:552`) y la cabecera dice «Versión 2024-001» (`:39`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: en la web el tipo de alta puede ser un parámetro de ruta, sin formulario intermedio cuando solo hay uno permitido.

## Criterios de aceptación

Propuestos, no aprobados:

1. Un administrador ve «Alta de <tipo>» para cada tipo permitido en su contexto; un no administrador no puede iniciar el alta.
2. Con un único tipo permitido se salta la selección.
3. Una alta derivada hereda del padre solo los campos listados y asigna el siguiente `Ordinal` de dos dígitos.
4. Elegir HPS abre el alta específica de HPS en lugar del formulario general.

## Pendientes

- Llamadores de este formulario y valores reales de `OpenArgs`.
- Contenido de `SetDatosDeExpedienteTipo`, `EstablecerDTODeExpediente` y de `m_ObjEntorno.ColTiposExpedientes*`.
- Comportamiento real del cierre con un único tipo.
