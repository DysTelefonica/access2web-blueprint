# Form_frmBuscarExpediente

> **Commit fijado:** `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · **Fuente:** `src/forms/Form_frmBuscarExpediente.cls` (242 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Buscador de expedientes con PECAL: filtra por palabra clave, responsable técnico y responsable de calidad, muestra el resultado en una lista y devuelve el expediente elegido al formulario que lo abrió en diálogo. Es **no enlazado** (sin `RecordSource` en el `.form.txt`). Las citas son `:línea` de `Form_frmBuscarExpediente.cls` salvo indicación. Abreviaturas: `ES.cls` = `classes/ExpedienteServicio.cls`, `US.cls` = `classes/UsuarioServicio.cls`, `FU.bas` = `modules/FUNCIONES UTILES.bas`.

## Entrada y salida

| Desde | Modo | Evidencia |
|---|---|---|
| `Form_frmAltaSolicitud`, búsqueda rápida con varios resultados | `acDialog`, `OpenArgs` = palabra clave | `Form_frmAltaSolicitud.cls:274-275` |
| `Form_frmAltaSolicitud`, «Búsqueda Avanzada» | `acDialog`, sin `OpenArgs` | `Form_frmAltaSolicitud.cls:312` |
| `Form_frm0OtrosAdmin`, `cmdBuscarExpedientes` | ventana normal, sin `OpenArgs` | `Form_frm0OtrosAdmin.cls:22-25` |

**Devolución:** al elegir, escribe el id en `Me.Tag` y se oculta (`:184, 188`); al ocultarse termina el diálogo, el llamador lee `Tag` y cierra el formulario (`Form_frmAltaSolicitud.cls:278-294`). Abre `frmDetalleExpediente` en diálogo (`:228`). Sin control de rol.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `txtPalabraClave` | texto | `AfterUpdate` | Relanza la búsqueda | `:48-50`, `.form.txt:728` |
| `cboResponsableTecnico` | combo, lista de valores | `AfterUpdate` | Filtro; opción `0` «(Todos)» | `:44-46, 98`; `FU.bas:554-555` |
| `cboResponsableCalidad` | combo, lista de valores | `AfterUpdate` | Filtro; opción `0` «(Todos)» | `:40-42, 99` |
| `cmdLimpiarPalabraClave`, `cmdLimpiarRespTecnico`, `cmdLimpiarRespCalidad` | botones | `OnClick` | Vacían el filtro y relanzan la búsqueda | `:60-73` |
| `lstResultados` | lista, 7 columnas, id oculto | `DblClick` | Elige el expediente | `:17-22, 56-58`, `.form.txt:641-658` |
| `cmdElegir` («Elegir») | botón | `OnClick` | Elige el expediente | `:52-54` |
| `cmdVerDetalle` («Ver detalle») | botón | `OnClick` | Abre el detalle del seleccionado | `:203-240` |
| `cmdSalir` («Salir») | botón | `OnClick` | Vacía `Tag` y cierra | `:76-87` |
| `lblBloqueo` | etiqueta oculta | — | Aviso de espera | `:120, 156`, `.form.txt:1420` |

Columnas: `IDExpediente` (ancho 0), código, nemotécnico, inicio, fin, responsable técnico y RAC (`:21, 132`). `ControlBox` y `CloseButton` con valor no predeterminado (`.form.txt:10, 13`; a verificar en el binario). El evento `ExpedienteElegido` está declarado y no se lanza (`:9`).

## Flujos

### Apertura (`Form_Load`, `:11-38`)

1. `Ajustar Me`, título «Búsqueda de Expedientes Registrados con PECAL» y configuración de la lista (`:14-22`).
2. Si hay `OpenArgs`, lo copia a la palabra clave (`:24-26`).
3. Carga los combos desde la caché del `Entorno` (`:90-104`; `classes/Entorno.cls:118-153`) y busca (`:29`).

### Búsqueda (`RealizarBusqueda`, `:111-163`)

1. Vacía la lista y lee los tres filtros; combo vacío equivale a `0` (`:124-128`).
2. `getExpedientes` (`ES.cls:18-87`): siempre `Pecal = 'Sí'`; palabra clave como `*texto*` contra nemotécnico o código; responsable de calidad por `IDResponsableCalidad`; responsable técnico por `EXISTS` sobre jefes de proyecto; orden por id descendente (`ES.cls:53-77`).
3. Construye la lista de valores con una primera fila de cabecera como dato (`:132`, `ColumnHeads = False` en `:20`); fechas a cero se muestran vacías y responsables nulos como «N/A» (`:141-147`).

### Elección y detalle

- **Elegir** (`:170-197`): ignora la fila 0 (cabecera) o sin selección (`:175`) y el id `0` (`:179`); guarda el id en `Tag` y se oculta.
- **Ver detalle** (`:203-240`): sin selección, «Por favor, seleccione un expediente de la lista para ver su detalle.» (`:213-216`); id vacío, «No se pudo obtener el identificador del expediente seleccionado.» (`:220-223`).
- **Salir** (`:76-87`): con el formulario cerrado, el llamador no lo encuentra y no carga nada (`Form_frmAltaSolicitud.cls:280-288`).

### Errores

Todos los procedimientos usan `ManejadorErroresFormulario` (mensaje por tipo y registro en `tbLogErrores`: `FU.bas:621-636, 789-798`); un error en `Form_Load` cierra el formulario (`:34-37`).

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| Solo expedientes con `Pecal = 'Sí'` | `getExpedientes` | `ES.cls:53` |
| Responsable técnico = primer jefe de proyecto (`FIRST`) | subconsulta | `ES.cls:37-39` |
| RAC asignado = primer RAC del expediente (`FIRST`) | subconsulta | `ES.cls:40-42` |
| Técnicos del combo: jefes de proyecto de expedientes PECAL sin fecha de baja | `getResponsablesTecnicos` | `US.cls:116-121` |
| Calidad del combo: usuarios sin baja con `EsUsuarioCalidad = 'Sí'` en la aplicación, desde Lanzadera | `getResponsablesCalidad` | `US.cls:134-141` |

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Expedientes | `TbExpedientes`, `TbExpedientesResponsables`, `TbUsuariosAplicaciones`, `TbExpedientesRACS`, `TbRACS` (backend de expedientes) | SQL dinámico parametrizado | No | `ES.cls:35-79` |
| Responsables técnicos | backend de expedientes | SQL dinámico | No | `US.cls:116-121` |
| Responsables de calidad | `TbUsuariosAplicaciones`, `TbUsuariosAplicacionesPermisos` (Lanzadera) | SQL dinámico | No | `US.cls:134-141` |

Sin escrituras propias.

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: `FIRST()` sin equivalente directo (`ES.cls:37, 40`); comodín `*` en `LIKE` (`ES.cls:57`); booleanos como texto `'Sí'` (`ES.cls:39, 53`); falta confirmar si la búsqueda por palabra clave debe seguir sin distinguir mayúsculas, como en Jet.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Abierto desde `frm0OtrosAdmin` (sin diálogo), «Elegir» oculta el formulario y nadie lo cierra ni lee el `Tag` (`:188`; `Form_frm0OtrosAdmin.cls:25`).
- La cabecera es una fila de datos seleccionable (`:20, 132`).
- La lista de valores concatena con `;` sin escapar (`:141-147`): un nemotécnico con `;` desplaza columnas; con muchos expedientes podría superarse el límite de longitud de `RowSource` (a verificar).
- Los combos de calidad salen de Lanzadera (`US.cls:141`) y el filtro compara contra `TbUsuariosAplicaciones` del backend de expedientes (`ES.cls:36, 61`); si no es la misma tabla, los id no casarían.
- Cada cambio de filtro relanza la consulta completa (`:40-73`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: selector modal que devuelva el id por contrato explícito en vez de `Tag`; paginación del resultado.

## Criterios de aceptación

Propuestos, no aprobados:

1. Solo aparecen expedientes con PECAL, del más reciente al más antiguo.
2. La palabra clave ficticia `ABC` encuentra expedientes cuyo nemotécnico o código contiene `ABC`.
3. Los filtros de responsable se combinan con la palabra clave y «(Todos)» los desactiva.
4. Elegir un expediente desde el alta de solicitud lo carga en ese formulario; «Salir» no carga nada.
5. «Ver detalle» sin selección muestra el aviso y no abre el detalle.

## Pendientes

- Comportamiento desde `frm0OtrosAdmin` previsto por el negocio (consulta o selección).
- Columnas de `TbExpedientes` y relación de `TbUsuariosAplicaciones` entre backends (D-CON-5).
