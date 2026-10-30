# Form_FormFacturaResultado

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormFacturaResultado.cls` (337 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Resumen de solo lectura de una factura («RESULTADO DE LA FACTURA», `Form_FormFacturaResultado.form.txt:351`): fechas, importes, impuestos y desglose por perfil. Es un formulario **no enlazado**: toma la factura de la global `m_ObjFacturaActiva` (`Variables Globales.bas:249`) y, según su bandera `ParaPorFacturar`, la trata como **simulacro** (importes «por facturar») o como **factura grabada** (importes «calculados»). Abreviatura: `Factura.cls` = `classes/Factura.cls`. Las citas `:línea` sin fichero son de `Form_FormFacturaResultado.cls`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormFacturaAlta` (generar) | Deja `m_ObjFacturaActiva` con `ParaPorFacturar = No`, cierra este formulario si estaba abierto y lo abre | `Form_FormFacturaAlta.cls:104-108` |
| `Form_FormFacturaAlta` (simulacro) | `m_ObjFacturaActiva` = factura sin grabar, `ParaPorFacturar = Sí` | `Form_FormFacturaAlta.cls:163, 183-187` |
| `Form_FormFacturacionGestion` (Ver Detalles) | Factura seleccionada como activa con `ParaPorFacturar = No` | `Form_FormFacturacionGestion.cls:443-448` |

No emite eventos. Se cierra con `cmdSalir` (`:10-13`) o por error de carga (`:278`).

## Controles y eventos

Controles: 31 (45 `Begin` en el `.form.txt`: 14 de formulario, secciones y plantillas por defecto + 31 controles).

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `FechaFinalEventos`, `FechaFactura` | texto bloqueado | — | Fechas en `dd/mm/yyyy` si son fecha válida | CAP-8 | `:106-111`; `.form.txt:394-423, 595-625` |
| `Recargo`, `TipoImpositivo` | texto bloqueado | — | Recargo y tipo con `#,##0.00` si son numéricos | CAP-8 | `:112-114, 136-138`; `.form.txt:511-542, 800-831` |
| `PrecioSuministro`, `PrecioSuministroConRecargo` | texto bloqueado | — | Importe de suministro sin y con recargo | CAP-8 | `:115-120`; `.form.txt:979-1011, 1072-1104` |
| `PrecioSubcontrataciones`, `PrecioSubcontratacionesConRecargo` | texto bloqueado | — | Importe de subcontrataciones sin y con recargo | CAP-8 | `:121-126`; `.form.txt:1039-1071, 1105-1137` |
| `PrecioDietasYGastos`, `PrecioIncurrido` | texto bloqueado | — | Dietas y gastos; precio incurrido | CAP-8 | `:127-132`; `.form.txt:451-483, 680-712` |
| `PrecioBaseImponible`, `PrecioImpuestos`, `ImporteAFacturar` | texto bloqueado | — | Base, impuestos y «TRABAJOS A FACTURAR» | CAP-8 | `:133-135, 139-144`; `.form.txt:740-772, 859-891, 919-951` |
| `ListaPerfilFactura` | lista de valores | — | Perfil, horas y precios laborales y extra, importes | CAP-9 | `:281-330`; `.form.txt:1138-1176` |
| `ComandoGenerarExcel` | botón | `OnClick` → `ComandoGenerarExcel_Click` | Genera el Excel y lo abre | CAP-8 | `:157-192`; `.form.txt:543-567` |
| `ComandoObtenerLI` | botón («LI»; oculto y deshabilitado en diseño) | `OnClick` → `ComandoObtenerLI_Click` | Genera la lista de importes (LI) del simulacro | CAP-8 | `:196-233`; `.form.txt:1244-1270` |
| `cmdSalir` | botón | `OnClick` → `cmdSalir_Click` | Cierra sin guardar | CAP-8 | `:10-13`; `.form.txt:1188-1211` |
| `lblEstado` | etiqueta (oculta en diseño) | — | Línea de avance | CAP-8 | `:246`; `.form.txt:1212-1243` |
| `lblTitulo1`, `lblTitulo2` | etiquetas de cabecera | — | Título dinámico (simulacro o factura n.º) | CAP-8 | `:253-260`; `.form.txt:340-382` |
| 11 etiquetas `Etiqueta52` … `Etiqueta84` | etiquetas | — | Rótulos de campo (`PrecioSuministroConRecargo` y `PrecioSubcontratacionesConRecargo` no tienen rótulo propio) | CAP-8 | `.form.txt:424-450, 484-510, 568-594, 626-679, 713-739, 773-799, 832-858, 892-918, 952-978, 1012-1038` |

## Flujos

### Apertura (`Form_Load`, `:236-279`)

1. Reloj de arena, `AjustarTamaño Me`, `lblEstado` oculto (`:243-246`).
2. Sin `m_ObjFacturaActiva`: «No hay ninguna factura activa» y se cierra (`:249-252, 276-278`).
3. Título: «RESÚMEN DEL SIMULACRO DE LA FACTURA» si `ParaPorFacturar = Sí`; si no, «RESÚMEN DE LA FACTURA » + `IDFactura` (`:253-260`).
4. `EstablecerDatos` (`:16-155`): oculta `ComandoObtenerLI` y solo lo muestra si `IDFactura = ""`, es decir, en simulacro (`:42-52`). Si `ParaPorFacturar = 0` lo trata como `Sí` (`:57-59`). Con `Sí` lee las propiedades `…PorFacturar`; en otro caso las `…Calculado` (`:60-104`). Formatea cada valor solo si es fecha o número (`:106-144`) y rellena la lista (`:145`).

### Generar Excel (`:157-192`)

Con `ParaPorFacturar = Sí` llama a `SalidaAExcel No`; si no, a `SalidaAExcel Sí` (`:171-176`). Si el fichero existe lo abre con ShellExecute (`:180-182`). `SalidaAExcel Sí` exige `IDFactura` numérico y reescribe el fichero si no está abierto (`Factura.cls:5170-5186, 5233-5235`).

### Obtener LI (`:196-233`)

`ObtenerLI(Sí)` sobre la factura activa (`:211`); sin `IDFactura` usa las colecciones «por facturar» y el directorio local del entorno (`Factura.cls:3474-3479, 3487-3489`). Si el fichero devuelto no existe, «El método ObtenerLI ha devuelto un resultado con formato desconocido» (`:215-218`). Error: `MsgBox` crítico, el formulario sigue abierto (`:226-232`).

## Reglas, validaciones y cálculos

- Todos los importes son de solo lectura: los 13 cuadros de texto llevan `Locked` en diseño (`.form.txt:395, 452, 512, 596, 681, 741, 801, 860, 920, 980, 1040, 1073, 1106`).
- El formulario no calcula: muestra las propiedades de `Factura` (`Factura.cls:200-1969`) con formato monetario `#,##0.00 €` (`:116-143`).
- Selección de fuente de datos: `ParaPorFacturar = Sí` ⇒ `…PorFacturar`; `No` ⇒ `…Calculado` (`:60-104, 301-305`).
- La lista de perfiles usa `ColPerfilesPorFacturar` en simulacro y `ColPerfilesEnFactura` en factura grabada (`:301-305`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Importes y perfiles de la factura grabada | tablas `TbFactura*` | lecturas diferidas de `Factura` (propiedades `…Calculado`, `ColPerfilesEnFactura`) | — | `Factura.cls:200-1969, 1632` |
| Importes del simulacro | `TbEventos` y relacionadas | propiedades `…PorFacturar` sobre SQL dinámico de eventos pendientes | — | `Constructor.bas:1206-1211`; `Factura.cls:262-1969` |
| Excel y LI | ficheros `.xlsx` | automatización de Excel; no escribe en la base | — | `Factura.cls:5155-5266, 3453-3590` |

El formulario no escribe en tablas.

## Coexistencia

Los importes dependen de la capa de `Factura` (consultas Jet con fechas `#mm/dd/yyyy#`, ver `Constructor.bas:1209`). El propio formulario solo requiere que las lecturas devuelvan números y fechas equivalentes; `Format` con `€` y separadores depende de la configuración regional del equipo (`:113-143`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Mensajes de error con nombre equivocado o sin detalle: `EstablecerDatos` informa «El método RellenarListaPerfiles…» (`:153`) y `RellenarPerfiles` solo devuelve el número de error (`:328`).
- El modo se codifica con `ParaPorFacturar` (0 se trata como `Sí`, `:57-59`): una factura recuperada sin fijar la bandera se mostraría como simulacro.
- `ComandoGenerarExcel` en factura grabada regenera y reemplaza el fichero existente (`Factura.cls:5233-5235`) sin avisar.
- Si `EstablecerDatos` falla en `RellenarPerfiles`, ya se han rellenado los importes (`:145-148`).
- El rótulo «TIPO IMPONIBLE» (`.form.txt:785`) designa el tipo impositivo.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: modo explícito (simulacro / factura grabada) en lugar de una bandera de tres estados.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir sin factura activa muestra «No hay ninguna factura activa» y cierra.
2. Tras un simulacro, el título es «RESÚMEN DEL SIMULACRO DE LA FACTURA», aparece `ComandoObtenerLI` y los importes son los «por facturar».
3. Tras generar o desde Gestión, el título incluye el número de factura, `ComandoObtenerLI` no se muestra y los importes son los calculados.
4. Ningún campo del resumen es editable.
5. La lista de perfiles muestra horas laborales y extra, precios e importe por perfil (`:309-321`).
6. `Generar Excel` abre el fichero de la factura; en simulacro lo genera sin `IDFactura` numérico.

## Pendientes

- Fórmulas de las propiedades `Precio*Calculado` y `*PorFacturar` (`Factura.cls:200-1969`), no desarrolladas aquí.
- Contenido de `AjustarTamaño` (`:245`).
- Contenido de `ObtenerLI` más allá de la selección de colecciones y directorio (`Factura.cls:3526-3590`).
