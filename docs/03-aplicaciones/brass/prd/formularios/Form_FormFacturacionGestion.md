# Form_FormFacturacionGestion

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormFacturacionGestion.cls` (665 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Gestión de facturas ya generadas («GESTIÓN DE LA FACTURACIÓN», `Form_FormFacturacionGestion.form.txt:40`). Es un formulario **no enlazado**: un combo lista todas las facturas de `TbFacturaPrincipal`, una lista de valores muestra la seleccionada y los botones actúan sobre el objeto `Factura` en `m_ObjFacturaSeleccionada` (`:6`). Permite crear una factura nueva, ver sus detalles, obtener la LI, generar los partes por centro, anexar partes firmados, anexar o quitar la autorización de materiales y eliminarla. Se abre desde el menú de `Form_Form0BDOpciones` (`Form_Form0BDOpciones.cls:146`). Abreviatura: `Factura.cls` = `classes/Factura.cls`. Las citas `:línea` sin fichero son de `Form_FormFacturacionGestion.cls`.

## Entrada y salida

| Hacia | Cómo | Evidencia |
|---|---|---|
| `FormFacturaAlta` | Cierra el formulario si estaba abierto, lo abre y guarda `m_Form` (`WithEvents`) | `:235-258` |
| `FormFacturaResultado` | `Ver Detalles`: factura activa con `ParaPorFacturar = No` | `:421-461` |
| `FormParteGestion` | `Anexar P.Firmados`: solo si no está abierto, fija su control `IDFactura` y llama a `IDFactura_Exit` | `:97-101` |
| Selector de fichero o carpeta | `Seleccionar` (de `Funciones Generales.bas`) para el anexo y la carpeta de partes | `:129, 329` |

Escucha el evento `Alta` de `FormFacturaAlta` (`:628-636`), pero ese formulario no lo lanza (`Form_FormFacturaAlta.cls:15`). Al generar una factura, `FormFacturaAlta` llama directamente a `RellenarComboIDFactura` (`Form_FormFacturaAlta.cls:109-111`).

## Controles y eventos

Controles: 17 (29 `Begin` en el `.form.txt`: 12 de formulario, secciones y plantillas por defecto + 17 controles).

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `ComboIDFactura` | combo de lista de valores, 2 columnas | `OnEnter` → `Requery`; `OnChange` → `ComboIDFactura_Change` | Elige la factura; el cambio recarga la lista | CAP-8 | `:463-484`; `.form.txt:518-555` |
| `ListaFiltrados` | lista de valores con cabecera, `ColumnCount` 3 | `OnClick` → `ListaFiltrados_Click` | Muestra ID, título y fecha; al pulsar habilita botones | CAP-8 | `:13-70, 555-626`; `.form.txt:358-395` |
| `ComandoActualizar` | botón (rótulo «Comando4», ayuda «Actualizar») | `OnClick` | Recarga el combo con `On Error Resume Next` | CAP-8 | `:72-75`; `.form.txt:944-1005` |
| `ComandoNuevaFactura` | botón | `OnClick` | Abre `FormFacturaAlta` | CAP-8 | `:235-258`; `.form.txt:396-433` |
| `ComandoVerDetalles` | botón, deshabilitado en diseño | `OnClick` | Abre el resultado de la factura | CAP-8 | `:421-461`; `.form.txt:746-783` |
| `ComandoObtenerLI` | botón («LI»), deshabilitado en diseño | `OnClick` | Genera y abre la LI de la factura | CAP-8 | `:259-304`; `.form.txt:583-622` |
| `ComandoPartesParaCentros` | botón («Informe Para Centros»), deshabilitado | `OnClick` | Genera partes y documentos por centro | CAP-8 | `:306-359`; `.form.txt:707-745` |
| `ComandoAnexarFirmados` | botón, deshabilitado | `OnClick` | Abre `FormParteGestion` con la factura | CAP-8 | `:77-108`; `.form.txt:434-473` |
| `ComandoAnexoAutorizacionMateriales` | botón, deshabilitado | `OnClick` | Anexa un documento de autorización a los materiales | CAP-8 | `:110-187`; `.form.txt:623-662` |
| `ComandoQuitarAnexoAutorizacionMateriales` | botón, deshabilitado | `OnClick` | Elimina esos anexos | CAP-8 | `:361-419`; `.form.txt:663-706` |
| `ComandoEliminarFactura` | botón, deshabilitado | `OnClick` | Elimina la factura | CAP-8 | `:189-232`; `.form.txt:474-517` |
| `ImagenFacturacionAnexada` | imagen (oculta; «Sobre.png») | `OnClick` | Abre el fichero de la factura | CAP-8 | `:519-554`; `.form.txt:784-943` |
| `cmdSalir` | botón | `OnClick` | Cierra sin guardar | CAP-8 | `:515-518`; `.form.txt:1017-1053` |
| `lblEstado` | etiqueta (oculta) | — | Línea de avance | CAP-8 | `:498`; `.form.txt:1054-1076` |
| `lblTitulo1`, `lblTitulo2` | etiquetas | — | Título; `lblTitulo1` se copia a `Caption` | CAP-8 | `:499`; `.form.txt:300-346` |
| `Etiqueta142` | etiqueta «ID FACT» | — | Rótulo del combo | CAP-8 | `.form.txt:556-582` |

## Flujos

### Apertura (`Form_Open`, `:486-514`)

`AjustarTamaño`, `RellenarComboIDFactura` (todas las facturas, de mayor a menor ID, `:637-658`), título de ventana. Error: `MsgBox` crítico y cierre (`:507-513`).

### Selección (`ComboIDFactura_Change` → `ActualizarLista`, `:13-70, 463-479`)

1. Deshabilita todos los botones de acción y oculta la imagen (`:29-36`).
2. Con combo no vacío, `Constructor.getFactura` (`:38-43`); añade una fila «ID;Título;Fecha» con `;` sustituido por `:` en el título (`:53`).
3. Si la lista tiene dos filas (cabecera y una factura), la selecciona y llama a `ListaFiltrados_Click` (`:56-59`).
4. `ListaFiltrados_Click` (`:555-626`) recarga `m_ObjFacturaSeleccionada` si cambió el ID y habilita: `Partes para centros`, `Eliminar`, `LI`, `Ver Detalles` siempre (`:604-612`); `Anexar P.Firmados` solo si la factura tiene partes (`:600-602`); `Quitar Anexo` si hay anexos de autorización y, si no, `Anexo Autorización Mat.` (`:607-611`).

### Acciones (todas reutilizan la factura cargada o la recargan de la lista)

- **Eliminar** (`:189-232`): confirmación «¿Desea eliminar la factura?» con «No» por defecto; `Factura.Eliminar`; vacía el combo, lo recarga y lanza `ComboIDFactura_Change`.
- **LI** (`:259-304`): `ObtenerLI(No)` → fichero en el directorio de certificaciones del entorno (`Factura.cls:3474-3476`); lo abre.
- **Partes por centro** (`:306-359`): pide carpeta; `GenerarPartesDeFactura` crea un `Parte` por BUI y los documentos Word (`Factura.cls:4013-4070`); vacía `ColPartes`, refresca botones y abre la carpeta.
- **Anexo de autorización** (`:110-187`): pide fichero, comprueba que no está abierto (`:139-143`), pide confirmación y anexa a cada material de la factura que aún no tenga anexo (`Factura.cls:3945-3962`).
- **Quitar anexo** (`:361-419`): cuenta los anexos, confirma, y los elimina (`Factura.cls:3974-4012`).
- Errores: `MsgBox` crítico con `m_Error`; el formulario sigue abierto.

## Reglas, validaciones y cálculos

- **Selección obligatoria:** sin factura en la lista, «Se ha de seleccionar una factura de la lista» (p. ej. `:88-91`).
- **Eliminar:** exige `IDFactura` numérico y el fichero Excel cerrado («Ha de cerrar la factura primero») (`Factura.cls:3398-3407`).
- **Anexo de autorización:** solo para materiales sin anexo previo (`Factura.cls:3951-3959`); los anexos de una factura se reconocen por la descripción «Anexado automáticamente desde la factura N» (`Constructor.bas:3357-3360`).
- **Partes:** exige una carpeta existente (`Factura.cls:4027-4030`) y al menos un evento por centro (`Factura.cls:4041-4044`).
- **Control de acceso:** el formulario no comprueba permisos al abrirse (`:486-514`); el menú lo abre con `On Error Resume Next` (`Form_Form0BDOpciones.cls:144-147`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Lista de facturas | `TbFacturaPrincipal` | SQL fija `SELECT *` | — | `Constructor.bas:810-812` |
| Factura por ID | `TbFacturaPrincipal` | SQL concatenado | — | `Constructor.bas:2440-2441` |
| Partes de la factura | `TbPartesPpal` | SQL concatenado | — | `Constructor.bas:919-921` |
| Anexos de autorización | `TbAnexos` | SQL concatenado con texto de descripción | — | `Constructor.bas:3357-3360` |
| Eliminar factura | `TbFacturaEventosInvolucrados`, `…ActividadesInvolucradas`, `…MaterialesInvolucrados`, `…GastosInvolucrados`, `…SubcontratacionesInvolucradas`, `TbFacturaPrincipalPerfiles`, `TbFacturaPrincipal` | 7 `DELETE` concatenados, y borrado del fichero Excel | no | `Factura.cls:3408-3441` |
| Alta de partes y anexos | `TbPartesPpal`, `TbPartesDetalle`, `TbAnexos` | clases `Parte` y `Anexo` | — | `Factura.cls:4051-4057, 3952-3954` |

## Coexistencia

Adaptar al PostgreSQL vinculado: `DELETE tabla.campo FROM` de dialecto Jet (`Factura.cls:3408-3436`); siete borrados consecutivos sin transacción; identificación de anexos por texto de descripción (`Constructor.bas:3359-3360`); `ORDER BY … Desc` y conversión a `String` del ID en claves de diccionario (`Constructor.bas:832-833`); ficheros en red, Word y Excel quedan fuera de la base.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Lógica invertida tras quitar anexos: si ya no quedan, `ComandoQuitarAnexoAutorizacionMateriales` se **habilita** y si quedan se deshabilita (`:401-405`).
- Tras anexar, `ComandoAnexoAutorizacionMateriales` sigue habilitado (`:167-171`).
- `ImagenFacturacionAnexada` nunca se muestra (se oculta en `:36, 571, 614`), por lo que `ImagenFacturacionAnexada_Click` (`:519-554`) es inaccesible.
- Eliminar no borra partes (`TbPartesPpal`) ni anexos de autorización de la factura (`Factura.cls:3408-3436`) y no es atómico; si falla a mitad quedan datos parciales.
- Tras eliminar, `m_ObjFacturaSeleccionada` no se pone a `Nothing` (`:212-221`); un botón posterior que reutilice el objeto podría actuar sobre la factura borrada.
- `ComandoAnexarFirmados`: el mensaje «no tiene ningún parte de trabajo generado» (`:93-96`) es inalcanzable tras `:88-91`; si `FormParteGestion` ya está abierto no hace nada (`:97`).
- `m_Form_Alta` (`:628-636`) es código muerto y desreferencia `m_ObjFacturaSeleccionada`, que puede ser `Nothing`.
- `RowSource` de diseño de `ListaFiltrados` declara 4 columnas y `ColumnCount` vale 3 (`.form.txt:365, 376`).
- Rótulo del botón de refresco «Comando4» (`.form.txt:952`).
- Las acciones aprovechan un objeto en caché sin releer la factura: datos modificados por otra sesión no se reflejan (`:81-92`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: borrado atómico con cascada; control de permisos explícito; mostrar el estado de «factura anexada».

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir, el combo lista todas las facturas existentes y los botones de acción están deshabilitados.
2. Elegir una factura la muestra en la lista y habilita `Ver Detalles`, `LI`, `Informe Para Centros` y `Eliminar`.
3. `Anexar P.Firmados` solo se habilita si la factura tiene partes generados.
4. Eliminar pide confirmación; con «No» no cambia nada; con «Sí» desaparece de la lista y se borran sus filas relacionadas y el Excel.
5. Eliminar con el Excel abierto muestra «Ha de cerrar la factura primero» y no borra.
6. Anexar autorización se aplica solo a materiales sin anexo y deja habilitado `Quitar Anexo`.
7. Quitar anexo pide confirmación e informa del número de anexos a eliminar.
8. Una factura recién generada aparece en el combo sin reabrir el formulario.

## Pendientes

- Contenido de `Seleccionar`, `AbrirEnLocal` y `Avance` (`Funciones Generales.bas:1117, 218, 7263`) y de `Form_FormParteGestion.IDFactura_Exit`.
- Qué ocurre con los ficheros de partes y Word ya generados al eliminar la factura.
- Declaración y uso de `m_FormPartes` (`:8`), sin handlers en este formulario.
