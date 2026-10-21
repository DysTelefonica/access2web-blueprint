# Form_FormEleccionTipoConsulta

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/forms/Form_FormEleccionTipoConsulta.cls` (390 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Diálogo que pide el tipo de consulta («Consulta Completa» o «Consulta Simple») y genera un Excel con los expedientes que muestra en ese momento una lista del formulario de origen (`:129-131`, `:300-330`). Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`); lee expedientes de la base y, con «Actualizar», **escribe** `TbExpedientesConEntidades` (`:357-367`). Las citas son `Form_FormEleccionTipoConsulta.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormExpedientesGestion` (botones de exportar AM, lotes, basados e informe unificado) | Fija la global `m_NombreListaParaInforme` (nombre de la lista o «Todos») y abre sin `OpenArgs`; cierra antes una instancia abierta | `Form_FormExpedientesGestion.cls:1112-1116, 1143-1147, 1175-1179, 1206-1210` |
| `Form_FormExpedientesGestionTecnica` (`ComandoExportarExcel`) | Abre con `OpenArgs = "ListaFiltrados|FormExpedientesGestionTecnica"`; no toca la global | `Form_FormExpedientesGestionTecnica.cls:239-242` |

Abre el Excel generado con `Ejecutar` (`ShellExecuteA`, `Variables Globales.bas:24`) y se cierra (`:77-80`). `ComandoAyuda` abre un PDF de ayuda vía `AbrirAyuda` (`:104`; `FUNCIONES UTILES.bas:3076-3100`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `ListaTipoConsulta` | ListBox, lista de valores | `OnClick` | Elige la colección de campos del entorno y, la primera vez, carga los expedientes de la lista de origen | `:194-291`, `.form.txt:364-371` |
| `ListaTipoConsulta` | — | `OnDblClick` | Ejecuta la consulta si `ComandoEjecutar` está habilitado | `:293-299` |
| `ComandoEjecutar` | botón (deshabilitado en diseño) | `OnClick` | Genera el Excel, lo abre y cierra el formulario | `:56-94`, `.form.txt:483, 492-494` |
| `ComandoActualizarCompleto` | botón, tooltip «Actualizar para informe» | `OnClick` | Re-registra entidades de cada expediente cargado y lo relee | `:27-54`, `.form.txt:433-435` |
| `ComandoAyuda` | botón con imagen, tooltip «Ayuda» | `OnClick` | `AbrirAyuda` | `:98-120`, `.form.txt:273-275` |
| `cmdSalir` | botón | `OnClick` | Cierra sin guardar, con `On Error Resume Next` | `:19-23`, `.form.txt:580-582` |
| `lblEstado` | etiqueta oculta en diseño | — | `Avance` escribe ahí el progreso del formulario activo | `.form.txt:403, 411`; `FUNCIONES UTILES.bas:2729-2738` |

El `RowSource` de diseño de `ListaTipoConsulta` (`.form.txt:366-367`) es un resto de otro formulario: `EstablecerLista` lo sustituye al abrir por `"Consulta Completa;Consulta Simple"` (`:130`).

## Flujos

### Apertura (`Form_Open`, `:140-193`)

1. Deshabilita `ComandoEjecutar` y `ComandoActualizarCompleto`; `AllowEdits = False` (`:150-152`).
2. Sin `OpenArgs`: origen `FormExpedientesGestion` y lista = `m_NombreListaParaInforme` (`:154-156`). Con `OpenArgs`: `Split` por «|»; si no hay exactamente dos partes, error «El formulario se ha abierto con parámetros insuficientes» (`:158-164`).
3. Título `m_TituloFormulario` (`:167`). Si el formulario de origen no está abierto, error «No está abierto el formulario <nombre>» (`:168-171`; `FUNCIONES UTILES.bas:2821-2833`).
4. `EstablecerLista` carga las dos opciones, **selecciona la fila 1 («Consulta Simple»)** y dispara `ListaTipoConsulta_Click` (`:129-132`).

### Elección del tipo (`ListaTipoConsulta_Click`, `:194-291`)

1. «Consulta Completa» ⇒ `m_ObjEntorno.ColCamposInformeCompleto`; «Consulta Simple» ⇒ `ColCamposInformeSimple`; otro ⇒ «Tipo no codificado» (`:208-215`). Habilita los dos botones (`:216-217`).
2. Si los expedientes ya se cargaron (`m_ColObtenidas = Sí`), termina (`:218-224`).
3. Lista «Todos»: carga `ListaAMEIndividual`, `ListaLotes` y `ListaBasados` del origen (`:226-242`). Otra lista: vacía AM, lotes y basados y carga solo la indicada; `ListaFiltrados` va a `m_ColTecnica` (`:244-270`).
4. `getColExpedientesDeLista` salta la fila 0 (cabecera) y lee cada `IDExpediente` con `constructor.getExpedienteCompleto` (`FUNCIONES UTILES.bas:3324-3341`; `constructor.bas:1208-1219`). Las listas de origen tienen `ColumnHeads` (`Form_FormExpedientesGestion.form.txt:1627, 1740, 1782`; `Form_FormExpedientesGestionTecnica.form.txt:1067`).

### Ejecución (`ComandoEjecutar_Click`, `:56-94`)

1. `getExcelConsulta` decide `p_IncluirDerivados = (m_NombreListaParaInforme = "Todos")` (`:314-318`) y llama a `GenerarConsultaExpedientes` (`:319-326`).
2. El generador crea un libro con nombre temporal en `URLDirectorioLocal`, escribe la cabecera con los textos de la colección de campos, una fila por expediente (y derivados, sin repetir ID) y lo convierte en tabla «Tabla1» (`FUNCIONES UTILES.bas:1803-1877, 1916-1963, 1980`).
3. Si el fichero no existe, «No se ha podido obtener la consulta» (`:73-76`). Si existe, lo abre y cierra el diálogo (`:77-80`).

### Actualización (`ComandoActualizarCompleto_Click`, `:27-54`)

Para AM, lotes y basados, `geColActualizada` llama por expediente a `ExpedienteOperaciones.RegistrarExpEntidades` y relee el expediente (`:358-378`). `RegistrarExpEntidades` refresca la fila de `TbExpedientesConEntidades` (`classes/ExpedienteOperaciones.cls:609-621`; `classes/ExpedienteEntidadOperaciones.cls:54-67`).

### Errores

Los handlers muestran `MsgBox` crítico y, salvo el error interno 1000, envían `CorreoAlAdministrador` (`:86-93, 113-119, 184-191, 283-290`). Un error en la apertura cierra el formulario (`:192`). `ComandoActualizarCompleto` no envía correo (`:48-53`).

## Reglas, validaciones y cálculos

- **Campos de cada consulta:** Simple, 13 columnas (`classes/Entorno.cls:194-208`); Completa, 54 columnas (`classes/Entorno.cls:229-284`). La clave es la propiedad del expediente y el valor, el texto de cabecera (`FUNCIONES UTILES.bas:1821-1825`).
- **Formato de celda:** `ImporteContratacion` numérico cambia coma por punto; campos con «Fecha» van a `mm/dd/yyyy`; `ResponsableSeguridad` «0» o vacío ⇒ «N/A»; `Estado` usa `ESTADOCalculadoTexto` (`modInformeExpedientesHelper.bas:20-31`).
- **Derivados:** solo con «Todos» se añaden recursivamente los derivados de cada expediente (`:314-318`; `FUNCIONES UTILES.bas:1949-1962`).
- **Lista técnica:** si llega `p_ColTecnica`, sustituye a la colección de basados (`FUNCIONES UTILES.bas:1799-1801`).
- **Sin datos:** sin campos o sin ninguna colección, el generador devuelve cadena vacía (`FUNCIONES UTILES.bas:1793-1798`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Carga de cada expediente | `TbExpedientes` LEFT JOIN `TbExpedientesConEntidades` | SQL dinámico, una consulta por ID | No | `constructor.bas:1208-1219` |
| Re-registro de entidades | `TbExpedientesConEntidades` (AddNew/Edit) | DAO | No consta | `classes/ExpedienteEntidadOperaciones.cls:54-67` |
| Excel de salida | fichero temporal en `URLDirectorioLocal` | Automatización Excel | — | `FUNCIONES UTILES.bas:1803-1815` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: `IDExpediente` concatenado sin parámetros (`constructor.bas:1218`); una consulta por fila de la lista (N+1, `FUNCIONES UTILES.bas:3329-3332`); y una escritura en `TbExpedientesConEntidades` lanzada desde un diálogo de informes, que debe respetar las reglas de concurrencia que se acuerden.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `Err.Numbr` (`:50`): miembro inexistente de `ErrObject`; con `Option Explicit` y enlace temprano, el módulo podría no compilar o fallar en esa rama.
- `getExcelConsulta` lee la global `m_NombreListaParaInforme` (`:314`), no `m_NombreLista`: desde Gestión Técnica, que no la fija (`Form_FormExpedientesGestionTecnica.cls:239-242`), los derivados dependen del último botón pulsado en Gestión.
- `ComandoActualizarCompleto` no actualiza `m_ColTecnica` (`:33-44`): en la consulta técnica «Actualizar» no tiene efecto sobre el informe.
- «Actualizar» escribe en la base por cada expediente sin confirmación ni transacción (`:363`); un expediente que no se relee desaparece del informe sin aviso (`:371`).
- La rama de error de `ComandoActualizarCompleto` y la de `ComandoEjecutar` no llaman a `AvanceCerrar` (`:48-53, 85-93`).
- La opción preseleccionada es «Consulta Simple» (`Selected(1)`, `:131`); no consta si se pretendía la completa.
- Con la lista de origen vacía, el mensaje es «No se ha podido obtener la consulta» (`:73-75`), sin indicar que no hay expedientes.
- Si el formulario de origen se cierra tras la apertura, `Forms(m_NombreFormulario)` falla en el clic (`:227, 248`); solo se comprueba en `Form_Open` (`:168`).
- El Excel temporal no se borra nunca (`FUNCIONES UTILES.bas:1803-1815`).
- Cabecera «ESTADOCalculadoTexto» en ambas consultas (`classes/Entorno.cls:200, 277`); `ImporteLicitacion` no recibe la normalización de `ImporteContratacion` (`modInformeExpedientesHelper.bas:23`).
- `AbrirAyuda` ignora el nombre del formulario y abre siempre el mismo PDF genérico (`FUNCIONES UTILES.bas:3094-3095`).
- Restos: `Caption` «Salir» en `ComandoAyuda` y «Editar» en los botones de pie (`.form.txt:274, 434, 493`); `lst` sin uso en `ComandoEjecutar` (`:57`); `RowSource` de diseño ajeno (`.form.txt:366`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: exportar desde el filtro activo del listado (endpoint con los mismos criterios), sin diálogo intermedio; separar «Actualizar» del informe como acción administrativa explícita; generar el fichero en servidor en streaming, sin automatización de Excel.

## Criterios de aceptación

Propuestos, no aprobados:

1. La exportación ofrece «Consulta Completa» y «Consulta Simple», con las columnas y cabeceras de `ColCamposInformeCompleto` y `ColCamposInformeSimple`.
2. Exportar una lista concreta incluye solo sus expedientes; el informe unificado incluye AM, lotes y basados con sus derivados, sin filas repetidas.
3. Las fechas, el importe de contratación y el responsable de seguridad vacío se exportan con el formato descrito en Reglas.
4. Exportar una lista vacía informa de que no hay expedientes y no genera fichero.
5. La exportación desde la búsqueda técnica no depende de acciones previas en otra pantalla.

## Pendientes

- Si el proyecto compila con `Err.Numbr` (`:50`) o si esa rama nunca se ha ejecutado.
- Quién ve los botones de exportar en los formularios de origen (permisos no trazados; este formulario no comprueba rol).
- Contenido de `ActualizarExpedienteCompleto` y de las propiedades `Derivados` y `getPropiedad` de `ExpedienteCompleto` (`classes/ExpedienteOperaciones.cls:621`; `FUNCIONES UTILES.bas:1950`; `modInformeExpedientesHelper.bas:16`).
- Si `ExpedienteEntidadOperaciones.Registrar` abre transacción en otro tramo (`classes/ExpedienteEntidadOperaciones.cls:17`).
- Valor de `URLDirectorioLocal` y política de limpieza de ese directorio.
