# Form_FormExpedientesGestionTecnica

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/forms/Form_FormExpedientesGestionTecnica.cls` (773 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Buscador de expedientes para usuarios técnicos: cinco filtros (palabra clave, estado, empresa, número de expediente y miembro del proyecto), una lista de resultados y acceso a la ficha `FormExpediente`. Es un formulario **no enlazado**: la lista es de tipo «Value List» y se rellena por código (`.form.txt:1083-1085`; `:111, 128-129`). No escribe en la base de datos. Las citas son `Form_FormExpedientesGestionTecnica.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_Form0BDOpciones` | `ComandoGestionExpedientes_Click` abre este formulario si `EsTecnico = Sí`; si no, `FormExpedientesGestion` | `Form_Form0BDOpciones.cls:116-124` |
| `Form_Form0BDOpcionesTecnicos` | `ComandoBusquedaSimple_Click` lo cierra si estaba abierto y lo abre | `Form_Form0BDOpcionesTecnicos.cls:84-95` |

Abre `FormExpediente` sin `OpenArgs` (`:462-465`) y `FormEleccionTipoConsulta` con `OpenArgs = "ListaFiltrados|<nombre del formulario>"` (`:239-242`). No escucha eventos de los formularios que abre. Sale con `cmdSalir` (`:64-69`).

Estado compartido (variables globales de `Variables Globales.bas`): `m_ObjExpBusquedaTecnicaActiva` (`:270`), `m_DatosEnMemoria` (`:277`, fijado por `TempVars("DatosEnMemoria")` en `:376-380`) y `m_TituloFormulario`. Al abrir la ficha escribe `m_ObjExpedienteActivo`, `m_ObjExpedienteDTOActivo` y `m_ObjExpedienteAlInicio` (`:450-457`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `PalabraClave` | texto | `OnKeyDown` | Con Intro, si `ComandoPalabraClave` está habilitado, le da el foco y lo ejecuta | `:756-772`, `.form.txt:710-714` |
| `ComandoPalabraClave` (rótulo «Ver Todos») | botón | `OnClick` | `Filtrar` | `:410-431`, `.form.txt:617-619` |
| `Estado` | combo 2 columnas (ID oculto), `LimitToList` | `AfterUpdate` | `Filtrar`; la búsqueda usa el texto de `.Column(1)` | `:481-503`, `.form.txt:1200-1224` |
| `Juridica` (rótulo «Empresa») | combo | `AfterUpdate` | `Filtrar` | `:528-549`, `.form.txt:1395-1415` |
| `CodExp` | combo, `LimitToList` | `AfterUpdate` | `Filtrar` | `:150-171`, `.form.txt:1257-1294` |
| `JP` (rótulo «Miembro del proyecto») | combo, `LimitToList` | `AfterUpdate` | `Filtrar` | `:504-526`, `.form.txt:1327-1362` |
| `ComandoIniciarPalabraClave`, `ComandoIniciarEstado`, `ComandoIniciarJuridica`, `ComandoIniciarCodExp`, `ComandoIniciarJP` | botones | `OnClick` | Vacían un filtro con `IniciarCampo` y vuelven a filtrar | `:380-407, 291-315, 350-377, 261-288, 321-348`; `IniciarCampo` `:174-228` |
| `ComandoReiniciar` («Ver Todos») | botón | `OnClick` | `Reestablecer` + olvida la última búsqueda + `Filtrar` | `:719-753`, `.form.txt:529-533` |
| `ListaFiltrados` | lista 8 columnas con cabecera, ID oculto | `OnClick`, `OnDblClick` | Carga el expediente seleccionado y habilita `ComandoVerDetalle`; el doble clic lo ejecuta | `:625-689, 618-623`, `.form.txt:1066-1090` |
| `ComandoVerDetalle` («Detalle»; deshabilitado en diseño) | botón | `OnClick` | Prepara las globales del expediente y abre `FormExpediente` | `:433-479`, `.form.txt:1826-1842` |
| `ComandoExportarExcel` | botón | `OnClick` | Abre `FormEleccionTipoConsulta` sobre la lista | `:230-259`, `.form.txt:1120-1122` |
| `ComandoAyuda` (rótulo «Salir» en `.form.txt`) | botón | `OnClick` | `AbrirAyuda` (`FUNCIONES UTILES.bas:3076`) | `:551-576`, `.form.txt:949-953` |
| `cmdSalir` | botón | `OnClick` | Cierra sin guardar | `:64-69` |
| `lblLista` | etiqueta | — | Título de la lista con el número de resultados | `:109, 135` |

Los `RowSource` de diseño de los combos y de la lista (p. ej. `.form.txt:1222, 1277, 1349, 1085`) se sustituyen en ejecución por `RellenarCombos` y `Filtrar`; contienen datos reales y no se reproducen aquí.

## Flujos

### Apertura (`Form_Open`, `:579-616`)

1. `Ajustar` dimensiona la ventana, `AllowEdits = False` y el título pasa a `m_TituloFormulario` (`:588-590`).
2. `EstablecerDatos` rellena los combos y fija los filtros iniciales con `setExpBusquedaTecnica` (`:12-33`). Como `m_ObjExpBusquedaTecnicaActiva` llega vacía, los filtros quedan en «Todos», `Estado = "0"` y palabra clave nula (`FUNCIONES UTILES.bas:1595-1603`).
3. `Filtrar` carga la lista (`:596`) y `AllowEdits = True` (`:600`).

### Filtrado (`Filtrar`, `:71-148`)

1. Deshabilita `ComandoVerDetalle` y lee los cinco filtros con `getExpTecnicoBusqueda` (`:93-95`; `FUNCIONES UTILES.bas:1269-1296`).
2. Si los filtros no cambian frente a la última búsqueda (`HaHabidoCambiosBusquedaTecnica`, `FUNCIONES UTILES.bas:1392-1449`), sale sin tocar la lista (`:99-108`).
3. Reinicia la lista con la fila de cabecera (`:109-111`) y obtiene los expedientes con `getColBusquedaTecnica` (`:113`): de la caché `m_ObjEntorno.ColExpedientesCompletos` si `m_DatosEnMemoria = Sí`, o de la base en otro caso (`FUNCIONES UTILES.bas:790-795`).
4. Añade una fila por expediente: ID, tipo, número, título, inicio, fin, fin de garantía y estado; el estado «NoAplica» se muestra «--» (`:121-132`).
5. Rotula «Lista de Expedientes Filtrados (n)», selecciona la fila de índice 2, ejecuta `ListaFiltrados_Click` y guarda los filtros como última búsqueda (`:135-139`).

### Ver detalle (`ComandoVerDetalle_Click`, `:433-479`)

1. Si no hay expediente cargado, lo lee con `constructor.getExpedienteCompleto`; si no existe, «Ha de seleccionar un expediente del tipo adecuado de la lista» (`:440-449`).
2. Fija `m_ObjExpedienteActivo`, construye el DTO con `EstablecerDTODeExpediente` (`FUNCIONES UTILES.bas:2105`) y relee el expediente para `m_ObjExpedienteAlInicio` (`:450-460`).
3. Cierra `FormExpediente` si estaba abierto y lo abre sin `OpenArgs`, es decir, en modo edición sujeto a `EsAdministrador` (`:462-465`; `Form_FormExpediente.cls:121-140`).

### Errores

Los handlers acumulan el mensaje en `m_Error`, envían `CorreoAlAdministrador` salvo el error interno 1000 y muestran `MsgBox` crítico (p. ej. `:162-170`). En `Form_Open` el correo se envía sin `On Error Resume Next` y el formulario se cierra (`:607-615`).

## Reglas, validaciones y cálculos

- **Filtros** (`ColPasanCriterioBusquedaTecnica`, `FUNCIONES UTILES.bas:1070-1143`); «Todos» o vacío no filtra:
  - Miembro del proyecto: subcadena dentro de `CadenaJPs` (`:1095-1099`).
  - Estado: igualdad con `ESTADOCalculadoTexto` (`:1100-1104`).
  - Empresa: subcadena dentro de `CadenaContratistas` (`:1105-1109`).
  - Número: igualdad con `CodExp` (`:1112-1116`).
  - Palabra clave: subcadena en `Titulo` o `Nemotecnico`, o igualdad con `CodExp` o `IDExpediente` (`:1117-1124`).
- **Abortar sin resultados parciales:** si algún expediente trae error previo, la colección completa se descarta con un mensaje que incluye su `IDExpediente` (`FUNCIONES UTILES.bas:1089-1094`).
- **Opciones de los combos:** estados de `ColEstadosTitulo` con «0;Todos» delante (`FUNCIONES UTILES.bas:83-107`; `classes/Entorno.cls:1539-1562`); empresas = valores distintos de `TbExpedientesConEntidades.CadenaContratistas` (`constructor.bas:3117-3119`); números = `CodExp` de todos los expedientes completos, vaciando antes la caché si no se trabaja en memoria (`FUNCIONES UTILES.bas:134-157`); miembros = nombres distintos de usuarios con `EsJefeProyecto = 'Sí'` en `TbExpedientesResponsables` (`constructor.bas:2580-2584`).
- **Tipo en la lista:** `TipoParaLista` vale «AM» si `EsAM = Sí`, «C» si `EsExpediente = Sí` y vacío en otro caso (`classes/ExpedienteCompleto.cls:392-396`).
- **Selección:** la lista solo recarga el expediente si cambia el ID seleccionado (`:646-663`); una fila sin ID deshabilita `ComandoVerDetalle` (`:636-645`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Todos los expedientes | `TbExpedientes` LEFT JOIN `TbExpedientesConEntidades` | SQL dinámico DAO | `constructor.bas:1254-1265` |
| Expediente seleccionado | mismas tablas, `WHERE IDExpediente=` concatenado | SQL dinámico DAO | `constructor.bas:1208-1219` |
| Empresas del combo | `TbExpedientesConEntidades` | SQL dinámico DAO | `constructor.bas:3117-3121` |
| Miembros del combo | `TbExpedientesResponsables` INNER JOIN `TbUsuariosAplicaciones` | SQL dinámico DAO | `constructor.bas:2580-2585` |
| DTO para la ficha | `constructor.getExpediente` y árbol de suministradores | DAO (no trazado aquí) | `FUNCIONES UTILES.bas:2105-2125` |

Sin escrituras propias.

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: filtrado en memoria sobre la caché de cadenas desnormalizadas de `TbExpedientesConEntidades` (`CadenaJPs`, `CadenaContratistas`), que debe estar al día para que la búsqueda sea correcta; `IDExpediente` concatenado sin parámetros (`constructor.bas:1218`); comparaciones `InStr` bajo `Option Compare Database` (`FUNCIONES UTILES.bas:2`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `Filtrar` lanza `Err.Raise 100` tras fallar `getExpTecnicoBusqueda` (`:97`): el manejador sobrescribe el detalle original con un mensaje genérico de error 100 (`:144-146`).
- `Filtrar` selecciona la fila de índice 2 (`:136`); con cabecera activa (`.form.txt:1067`) es el **segundo** resultado, y con un solo resultado el índice no existe.
- Sin resultados, `Filtrar` sale sin actualizar el contador ni la última búsqueda (`:117-120`).
- `ComandoIniciarEstado_Click` no apaga el reloj de arena ni muestra el error (`:309-315`).
- `IniciarCampo` asigna `Estado = "Todos"` (`:207`), pero la columna ligada es el ID («0;Todos», `FUNCIONES UTILES.bas:92`); `setExpBusquedaTecnica` asignaría el texto del estado al mismo combo (`FUNCIONES UTILES.bas:1614`).
- `m_ObjExpBusquedaTecnicaActiva` solo se pone a `Nothing` (`:699`); no se ha encontrado ninguna asignación, así que los filtros nunca se restauran.
- El combo «Empresa» ofrece cadenas completas de contratistas, no empresas (`constructor.bas:3117`); el filtro por subcadena (`FUNCIONES UTILES.bas:1106`) devuelve también expedientes cuyo nombre contiene el elegido. Lo mismo ocurre con miembros del proyecto (`:1096`).
- Los estados del combo («Ejecución», «Garantía»…) proceden de `ColEstadosTitulo` (`classes/Entorno.cls:1551-1560`), no del `RowSource` de diseño (`.form.txt:1222`); si `RellenarComboEstado` falla, quedarían textos que no coinciden.
- Lotes y basados aparecen con tipo vacío (`classes/ExpedienteCompleto.cls:392-396`).
- `ComandoVerDetalle` abre la ficha en modo edición (`:465`) y no escucha `ExpedienteEditado`, así que la lista no se refresca tras guardar.
- Rótulos de diseño incoherentes: `ComandoAyuda` «Salir» (`.form.txt:950`), `ComandoVerDetalle` con ayuda contextual de exportación a Excel (`.form.txt:1842`), `PalabraClave` anuncia campos distintos de los que se buscan (`.form.txt:714`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: fusionar con el buscador general de `FormExpedientesGestion` con un perfil técnico; filtrar en servidor por tablas relacionales en lugar de cadenas; conservar los filtros en la URL.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir, todos los filtros están en «Todos» y la lista muestra todos los expedientes con su número.
2. Cambiar cualquier filtro recarga la lista; repetir el mismo filtro no la recarga.
3. Cada botón «Ver Todos» vacía solo su filtro; «Reiniciar» vacía los cinco.
4. La palabra clave encuentra por título, nemotécnico, número exacto o ID exacto; Intro lanza la búsqueda.
5. El estado «NoAplica» se muestra como «--».
6. «Detalle» y el doble clic abren la ficha del expediente seleccionado; sin selección, el botón está deshabilitado.
7. Exportar abre la elección de consulta sobre los resultados visibles.

## Pendientes

- Quién asigna `EsTecnico` y `m_TituloFormulario`.
- Comportamiento de `ListBox.Selected(2)` fuera de rango y de asignar «Todos» a un combo con `LimitToList` ligado al ID (`:136, 207`).
- Contenido de `ESTADOCalculado` (`classes/ExpedienteCompleto.cls:321`) y de la ficha `Form_FormEleccionTipoConsulta`.
- Si `Avance` (`:94`) tiene efecto: el formulario no declara `lblEstado` en el `.form.txt`.
- Referencias dinámicas: `OpenArgs` con nombre de lista y formulario consumido por `Form_FormEleccionTipoConsulta.cls:158-171`.
