# Form_FormExpedientesBusqueda

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Fuente:** `src/forms/Form_FormExpedientesBusqueda.cls` (424 líneas) y `src/forms/FormExpedientesBusqueda.form.txt` · **PRD:** [README](../README.md)

Buscador de expedientes. Abre una lista de expedientes de la aplicación **ajena** de Expedientes, deja filtrar por palabra clave y por responsable de calidad, y devuelve el expediente elegido a quien lo abrió. También ofrece abrir la gestión de riesgos asociada al expediente seleccionado, si la hay. No escribe nada. Las citas son `Form_FormExpedientesBusqueda.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `FormGestionRiesgosDatosGenerales` | `DoCmd.OpenForm "FormExpedientesBusqueda"` tras cerrar el anterior, y suscripción al evento `Seleccionado` | `Form_FormGestionRiesgosDatosGenerales.cls:67-75` |

Abre siempre el mismo llamador, que se registra con `WithEvents` y reutiliza el formulario ya abierto (`Form_FormGestionRiesgosDatosGenerales.cls:13`). No acepta parámetros: el estado viaja en los controles del formulario y en el evento `Seleccionado` (`:18`), que entrega un objeto `Expediente`.

Salidas: `RaiseEvent Seleccionado(m_ExpSeleccionado)` y cierre de sí mismo (`:103-104`); o bien, desde `ComandoVerGR`, fija `m_ObjProyectoActivo` y `m_ObjEdicionActiva` (`Variables Globales.bas:75, 79`) y abre `FormRiesgosGestion` cerrando `FormGestionRiesgos` si estaba abierto (`:165-178`); o `ComandoSalir`, que solo cierra (`:140-143`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `Form_Open` | Ajusta tamaño, pone el título con el usuario conectado, fija las cabeceras de la lista y rellena el combo de responsables | `:231-262`, `.form.txt:41` (`OnOpen`), `.form.txt:2169` |
| `ListaFiltrados` | lista | `ListaFiltrados_Click` | Resuelve el expediente, muestra el título y decide si hay gestión de riesgos | `:313-387`, `.form.txt:1069` |
| `PalabraClave` | texto | `PalabraClave_Exit` | Refiltra solo si el texto cambió | `:390-422`, `.form.txt:1276` |
| `IDResponsableCalidad` | combo | `IDResponsableCalidad_BeforeUpdate` | Refiltra solo si cambió la selección | `:197-228`, `.form.txt:1241` |
| `Titulo` | etiqueta | — | Muestra el título del expediente elegido | `:350, 384`, `.form.txt:1667` |
| `lblRiesgosRegistrados` | etiqueta | — | Anuncia si el expediente ya tiene gestión de riesgos | `:353-364`, `.form.txt:1795` |
| `ComandoElegir` | botón | `ComandoElegir_Click` | Devuelve el expediente al llamador y cierra | `:83-117`, `.form.txt:1756` |
| `ComandoVerGR` | botón | `ComandoVerGR_Click` | Abre la gestión de riesgos del expediente | `:146-195`, `.form.txt:1826` |
| `ComandoActualizar` | botón | `ComandoActualizar_Click` | Vacía las memorizadas y vuelve a filtrar | `:22-53`, `.form.txt:920` |
| `ComandoAyuda` | botón | `ComandoAyuda_Click` | Abre la ayuda | `:55-81`, `.form.txt:987` |
| `ComandoLimpiarPalabraClave`, `ComandoLimpiarIDResponsableCalidad` | botones | `_Click` | Vacían el control y refiltran si había algo | `:120-138`, `.form.txt:1303, 1482` |
| `ComandoSalir` | botón | `ComandoSalir_Click` | Cierra sin devolver nada | `:140-144`, `.form.txt:1872` |

## Flujos

### Apertura (`Form_Open`, `:231-262`)

1. `AjustarTamaño Me` y `Me.Caption = m_ObjEntorno.TituloUsuarioConectado` (`:237-238`).
2. `Me.ListaFiltrados.RowSource = "IDExp;Cod_Exp;Nemotécnico;Título"`: fija las cabeceras, no una consulta (`.form.txt:1070`).
3. `RellenarComboRespCalidad` (`Funciones Generales.bas:2265`) carga el combo desde `m_ObjEntorno.ColUsuariosCalidad` en memoria; no consulta la base de datos.

No filtra en apertura: la lista nace vacía hasta que el usuario actúa o el llamador invoca `Filtrar` (`Form_FormGestionRiesgosDatosGenerales.cls:75`).

### Normal (filtrar y elegir)

1. `Filtrar` (`:266-311`) limpia el estado, pone de nuevo las cabeceras y llama a `Constructor.getExpedientesBusqueda(palabraClave, responsable, p_Error)` (`:286`).
2. Esa función (`Constructor.bas:5908`) abre `TbExpedientes` en la base **ajena** con `getdbExpedientes()` (`Constructor.bas:5929`), monta un `Expediente` por fila (`Constructor.bas:5934-5937`) y **descarta en memoria** las que no casan: descarta si la palabra clave no aparece ni como identificador, ni en el nemotécnico, ni en el título (`Constructor.bas:5941-5946`). El filtro por responsable, en cambio, va en el `WHERE` (`Constructor.bas:5920-5927`). Lo que sobrevive se guarda en un diccionario por identificador (`Constructor.bas:5949-5955`).
3. `Filtrar` vuelca el diccionario en la lista como lista de valores, sustituyendo el punto y coma del título por dos puntos (`Replace(.Titulo, ";", ":")`, `:295-297`), selecciona la primera fila (`:301`) y llama a `ListaFiltrados_Click` (`:302`).
4. `ListaFiltrados_Click` (`:313-387`) recarga el expediente si cambió de fila (`:337-348`), pide `Constructor.getProyectoPorExpediente` (`:356`) y, según haya proyecto o no, muestra `ComandoVerGR` con el texto «Gestión de riesgos ya registrada» o «Sin gestión de riesgos» (`:358-364`).
5. `ComandoElegir_Click` devuelve el expediente y se cierra (`:103-104`).

### Alternativo (ver la gestión de riesgos)

`ComandoVerGR_Click` (`:146-195`) resuelve el proyecto si aún no lo tiene, elige edición activa o última (`:167-171`) y salta a `FormRiesgosGestion` cerrando antes `FormGestionRiesgos` si estuviera abierto (`:173-179`).

### De error

Ocho manejadores repiten el mismo patrón: `Err.Raise 1000` con `m_Error` ya escrito; en el manejador, número distinto de 1000 se traduce a `CorreoAlministrador` más `MsgBox` crítico, y 1000 se muestra como advertencia con el texto que traía `m_Error` (`:41-51`, y equivalentes). La excepción es `Form_Open`, que además cierra el formulario (`:255-262`), y los tres botones de limpiar y salir, que usan `On Error Resume Next` sin aviso (`:121, 131, 141`).

## Reglas, validaciones y cálculos

- **El descarte de filas no depende del motor de datos.** La base ajena solo filtra por responsable; el texto libre se resuelve recorriendo todas las filas en memoria (`Constructor.bas:5939-5943`). Cada pulsación de «Actualizar» vuelve a traer la tabla entera.
- **El origen de los datos se decide por TempVar.** `getdbExpedientes` lee `DatosEnLocal`: si es «Sí» toma la copia local, si es «No» la remota, y cualquier otro valor es error «No se conoce el origen de los datos» (`Variables Globales.bas:518-526`). En ambos casos abre el mismo fichero con una contraseña incrustada en la llamada (`:529`).
- **Cada pulsación de filtro revalida contra la base de datos** solo a través de `m_PalabraClaveUltima` e `m_IDResponsableCalidadUltimo` (`:33-34`, `:194, 215`): son la única caché, y solo de la última entrada.
- **`Filtrar` es público y sin parámetros obligatorios**: el llamador lo invoca sin argumentos (`Form_FormGestionRiesgosDatosGenerales.cls:75`), y entonces el filtro actúa solo sobre lo que ya hay en los controles.
- **Los dos desplegables filtran por cadenas concatenadas**, no por columnas: el valor de la lista es `ID;Cod_Exp;Nemotécnico;Título` (`:239`) y los nombres se leen por posición.

## Lecturas y escrituras

| Operación | Origen | Evidencia |
|---|---|---|
| Lista de expedientes filtrable | Base ajena de Expedientes, tabla `TbExpedientes` | `Constructor.bas:5918-5929`, `Variables Globales.bas:518-530` |
| Expediente concreto de una fila | Base ajena, misma tabla | `Constructor.bas:2704-2722` |
| Proyecto asociado al expediente | Base **propia**, tabla `TbProyectos` | `Constructor.bas:736` (usa `getdb()`, no `getdbExpedientes()`) |
| Responsables de calidad | Caché de sesión `m_ObjEntorno.ColUsuariosCalidad` | `Funciones Generales.bas:2265-2284` |
| Proyecto y edición activos | Globales del proceso | `Variables Globales.bas:75, 79` |

Sin escrituras en base de datos.

## Coexistencia

Pendiente de `../coexistencia.md`. Aquí se ve el patrón que más pesa: el formulario elige entre **dos bases de datos distintas** según una TempVar del proceso (`Variables Globales.bas:520-526`) y sostiene en memoria objetos de negocio para no volver a consultar. En una web, el origen de datos es el servidor y esa TempVar desaparece; la lista, en cambio, encaja bien como consulta paginada al backend de Expedientes.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-GR-360.** «Actualizar» no filtra: pone a vacío `m_PalabraClaveUltima` y `m_IDResponsableCalidadUltimo` (`:33-34`) y después llama a `Filtrar`, que sí lee los controles. Si lo que se pretendía era invalidar la caché, la efecto es el contrario al necesario: fuerza un refiltrado manteniendo el criterio actual.
- **H-GR-361.** Cuando no hay resultados, `Filtrar` sale con `m_Col Is Nothing` (`:289-291`) sin avisar y sin tocar la lista, que conserva las filas del filtrado anterior: el usuario ve un resultado que ya no corresponde a su búsqueda.
- **H-GR-362.** `lst.Selected(1) = True` (`:301`) se ejecuta sin comprobar que la lista tenga al menos una fila; con una búsqueda sin coincidenciasAccess deja el índice fuera de rango y el error se silenciosa en `Filtrar`, que solo reescribe `p_Error` si el número no es 1000 (`:307-310`).
- **H-GR-363.** El descarte en memoria compara `p_PalabraClave` con el identificador completo (`Constructor.bas:5942`): buscar un fragmento numérico no coincide nunca con el `IDExpediente`, aunque el usuario lo espere.
- **H-GR-364.** La consulta del responsable concatena el identificador sin comillas (`Constructor.bas:5923`): depende de que el valor sea numérico, y el mismo patrón se repite en `getExpediente` (`Constructor.bas:2721`).
- **H-GR-365.** La lista descarta el detalle de las filas solo por las tres columnas montadas (`:295-297`); el resto del expediente se pierde y hay que releerlo fila a fila (`:337-348`).
- **H-GR-366.** `Err.Raise 1000` se lanza sin `vbObjectError` en los ocho manejadores (`Constructor.bas` incluido). Access puede tratarlo como error de usuario y **ejecutar el manejador de error del propio evento**, con lo que `Err.Number` valdría otra cosa y el mensaje mostraría una descripción vacía.
- **H-GR-367.** El texto del error se muestra solo porque el manejador consulta la cadena `m_Error` que ya estaba escrita, no la descripción de `Err` (`:41-51`). Es un patrón frágil que se repite en las ocho rutinas.
- **H-GR-368.** `lblRiesgosRegistrados` conserva en diseño el rótulo «Palabra clave» (`.form.txt:1796`), de donde parece copiado; solo se corrige en ejecución (`:361, 364`).
- **H-GR-369.** Varios botones guardan rótulo de diseño heredado, no su texto real: «Comando47» en actualizar (`.form.txt:921`), «Comando36» en los dos de limpiar (`.form.txt:1304, 1483`) y «Salir» tanto en ayuda (`.form.txt:988`) como en salir (`.form.txt:1873`). El título de la ventana, además, se fija en diseño con el usuario de ejemplo (`.form.txt:40`).
- **H-GR-370.** Los botones de limpiar y salir usan `On Error Resume Next` (`:121, 131, 141`) sin ninguna señal al usuario si algo falla.
- **H-GR-371.** `ComandoElegir_Click` no comprueba que la selección siga existiendo si `m_ExpSeleccionado` ya venía cacheado (`:90-101`): con la lista vacía tras un refiltrado, puede devolverse el expediente anterior.
- **H-GR-372.** `FormularioAbierto("FormRiesgosGestion")` y `FormularioAbierto("FormGestionRiesgos")` (`:173, 176`) se apoyan en la colección global de formularios de Access, que no tiene equivalente directo fuera de Access.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: mover el descarte por palabra clave a la consulta del backend y devolver solo el identificador, la columna y el título; sustituir la TempVar de origen de datos por un parámetro de conexión del servidor, dejando de existir la copia local; paginar la lista en lugar de traerla entera; y resolver el expediente y su proyecto en una sola llamada en vez de dos, para que elegir una fila no dependa de una segunda consulta sin caché.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir el buscador muestra la lista de expedientes del backend de Expedientes y el usuario conectado.
2. Escribir una palabra clave refiltra por identificador, nemotécnico y título.
3. Elegir responsable de calidad refiltra la lista por ese responsable.
4. Seleccionar una fila muestra su título e indica si ya tiene gestión de riesgos.
5. Elegir devuelve el expediente al llamador; ver gestión de riesgos abre el formulario de riesgos del expediente.
6. Si no hay coincidencias, la lista queda vacía y se avisa; si el backend no responde, se avisa y no se deja el cursor en modo espera.

## Pendientes

- Quién fija `DatosEnLocal` y en qué momento del arranque; si la copia local y la remota pueden coexistir en la misma sesión.
- Si el descarte por palabra clave es intencionadamente fiel a una versión anterior que filtraba en SQL, o es una regresión.
- El contenido de la ayuda que abre `ComandoAyuda_Click` (`Funciones Generales.bas:5327`).
- La periodicidad real de «Actualizar»: en el código no hay caché real, solo las dos cadenas de la última entrada.