# Form_FormComercialesGestion

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/forms/Form_FormComercialesGestion.cls` (416 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Listado del catálogo de comerciales con filtro por nombre, alta, edición y borrado. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): la lista se rellena por código desde la caché `Entorno.Comerciales`. La edición se delega en [Form_FormComercial](Form_FormComercial.md). Las citas son `Form_FormComercialesGestion.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_Form0BDGestorEntidades` | Opción con clave «1» → `ResolverFormDestino` → `DoCmd.OpenForm` sin `OpenArgs` | `Form0BDGestorEntidadesLogic.bas:13-14`; `Form_Form0BDGestorEntidades.cls:58-73` |

La opción y su formulario también figuran en `Entorno.OtrasOpciones` y `OtrasOpcionesFormularios` (`classes/Entorno.cls:690, 729`). Abre `FormComercial` y escucha sus eventos `Alta` y `Editado` (`:7, 372-416`). Declara un evento `Eliminar` (`:8`, lanzado en `:181`) que ningún formulario del repositorio escucha (búsqueda de `WithEvents` sin resultados). `cmdSalir` cierra sin guardar (`:12-15`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `Comercial` | texto (filtro) | `OnKeyDown` | Con Intro ejecuta `ComandoBuscar_Click` y anula la tecla | `:245-261`, `.form.txt:953-955` |
| `ComandoBuscar` (rótulo «Comando36») | botón | `OnClick` | `Filtrar` | `:120-144`, `.form.txt:992-994` |
| `ComandoLimpiar` (rótulo «Eliminar») | botón | `OnClick` | Vacía el filtro y busca | `:198-201`, `.form.txt:1057-1059` |
| `ListaFiltrados` | lista, 2 columnas (ID oculto, `ColumnWidths ="0;9072"`) | `OnClick` | Carga el seleccionado y habilita botones | `:303-360`, `.form.txt:903-919` |
| `ListaFiltrados` | — | `OnDblClick` | Si `ComandoEditar` está habilitado, lo ejecuta | `:363-370` |
| `ComandoAlta` | botón | `OnClick` | Abre `FormComercial` en alta | `:63-94`, `.form.txt:1268-1270` |
| `ComandoEditar` (deshabilitado en diseño) | botón | `OnClick` | Abre `FormComercial` con el seleccionado | `:203-243`, `.form.txt:1152-1165` |
| `ComandoEliminar` (deshabilitado en diseño) | botón | `OnClick` | Confirma y borra | `:146-196`, `.form.txt:1357-1370` |
| `ComandoAyuda` (rótulo «Salir» en `.form.txt`) | botón | `OnClick` | `AbrirAyuda` | `:96-118`, `.form.txt:806-808` |
| `cmdSalir` | botón | `OnClick` | Cierra | `:12-15`, `.form.txt:1557-1559` |

## Flujos

1. **Apertura** (`Form_Open`, `:264-300`): deshabilita Alta, Editar y Eliminar (`:274-276`), ejecuta `Filtrar` (`:278`) y habilita `ComandoAlta` solo si `EsAdministrador = Sí` (`:282-285`).
2. **Filtrar** (`:18-61`): vacía la lista (`:32`), recorre la caché (`:34, 42-52`), descarta los que no contienen el texto del filtro (`InStr`, `:44-48`), añade `ID;nombre` sustituyendo `;` por `:` (`:49`) y selecciona la fila de índice 1 (`:53-54`).
3. **Selección** (`:303-360`): relee el comercial por ID solo si cambia la selección (`:321-335`); habilita `ComandoEditar` para todos y `ComandoEliminar` solo para administradores (`:342-345`).
4. **Alta / edición** (`:63-94`, `:203-243`): fija la global `m_ObjComercialActivo` (a `Nothing` o al seleccionado, `:72, 221`), cierra `FormComercial` si estaba abierto y lo abre. Al recibir `Alta` o `Editado` invalida la caché y vuelve a buscar (`:379-380, 402-403`).
5. **Borrado** (`:146-196`): pregunta «¿Desea realmente borrar al comercial seleccionado?» con «No» por defecto (`:167-170`), llama a `ComercialOperaciones.Eliminar` (`:174-180`) y vuelve a buscar (`:183`).

**Errores:** mismo patrón que [Form_FormComercial](Form_FormComercial.md#errores); sin selección, «Se ha de seleccionar un Elemento de la lista» (`:159-162, 216-219`).

## Reglas y escrituras (`classes/ComercialOperaciones.cls`)

| Regla u operación | Detalle | Evidencia |
|---|---|---|
| Comercial estructural | Nombre «N/A»: «Es una opción Estructural que otras herramientas necesitan.No se puede borrar» | `:115-118` |
| En uso | Si aparece en `TbExpedientesComerciales`, lista los `CodExp` y no borra | `:119-135`; `constructor.bas:2476-2479` |
| Borrado | `DELETE * FROM TbComerciales WHERE IDComercial=<id>` (SQL dinámico, sin transacción) | `:136-139` |
| Lista | `SELECT * FROM TbComerciales ORDER BY Comercial` en caché | `constructor.bas:2674-2676`; `Entorno.cls:778-792` |

El patrón (filtro, lista de dos columnas, alta/edición por global y `WithEvents`) se repite en los demás formularios `*Gestion` de catálogo listados en `Entorno.cls:729-738`; esta ficha solo describe su instancia.

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: `DELETE *` y SQL concatenado de Jet; la integridad frente a expedientes se comprueba en VBA, no por clave foránea.

## Posibles defectos · Propuestas web pendientes de decisión

**Posibles defectos (sin corregir, a verificar):**

- `Filtrar` comprueba `p_Error` tras leer la caché (`:34-37`), pero `Entorno.Comerciales` deja el error en `m_ObjEntorno.Error` (`Entorno.cls:799`): un fallo de lectura mostraría una lista vacía sin aviso.
- `lst.Selected(1)` (`:53`) selecciona la segunda fila; con un único resultado podría fallar o no seleccionar nada.
- `AddItem` protege `;` pero no `,` (`:49`): un nombre con coma podría partirse en columnas.
- La protección de «N/A» depende del nombre (`ComercialOperaciones.cls:115`): si se renombra, el comercial estructural se podría borrar.
- El aviso «en uso» solo mira `TbExpedientesComerciales` (`constructor.bas:2477`); otras referencias, si existen, no se comprueban.
- El manejador de `ComandoAlta_Click` rotula sus errores como «ComandoEditar_Click» (`:87`).
- Rótulos de diseño heredados: ventana «GESTIÓN DE LOS PEDIDOS» (`.form.txt:36`, sobrescrita en `:277`), `ComandoLimpiar` «Eliminar» (`.form.txt:1058`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: marcar «N/A» como estructural por atributo, no por nombre; clave foránea con `RESTRICT` en lugar de la comprobación en VBA; un único componente de gestión de catálogo para toda la familia.

## Criterios de aceptación

Propuestos, no aprobados:

1. Solo un administrador ve habilitados Alta y Eliminar; cualquier usuario puede abrir un comercial en consulta.
2. Filtrar por un texto muestra solo los comerciales cuyo nombre lo contiene; limpiar el filtro muestra todos.
3. Borrar un comercial usado en expedientes muestra la lista de códigos y no borra.
4. «N/A» no se puede borrar.
5. Tras un alta o una edición la lista se refresca.

## Pendientes

- Si `ComandoBuscar` y `ComandoLimpiar` muestran un icono en ejecución (rótulos de diseño «Comando36» y «Eliminar»).
- Si la comparación de `InStr` (`:45`) distingue tildes con `Option Compare Database` (`:2`).
- Límite de longitud de la lista de valores con catálogos grandes.
