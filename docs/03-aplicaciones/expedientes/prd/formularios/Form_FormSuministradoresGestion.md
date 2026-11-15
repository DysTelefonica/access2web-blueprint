# Form_FormSuministradoresGestion

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/forms/Form_FormSuministradoresGestion.cls` (459 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Listado del catálogo de empresas (`TbSuministradores`; título «EMPRESAS», `.form.txt:480-481`) con filtro por nombre o CIF, alta, edición y borrado. Es un formulario **no enlazado**: la lista se rellena desde la caché `m_ObjEntorno.Suministradores`. Alta y edición se delegan en [Form_FormSuministrador](Form_FormSuministrador.md). Sigue el patrón de gestor de catálogo descrito en [Form_FormEjercitosGestion](Form_FormEjercitosGestion.md); aquí solo se detalla lo propio. Las citas son `Form_FormSuministradoresGestion.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_Form0BDGestorEntidades` (opción 4) | `ResolverFormDestino` devuelve `FormSuministradoresGestion`; se abre sin `OpenArgs` | `Form0BDGestorEntidadesLogic.bas:19-20`; `Form_Form0BDGestorEntidades.cls:58-71` |
| `Form_FormExpedienteSuministradores` («Añadir Aquí...») | Lo abre en diálogo y lee el elegido de `TempVars("SelectedIDSuministrador")` | `Form_FormExpedienteSuministradores.cls:560-564` |

Abre `FormSuministrador` en alta (`:76-84`) o edición (`:227-235`) y escucha `Alta` y `Editado` con `WithEvents` (`:7, 398-444`). Declara `Seleccionado`, `Eliminar`, `Alta` y `Editado` (`:5, 8-10`); no se ha encontrado ningún `WithEvents` sobre este formulario.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `PalabraClave` | texto (filtro) | `OnKeyDown` | Con Intro busca y anula la tecla | `.form.txt:647-649`; `:446-459` |
| `ComandoBuscar` | botón | `OnClick` | `Filtrar` | `.form.txt:687-691`; `:124-148` |
| `ComandoLimpiar` (rótulo «Eliminar», tip «Limpiar campo Búsqueda») | botón | `OnClick` | Vacía el filtro y busca | `.form.txt:751-755`; `:203-206` |
| `ListaFiltrados` | lista 3 columnas (ID oculto, CIF, Nombre) | `OnClick`, `OnDblClick` | Selecciona; doble clic equivale a Editar | `.form.txt:591-614`; `:328-396` |
| `ComandoEditar` | botón, deshabilitado en diseño | `OnClick` | Abre `FormSuministrador` con el seleccionado | `.form.txt:846-860`; `:208-249` |
| `ComandoAlta` | botón | `OnClick` | Abre `FormSuministrador` en alta | `.form.txt:961-965`; `:67-98` |
| `ComandoEliminar` | botón, deshabilitado en diseño | `OnClick` | Borra el seleccionado tras confirmar | `.form.txt:1050-1065`; `:150-201` |
| `ComandoSeleccionar` | botón **oculto** en diseño | `OnClick` | Lanza `Seleccionado` y cierra | `.form.txt:1347-1360`; `:252-286` |
| `ComandoAyuda`, `cmdSalir` | botones | `OnClick` | Ayuda y cierre | `:100-122, 14-17` |

## Flujos

### Apertura y filtrado (`Form_Open`, `:288-325`; `Filtrar`, `:20-65`)

1. Deshabilita Alta, Editar y Eliminar; con `EsAdministrador = Sí` habilita Alta (`:299-306`). `Filtrar` carga la lista; un error cierra el formulario (`:307-324`).
2. `Filtrar` vacía la selección y deja la fila de encabezados `IDSuministrador;CIF;Nombre` (`:30-34`). Recorre la caché, ordenada por nombre (`classes/Entorno.cls:1016-1030`; `constructor.bas:3013-3014`).
3. Con filtro, conserva las empresas cuyo **nombre o CIF** lo contienen (`:46-51`). Sustituye `;` por `:` (`:52-53`), selecciona la primera fila y ejecuta `ListaFiltrados_Click` (`:57-58`).

### Selección, alta y edición

Sin ID limpia la selección; si cambia el ID relee con `getSuministrador` (`:339-360`). Habilita Editar para todos y Eliminar solo para administradores (`:367-370`). Al recibir `Alta` o `Editado`, vuelve a buscar y reenvía el evento (`:406-407, 430-431`); la caché ya la invalidó `Registrar` (`SuministradorOperaciones.cls:157`).

### Borrado (`ComandoEliminar_Click`, `:150-201`)

1. Sin selección, «Se ha de seleccionar un Elemento de la lista» (`:159-168`). Confirmación con «No» por defecto (`:172-175`).
2. `SuministradorOperaciones.Eliminar`: si `MotivoEliminarNoOK` devuelve motivo («Este Suministrador ha sido usado en algún Expediente»), rechaza; si no, `DELETE` e invalida la caché (`SuministradorOperaciones.cls:199-205, 228-241`).
3. Lanza `Eliminar` y vuelve a buscar (`:186-188`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Listado | `TbSuministradores` (caché de `Entorno`) | SQL dinámico | no | `constructor.bas:3013-3014` |
| Selección | `TbSuministradores` por ID | SQL dinámico concatenado | no | `constructor.bas:5573-5576` |
| Uso en expedientes | `TbSuministradores` ⨝ `TbExpedientesSuministradores` | SQL dinámico concatenado | no | `constructor.bas:3068-3071` |
| Borrado | `DELETE` en `TbSuministradores` | SQL dinámico (`Execute`) | no | `SuministradorOperaciones.cls:237-240` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: comprobación de uso y `DELETE` sin transacción; la integridad con `TbExpedientesSuministradores` la aplica el código (a confirmar en el origen del modelo, que no está en el repositorio).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `MotivoEliminarNoOK` asigna con `Set` el `Dictionary` de `getSuministradoresEnExpedientes` a una variable `Suministrador` (`SuministradorOperaciones.cls:184, 199`; `constructor.bas:3056-3059`). Con la empresa en uso daría error de tipo: el borrado se rechazaría con un mensaje técnico, no con «ha sido usado en algún Expediente».
- `CadenaHPSUsuariosParaSuministrador` y su versión histórica, que comprueban el uso en `TbUsuarios_HPS`, son privadas y no se llaman (`SuministradorOperaciones.cls:254, 319`): el borrado no mira HPS.
- «Añadir Aquí...» espera un `TempVar` que este formulario no asigna y `ComandoSeleccionar` está oculto (`.form.txt:1347`); ya recogido en [Form_FormExpedienteSuministradores](Form_FormExpedienteSuministradores.md).
- Restos copiados: aviso «Obteniendo los RACes Filtrados» (`:36`), `Caption` «GESTIÓN DE LOS PEDIDOS» (`.form.txt:36`) y `RowSource` de estados (`.form.txt:610`, sustituido en `:34`).
- Patrón de familia ya descrito en [Form_FormEjercitosGestion](Form_FormEjercitosGestion.md): error de caché sin aviso (`:37-40`) y `Selected(1)` con lista vacía (`:57`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: borrado protegido por clave foránea; selector reutilizable para la pestaña de empresas del expediente.

## Criterios de aceptación

Propuestos, no aprobados:

1. Solo un administrador puede dar de alta o eliminar; cualquier usuario puede abrir una ficha.
2. El filtro encuentra empresas por parte del nombre o del CIF; Limpiar las muestra todas.
3. Eliminar pide confirmación con «No» por defecto.
4. No se puede eliminar una empresa usada en un expediente y se informa del motivo.
5. Tras un alta, edición o borrado, el listado refleja el cambio.

## Pendientes

- Comportamiento real de `MotivoEliminarNoOK` con una empresa en uso (prueba en el binario).
- Si el uso en `TbUsuarios_HPS` debe impedir el borrado.
- Si `InStr` del filtro distingue mayúsculas con `Option Compare Database` (`:2, 47-48`).
