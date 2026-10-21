# Form_FormOficinasProgramaGestion

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/forms/Form_FormOficinasProgramaGestion.cls` (425 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Listado del catálogo de oficinas de programa con filtro por nombre, alta, edición y borrado. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): la lista se rellena por código desde la caché `m_ObjEntorno.OficinasDePrograma`. La edición se delega en [Form_FormOficinaPrograma](Form_FormOficinaPrograma.md). Sigue el patrón común de los `*Gestion` de catálogo descrito en [Form_FormComercialesGestion](Form_FormComercialesGestion.md); esta ficha solo recoge lo propio. Las citas son `Form_FormOficinasProgramaGestion.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_Form0BDGestorEntidades` | Opción con clave «6» → `ResolverFormDestino` → `DoCmd.OpenForm` | `Form0BDGestorEntidadesLogic.bas:23-24`; `Form_Form0BDGestorEntidades.cls:58-73` |

La opción figura también en `OtrasOpcionesFormularios` (`classes/Entorno.cls:735`). Abre `FormOficinaPrograma` y escucha `Alta` y `Editado` (`:7, 381-425`). Declara un evento `Eliminar` (`:8`, lanzado en `:183`) sin oyentes en el repositorio (ningún `WithEvents` de este formulario). `cmdSalir` cierra (`:12-15`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `OficinaPrograma` | texto (filtro) | `OnExit` | Al salir del campo ejecuta `ComandoBuscar_Click` | `:248-270`, `.form.txt:642-643` |
| `ComandoBuscar` (rótulo «Comando36») | botón | `OnClick` | `Filtrar` | `:122-146`, `.form.txt:682-684` |
| `ComandoLimpiar` (rótulo «Eliminar») | botón | `OnClick` | Vacía el filtro y busca | `:200-203`, `.form.txt:746-748` |
| `ListaFiltrados` | lista, 3 columnas con encabezados (ID oculto, `ColumnWidths ="0;2835;6237"`) | `OnClick` | Carga la seleccionada y habilita botones | `:312-369`, `.form.txt:586-609` |
| `ListaFiltrados` | — | `OnDblClick` | Si `ComandoEditar` está habilitado, lo ejecuta | `:372-379` |
| `ComandoAlta` | botón | `OnClick` | Abre `FormOficinaPrograma` en alta | `:65-96`, `.form.txt:956-958` |
| `ComandoEditar` (deshabilitado en diseño) | botón | `OnClick` | Abre `FormOficinaPrograma` con la seleccionada | `:205-245`, `.form.txt:841-853` |
| `ComandoEliminar` (deshabilitado en diseño) | botón | `OnClick` | Confirma y borra | `:148-198`, `.form.txt:1045-1058` |
| `ComandoAyuda` (rótulo «Salir») | botón | `OnClick` | `AbrirAyuda` | `:98-120`, `.form.txt:495-499` |
| `cmdSalir` | botón | `OnClick` | Cierra | `:12-15`, `.form.txt:1246-1248` |

## Flujos

1. **Apertura** (`Form_Open`, `:272-309`): deshabilita Alta, Editar y Eliminar (`:283-285`), habilita y enfoca `ComandoAlta` solo si `EsAdministrador = Sí` (`:287-290`) y ejecuta `Filtrar` (`:291`).
2. **Filtrar** (`:18-63`): reinicia la lista con la fila de encabezado `IDOficinaPrograma;OFICINA;DESCRIPCIÓN` (`:32`), recorre la caché (`:35, 43-54`), descarta las que no contienen el filtro (`InStr`, `:45-49`), añade `ID;nombre;descripción` sustituyendo `;` por `:` (`:50-51`) y selecciona la primera fila de datos (`:55-56`).
3. **Selección** (`:312-369`): relee la oficina por ID solo si cambia (`:330-344`); habilita `ComandoEditar` para todos y `ComandoEliminar` solo para administradores (`:351-354`).
4. **Alta / edición** (`:65-96`, `:205-245`): fija la global `m_ObjOficinaProgramaActiva` (`:74, 223`), cierra el formulario de edición si estaba abierto y lo abre. Al recibir `Alta` o `Editado` invalida la caché y vuelve a buscar (`:388-389, 411-412`).
5. **Borrado** (`:148-198`): «¿Desea realmente borrar la Oficina de Programa seleccionada?» con «No» por defecto (`:169-172`), `OficinaProgramaOperaciones.Eliminar` (`:176-182`) y nueva búsqueda (`:185`).

**Errores:** patrón de la familia; sin selección, «Se ha de seleccionar un Elemento de la lista» (`:161-164, 218-221`).

## Reglas y escrituras (`classes/OficinaProgramaOperaciones.cls`)

| Regla u operación | Detalle | Evidencia |
|---|---|---|
| Oficina estructural | Nombre «N/A»: «Es una opción Estructural que otras herramientas necesitan.No se puede borrar» | `:108-111` |
| En uso | Si la caché `ColOficinaProgramaSEnExpedientes` contiene el ID: «Al menos un expediente usa este OficinaPrograma» | `:112-117`; `Entorno.cls:1222-1243` |
| Borrado | `DELETE * FROM TbOficinasPrograma WHERE IDOficinaPrograma=<id>` (SQL dinámico, sin transacción) e invalida la caché | `:118-123` |
| Lista | `SELECT * FROM TbOficinasPrograma` sin `ORDER BY`, en diccionario por ID | `constructor.bas:3279-3308`; `Entorno.cls:1075-1097` |

## Coexistencia

Pendiente de `../coexistencia.md`. `DELETE *` y SQL concatenado de Jet; la integridad frente a expedientes se comprueba en VBA, no por clave foránea; la lista sale en el orden que devuelva el motor, sin `ORDER BY`.

## Posibles defectos · Propuestas web pendientes de decisión

**Posibles defectos (sin corregir, a verificar):**

- **La comprobación «en uso» no se ejecuta:** `Entorno.cls:1231` pasa `m_Error` (vacío) como primer argumento posicional, que es `p_IDOficinaPrograma` (`constructor.bas:3544-3546`); la función sale sin datos (`:3556-3557`), la caché queda en `Nothing` y `Eliminar` salta la comprobación (`OficinaProgramaOperaciones.cls:112`). Una oficina usada por expedientes se borraría.
- Aunque recibiera el ID, la consulta une por nombre (`TbExpedientes.OficinaPrograma = TbOficinasPrograma.OficinaPrograma`, `constructor.bas:3561`), mientras que `TbExpedientes` guarda `IDOficinaPrograma` (`ExpedienteOperaciones.cls:203-206`).
- La caché «en uso» nunca se invalida tras altas o borrados en este formulario (solo se invalida `OficinasDePrograma`, `:388, 411`; `OficinaProgramaOperaciones.cls:123`).
- Sin `ORDER BY`, el orden de la lista no está garantizado (`constructor.bas:3280-3281`).
- Patrones de familia ya descritos en [Form_FormComercialesGestion](Form_FormComercialesGestion.md#posibles-defectos--propuestas-web-pendientes-de-decisión): `p_Error` comprobado en lugar de `m_ObjEntorno.Error` (`:35-38`; `Entorno.cls:1096`), `Selected(1)` con lista vacía (`:55`), protección de «N/A» por nombre, `Eliminar` sin oyentes y rótulos heredados (`.form.txt:33, 605, 747`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata propia: clave foránea `TbExpedientes.IDOficinaPrograma` con `RESTRICT` en lugar de la caché «en uso».

## Criterios de aceptación

Propuestos, no aprobados:

1. Solo un administrador ve habilitados Alta y Eliminar; cualquiera puede abrir una oficina en consulta.
2. Salir del campo de filtro filtra por nombre; limpiar muestra todas.
3. Borrar una oficina usada por al menos un expediente se rechaza y no borra.
4. «N/A» no se puede borrar.
5. La lista sale ordenada por nombre y se refresca tras un alta, una edición o un borrado.

## Pendientes

- Si `TbExpedientes` tiene una columna `OficinaPrograma` (requiere el ERD de origen, `erd-origen.md`).
- Si el backend declara relación entre `TbExpedientes.IDOficinaPrograma` y `TbOficinasPrograma` que impida el borrado.
- Si `InStr` (`:46`) distingue mayúsculas y tildes con `Option Compare Database` (`:2`).
