# Form_FormLugarEjecucionGestion

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/forms/Form_FormLugarEjecucionGestion.cls` (416 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Listado, búsqueda, alta, edición y borrado de lugares de ejecución. Es un formulario **no enlazado**: la lista se rellena por código desde la caché `m_ObjEntorno.LugaresEjecucion` (`classes/Entorno.cls:1045-1067`). Sigue el patrón común de las gestiones de catálogo (misma estructura en `Form_FormCPVsGestion.cls` y `Form_FormEjercitosGestion.cls`, p. ej. `:54` y `:182` en ambos); esta ficha cita solo sus propias líneas. Las citas son `Form_FormLugarEjecucionGestion.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_Form0BDGestorEntidades` | opción «5» de la lista de otras opciones → `DoCmd.OpenForm` | `Form0BDGestorEntidadesLogic.bas:21-22`; `classes/Entorno.cls:734`; `Form_Form0BDGestorEntidades.cls:58-73` |

Abre [Form_FormLugarEjecucion](Form_FormLugarEjecucion.md) para alta (`:73-81`) y edición (`:222-230`) y escucha sus eventos `Alta` y `Editado` con `WithEvents` (`:7, 373-417`). Declara y lanza un evento `Eliminar` (`:8, 182`) sin suscriptores encontrados en `src`. `cmdSalir` cierra (`:12-15`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `LUGAREJECUCION` | texto de búsqueda (`Tag = "BUSQUEDA"`) | `OnKeyDown` | Intro lanza la búsqueda y anula la tecla | `:358-371`, `.form.txt:939-956` |
| `ComandoBuscar` | botón | `OnClick` | `Filtrar` | `:121-145`, `.form.txt:983-995` |
| `ComandoLimpiar` (rótulo «Eliminar» en `.form.txt`) | botón | `OnClick` | Vacía el filtro y busca | `:199-202`, `.form.txt:1046-1059` |
| `ListaFiltrados` | lista 2 columnas (ID oculto, cabeceras) | `OnClick`, `OnDblClick` | Selecciona el lugar; doble clic edita | `:288-356`, `.form.txt:897-920` |
| `ComandoAlta` | botón | `OnClick` | Abre el formulario de alta | `:64-95`, `.form.txt:1256-1269` |
| `ComandoEditar` | botón | `OnClick` | Abre el formulario con el lugar seleccionado | `:204-244`, `.form.txt:1151-1164` |
| `ComandoEliminar` | botón | `OnClick` | Confirma y borra | `:147-197`, `.form.txt:1355-1369` |
| `ComandoAyuda` (rótulo «Salir» en `.form.txt`) | botón | `OnClick` | `AbrirAyuda` (`FUNCIONES UTILES.bas:3076`) | `:97-119`, `.form.txt:796-808` |

Valores de diseño que el código sustituye: `Caption` del formulario (`.form.txt:36`, sustituido en `:262`) y `RowSource` de la lista (`.form.txt:916`, sustituido en `:32`).

## Flujos

### Apertura y búsqueda

1. `Form_Open`: `Ajustar Me`; `ComandoAlta`, `ComandoEditar` y `ComandoEliminar` deshabilitados; solo `EsAdministrador = Sí` habilita `ComandoAlta` (`:257-266`). Llama a `Filtrar` (`:267`); un error cierra el formulario (`:277-284`).
2. `Filtrar` (`:18-62`): vacía la selección, deshabilita editar y eliminar, deja solo la cabecera de la lista (`:28-32`), recorre la caché y añade `ID;nombre` de cada lugar cuyo nombre contenga el texto del filtro, con `;` del nombre cambiado a `:` (`:43-53`). Selecciona la primera fila y simula su clic (`:54-55`).
3. `ListaFiltrados_Click` (`:288-346`): relee el lugar seleccionado con `constructor.getLugarEjecucion` si cambió (`:306-320`); habilita `ComandoEditar` para cualquier usuario y `ComandoEliminar` solo para administradores (`:327-330`).

### Alta y edición

`ComandoAlta` vacía `m_ObjLugarEjecucionActiva` (`:73`); `ComandoEditar` la apunta al seleccionado (`:222`). Ambos cierran el formulario de detalle si estaba abierto y lo abren (`:75-81, 224-230`). Un no administrador puede abrir la edición, pero allí el guardado está deshabilitado (`Form_FormLugarEjecucion.cls:175-179`). Al recibir `Alta` o `Editado`, invalida la caché y repite la búsqueda (`:380-381, 403-404`).

### Borrado (`ComandoEliminar_Click`, `:147-197`)

1. Sin selección: «Se ha de seleccionar un Elemento de la lista» (`:155-163`).
2. Confirmación «¿Desea realmente borrar el lugar de ejecución seleccionado?», con «No» por defecto (`:168-171`).
3. `LugarEjecucionOperaciones.Eliminar` (`:175-183`): rechaza el lugar «N/A» por estructural (`LugarEjecucionOperaciones.cls:112-115`) y el que esté en algún expediente, listando sus `CodExp` (`LugarEjecucionOperaciones.cls:116-132`); si no, `DELETE` (`LugarEjecucionOperaciones.cls:133-136`).
4. Repite la búsqueda (`:184`). Errores: correo salvo el 1000 y `MsgBox` crítico (`:189-196`).

## Reglas, validaciones y cálculos

- **Filtro:** `InStr` sobre el nombre, sin buscar en la descripción; la comparación sigue `Option Compare Database` (`:2, 46`).
- **Permisos:** alta y borrado solo para administradores; edición abierta a todos, en modo consulta para no administradores (`:263-266, 327-330`).
- **Integridad al borrar:** comprobada en código contra `TbExpedientesLugaresEjecucion` (`constructor.bas:3504-3507`).

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Listado (caché de sesión) | `TbLugaresEjecucion` ordenado por nombre | SQL dinámico | — | `constructor.bas:3222-3224` |
| Lectura del seleccionado | `TbLugaresEjecucion` | SQL dinámico concatenado | — | `constructor.bas:636-639` |
| Expedientes que lo usan | `TbExpedientes` + `TbExpedientesLugaresEjecucion` | SQL dinámico | — | `constructor.bas:3504-3507` |
| Borrado | `TbLugaresEjecucion` | `DELETE` con `Execute` | no | `LugarEjecucionOperaciones.cls:133-136` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: comprobación de uso y `DELETE` sin transacción ni clave foránea declarada conocida; caché de sesión que no ve cambios de otros usuarios.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `Eliminar` solo invalida la caché en su rama de error (`LugarEjecucionOperaciones.cls:140-146`): tras un borrado correcto, la búsqueda (`:184`) podría seguir mostrando el lugar borrado.
- `Filtrar` comprueba `p_Error` tras leer la caché, pero `Entorno.LugaresEjecucion` deja su error en `.Error` (`:35-38`; `classes/Entorno.cls:1061-1066`): un fallo de lectura se vería como lista vacía.
- `lst.Selected(1) = True` sin comprobar que haya filas (`:54`): con un filtro sin resultados podría fallar.
- La comprobación de uso y el `DELETE` no son atómicos (`LugarEjecucionOperaciones.cls:116-136`).
- Rótulos de diseño incoherentes (`ComandoLimpiar` «Eliminar», `ComandoAyuda` «Salir», `ComandoBuscar` «Comando36»; `.form.txt:994, 1058, 807`); no se ha comprobado si son botones con imagen.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: búsqueda en servidor sobre nombre y descripción; clave foránea en lugar de la comprobación en código; marcar «N/A» como registro protegido por dato y no por texto.

## Criterios de aceptación

Propuestos, no aprobados:

1. Escribir parte de un nombre y pulsar Intro lista solo los lugares que lo contienen.
2. Un no administrador no ve habilitados alta ni borrado.
3. Borrar un lugar usado en expedientes lo rechaza indicando sus códigos.
4. «N/A» no se puede borrar.
5. Tras un alta, edición o borrado, la lista refleja el cambio.

## Pendientes

- Si existe relación declarada entre `TbExpedientesLugaresEjecucion` y `TbLugaresEjecucion` (`erd-origen.md`).
- Comportamiento real de `Selected(1)` con la lista sin filas y del rótulo de los botones en ejecución.
- Contenido de `Ajustar` (`FUNCIONES UTILES.bas:2703`).
