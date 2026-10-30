# Form_FormTareasTramitadorPendientes

> **Commit fijado:** `cfa1937ab05a9e40150ea9236243898de96f84da` · **Fuente:** `src/forms/Form_FormTareasTramitadorPendientes.cls` (652 líneas) y `.form.txt` (26 bloques `Begin`) · **PRD:** [README](../README.md)
>
> **Abreviaturas:** `FU.bas` = `modules/FUNCIONES UTILES.bas`, `VG.bas` = `modules/Variables Globales.bas`, `constructor.bas` = `modules/constructor.bas`; las clases de `src/classes/` se citan por el nombre del fichero. Las citas son `Form_FormTareasTramitadorPendientes.cls:línea` salvo indicación.

Cola de trabajo del tramitador: lista las solicitudes que esperan una actuación concreta y abre la ventana de edición correspondiente. El tipo de tarea se elige en una lista y **selecciona qué consulta se ejecuta**: hay doce, una por estado (`:41-70`). No escribe nada; su efecto es abrir `FormSolicitudAltaTraspaso` o `FormSolicitudRenovacion` con la solicitud de sesión ya fijada.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `FU.bas:41-72` (`PintarIndicadores`) | Tras cualquier modification de datos, el módulo llama a `Form_FormTareasTramitadorPendientes.EstablecerDatos` si el formulario está abierto | `FU.bas:59-63` |
| `Form_Form0BDOpciones` | Lo abre desde el menú principal | `.form.txt:37` (rótulo de diseño) |

| Hacia | Cómo | Evidencia |
|---|---|---|
| `FormSolicitudAltaTraspaso` | Con `OpenArgs:="Alta"` o `"Traspaso"`, cerrando antes la instancia previa con `acSaveNo` | `:177-186` |
| `FormSolicitudRenovacion` | Con `OpenArgs:=.TIPO`, previa asignación de `m_ObjUsuarioHPSActivo` | `:187-197` |
| Suscriptores de `SolicitudModificada` | `Public Event SolicitudModificada()`, lanzado desde los eventos de los formularios hijos | `:5, 647, 651` |

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| Formulario | — | `OnOpen` → `Form_Open(Cancel)` | Dimensiona, fija el tramitador por defecto y carga la lista | `:531-580` |
| `ListaTipoTareas` | lista | `OnClick` → `ListaTipoTareas_Click` | Rellena la lista de resultados para el tipo elegido | `:260-306` |
| `Gestor` | combo | `OnBeforeUpdate` → `Gestor_BeforeUpdate` | Cambia el tramitador y recarga | `:221-248`, `.form.txt:1203-1205` |
| `ListaFiltrados` | lista de 5 columnas | `OnClick` → `ListaFiltrados_Click` | Resuelve la solicitud seleccionada y habilita «Editar» | `:582-645`, `.form.txt:915-917` |
| `ListaFiltrados` | lista | `OnDblClick` → `ListaFiltrados_DblClick` | Equivale a «Editar» | `:250-258` |
| `ComandoEditar` | botón | `OnClick` → `ComandoEditar_Click` | Abre el contenedor adecuado según el tipo de solicitud | `:160-215`, `.form.txt:1033` |
| `ComandoActualizar` | botón | `OnClick` → `ComandoActualizar_Click` | Recarga la cola | `:108-131`, `.form.txt:1102` |
| `ComandoAyuda` | botón | `OnClick` → `ComandoAyuda_Click` | Abre el PDF de ayuda | `:132-159`, `.form.txt:808` |
| `cmdSalir` | botón | `OnClick` → `cmdSalir_Click` | Cierra el formulario | `:10-13`, `.form.txt:1267` |
| `lblTitulo`, `lblListaDeResultados`, `lblEstado`, `lblUltimaModificacion` | etiquetas | — | Rótulos del formulario, del contador de resultados y de estado | `.form.txt:789, 1081, 1246, 1372` |

## Flujos

### Normal — carga

1. `Form_Open` dimensiona y copia el rótulo de `lblTitulo` al del formulario (`:541-545`).
2. Rellena el combo de tramitadores con `RellenarComboTramitadores` (`FU.bas`) y, si el usuario conectado está en la colección, lo preselecciona; si no, pone «Todos» (`:546-558`).
3. `EstablecerDatos` (`:473-529`) lanza `RellenarListas` y `Filtrar`, de modo que la ventana ya aparece con contenido.

### Normal — filtrado por tipo de tarea

1. `ListaTipoTareas_Click` llama a `Filtrar` (`FU.bas`), que lee el identificador del tipo seleccionado (`:35-38`).
2. `Filtrar` deshabilita «Editar» (`:33`) y elige la consulta con una cadena de doce `ElseIf` sobre `m_TipoTarea` (`:41-70`), todas de la forma `constructor.getSolicitudesPendientes<Estado>(p_Tramitador:=Me.Gestor, p_Error:=p_Error)`.
3. Añade una fila por solicitud con `id & ";" & TIPO & ";" & DNI & ";" & NombreCompletoCapitalizado & ";" & EstadoCalculadoTitulo` (`:87`), actualiza el contador (`:92`) y, si hay resultados, selecciona la primera fila y dispara `ListaFiltrados_Click` (`:93-97`).

### Normal — edición

1. `ComandoEditar_Click` exige selección (`:171-174`) y copia la solicitud a `m_ObjSolicitudActiva` (`:175`).
2. Según el tipo: «Alta» y «Traspaso» abren `FormSolicitudAltaTraspaso` con el `OpenArgs` correspondiente (`:177-186`); cualquier otro tipo exige que la solicitud tenga `UsuarioHPS` (`:188-191`), lo asigna a `m_ObjUsuarioHPSActivo` (`:195`) y abre `FormSolicitudRenovacion` (`:196`).

### Alternativos

- **Sin tipo seleccionado:** `Filtrar` sale antes de consultar si el identificador no es numérico (`:35-37`).
- **Tipo no contemplado:** la cadena de `ElseIf` cae en `Exit Function` sin mensaje (`:68-69`), con lo que la lista queda vacía sin explicación.
- **Renovación sin usuario HPS:** se aborta con «No se puede identificar el usuario HPS al ser una renovación» (`:188-191`).
- **Doble clic en la lista:** equivalente a «Editar», sólo si el botón está habilitado (`:253-255`).

### Errores

- Patrón uniforme de `On Error GoTo errores` con el centinela `1000`; los errores reales se envían a `CorreoAlAdministrador` (`FU.bas:585`) y se muestran en `MsgBox` crítico.
- `Form_Open` cierra el formulario en su rama de error (`:579`), dejando la coherencia con el resto de formularios del menú.

## Reglas, validaciones y cálculos

- **El tipo de tarea decide la consulta:** doce funciones distintas de `constructor`, una por `EnumEstado` (`:41-70`). No hay tabla ni vista única detrás.
- **El tramitador es un filtro obligatorio:** todas las consultas lo reciben como `p_Tramitador` (`:42-66`), y el valor por defecto es «Todos» si el usuario no está en la colección (`:554, 557`).
- **No se escribe nada:** este formulario sólo consulta y abre ventanas; la persistencia ocurre en `SolicitudOperaciones`, dentro de los formularios abiertos.
- **`Filtrar` es `Public`** (`:15`) y limpia la selección previa (`:28`), lo que obliga a volver a resolver la solicitud desde la base en cada filtro.
- **La fila lleva el DNI** como tercera columna (`:87`), dato personal visible en la lista.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Cola por estado de tarea | doce funciones `constructor.getSolicitudesPendientes*` | DAO, no acotadas en este fichero | `:42-66`, `constructor.bas` |
| Tramitadores del combo | `RellenarComboTramitadores` sobre `m_ObjEntorno.Coltramitadores` | DAO, cacheado | `:546, 550-551` |
| Resolución de la selección | `constructor.getSolicitud` | DAO | `ListaFiltrados_Click`, `:582-645` |
| Totales del menú | `Entorno.CalcularTotalesPendientes`, invocada desde `PintarIndicadores` antes de llamar a `EstablecerDatos` | DAO | `FU.bas:52, 62` |

## Coexistencia

Pendiente de `../coexistencia.md`. El patrón «tipo de tarea → consulta distinta» es el que más cuesta migrar: en web son doce consultas que pueden consolidarse en una con el estado como parámetro, y la lista necesita paginación porque hoy trae todas las solicitudes del tramitador de una vez. La dependencia inversa —`FU.bas:59-63` repinta esta ventana tras cada modificación— también debe convertirse en un evento de suscripción, no una llamada de ida y vuelta.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-HPS-S-58.** `ComandoEditar_Click` cierra `FormSolicitudAltaTraspaso` (`:179, 184`) o `FormSolicitudRenovacion` (`:193`) con `acSaveNo` antes de abrirlos de nuevo, sin preguntar: si el usuario tenía una solicitud a medias en esa ventana, el trabajo se descarta en silencio y el doble clic sobre la lista (`:250-258`) lo dispara igual de inmediato.
- **H-HPS-S-59.** La rama de error de `Gestor_BeforeUpdate` **no informa al usuario**: sólo envía el correo al administrador (`:240-245`) y reactiva la edición (`:246`), sin `MsgBox`. Cambiar de tramitador y que el filtrado falle deja la lista con el contenido anterior y ninguna explicación visible, a diferencia de los cinco handlers hermanos del mismo `.cls`.
- **H-HPS-S-60.** `Filtrar` sobrescribe el error que acaban de devolver las doce consultas: cada `constructor.getSolicitudesPendientes*` recibe `p_Error:=p_Error` (`:42-66`) y, acto seguido, la línea `p_Error = m_ObjEntorno.Error` (`:71`) lo reemplaza por el de la última operación que hizo `Entorno`. Si la consulta falla, el mensaje que sube a la interfaz no es el de la consulta.
- **H-HPS-S-61.** `FU.bas:59-63` invoca `Form_FormTareasTramitadorPendientes.EstablecerDatos` **sobre el nombre de clase**, no sobre la instancia, para forzar el repintado desde un módulo. Funciona por la instancia abierta, pero cualquier segundo formulario del mismo tipo recibiría la llamada y ambos se repintarían, y un acoplamiento de módulo a formulario que no aparece al leer el formulario.
- `cmdSalir_Click` (`:11`) y `ListaFiltrados_DblClick` (`:251`) arrancan con `On Error Resume Next` sin aviso ni correo.
- La cadena de doce `ElseIf` (`:41-70`) no tiene registro de casos sin cubrir: un `EnumEstado` nuevo deja la lista vacía sin diagnóstico (`:68-69`).
- `Me.Caption` se escribe dos veces seguidas en `Form_Open`, con `m_TituloFormulario` (`:542`) y con el rótulo de la etiqueta (`:545`); la primera no sirve para nada.
- `m_SolicitudSeleccionada` y `m_TipoTareaSeleccionada` son `Public` (`:6-7`) pero no hay ningún lector ni escritor fuera de este `.cls`: las dos únicas referencias externas son líneas comentadas (`:463-464`).
- Las filas se arman concatenando cinco campos con `;`, con el DNI en medio (`:87`): misma fragilidad de separador que el resto de listas del legacy (ver `H-HPS-S-39` y `H-HPS-S-47`).
- **Defectos de interfaz heredados (a verificar):** el combo de tramitadores lleva `RowSourceType ="Value List"` con nombres reales (`Todos;Martina Torralba Rodríguez;Esperanza del Álamo Arriba`, `.form.txt:1205`) y la lista de resultados declara su `RowSource` de diseño como `IDSolicitud;DNI;Nombre;Estado` (`.form.txt:917`), con cuatro columnas, mientras el código añade cinco (`:30, 87`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: una sola consulta de cola con el estado como parámetro en lugar de doce; paginación y orden explícito en el listado de tareas; suscripción a los cambios de solicitud en lugar del repintado imperativo desde `PintarIndicadores`; y confirmación antes de descartar una solicitud en curso.

## Criterios de aceptación

Propuestos, no aprobados:

1. La lista muestra las solicitudes del tramitador seleccionado para el tipo de tarea elegido.
2. Cambiar de tipo de tarea recarga la lista y el contador refleja el número real de resultados.
3. Cambiar de tramitador recarga la lista, y si el filtrado falla se informa del motivo.
4. Sin fila seleccionada, «Editar» no hace nada.
5. Editar una solicitud de alta o traspaso abre el contenedor con ese tipo; editar una renovación exige usuario HPS y, si falta, se informa.
6. Abrir una solicitud descarta cualquier otra ventana de solicitud abierta sólo después de confirmarlo con el usuario.
7. Un tipo de tarea sin consulta asociada no deja la lista vacía sin explicación.
8. La fila muestra la DNI del solicitante sin que un `;` en los datos desplace las columnas.

## Pendientes

- Qué devuelve cada una de las doce funciones `constructor.getSolicitudesPendientes*`: la elección de consulta es legible, pero su SQL no está en este `.cls`.
- Valores concretos de `EnumEstado` que corresponden a cada fila de `ListaTipoTareas`: el `.form.txt` no declara el `RowSource` de esa lista y el código sólo lee `Column(0)`.
- Si `constructor.getSolicitudesPendientes*` admite un error por parámetro y lo rellena: de ello depende la gravedad de H-HPS-S-60.
- Recuento de controles: el inventario fija 29 y el `.form.txt` contiene 26 bloques `Begin <Tipo>`; la diferencia no se reproduce con una regla simple de recuento.