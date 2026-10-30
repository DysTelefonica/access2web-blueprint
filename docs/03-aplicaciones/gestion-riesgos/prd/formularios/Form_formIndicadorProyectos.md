# Form_formIndicadorProyectos

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Fuente:** `src/forms/Form_formIndicadorProyectos.cls` (277 líneas) y `src/forms/formIndicadorProyectos.form.txt` · **PRD:** [README](../README.md) · **Formularios:** [forms.md](../../forms.md)

Diálogo modal de selección de proyectos para los indicadores. Carga la lista de proyectos «abiertos» en el periodo filtro, deja quitar los que el usuario no quiera y devuelve la selección al panel. No escribe nada en la base de datos: todo el estado vive en memoria. Las citas son `Form_formIndicadorProyectos.cls:línea` salvo indicación.

**Aviso sobre el diseño.** El directorio contiene dos ficheros de diseño que solo difieren en mayúsculas: `formIndicadorProyectos.form.txt` (1041 líneas) y `Form_formIndicadorProyectos.form.txt` (1043). Difieren en el checksum, en una propiedad de guardado y en un bloque de código embebido al final. En un sistema de ficheros sin distinción de mayúsculas ambos nombres colisionan; en Linux conviven y **pueden divergir del diseño realmente desplegado**. Las citas de diseño de esta ficha usan el fichero sin prefijo.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormIndicador` | `DoCmd.OpenForm ... acDialog` con `OpenArgs` `ANIO=<año>;SEM=<1\|2\|>` | `Form_FormIndicador.cls:73` |

Es modal y se cierra a sí misma: no hay ningún `DoCmd.OpenForm` que la abra desde otro sitio. Acepta el contrato con nombre que declara el comentario de `:88-97`, que delega en `ParseIndicadorOpenArgs` (`modIndicadorOpenArgsParser.bas:49`); ese comentario afirma además que el formato heredado con tubería se acepta por compatibilidad. Si la entrada no es válida, se lanza error 1000 con mensaje en castellano y **no hay respaldo silencioso al año en curso** (`:101`, `:88-97`).

Salidas: al aceptar, deja la selección y la marca confirmada en el estado de sesión (`:173-174`) y cierra (`:176`); al cancelar, solo cierra (`:187`). Ninguna de las dos communicates nada al panel salvo a través de ese estado.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `Form_Open` | Prepara el estado, interpreta los argumentos y carga la lista | `:17-30`, `.form.txt:41` (`Caption`) |
| (formulario) | — | `Form_Close` | Marca el formulario como cerrado | `:12-15`, `.form.txt` (`OnClose`) |
| `lstProyectos` | lista | `lstProyectos_Click` | Habilita Quitar si el identificador de la fila es mayor que cero | `:107-110`, `.form.txt:554` |
| `cmdQuitar` | botón | `cmdQuitar_Click` | Excluye el proyecto seleccionado de la lista | `:128-152`, `.form.txt:584` |
| `cmdRecargarTodos` | botón | `cmdRecargarTodos_Click` | Borra las exclusiones y vuelve a cargar | `:157-166`, `.form.txt:627` |
| `cmdAceptar` | botón | `cmdAceptar_Click` | Confirma la lista y cierra | `:169-181`, `.form.txt:681` |
| `cmdCancelar` | botón | `cmdCancelar_Click` | Cierra sin confirmar | `:185-188`, `.form.txt:720` |

## Flujos

### Apertura (`Form_Open`, `:17-30`)

1. `IndicadorState_Init` y marca el formulario como abierto (`:20-21`).
2. `ParseOpenArgs` (`:24`, `:88-102`). Si el año falta, se lanza error 1000 con mensaje explícito (`:101`).
3. `CargarLista_RespetandoExcluidos` (`:25`, `:32-84`).
4. Deja `cmdQuitar` deshabilitado (`:26`).

El manejador de error muestra un `MsgBox` con número y descripción (`:30`) y **no asigna `Cancel`**: el diálogo queda abierto, con la lista vacía o a medias.

### Normal — carga de la lista (`CargarLista_RespetandoExcluidos`, `:32-84`)

1. Escribe la fila de cabecera como origen de la lista: `IDProyecto;Código;Proyecto;Cierre;Alta` (`:48`) y hace `Requery` (`:49`).
2. Si el año no es numérico, lanza error 1000 (`:52`).
3. Pide los proyectos a `getProyectosAbiertos` (`:56`, `:190-277`). Si la colección es `Nothing`, **sale sin avisar y sin llenar la lista** (`:58-60`).
4. Recorre la colección, divide cada elemento por tubería en cinco campos (`:66`), sustituye los puntos y comas de los textos por dos puntos para no romper la separación por punto y coma del `ListBox` (`:70-71`), formatea las fechas a día/mes/año (`:73-74`) y añade la fila solo si el identificador es numérico y **no está excluido** (`:71-77`).

### Normal — quitar un proyecto (`cmdQuitar_Click`, `:128-152`)

Si el índice seleccionado es menor o igual que cero, avisa de que hay que elegir un proyecto y no la cabecera, y deshabilita el botón (`:135-139`). Si no, marca el identificador como excluido, lo borra de la lista y deshabilita el botón (`:141-147`).

### Alternativo — recargar todos (`cmdRecargarTodos_Click`, `:157-166`)

Borra las exclusiones, vuelve a cargar la lista y deja Quitar deshabilitado (`:159-162`). El proyecto vuelve a aparecer, pero **el orden y la composición son los de una consulta nueva**, no los que tenía el usuario delante.

### Normal — aceptar (`cmdAceptar_Click`, `:169-181`)

`IndicadorState_Init`, `IndicadorState_SetSeleccionadosDesdeLista`, `IndicadorState_MarcarConfirmado True` y cierre (`:172-176`).

### Alternativo — cancelar (`cmdCancelar_Click`, `:185-188`)

`On Error Resume Next` y cierre (`:187`). No toca la selección ni el flag de confirmación.

### Errores

Los cuatro manejadores de evento muestran `MsgBox` con prefijo (`:30`, `:84`, `:152`, `:165`, `:180`) y no notifican por correo. `getProyectosAbiertos` es distinto: distingue el error interno 1000 de los demás, compone el detalle para esos últimos y hace `Resume salir` (`:273-277`), de modo que **un fallo no previsto devuelve una colección vacía sin mensaje**.

## Reglas, validaciones y cálculos

- **«Abierto en el periodo» es un solape de intervalos, no una pertenencia**: alta anterior al final del periodo y, o bien sin fecha de cierre, o bien cierre posterior o igual al inicio del periodo (`:238-239`). La consulta va contra la tabla de proyectos con parámetros de fecha (`:234-240`).
- **El rango del filtro es exclusivo en el final y el de los indicadores es inclusivo.** Aquí `dFinExcl` es el primer día del periodo siguiente (`:216-230`); `CalcularRangoSemestre`, que usa el panel, devuelve el último día del semestre (`modIndicadorDashboard.bas:41-55`). Mismo periodo, dos convenciones.
- **Quitar no es seleccionar.** La lista es «proyectos abiertos menos exclusiones» (`:72`) y la selección final se reconstruye entera al aceptar (`:173`). Una exclusión es reversible hasta que se acepta.
- **La cabecera es una fila de datos.** No hay propiedad de encabezados: se inyecta una fila falsa por `RowSource` (`:48`), y el código la trata como tal en dos sitios distintos, con resultados distintos: `SelectedIDProyecto` no la filtra (`:113-125`) y `cmdQuitar` sí (`:135`). La columna del identificador tiene además ancho cero (`.form.txt:556`).
- **La consulta es efímera a propósito.** `CreateQueryDef` con nombre vacío (`:243`) crea una consulta no guardada, cuyo objeto se libera al terminar (`:270`). No deja nada en la base de datos.

## Lecturas y escrituras

| Operación | Origen | Evidencia |
|---|---|---|
| Proyectos del periodo | Consulta efímera sobre la tabla de proyectos, por solape de intervalos | `:234-247` |
| Exclusiones y selección | Estado de sesión en memoria | `mIndicadorProyectosState.bas:54-131` |
| Año y semestre de entrada | `OpenArgs` con nombre | `:88-102`, `modIndicadorOpenArgsParser.bas:49` |

Sin escrituras en base de datos ni en disco: solo lectura y estado en memoria.

## Coexistencia

Pendiente de decisión de coexistencia. Es la pantalla que más se apoya en una peculiaridad del motor de Access: el `ListBox` con origen de valores, la cabecera inyectada como fila y la anchura cero de la columna clave. En la web esa lista se sustituye por una tabla con columna de identificador visible o identificable, y la selección se resuelve en el servidor. El punto a decidir es el estado de sesión: hoy la lista vive en memoria del proceso y se pierde si el proceso se reinicia entre el diálogo y la confirmación.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** diálogo modal que carga los proyectos abiertos en el periodo filtro, permite excluir los innecesarios y devuelve la selección confirmada al panel de indicadores.

**Posibles defectos (sin corregir, a verificar):**

- **H-GR-286.** Los dos ficheros de diseño que solo difieren en mayúsculas son un riesgo real de despliegue: en un sistema de ficheros que no distingue mayúsculas no pueden coexistir, y en Linux el que se lea puede no ser el compilado. Hay que determinar cuál refleja el binario.
- **H-GR-287.** Si `getProyectosAbiertos` devuelve una colección vacía o `Nothing`, `Form_Open` y `CargarLista_RespetandoExcluidos` **no avisan** (`:58-60`, `:273-277`): el usuario ve un diálogo con cabecera y sin filas, indistinguible de «no hay proyectos en ese periodo».
- **H-GR-288.** El bucle parte cada elemento en cinco campos sin comprobar la longitud (`:66`, `:67-71`). Si el nombre del proyecto contiene una tubería, el acceso a un campo inexistente falla con error en tiempo de ejecución y el usuario ve el `MsgBox` del manejador.
- **H-GR-289.** La cabecera es una fila de datos real (`:48`): se puede seleccionar, `SelectedIDProyecto` intenta leer su identificador (`:113-125`) y el botón Quitar se habilita si el valor numérico es mayor que cero (`:109`). Cuando está habilitada y se pulsa, la respuesta es el aviso de la `:137`.
- **H-GR-290.** **Cancelar no deshace nada**: no limpia la selección ni desactiva el flag de confirmación (`:185-188`). Si el usuario ya había confirmado antes en el mismo proceso, al volver al panel la selección anterior sigue viva y Calcular queda habilitado con proyectos que no ve en pantalla.
- **H-GR-291.** El ancho cero de la columna del identificador (`.form.txt:556`) oculta la clave de cada fila: el usuario ve código, proyecto y fechas, pero no qué proyecto es en términos del modelo de datos.
- **H-GR-292.** `Err.Raise 1000` en `:214` y `:229` se lanza sin descripción; el texto lo aporta `p_Error`. Si algún camino futuro depende de `Err.Description` en lugar de `p_Error`, el mensaje se pierde.
- **H-GR-293.** La sustitución de punto y coma por dos puntos (`:70-71`) **altera el texto mostrado** de proyectos y nombres que contengan ese carácter. Es una corrección de formato aplicada a un dato de negocio.
- **H-GR-294.** En `getProyectosAbiertos` el recordset se cierra (`:267-268`) pero el `QueryDef` solo se libera con `Set qd = Nothing` (`:270`), sin cierre explícito.
- **H-GR-295.** Hay variables de módulo (`flag` y `dato`, `:6-7`) **sombreadas** por variables locales de igual nombre (`:39-40`), que quedan sin uso. El mismo nombre para el estado global y el local invita a errores en cambios futuros.
- **H-GR-296.** `cmdRecargarTodos` vuelve a ejecutar la consulta y puede devolver un conjunto distinto del que tenía el usuario delante (`:159-162`), sin avisar de si algo ha cambiado.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: sustituir la lista con cabecera inyectada por una tabla con columna de identificador real y selección visible; devolver la selección por parámetro en la ruta en lugar de por estado en memoria; hacer que cancelar descarte explícitamente la confirmación anterior; y mover el filtro de proyectos a una consulta de servidor reutilizable por otroslistados.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir el diálogo carga los proyectos abiertos en el periodo indicado.
2. Sin proyectos en el periodo se avisa, en vez de mostrar una lista vacía sin explicación.
3. Quitar un proyecto lo desaparece de la lista; Recargar lo devuelve.
4. Aceptar devuelve exactamente los proyectos visibles y habilita el cálculo en el panel.
5. Cancelar no cambia lo confirmado anteriormente.
6. Un identificador de proyecto no numérico o mal formado no rompe la carga.

## Pendientes

- Qué versión del diseño está realmente en el binario, dados los dos ficheros que solo difieren en mayúsculas.
- Si una colección de proyectos vacía significa «sin resultados» o «consulta fallida» en producción, ya que el código no lo distingue.
- Qué consultas del resto de la aplicación reutilizan el criterio de proyecto abierto en el periodo.