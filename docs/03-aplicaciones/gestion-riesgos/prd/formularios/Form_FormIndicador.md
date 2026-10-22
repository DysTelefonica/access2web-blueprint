# Form_FormIndicador

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Fuente:** `src/forms/Form_FormIndicador.cls` (315 líneas) y `src/forms/FormIndicador.form.txt` · **PRD:** [README](../README.md) · **Formularios:** [forms.md](../../forms.md)

Panel de indicadores con seis tarjetas (tiles), una por métrica, cada una con su botón de descarga a Excel. Es una pantalla de **lectura y exportación**: no escribe nada en la base de datos. Para obtener cifras exige elegir antes los proyectos en un diálogo modal. La salida en HTML es la ayuda, que se abre en el navegador del sistema. Las citas son `Form_FormIndicador.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_Form0BDOpciones` | `DoCmd.OpenForm "FormIndicador"`, sin argumentos | `Form_Form0BDOpciones.cls:113` |

No acepta `OpenArgs`. Su contexto es íntegramente estado de sesión, que reinicia al abrirse.

Salidas: abre `FormIndicadorProyectos` como modal con `OpenArgs` con nombre (`:73`); exporta un Excel por tarjeta (`ExportarSQLaExcel`, `:310`, que instancia Excel en `modIndicadorDashboard.bas:127`); y abre la ayuda en HTML contra una ruta construida en el entorno (`Entorno.cls:876`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `Form_Open` | Fija valores por defecto, limpia estado y deja los tiles a cero | `:11-40`, `.form.txt:40` (`Caption`) |
| `txtAnio` | cuadro de texto | — | Año; por defecto el año en curso con cuatro dígitos | `:20`, `.form.txt:693` |
| `cboSemestre` | desplegable | — | `S1;S2;Anual`, por defecto `S1` | `:18-19`, `.form.txt:664-666` |
| `cmdCargarProyectos` | botón | `cmdCargarProyectos_Click` | Abre el diálogo modal de proyectos | `:43`, `.form.txt:1847` |
| `lblProyectosCargados` | etiqueta | — | Recuento de proyectos confirmados | `:192-204`, `.form.txt:1889` |
| `cmdCalcular` | botón | `cmdCalcular_Click` | Calcula el resumen del periodo | `:86`, `.form.txt:611` |
| `lblTile1..6Valor` | etiquetas | — | Las seis cifras mostradas | `:866` y ss., `.form.txt` |
| `cmdTile1..6Excel` | botones | `cmdTile1..6Excel_Click` | Una línea cada uno: delega en `ExportarTile` | `:171-176`, `.form.txt:888` |
| `cmdAyuda` | botón | `cmdAyuda_Click` | Abre el HTML de ayuda | `:180-186`, `.form.txt:538` |
| `btnSalir` | botón | `btnSalir_Click` / `CmdSalir_Click` | Cierra sin guardar | `:7-9` y `:188-190`, `.form.txt:2194` |

## Flujos

### Apertura (`Form_Open`, `:11-40`)

1. `AjustarTamaño Me` y copia `lblTitulo` al título de la ventana (`:13-15`).
2. **Reinicia el estado**: `IndicadorState_ResetAll` (`:17`), `LimpiarTempVarsIndicador` (`:24`) y `m_IdsCsvUltimoCalculo` a vacío (`:16`). Tres mecanismos distintos de limpieza.
3. Deja semestre `S1` y el año en curso (`:18-20`); pone los seis tiles a cero y deshabilita Calcular (`:29-35`).
4. `ActualizarBotonesExport` y `RefrescarEstadoProyectos` (`:36-37`).

El manejador de error avisa con `MsgBox` y **no asigna `Cancel`** (`:39-40`): el formulario queda abierto con valores a medio inicializar.

### Normal — cálculo (`cmdCalcular_Click`, `:86-164`)

1. Exige confirmación previa del diálogo (`:100`) y una lista de identificadores no vacía (`:107-110`); guarda esa lista en `m_IdsCsvUltimoCalculo` (`:112`).
2. Valida el año y resuelve el rango con `CalcularRangoSemestre` (`:115-127`).
3. Llama a `IndicadorRiesgosV2_GetResumen` (`:131`, `IndicadorRiesgosRepositorio.bas:185`) y pinta los seis campos en las etiquetas (`:135-140`). Si el recordset llega vacío, pone los seis a cero (`:143-148`).
4. `ActualizarBotonesExport` habilita la descarga solo de los tiles con valor mayor que cero (`:155`, `:251-268`).

### Alternativo — carga de proyectos (`cmdCargarProyectos_Click`, `:43-80`)

1. Valida el año y normaliza el semestre a `1`, `2` o vacío (`:51-65`).
2. `IndicadorState_Init` y marca el formulario como abierto (`:67-68`).
3. Abre `FormIndicadorProyectos` en modal con `ANIO=...;SEM=...` (`:73`).
4. Al volver, `RefrescarEstadoProyectos` (`:76`): habilita Calcular solo si hay confirmación y al menos un proyecto.

### Alternativo — exportación (`ExportarTile`, `:270-315`)

1. Usa los identificadores del **último cálculo** (`:283`); si están vacíos, cae a la selección actual en estado (`:285-290`).
2. Revalida el año y **vuelve a resolver el rango con los valores actuales del filtro** (`:292-301`).
3. Pide el SQL de detalle (`IndicadorRiesgosV2_SqlDetalle`, `:304`) y lo exporta (`:310`). Si el SQL viene vacío, avisa (`:305-307`).

### Errores

Cada manejador termina en `MsgBox` con su propio prefijo (`:40`, `:80`, `:163`, `:315`). `cmdCalcular` cierra el recordset antes de mostrar el mensaje (`:157-159`). Ninguno notifica por correo ni deja registro.

## Reglas, validaciones y cálculos

- **El botón Calcular no se habilita solo, sino por un flag de sesión.** `RefrescarEstadoProyectos` exige `IndicadorState_HaConfirmado` y un recuento mayor que cero (`:198`).
- **El semestre viaja en dos formatos.** Al diálogo sale como `1`, `2` o vacío (`:57-65`); el rango lo resuelve `CalcularRangoSemestre`, que además acepta `S1`, `S2`, `ANUAL` y sus formas cortas (`modIndicadorDashboard.bas:41-55`).
- **La habilitación de la descarga se decide leyendo el texto del tile.** `ToLongSafe` (`:232-249`) filtra los dígitos de la etiqueta y compara con cero (`:255-268`); un número con signo negativo o decimal se lee como su parte entera.
- **`m_IdsCsvUltimoCalculo` es una caché de sesión, no un estado derivado del filtro** (`:5`, `:283`).

## Lecturas y escrituras

| Operación | Origen | Evidencia |
|---|---|---|
| Resumen de los seis indicadores | Consulta de resumen sobre el periodo y los proyectos | `:131`, `IndicadorRiesgosRepositorio.bas:185` |
| Detalle para la descarga | SQL de detalle por tipo de tile | `:304`, `IndicadorRiesgosRepositorio.bas:317` |
| Selección de proyectos y confirmación | Estado de sesión | `mIndicadorProyectosState.bas:85-131` |
| Rango de fechas del semestre | Cálculo en módulo | `modIndicadorDashboard.bas:25` |
| Ayuda en HTML | Fichero en disco | `:181-185`, `Entorno.cls:876` |

Sin escrituras en base de datos. Escribe un fichero Excel por descarga y deja estado en memoria.

## Coexistencia

Pendiente de decisión de coexistencia. Aquí la dependencia de ficheros es doble: la ayuda se abre desde una ruta local construida por el entorno (`Entorno.cls:876`) y la exportación produce un fichero en la carpeta del usuario. En la web, la ayuda debe servirse desde la propia aplicación y la descarga debe pasar por una ruta autenticada en lugar del disco del cliente. El estado de sesión (selección de proyectos y confirmación) es el segundo punto que hay que decidir: hoy vive en memoria del proceso de Access.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** panel de seis indicadores calculados sobre los proyectos confirmados, con filtro de año y semestre, descarga individual a Excel y ayuda en HTML.

**Posibles defectos (sin corregir, a verificar):**

- **H-GR-276.** Hay **dos procedimientos de salida idénticos** (`btnSalir_Click` en `:7-9` y `CmdSalir_Click` en `:188-190`) con la misma acción y distinto casing. En Access ambos son el mismo nombre, así que solo uno puede quedar enlazado y el otro es código muerto; conviene confirmar en el binario cuál está vivo.
- **H-GR-277.** La exportación reutiliza los proyectos del último cálculo (`:283`) pero **recalcula el rango con el año y el semestre que haya en pantalla en ese momento** (`:292-301`). Si el usuario cambia el filtro después de calcular, el fichero sale con datos de dos periodos distintos sin ningún aviso.
- **H-GR-278.** `ActualizarBotonesExport` decide la habilitación solo mirando la cifra de cada etiqueta (`:251-268`), no la selección ni el filtro. Tras cambiar el filtro, los botones siguen activos y el usuario puede descargar un detalle que no corresponde a lo que ve.
- **H-GR-279.** `Form_Open` no asigna `Cancel` en su manejador de error (`:39-40`): si el reinicio de estado falla a mitad, el formulario queda abierto con tiles a cero, Calcular deshabilitado y el diálogo como único camino de salida.
- **H-GR-280.** La limpieza de estado se produce por tres vías distintas: `IndicadorState_ResetAll` (`:17`), `LimpiarTempVarsIndicador` (`:24`) y el borrado de `m_IdsCsvUltimoCalculo` (`:16`). Conviene verificar que no se pisan entre sí y que ninguno deja restos.
- **H-GR-281.** `Err.Raise 1000` de `:132` se lanza sin descripción, y el mensaje final depende de `errTxt` (`:163`). Si el resumen falla por otra causa, el texto mostrado sería el de Access y no el del módulo.
- **H-GR-282.** `ToLongSafe` (`:232-249`) descarta cualquier carácter no numérico: un tile con formato monetario o con separador de miles se leería como un número distinto del mostrado.
- **H-GR-283.** `cmdAyuda` muestra el literal «Pendiente de Integrar» (`:182`) si el HTML no está en disco: es un mensaje de estado interno de desarrollo expuesto al usuario final, en un título de diálogo que dice «Sin Ayuda».
- **H-GR-284.** Cada exportación crea una instancia de Excel (`modIndicadorDashboard.bas:127`); conviene verificar que la función cierra libro y aplicación al salir, o se acumulan procesos huérfanos por cada descarga.
- **H-GR-285.** El contenido descargable depende de que el estado de sesión siga vivo. Si el proceso se reinicia entre el cálculo y la descarga, `m_IdsCsvUltimoCalculo` se pierde y la exportación cae a la selección actual, que puede ser otra.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: convertir las seis tarjetas en una tabla de indicadores con descarga bajo demanda en lugar de seis botones; recalcular siempre la descarga desde los filtros visibles y no desde una caché del último cálculo; servir la ayuda desde una ruta de la propia web en lugar de un fichero en disco; mantener la selección de proyectos en el servidor, con identificador de sesión, en vez de en memoria del proceso.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir el panel no hay proyectos seleccionados y Calcular está deshabilitado.
2. Elegir proyectos y aceptar habilita Calcular y muestra el recuento.
3. Calcular muestra las seis cifras del periodo filtrado y activa la descarga solo de las que tengan valor.
4. Cancelar el diálogo deja el panel como estaba.
5. Cada descarga produce un fichero con el detalle del indicador y el periodo que se está viendo.
6. Si no hay ayuda disponible, se avisa de forma comprensible para el usuario.

## Pendientes

- Qué consultas exactas ejecuta el resumen y el detalle, y cómo se relacionan con el resto de indicadores de la aplicación.
- Si la descarga se cierra correctamente la instancia de Excel, o si depende de que el usuario la cierre.
- Dónde vive la ayuda en HTML y si esa ruta es alcanzable desde el despliegue real.
- Si existe alguna otra entrada a esta pantalla además del menú de opciones.