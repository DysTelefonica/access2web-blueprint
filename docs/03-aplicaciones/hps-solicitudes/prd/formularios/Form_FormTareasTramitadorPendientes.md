# Form_FormTareasTramitadorPendientes

> **Commit fijado:** `cfa1937a` · **Fuente:** `src/forms/Form_FormTareasTramitadorPendientes.cls` (651 líneas) y `.form.txt` (2046 líneas) · **PRD:** [README](../README.md)

Worklist del tramitador. Muestra once estados (`EnumEstado`) con su recuento y, al elegir uno, lista las solicitudes de ese estado que el tramitador seleccionado tiene pendientes. Desde la lista se abre el formulario de detalle de la solicitud (`Form_FormSolicitudAltaTraspaso` o `Form_FormSolicitudRenovacion`) según `TIPO`. Las citas son `Form_FormTareasTramitadorPendientes.cls:línea` salvo que se indique el `.form.txt`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| Menú principal | `DoCmd.OpenForm "FormTareasTramitadorPendientes"` | global `m_TituloFormulario` (`Variables Globales.bas:196, 388-401`) |
| `FormSolicitudAltaTraspaso`, `FormSolicitudRenovacion` | `RaiseEvent SolicitudModificada` ⇒ `m_Form_Alta/Editado` ⇒ `RaiseEvent SolicitudModificada` | `:11, :627-633` |
| Consumidores del evento | tras refrescar `Form_Open` o un click explícito de «Contador» | `:117-118, :621-625` |

Estado compartido: `m_SolicitudSeleccionada` y `m_TipoTareaSeleccionada` (`:9-10`), `m_ObjUsuarioConectado` y `m_ObjUsuarioHPSActivo` (`Variables Globales.bas:189, 204-205`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `lblTitulo` | etiqueta | — | «TAREAS PENDIENTES DEL TRAMITADOR» | `.form.txt:789-790` |
| `Gestor` | combo (value list en diseño) | `BeforeUpdate` ⇒ `EstablecerDatos` | Cambia el tramitador; repinta contadores y, si hay pendientes, auto-selecciona estado | `:255-280`, `.form.txt:1203-1206` |
| `ListaTipoTareas` | lista 3 columnas (ID oculto, contador, título) | `OnClick` | Selecciona el estado y lanza `Filtrar` y, si procede, `ListaFiltrados_Click` | `:283-329`, `.form.txt:955-969` |
| `Explicacion` | texto | — | Leyenda del estado seleccionado (`m_ObjEntorno.colExplicacionTipoSeleccionTarea`) | `:316-318`, `.form.txt:998-1002` |
| `ListaFiltrados` | lista 5 columnas (ID oculto, Tipo, DNI, Nombre, Estado) | `OnClick` y `OnDblClick` | Lista de solicitudes del estado y tramitador actuales | `:374-389, :396-405`, `.form.txt:906-921` |
| `lblListaDeResultados` | etiqueta | — | Recuento «Registros filtrados (n)» | `:99-101, :120`, `.form.txt:1073-1082` |
| `lblUltimaModificacion` | etiqueta (oculta para no administradores) | — | Texto del último cambio global | `:486-495`, `.form.txt:1364-1373` |
| `ComandoEditar` | botón | `OnClick` | Abre `FormSolicitudAltaTraspaso` o `FormSolicitudRenovacion` | `:158-219`, `.form.txt:1023-1037` |
| `ComandoActualizar` | botón | `OnClick` | Repinta indicadores y contadores | `:109-130`, `.form.txt:1102-1106` |
| `ComandoAyuda` | botón | `OnClick` | Llama a `AbrirAyuda` | `:69-93`, `.form.txt:808-812` |
| `cmdSalir` | botón | `OnClick` | Cierra (`acSaveNo`) | `:21-24`, `.form.txt:1267-1269` |
| `lblEstado` | etiqueta | — | Rótulo heredado «ColTiposCorreos», no se usa | `.form.txt:1246-1247` |
| `Etiqueta57` | etiqueta | — | «Tramitador» | `.form.txt:1174-1175` |

## Flujos

### Apertura (`Form_Open`, `:541-589`)

1. `Ajustar Me`, fija `Caption = m_TituloFormulario`, deshabilita edición (`:548-553`).
2. `RellenarComboTramitadores Me.Gestor, EnumSiNo.Sí` (`FU.bas:370-410`) añade «Todos» y el resto de tramitadores desde `m_ObjEntorno.Coltramitadores` (`:556-560`).
3. Si el usuario conectado está en `Coltramitadores`, `Me.Gestor = m_ObjUsuarioConectado.Nombre`; en otro caso, «Todos» (`:561-572`).
4. `EstablecerDatos` repinta la lista de estados (`:574-586`).

### Establecer datos (`EstablecerDatos`, `:483-538`)

1. `RellenarListas` carga once estados con su contador (`:487-491`).
2. Si el usuario es administrador y hay `UltimoCambio`, muestra `lblUltimaModificacion`; en otro caso, lo oculta (`:492-504`).
3. Lee los contadores del primer y segundo estado (`m_NPendientesEnvioExcel`, `m_NPendientesRegistroMarga`) de la columna 1 (`:507-515`).
4. Si hay pendientes de envío Excel, auto-selecciona la primera fila; si no, la segunda (registro Marga) (`:521-530`). Dispara `ListaTipoTareas_Click`.

### Recuento y listado (`RellenarListas`, `:283-329`)

1. Detecta si el filtro es «Todos» o un tramitador concreto (`:301-306`).
2. Si es «Todos» y la colección existe en `m_ObjEntorno`, usa `m_objColPendientes*.Count`; si no, llama al `constructor` con `p_Tramitador` y suma lo devuelto (`:312-553`).
3. Las colecciones que no tienen caché en `m_ObjEntorno` (`PendienteExcelRelleno`, `PendienteEnvioTraspasoONS`) consultan siempre al constructor (`:368-380, :424-431`).
4. Pinta once filas combinando `EnumEstado`, contador y `m_ObjEntorno.ColEstadosParaTitulo(CStr(...))` (`:485-557`).

### Filtrar por estado y tramitador (`Filtrar`, `:27-156`)

1. Reinicia la lista con cabecera `ID;Tipo;DNI;Nombre;Estado` (`:38-41`).
2. Lee el `ID` de `ListaTipoTareas.Column(0)` y aborta si no es numérico (`:43-46`).
3. Según el `EnumEstado`, delega en el método del constructor (`:48-105`):

   | Estado | Método | Evidencia |
   |---|---|---|
   | `PendienteEnvioExcel` | `getSolicitudesPendientesEnvioExcel` | `constructor.bas:668-778`, `:48-50` |
   | `PendienteRecordatorioExcel1/2` | `…RecordatorioExcel1/2` | `constructor.bas:781-928`, `:51-56` |
   | `PendienteExcelRelleno` | `…RellenoExcel` | `constructor.bas:931-1003`, `:57-59` |
   | `PendienteRegistroMarga` | `…AltaMarga` | `constructor.bas:1006-1077`, `:60-62` |
   | `PendienteRecordatorioEnvioDPS1/2` | `…RecordatorioEnvioDPS1/2` | (consultar `constructor.bas`), `:63-68` |
   | `PendienteEnvioDPS` | `…EnvioDPS` | `constructor.bas:1079-1149`, `:69-71` |
   | `PendienteRegistroONS` | `…RegistroONS` | `constructor.bas:1303-1371`, `:72-74` |
   | `PendienteEnvioTraspasoONS` | `…EnvioTraspasoONS` | `constructor.bas:1443-1513`, `:75-77` |
   | `PendienteRegistroEnHPS` | `…RegistroEnHPS` | `constructor.bas:1373-1441`, `:78-80` |

4. Si `m_ObjEntorno.Error <> ""`, aborta (`:107-110`).
5. Por cada `IDSolicitud` añade `<ID>;<Tipo>;<DNI>;<NombreCompletoCapitalizado>;<EstadoCalculadoTitulo>` (`:114-126`).
6. Actualiza el contador y, si hay un único resultado, preselecciona (`:127-131`).

### Selección (`ListaFiltrados_Click`, `:556-616`)

1. Si la columna 0 está vacía, `Set m_SolicitudSeleccionada = Nothing` (`:578-579`).
2. Si cambió el `IDSolicitud`, recarga con `constructor.getSolicitud(p_IDSolicitud)` (`constructor.bas:483-521`) (`:586-607`).
3. Habilita `ComandoEditar` (`:613-614`).

### Editar selección (`ComandoEditar_Click`, `:158-219`)

1. Aborta si no hay selección (`:175-178`).
2. Si `TIPO = "Alta"` o `"Traspaso"`, cierra un `FormSolicitudAltaTraspaso` previo y lo reabre con `OpenArgs` (`:181-195`).
3. Para cualquier otro `TIPO` (renovaciones), exige `UsuarioHPS` no nulo y abre `FormSolicitudRenovacion` con `TIPO` como `OpenArgs`, fijando antes `m_ObjUsuarioHPSActivo` (`:196-216`).

### Errores

Cada handler acumula `m_Error`, envía `CorreoAlAdministrador` cuando procede y muestra `MsgBox vbCritical`. `Form_Open` cierra tras el aviso (`:581-589`); el resto no cierra (defecto observado en `Gestor_BeforeUpdate`, que sólo notifica).

## Reglas, validaciones y cálculos

- **Caché en RAM**: `m_ObjEntorno.ColPendientes*` (`Entorno.cls:52-61, 1142-1625, 2410-2440`) alimentan `CalcularTotalesPendientes` (`Entorno.cls:2442-2533`) y se invalidan al guardar una solicitud. Con «Todos», se aprovechan estas colecciones; con tramitador concreto, se consulta a la base (`:329-553`).
- **Auto-selección**: si el primer estado tiene pendientes, se selecciona automáticamente (`:521-528`); en caso contrario, el segundo (registro Marga). Decisión implícita del autor.
- **Última modificación**: sólo se muestra si `EsAdministrador = Sí` (`Variables Globales.bas:192`); `m_ObjEntorno.UltimoCambio.texto` (`Entorno.cls:UltimoCambio`).
- **Apertura por defecto del gestor**: si el usuario conectado figura en `Coltramitadores`, se selecciona; si no, «Todos» (`:564-571`).
- **Edición según `TIPO`**: «Alta» y «Traspaso» van a `FormSolicitudAltaTraspaso`; el resto exige `UsuarioHPS` no nulo (`:200-204`).

## Lecturas y escrituras

Once métodos `getSolicitudesPendientes*` (entre `constructor.bas:668-1513`) componen un `Scripting.Dictionary` con `CompareMode = TextCompare` indexado por `IDSolicitud`; las fichas de `Solicitud`, `SolicitudOperaciones`, `TbSolicitudes` y las tablas asociadas se documentan por separado. La escritura es indirecta: el formulario no muta la base; sólo abre el detalle correspondiente.

## Coexistencia

Pendiente de `../coexistencia.md`. Observaciones relevantes:

- Once consultas SQL casi idénticas (`:668-1513` `constructor.bas`) por estado, cada una con su propio `m_Solicitud` y poblado manual. En PostgreSQL podrían sustituirse por una vista materializada parametrizada o por una función de tabla.
- `RellenarListas` decide «cache vs consulta directa» con un `If m_EsTodos` repetido once veces (`:312-553`); un ORM eliminaría la duplicación.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `Gestor` lleva en el diseño un `RowSourceType = "Value List"` con dos nombres reales en `RowSource` (`.form.txt:1204-1205`): el binario expone datos personales incluso cuando `RellenarComboTramitadores` lo sobrescribe al abrir. **Defecto de seguridad sensible.**
- `lblUltimaModificacion` quedó con el literal inicial «Última modificación por Administrador para el Exp <Nº/YY> el <dd/mm/aaaa hh:mm:ss>» (`.form.txt:1373`): deja huella de un expediente real y de una marca temporal concreta en el binario, aunque sólo sea el caption de diseño. **Defecto de seguridad sensible.**
- `lblEstado` mantiene el rótulo heredado «ColTiposCorreos» (`.form.txt:1247`) y `ComandoAyuda.Caption = "Comando46"` (`.form.txt:809`).
- `ComandoActualizar.Caption = "Contador"` (`.form.txt:1103`) no se corresponde con su `ControlTipText = "Actualizar Tareas"`; el código lo trata como refresco (`:109-130`).
- `ListaFiltrados` se diseña con `ColumnCount = 5` (`.form.txt:906`) pero `RowSource = "IDSolicitud;DNI;Nombre;Estado"` (`.form.txt:917`); el filtro runtime añade cinco columnas (`:114-126`), por lo que el binario arrastra un esquema de cuatro columnas que nunca se aplica.
- `EstablecerDatos` decide auto-selección de «PendienteEnvioExcel» o «PendienteRegistroMarga» mirando las dos primeras filas (`:521-528`); cualquier reordenación futura del `ColEstadosParaTitulo` rompería la selección esperada.
- `Gestor_BeforeUpdate` no cierra el formulario tras error (`:269-279`); deja al usuario con un combo inconsistente y sin notificar al contenedor.
- `ComandoEditar_Click` exige `UsuarioHPS` no nulo sólo en renovaciones (`:200-204`); si la solicitud no carga `UsuarioHPS` por una migración de datos, salta con mensaje y no se ofrece corrección en línea.
- Once caminos casi idénticos en `Filtrar` (`:48-105`) listos para parametrizar mal un estado y no compilan el error en tiempo de ejecución (la rama `Else` solo sale silenciosamente, `:103-105`).
- `PintarIndicadores` se invoca en `ComandoActualizar_Click` (`:117`) pero no está implementado en este módulo; la búsqueda del símbolo cae en otra clase. Riesgo de regresión si se elimina sin tocar quien lo llama.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: worklist por rol tramitador; vista única `trabajos pendientes` filtrable por estado en el cliente; contador por estado obtenido de una vista materializada; selector de gestor con búsqueda por iniciales.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir, el combo muestra «Todos» y los tramitadores dados de alta sin nombres ajenos a la lista (override del `RowSource` de diseño).
2. Si el usuario conectado es tramitador, queda seleccionado por defecto; si no, queda «Todos».
3. La lista de la izquierda muestra once estados con el recuento entre paréntesis y oculta «Última modificación» al usuario no administrador.
4. Si el primer estado tiene pendientes, queda auto-seleccionado y la derecha muestra las solicitudes; si no, se selecciona el segundo (registro Marga).
5. Cambiar el gestor repinta los recuentos sin reiniciar la sesión.
6. Cambiar el estado actualiza el listado; con un único resultado, queda seleccionado.
7. Doble clic en una solicitud abre el formulario de detalle según `TIPO`.

## Pendientes

- Forma exacta de las once consultas (`constructor.bas:668-1513`): filtros por fechas, recordatorios y traspasos pendientes; qué columna de cada `Tb*` discrimina cada estado.
- Cómo se invalida el caché `m_ObjEntorno.ColPendientes*` desde el alta/edición de solicitudes: invariante descrita de pasada en `ResponsableOperaciones.cls:114`, falta confirmar el resto de `*Operaciones`.
- Cómo decide `ColEstadosParaTitulo` el orden de los once estados (`Entorno.cls:1781-1816`): relevante para la auto-selección implícita (`:521-528`).
- Qué hace `PintarIndicadores` y desde qué clase/módulo: el símbolo no aparece en este archivo.
- Política de visibilidad de `lblUltimaModificacion` para roles distintos de administrador y técnico.
