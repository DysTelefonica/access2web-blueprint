# Form_FormPlanificacionAltaEquipo

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormPlanificacionAltaEquipo.cls` (85 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Diálogo para marcar un equipo como planificable para mantenimiento preventivo: pide un alias y una periodicidad en meses y da de alta el equipo en `TbPlanificacionEquipos`. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): el equipo llega por `OpenArgs` y la escritura la hace `AltaEquipoEnPlanificacion` (`FG.bas:2781-2849`). Abreviatura: `FG.bas` = `modules/Funciones Generales.bas`. Las citas sin fichero son `Form_FormPlanificacionAltaEquipo.cls:línea`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormEquipoGestion` (`ComandoEstablecerParaPlanificacion_Click`) | Exige un equipo numérico seleccionado, cierra el formulario si ya estaba abierto y lo abre con `OpenArgs` = `ID|nombre` | `Form_FormEquipoGestion.cls:194-213` |

No lanza eventos ni devuelve valores al llamador. Tras registrar se cierra solo (`:28`); `cmdSalir` cierra sin guardar (`:41-44`). El llamador no se entera del alta (no hay `WithEvents` sobre este formulario).

## Controles y eventos

Controles: 12 (12 `Begin` de controles dentro de la sección Detalle del `.form.txt`, líneas 254-567; excluye `Begin Form`, `Begin Section` y las plantillas de valores por defecto).

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `CuadroSuperior`, `Cuadro2` | rectángulos (2) | — | Marcos decorativos | CAP-6 Planificación | `.form.txt:254, 276` |
| `lblTitulo1`, `lblTitulo2` | etiquetas (2) | — | Rótulo «ALTA DE EQUIPO PARA MTO PREVENTIVO»; `lblTitulo1` da título a la ventana | CAP-6 Planificación | `:54`, `.form.txt:545-578` |
| `Etiqueta235`, `Etiqueta199`, `Etiqueta200` | etiquetas (3) | — | Rótulos «EQUIPO», «ALIAS», «PERIODICIDAD EN MESES RECOMENDADA» | CAP-6 Planificación | `.form.txt:377, 431, 458` |
| `ComboEquipoPlanificable` | cuadro de texto bloqueado (`Locked`) | — | Muestra el nombre del equipo recibido; no editable | CAP-6 Planificación | `:72`, `.form.txt:515-529` |
| `Alias` | cuadro de texto no enlazado | — | Alias opcional | CAP-6 Planificación | `:11`, `.form.txt:404-415` |
| `PERIODICIDADMESES` | cuadro de texto, formato «General Number» | — | Meses de periodicidad recomendada; opcional | CAP-6 Planificación | `:12-15`, `.form.txt:485-498` |
| `ComandoAlta` (rótulo «Registrar») | botón | `OnClick` → `ComandoAlta_Click` | Da de alta el equipo y cierra | CAP-6 Planificación | `:5-39`, `.form.txt:300-310` |
| `cmdSalir` | botón | `OnClick` → `cmdSalir_Click` | Cierra sin guardar | CAP-6 Planificación | `:41-44`, `.form.txt:338-348` |

## Flujos

### Apertura (`Form_Load`, `:45-86`)

1. Cursor de espera, `AjustarTamaño Me` y título de ventana tomado de `lblTitulo1` (`:49-54`).
2. Lee `OpenArgs`; si no contiene `|` ⇒ «El formulario se ha abierto con datos insuficientes» (`:56-60`). Separa `strIDEquipo` y `strNombreEquipo` (`:61-63`); exige ID numérico y nombre no vacío con el mismo mensaje (`:64-71`).
3. Vuelca el nombre en `ComboEquipoPlanificable` (`:72`).
4. Cualquier error: `MsgBox` crítico y cierra el formulario (`:79-85`).

### Alta (`ComandoAlta_Click`, `:5-39`)

1. Lee `Alias` con `Nz`; si `PERIODICIDADMESES` no es numérico, usa 0 (`:11-15`).
2. Llama a `AltaEquipoEnPlanificacion(strIDEquipo, strAlias, , CInt(strNumeroMeses))` (`:22`); sin fecha de alta explícita.
3. Si el resultado contiene `|` se trata como error y se muestra «El método AltaEquipoEnPlanificacion ha devuelto el error: …» (`:23-27`).
4. Éxito: cierra el formulario sin guardar (`:28`). No muestra confirmación de éxito.
5. Error: restablece el cursor y `MsgBox` crítico; el formulario queda abierto (`:33-38`).

## Reglas, validaciones y cálculos

- **ID numérico:** «El Equipo lo identifica un identificador numérico» (`FG.bas:2803-2806`).
- **Fecha de alta:** al no pasarse, vale la fecha del sistema (`FG.bas:2807-2814`; la validación `IsDate` del otro camino no se ejercita desde este formulario).
- **Unicidad:** si ya existe un registro con ese `IDEquipo`, «No se puede dar de alta a un equipo que ya  forma parte de los planificables» (`FG.bas:2815-2823`).
- **Campos opcionales:** `Alias` solo se escribe si no está vacío y `PeriodicidadEnMesesRecomendada` solo si es distinta de 0 (`FG.bas:2827-2832`). Un valor no numérico equivale a «sin periodicidad» (`:13-15`).
- **Conversión:** `CInt` sobre el texto numérico (`:22`); un decimal o un valor fuera de rango de `Integer` no se valida aquí.

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Comprobar si el equipo ya es planificable | `TbPlanificacionEquipos` | SQL dinámico concatenado (`IDEquipo=` + texto), `OpenRecordset` | no | `FG.bas:2815-2818` |
| Alta (`IDEquipo`, `FechaAltaParaPlanificacion`, `Alias`, `PeriodicidadEnMesesRecomendada`) | `TbPlanificacionEquipos` | DAO `AddNew`/`Update` | no | `FG.bas:2824-2833` |

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: SQL con el ID concatenado sin comillas; asignación de una cadena de fecha (`Date` convertido a `String`) a un campo de fecha (`FG.bas:2813, 2826`), sensible al formato regional; comprobación de existencia y alta separadas, sin transacción; la columna `PeriodicidadEnMesesRecomendada` se escribe como entero.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Carrera entre comprobar y escribir (`FG.bas:2815-2833`): dos altas simultáneas del mismo equipo podrían insertar duplicados si la tabla no tiene índice único.
- `strFechaAlta = Date` guarda la fecha como texto (`FG.bas:2813`); según la configuración regional podría interpretarse al revés día y mes.
- Un valor no numérico en la periodicidad se convierte en silencio a 0 y no se guarda (`:13-15`, `FG.bas:2830`), sin avisar al usuario.
- `CInt` desborda con valores mayores de 32767 o con decimales mal formados; el error se captura como genérico (`:22`, `:34-36`).
- `flag`, `dato` y `pregunta` son globales compartidas (`Variables Globales.bas:204-206`; también `m_SQL`, `Variables Globales.bas:203`), usadas aquí sin declarar (`:22, 24, 38`): estado compartido entre formularios.
- El formulario no avisa de éxito ni informa al llamador; `Form_FormEquipoGestion` no refresca nada tras el cierre (`Form_FormEquipoGestion.cls:213`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: restricción única sobre `IDEquipo`; validación explícita de la periodicidad (entero positivo); confirmación de éxito.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir sin `OpenArgs`, con ID no numérico o sin nombre muestra «El formulario se ha abierto con datos insuficientes» y se cierra (`:56-71, 84-85`).
2. El campo del equipo muestra el nombre recibido y no es editable (`:72`, `.form.txt:516`).
3. Registrar un equipo nuevo crea una fila en `TbPlanificacionEquipos` con la fecha del día y cierra el formulario (`FG.bas:2824-2833`, `:28`).
4. Registrar un equipo ya planificable muestra el mensaje de «ya forma parte de los planificables» y no escribe (`FG.bas:2820-2823`).
5. Alias y periodicidad vacíos o 0 se registran sin esos campos (`FG.bas:2827-2832`).
6. Una periodicidad no numérica se trata como 0 (`:13-15`).
7. `Salir` cierra sin escribir (`:41-44`).

## Pendientes

- Contenido de `AjustarTamaño` (`:51`; definición en `FG.bas:261`).
- Si `TbPlanificacionEquipos` tiene índice único sobre `IDEquipo` y tipos/longitudes de columna (ERD de origen).
- Qué hace `Form_FormPlanificacionEquipos` con el mismo método y 12 meses fijos (`Form_FormPlanificacionEquipos.cls:839`), por si debe unificarse en la web.
