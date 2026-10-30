# Form_FormFacturaAlta

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_FormFacturaAlta.cls` (248 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Generador de facturas («GENERADOR DE FACTURAS», `Form_FormFacturaAlta.form.txt:40`). Es un formulario **no enlazado** (sin `RecordSource`): recoge fecha límite de eventos, títulos, tipo impositivo y recargo, y delega todo el cálculo y la escritura en la clase `Factura` (`classes/Factura.cls`). Ofrece dos acciones: **Generar Factura** (graba en `TbFacturaPrincipal` y tablas hijas, crea el Excel y abre el resultado) y **Simulacro** (calcula y muestra sin grabar). Abreviaturas: `FG.bas` = `modules/Funciones Generales.bas`; `Factura.cls` = `classes/Factura.cls`. Las citas `:línea` sin fichero son de `Form_FormFacturaAlta.cls`.

## Entrada y salida

| Desde / hacia | Cómo | Evidencia |
|---|---|---|
| `Form_FormFacturacionGestion` → este formulario | `ComandoNuevaFactura_Click` cierra el formulario si estaba abierto, lo abre y guarda la referencia `m_Form` | `Form_FormFacturacionGestion.cls:235-258` |
| Este formulario → `FormTipoTecnicoPreciosGestion` | `ComandoPerfiles_Click` | `:125-137` |
| Este formulario → `FormFacturaResultado` | Tras generar o simular, cierra el resultado si estaba abierto y lo abre; la factura viaja en la global `m_ObjFacturaActiva` | `:70, 105-108, 163, 184-187`; `Variables Globales.bas:249` |
| Este formulario → `FormFacturacionGestion` | Si está abierto, tras generar llama a `RellenarComboIDFactura` | `:109-111` |

Declara `Event Alta` (`:15`) pero **ningún `RaiseEvent`** lo lanza; el manejador de Gestión (`Form_FormFacturacionGestion.cls:628-636`) no se ejecuta nunca.

## Controles y eventos

Controles: 24 (36 `Begin` en el `.form.txt`: 12 de formulario, sección y plantillas por defecto + 24 controles).

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `FechaFinalEventos` | texto, formato fecha corta | — | Fecha límite de los eventos a facturar; se inicializa a la fecha del día; obligatoria (la valida `Factura`) | CAP-8 | `:224`; `.form.txt:429-457` |
| `Titulo1`, `Titulo2` | texto | — | Títulos opcionales; si `Titulo1` queda vacío se calcula al grabar | CAP-8 | `:74-75`; `.form.txt:543-627`; `Factura.cls:2472-2474` |
| `TipoImpositivo` | texto, número general | — | Se inicializa a 21; si no es numérico se sustituye por el calculado | CAP-8 | `:222`; `.form.txt:655-685`; `Factura.cls:2086-2099` |
| `RecargoSobreSuministros` | texto, número general | — | Se inicializa a 15; si queda vacío se usa el calculado | CAP-8 | `:223`; `.form.txt:713-743`; `Factura.cls:2071-2084` |
| `ListaPreciosPerfil` | lista de valores | — | Muestra los perfiles activos con precio hora laboral/extra y vigencia | CAP-9 | `:23-55`; `.form.txt:813-853` |
| `ComandoGenerarFactura` | botón | `OnClick` → `ComandoGenerarFactura_Click` | Genera y graba la factura | CAP-8 | `:59-122`; `.form.txt:788-812` |
| `ComandoSimulacro` | botón | `OnClick` → `ComandoSimulacro_Click` | Simula la factura sin grabar | CAP-8 | `:139-201`; `.form.txt:854-878` |
| `ComandoPerfiles` | botón | `OnClick` → `ComandoPerfiles_Click` | Abre la gestión de precios por perfil | CAP-9 | `:125-137`; `.form.txt:879-902` |
| `cmdSalir` | botón | `OnClick` → `cmdSalir_Click` | Cierra sin guardar | CAP-8 | `:17-20`; `.form.txt:490-515` |
| `lblEstado` | etiqueta (oculta en diseño) | — | Línea de avance; se oculta al abrir y antes de abrir el Excel | CAP-8 | `:102, 181, 216`; `.form.txt:458-489` |
| `lblTitulo1`, `lblTitulo2`, `lblNumeroPartes` | etiquetas | — | Título (su rótulo se copia a `Caption`) y rótulo «Precios por Perfil» | CAP-8 | `:215`; `.form.txt:386-427, 903-930` |
| `Etiqueta52`, `54`, `56`, `58`, `70` | etiquetas | — | Rótulos de campo | CAP-8 | `.form.txt:516-542, 572-598, 628-654, 686-712, 931-957` |
| `Etiqueta99`, `Etiqueta61`, `Etiqueta71` | etiquetas «*» | — | Marcas de campo (sin lógica asociada) | CAP-8 | `.form.txt:744-787, 958-1000` |
| `Cuadro1`, `CuadroSuperior` | rectángulos | — | Decoración | CAP-8 | `.form.txt:340-385` |

## Flujos

### Apertura (`Form_Load`, `:204-239`)

1. Reloj de arena, `AjustarTamaño Me` (`:210-212`), título de ventana y `lblEstado` oculto (`:215-216`).
2. `RellenarListaPerfiles` lee `m_ObjEntorno.ColPerfilesActivos` y rellena la lista con `Format` monetario (`:31-48`).
3. Valores por defecto: tipo 21, recargo 15, fecha del día (`:222-224`).
4. Error: `MsgBox` crítico y cierre del formulario (`:232-238`).

### Generar factura (`:59-122`)

1. Aviso «El proceso puede tardar varios minutos» (`:66`) y `m_ObjFacturaActiva = Nothing` (`:70`).
2. Crea `Factura` con `ParaPorFacturar = Sí`, copia los controles con `Nz` y fija `FechaFactura = FechaFinalEventos` (`:71-79`).
3. `ValidacionFactura No` (`:80`): comprueba fecha, resuelve recargo e impuesto y recorre las colecciones «por facturar» (eventos sin equipo, actividades sin técnico, materiales sin actividad) (`Factura.cls:2006-2300`).
4. `GrabarFactura` (`:84`; `Factura.cls:3311-3387`): sin nada pendiente, «No hay nada que facturar para esa fecha» (`Factura.cls:3359-3367`); si no, graba cabecera, eventos, actividades, materiales, gastos, subcontrataciones y perfiles (`Factura.cls:3218-3246`).
5. Relee la factura con `Constructor.getFactura` (`:90-97`), genera el Excel con `SalidaAExcel Sí` (`:98`) y lo abre con ShellExecute (`:103`); marca `ParaPorFacturar = No` (`:104`), abre el resultado y refresca Gestión (`:105-111`).
6. Error: oculta el reloj y `MsgBox` crítico con `m_Error`; el formulario sigue abierto (`:116-121`).

### Simulacro (`:139-201`)

Igual que Generar hasta la validación (`:150-162`), pero **no** llama a `GrabarFactura`: usa `SalidaAExcel No` (`:165`), abre el Excel, marca `ParaPorFacturar = Sí` y abre el resultado (`:182-187`). Sin `IDFactura`, el fichero se llama con prefijo `000` (`Factura.cls:971-974`).

## Reglas, validaciones y cálculos

- **Fecha obligatoria:** sin fecha válida, «No se ha indicado la fecha final de los eventos» (`Factura.cls:2039-2042`).
- **Eventos facturables:** `TbEventos` con `FechaFinal` ≤ fecha, `Franqueado = True` y sin fila en `TbFacturaEventosInvolucrados` (`Constructor.bas:1206-1211`). Materiales: los de esas actividades sin fila en `TbFacturaMaterialesInvolucrados` (`Constructor.bas:1483-1486`).
- **Identificador:** `DameID("TbFacturaPrincipal", "IDFactura")` = máximo numérico + 1 (`Factura.cls:1001`; `FG.bas:408-427`).
- **Fichero:** `Format(IDFactura,"000")` + `FACTURA_` + `aammdd` + `.xlsx` en el directorio de facturas del entorno (`Factura.cls:976-978, 1029-1035`). Si existe y está abierto, «Ya existe una factura y parece estar abierta» (`Factura.cls:5181-5186`); si existe y no, se borra y se reescribe (`Factura.cls:5233-5235`).
- **Libro:** hojas DETALLE GETRONICS, MATERIALES, ACTIVIDADES, EVENTOS y FACTURA (`Factura.cls:5199-5230`).
- Los importes (horas, suministro, recargo, base, impuestos) los calculan las propiedades `Precio*PorFacturar` de `Factura` (`Factura.cls:262-1969`); su fórmula queda fuera de esta ficha.

## Lecturas y escrituras

| Operación | Tabla | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Eventos pendientes | `TbEventos`, `TbFacturaEventosInvolucrados` | SQL dinámico concatenado (fecha como literal `#mm/dd/yyyy#`) | — | `Constructor.bas:1206-1211` |
| Materiales pendientes | `TbMaterial`, `TbActividades`, `TbFacturaMaterialesInvolucrados` | SQL dinámico con subconsulta de eventos incrustada | — | `Constructor.bas:1483-1486` |
| Siguiente `IDFactura` | `TbFacturaPrincipal` | recorrido de recordset (máximo + 1) | no | `FG.bas:408-427` |
| Cabecera | `TbFacturaPrincipal` | DAO `AddNew`/`Update` | no | `Factura.cls:2465-2533` |
| Eventos, actividades, materiales, gastos, subcontrataciones | `TbFacturaEventosInvolucrados`, `TbFacturaActividadesInvolucradas`, `TbFacturaMaterialesInvolucrados`, `TbFacturaGastosInvolucrados`, `TbFacturaSubcontratacionesInvolucradas` | DAO `AddNew` por elemento | no | `Factura.cls:2571-2580, 2652-2671, 2834-2881, 2983-2996, 3058-3066` |
| Perfiles | `TbFacturaPrincipalPerfiles` | DAO `AddNew` (ID con `DameID`) | no | `Factura.cls:3282-3297` |
| Compensación de fallo | `TbFacturaPrincipal` | `DELETE` concatenado de la cabecera | no | `Factura.cls:3250-3253` |
| Relectura | `TbFacturaPrincipal` | SQL concatenado por ID | — | `Constructor.bas:2440-2441` |

## Coexistencia

Adaptar al PostgreSQL vinculado: literal de fecha `#mm/dd/yyyy#` (`Constructor.bas:1209`); `DELETE tabla.campo FROM` y `RIGHT JOIN` de dialecto Jet (`Factura.cls:2419-2445`); `Max + 1` en cliente en lugar de secuencia; dos escrituras sin transacción (cabecera y ~6 tablas hijas); `Excel.Application` y ShellExecute (`Factura.cls:5195`; `:103`) quedan fuera de la base de datos.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Sin transacción: si falla una tabla hija, solo se borra la cabecera (`Factura.cls:3250-3253`) y pueden quedar filas hijas con ese `IDFactura`.
- Si falla `SalidaAExcel` tras `GrabarFactura` (`:98-101`), la factura queda grabada y sus eventos pasan a «involucrados»: no volverán a salir como pendientes hasta eliminarla desde Gestión.
- `EliminarFacturasAnteriores` escribe el error en `m_Error` y no en `p_Error` (`Factura.cls:2447-2450`) y no borra `TbFacturaPrincipalPerfiles` (`Factura.cls:2419-2445`); con `IDFactura` vacío no actúa (`Factura.cls:2416-2418`).
- `Max + 1` sin bloqueo (`FG.bas:408-427`): dos generaciones simultáneas podrían obtener el mismo ID.
- Los simulacros del mismo día comparten nombre de fichero (`000FACTURA_aammdd`) y se sobrescriben (`Factura.cls:971-974, 5233-5235`).
- `Event Alta` sin `RaiseEvent` (`:15`): código muerto en Gestión (`Form_FormFacturacionGestion.cls:628-636`).
- `FechaFactura` siempre igual a `FechaFinalEventos` (`:77, 157`).
- Rótulo de diseño de `ListaPreciosPerfil` con una fila de ejemplo (`.form.txt:831`), sustituida al cargar (`:32`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: transacción única para cabecera e hijas; identificador generado por la base de datos; fecha de factura independiente; generación del documento desacoplada del grabado.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrir, tipo impositivo 21, recargo 15 y fecha del día; la lista muestra los perfiles activos con sus precios (`:218-224`).
2. Generar con la fecha vacía muestra «No se ha indicado la fecha final de los eventos» y no escribe.
3. Generar sin eventos, actividades, materiales, gastos ni subcontrataciones pendientes muestra «No hay nada que facturar para esa fecha» y no escribe.
4. Generar con datos crea la cabecera y las filas hijas, produce el Excel y abre `FormFacturaResultado` con la factura.
5. Una vez generada, los eventos incluidos no vuelven a ser candidatos a facturar.
6. Simulacro muestra el resultado y el Excel sin crear filas en `TbFacturaPrincipal`.
7. Un evento sin equipo, o una actividad sin técnico, aborta con el mensaje del elemento afectado.

## Pendientes

- Fórmulas de `Precio*PorFacturar` y de `Constructor.getActividadesDeEventosParaFactura` (`Constructor.bas:1527`), `getActividadesNoDeEventosParaFactura` (`Constructor.bas:1402`), `getGastosParaFactura` (`Constructor.bas:1135`) y `getSubContratacionesParaFactura` (`Constructor.bas:1073`), no leídas aquí.
- Valor de `m_ObjEntorno.URLDirectorioFacturas` (configuración, `Entorno.cls:27`) y contenido de `AjustarTamaño` (`:212`).
- Destino de `Titulo3` (campo de `Factura`, sin control en este formulario; `Factura.cls:20, 2481`).
