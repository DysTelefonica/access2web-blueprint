# Form_FormExpedienteGeneral

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/forms/Form_FormExpedienteGeneral.cls` (905 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Pestaña General de la ficha de expediente: datos de cabecera (códigos, título, objeto, importes, clasificación, órganos, responsables, ámbito, AGEDYS, HPS, S4H, Sharepoint). Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`): lee y escribe el DTO global `m_ObjExpedienteDTOActivo.Expediente`; no escribe en la base de datos. Las citas son `Form_FormExpedienteGeneral.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormExpediente` | `FrmDetalle.SourceObject = "FormExpedienteGeneral"` al abrir; `NavGeneral` como destino de navegación | `Form_FormExpediente.cls:146, 450` |

Abre `FormOrganoContratacion`, `FormEjercito`, `FormGradoClasificacion` y `FormOficinaPrograma` (alta de maestros) y escucha su evento `Alta` mediante `WithEvents` (`:6-9`, `:45-160`). Abre la URL de `AccesoSharepoint` con `ShellExecute` (`Ejecutar`, `Variables Globales.bas:24`; uso en `:186-190`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| `Ambito` | combo (`DEFENSA;FUERA;HPS`) | `AfterUpdate` | Con «HPS» fija `HPSAplica = "Sí"` | `:21-30`, `.form.txt:2335-2338` |
| `Ordinal` | combo libre (sin `LimitToList`) | `AfterUpdate` | Avisa si el ordinal ya existe bajo el mismo padre | `:779-835`, `.form.txt:2292-2297` |
| `AccesoSharepoint` | texto | `OnExit` | Habilita `ComandoVerSharepoint` solo si no está vacío | `:13-19`, `.form.txt:1496-1497` |
| `ComandoVerSharepoint` | botón | `OnClick` | Abre la URL si no está vacía | `:186-190` |
| `ComandoRellenarNemotecnico` | botón | `OnClick` | `Form_FormExpediente.RellenarDTO` y escribe `NemotecnicoCalculado` en `Nemotecnico` | `:162-184` |
| `ComandoNuevoOC`, `ComandoNuevoEjercito`, `ComandoNuevaClasificacion`, `ComandoNuevaOficinaPrograma` | botones | `OnClick` | Cierran el maestro si estaba abierto, vacían su global activa y lo abren | `:45-160` |
| `IDOrganoContratacion`, `IDEjercito`, `IdGradoClasificacion`, `IDOficinaPrograma` | combos 2 columnas (ID oculto) | — | Rellenados desde `m_ObjEntorno`; el DTO toma `.Column(0)` | `:435-505, 880-885` |
| `IDResponsableCalidad`, `IDResponsableSeguridad` | combos 2 columnas | — | Usuarios por rol; Seguridad añade `0;N/A` y `DefaultValue = 0` | `:507-551` |
| `AGEDYSAplica` | combo (1 No · 2 Sí normal · 3 Sí genérico) | — | Se traduce al par `AGEDYSAplica` + `AGEDYSGenerico` | `:280-297, 856-872` |
| `HPSAplica`, `APLICAESTADO`, `EnListas`, `AplicaTareaS4H` | combos `Sí;No` | — | Valores de cabecera; `EnListas` representa `POSTAGEDO` | `:299-301, 276-278, 409-418` |
| `TIPO` | texto bloqueado | — | Muestra `TipoCalculadoTexto` | `:393-395`, `.form.txt:2093-2108` |
| `OrdinalE2E` + `lblOrdinalE2E` | texto deshabilitado | — | Visibles solo con enlace E2E activo | `:262-273` |
| `Titulo`, `ObjetoContrato`, `Nemotecnico`, `CodExp`, `CodExpLargo`, `ImporteLicitacion`, `ImporteContratacion`, `CodS4H`, `NPedido`, `CodProyecto`, `CodigoActividad`, `Observaciones` | textos | — | Copia directa DTO ↔ control con `Nz` | `:303-386, 874-893` |

Los `RowSource` estáticos de diseño en el `.form.txt` (p. ej. `.form.txt:1023, 1432, 2837`) se sobrescriben en ejecución por `EstablecerCombos`, salvo los combos `Sí;No`, `Ambito` y `AGEDYSAplica`, que se rellenan con la lista del entorno (`:553-564`).

## Flujos

### Apertura (`Form_Load`, `:194-227`)

1. `AllowEdits = False`; pasa a `True` solo si `Form_FormExpediente.m_AbiertoParaEditar` y `EsAdministrador = Sí` (`:202-209`).
2. `EstablecerDatos`: sin `Expediente` en el DTO, error «El formulario se ha abierto con parámetros insuficientes» (`:237-240`). Con `IDExpediente` informado, **sustituye** el expediente del DTO por uno recién leído con `constructor.getExpediente` (`:241-246`).
3. `EstablecerCombos` rellena los combos (`:248-251`; detalle en Reglas).
4. Visibilidad E2E con `E2EEnlacesUIHelper_OrdinalE2EVisible` (`E2EEnlacesUIHelper.bas:26-41`); cualquier error, incluido `CLng("")` en un alta, se traga y oculta el campo (`:262-268`).
5. Copia cada campo no vacío del DTO a su control (`:274-418`). Importes con `Format(…, "#,##0.00 €")` (`:309-314`). `POSTAGEDO` vacío muestra `EnListas = "Sí"` (`:409-413`).
6. En un alta con `TIpo` informado, `EstablecerDatosConTipo` fija `EsAM`/`EsLote`/`EsExpediente` en el DTO según el texto de `TIpo` (`:388-392, 725-778`).

### Salida (`Form_Unload`, `:580-599`)

Llama a su propio `RellenarDTO` (`:587`), que copia 23 propiedades al DTO (`:846-904`).

### Errores

Los handlers acumulan el error en `m_Error`, envían `CorreoAlAdministrador` salvo el error interno 1000 y muestran `MsgBox` crítico (p. ej. `:36-43`). `EstablecerDatos`, `EstablecerCombos` y `RellenarDTO` solo devuelven el texto en `p_Error` (`:426-430, 570-574, 900-903`).

## Reglas, validaciones y cálculos

- **Tipo y número:** `AM` o `EXPIndividual` ocultan `lblNumero`, `Ordinal` y `ComandoRellenarNemotecnico`; `Lote` rotula «Nº DE LOTE»; `BasadoDeAM`/`BasadoDeLote` rotulan «Nº BASADO» (`:327-360`).
- **Ordinal propuesto:** en un alta con ordinal no numérico se propone `OrdinalCalculado` = último derivado del padre + 1, o «01» si no hay (`:335-340`; `classes/Expediente.cls:1721-1745`). El combo ofrece 01-20 (`:566-568`).
- **Ordinal repetido:** solo con padre, y en alta o si cambia frente a `m_ObjExpedienteAlInicio` (`:795-808`); `ExpedienteOrdinalUsado` busca en `TbExpedientes` (`FUNCIONES UTILES.bas:4156-4196`) y avisa «El número ya ha sido usado para el Expediente : <CodExp, Nemotécnico o Título>» (`:813-829`). **Solo avisa; no impide el valor.**
- **HPS:** `Ambito = "HPS"` ⇒ `HPSAplica = "Sí"` en pantalla (`:28-30`); `Registrar` repite la regla (`classes/ExpedienteOperaciones.cls:66`).
- **AGEDYS** (carga): `AGEDYSAplicaCalculado` No ⇒ 1; Sí y genérico No ⇒ 2; Sí y genérico Sí ⇒ 3; otro ⇒ vacío (`:283-297`). Al volcar, el número vuelve al par de textos (`:857-872`).
- **Responsables:** `GetUsuariosPorRol` lee `TbResponsablesPorRol` y `TbUsuariosAplicaciones` (usuarios sin `FechaBaja`) y añade el responsable guardado aunque ya no tenga el rol (`modResponsablePorRolHelper.bas:78-140`).
- **Maestros:** combos desde las cachés `ColOrganosContrataciones`, `GradosClasificaciones`, `Ejercitos`, `OficinasDePrograma` y `ColTiposAplicaAgedys` de `Entorno` (`classes/Entorno.cls:1164, 1134, 895, 1075, 417`). Tras el `Alta` de un maestro se invalida su caché, se recargan los combos y se selecciona el nuevo (`:601-723`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Recarga del expediente | vía `constructor.getExpediente` | DAO (no trazado aquí) | `:242` |
| Ordinal usado | `TbExpedientes` (`IDExpedientePadre`, `Ordinal`) | SQL dinámico concatenado | `FUNCIONES UTILES.bas:4172-4176` |
| Responsables por rol | `TbResponsablesPorRol`, `TbUsuariosAplicaciones` | DAO, dos pasos | `modResponsablePorRolHelper.bas:120-140` |
| Enlace E2E activo | `TbE2EEnlaces` vía `E2EEnlacesService.EstaVinculado` | DAO | `E2EEnlacesUIHelper.bas:38` |

Sin escrituras propias: la persistencia la hace `ExpedienteOperaciones.Registrar` desde [Form_FormExpediente](Form_FormExpediente.md).

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: SQL concatenado con `Ordinal` entre comillas e `IDExpedientePadre` sin comillas (`FUNCIONES UTILES.bas:4172-4175`); importes formateados como texto con «€» antes de volver al DTO.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `HPSAplica` no se copia al DTO en ningún `RellenarDTO` (ni `:846-904` ni `Form_FormExpediente.cls:502-546`): un cambio manual no se guardaría, salvo la regla HPS de `Registrar`.
- El `RellenarDTO` propio (usado en `Form_Unload`, `:587`) omite `POSTAGEDO` (`EnListas`) y `AplicaTareaS4H`, que sí copia el del contenedor (`Form_FormExpediente.cls:545-546`).
- Cada carga de la pestaña con `IDExpediente` relee el expediente de la base y reemplaza el DTO (`:241-246`): los cambios sin guardar de General o Fechas podrían perderse al volver a General.
- `m_FormOC_Alta` invalida `m_ObjEntorno.OrganosContratacion` (`Form_FormExpedienteGeneral.cls:671`) mientras el combo lee `ColOrganosContrataciones` (`:449`), pero `OrganoContratacionOperaciones.Registrar` ya vacía `ColOrganosContrataciones` al guardar (`OrganoContratacionOperaciones.cls:82`), así que el órgano nuevo sí aparecería: la invalidación de `OrganosContratacion` es redundante, no un defecto. Corregido tras leer la ficha de `Form_FormOrganoContratacion`.
- `EstablecerCombos` añade 01-20 a `Ordinal` sin vaciar su `RowSource` (`:565-568`); cada recarga tras un alta duplica opciones (el diseño ya guarda duplicados, `.form.txt:2294`).
- Literal «Lote sin Acuero Marco» (`:741`): si el texto real es «Acuerdo», ese tipo no fija `EsAM`/`EsLote`/`EsExpediente`.
- Los botones de alta de maestros y `ComandoRellenarNemotecnico` no se deshabilitan en solo lectura (`AllowEdits` no afecta a `OnClick`, `:202-209`).
- Restos sin efecto: `m_SiguienteOrdinal` nunca se asigna (`:233, 349-351`); `Debug.Print` (`:422`); `ObjetoContrato` se asigna dos veces (`:318-320, 403-405`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: validar el ordinal repetido como error bloqueante; guardar importes como número y formatear solo en la vista; cargar los combos con un endpoint de maestros con invalidación explícita.

## Criterios de aceptación

Propuestos, no aprobados:

1. Elegir ámbito «HPS» marca «Aplica HPS = Sí».
2. En un lote o basado, introducir un ordinal ya usado bajo el mismo padre muestra el aviso con el código del expediente que lo usa.
3. En AM o expediente individual no se muestra el campo de número ni el botón de nemotécnico.
4. El campo Ordinal E2E solo se ve con enlace E2E activo.
5. Crear un maestro desde la pestaña lo deja seleccionado en su combo.
6. Un responsable guardado que ya no tiene el rol sigue apareciendo en su combo.

## Pendientes

- Si Access permite asignar valores por código con `AllowEdits = False` en controles no enlazados (`:202, 274-418`).
- Si `IsNumeric` acepta el importe formateado con «€» al guardar (`classes/ExpedienteOperaciones.cls:124`).
- Valores reales de `TIpo` en un alta (origen en `Form_FormExpedienteAltaTipo`).
- Orden entre `Form_Unload` de la pestaña y `RellenarDTO` del contenedor al navegar.
- Contenido de `GetNemotecnicoCalculado` y `getTipoCalculado` (`classes/Expediente.cls:675, 1604`).
