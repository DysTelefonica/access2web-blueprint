# Capacidades — Gestión de Riesgos

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Issue:** #1079 · **Fase:** 5 · **PRD:** [README](README.md)

Clasificación de las capacidades que enumera el [estudio previo](../capabilities.md) frente al código del commit fijado. Cada fila dice si la capacidad **sigue igual**, se ha **modificado**, se ha **añadido** o se ha **retirado**, con la evidencia que lo sostiene.

No se propone una lista de capacidades para la web: eso es el resultado del PRD, no su punto de partida. Aquí solo se reconcilia lo observado con lo que el estudio previo afirmaba.

## Método y límites

- **Fuente de contraste:** el [estudio previo](../capabilities.md) del snapshot `data/staging` del 2026-08-06, con sus 28 dominios.
- **Comprobación:** presencia de la clase, el módulo o el formulario que la sostiene en el commit fijado, y lectura del código donde la afirmación es de comportamiento.
- **«Modificada» es diferencia de texto, no de comportamiento.** Cuando el objeto sigue existiendo con la misma forma pero distinto contenido, se dice «modificada» y se explica qué cambió, sin afirmarse que cambie lo que el usuario ve.
- **Límite de este documento:** solo cuatro de los 63 formularios tienen ficha. Las capacidades cuya evidencia es un formulario **sin ficha todavía** se marcan con el nombre del formulario y quedan **pendientes de confirmar en su ficha**. No se afirma comportamiento sin haberlo leído.

## Resumen

| Resultado | Capacidades | Detalle |
|---|---|---|
| Sin cambios | 14 | La forma del objeto es idéntica y nada sugiere cambio de comportamiento |
| Modificadas | 9 | El objeto existe pero su contenido difiere del snapshot |
| Añadidas | 1 | Existe en el commit fijado y no en el estudio previo |
| Retiradas | 0 | Ninguna capacidad del estudio previo ha desaparecido |
| Pendientes de ficha | 4 | Requieren leer el formulario antes de afirmar comportamiento |

**Ninguna capacidad se retira.** Es el hallazgo de más peso para la migración: la aplicación **no ha reducido su superficie funcional** respecto al snapshot; ha cambiado de forma interna.

## Clasificación

### Dominio

| Capacidad (estudio previo) | Objeto en `36fcf60` | Resultado | Evidencia |
|---|---|---|---|
| Arranque e identidad | `EVE` (`Variables Globales.bas:204`), `Usuario.cls`, `Entorno.cls` | sin cambios | Los tres objetos existen; `EVE` se sigue invocando desde el arranque (`Form_frmSplash.cls:61`) |
| Roles (administrador, calidad, técnico) | `Usuario.cls`, `UsuarioAplicacionPermisos.cls` | modificada | La identidad se resuelve leyendo el backend de **Lanzadera** (`getdbLanzadera`, 10 llamadas), no un servicio; el mecanismo es el descrito y la dependencia es externa |
| Configuración temporal | `Variables Globales.bas` | modificada | Las TempVar siguen siendo el mecanismo (126 usos), pero `CadenaJerarquicaModelo` se **fuerza a «nuevo»** en el arranque (`Form_frmSplash.cls:25-29`): el modelo conmutable que documentaba el estudio previo ya no es elegible desde ahí |
| Proyecto | `Proyecto.cls`, `ProyectoSuministrador.cls` | sin cambios | Ambas clases existen; `ProyectoSuministrador` además no está en el snapshot (ver [README](README.md#anexo-cambios-no-liberados)) |
| Edición | `Edicion.cls` | sin cambios | Clase presente y usada por el grueso de la aplicación |
| Riesgo (alta, edición, borrado) | `Riesgo.cls` | sin cambios | Clase presente |
| Estados del riesgo (14 estados) | `Riesgo.cls`, `RiesgosGestionRiesgo` | **pendiente de ficha** | El objeto existe; el comportamiento se afirmará en la ficha de `Form_FormRiesgosGestionRiesgo`, todavía no escrita |
| Causa raíz y biblioteca | `RiesgoBiblioteca.cls` | sin cambios | Clase presente |
| Riesgo externo | `RiesgoExterno.cls` | sin cambios | Clase presente |
| Materialización | `RiesgoMaterializacion.cls` | **modificada** | El formulario `Form_FormMaterializacionPlanContingencia`, que el snapshot tenía, **no está en `main`**; la clase sobrevive pero su formulario de alta desaparece (ver [README](README.md#inventario-reconciliado)) |
| Planes de mitigación (PM) | `PM.cls`, `PMAccion.cls`, `PMAccionReversa.cls` | sin cambios | Las tres clases existen |
| Planes de contingencia (PC) | `PC.cls`, `PCAccion.cls`, `PCAccionReversa.cls` | sin cambios | Las tres clases existen |
| Evidencias y anexos | `Anexo.cls`, `AnexoAntiguo.cls`, `EdicionSuministrador.cls` | sin cambios | Las tres clases existen; `EdicionSuministrador` no está en el snapshot |
| Suministradores por edición | `EdicionSuministrador.cls`, `Suministrador.cls` | sin cambios | Ambas clases existen |
| Control de cambios | `CCVersion.cls`, `CCCambio.cls`, `CCDocumentoCambio.cls` | sin cambios | Las tres clases existen, más `ControlCambios.bas` |
| Publicabilidad | `PublicabilidadEdicion.bas`, `PublicabilidadRiesgo.bas`, `PublicacionLog.cls` | sin cambios | Los dos módulos y la clase existen |
| Calidad — tareas | `TareasCalidad.cls`, `ArbolTareasCalidad.cls` | sin cambios | Ambas clases existen |
| Técnico — tareas | `TareasTecnico.cls`, `ArbolTareasTecnico.cls` | sin cambios | Ambas clases existen |
| Indicadores | `mIndicador.bas`, `modIndicadorDashboard.bas`, `IndicadorRiesgosRepositorio.bas` | modificada | Los tres módulos existen; el indicador se apoya en **consultas efímeras** (`CreateQueryDef` con nombre vacío, `Form_formIndicadorProyectos.cls:243`), sin consulta guardada que replicar |
| Informes | `InformeRiesgoHTML.bas`, `InformeRiesgoPDFServicio.cls`, `ExcelInforme.bas` | modificada | Los tres objetos existen; la salida es un **fichero en disco** que se abre con el visor del sistema (`Form_FormInformeTipoSalida.cls:81, 85`) |
| Histórico de estados del riesgo | `Riesgo.cls`, `RiesgosGestionRiesgo` | **pendiente de ficha** | Requisito dentro del detalle del riesgo; se confirmará en su ficha |
| Catálogos | `AreaImpacto`, `CarenciasExplicacion` (vía `CambioExplicacion`), `Juridica`, `MitigacionValor`, `NC`, `OrganoContratacion`, `Pedido`, `RAC` | sin cambios | Las ocho clases existen |
| Exportación e impresión | `ExcelInforme.bas`, `InformeRiesgoHTML.bas`, `InformeRiesgoPDFServicio.cls` | modificada | Ver «Informes»: la salida es un fichero local, no una descarga autenticada |
| Integraciones | `Constructor.bas`, `Expediente.cls`, `ExpedienteResponsable.cls`, `NC.cls` | modificada | **Se abre un backend por integración**: ocho funciones de apertura en `Variables Globales.bas:457-653` y 15 tablas vinculadas de seis backends distintos (ver [coexistencia](coexistencia.md#1-los-ocho-backends-que-abre)) |
| Acoplamiento DAO directo | `getdb` (`Variables Globales.bas:485`) | modificada | **362 llamadas en 42 ficheros de producción**, no 308 ni 372: el recuento del estudio previo y el de la épica no coinciden con la medición, y el alcance de cada una es distinto (D-GR-2) |

### Capacidad no declarada por el estudio previo

| Capacidad | Resultado | Evidencia |
|---|---|---|
| **Concurrencia sin atomicidad** | añadida (al conjunto de lo que hay que decidir) | Un único `BeginTrans` y un único `CommitTrans` en todo `src`, en `PublicabilidadEdicion.bas`, y **ningún `RollbackTrans`**; identificadores por `Max + 1` (`Funciones Generales.bas:2191, 2223`) |

No es una capacidad de negocio: es una propiedad del comportamiento que el estudio previo no recogía y que **condiciona cualquier diseño de migración**. Aparece aquí porque el [README](README.md) la cuenta entre los hallazgos y porque la decisión sobre ella no puede tomarse sin las fichas.

## Reglas de conservación

Las que el estudio previo fijó siguen en pie y esta tanda las confirma o precisa:

1. **La identidad y los permisos se resuelven contra el backend de Lanzadera**, no contra un servicio. Confirmado, y con más fuerza de la que registró el estudio previo: son 10 llamadas repartidas en cuatro ficheros y **tres tablas** de Lanzadera.
2. **`IDAplicacion` se asigna por TempVar.** Confirmado; sigue siendo configuración de sesión, no de despliegue.
3. **Los dos modelos de carga del árbol.** **Precisión:** desde `Form_frmSplash.cls:25-29` solo se alcanza «nuevo». Que «antiguo» siga vivo depende del binario y de quién lo active; queda abierto.
4. **Las worklists de calidad y el árbol de técnico** se rigen por el mismo patrón `PintarTareas`. Las clases que lo sostienen existen; el detalle está en las fichas, pendientes.
5. **El inventario incluye capacidades raras**: `InformeRiesgoPDFServicio`, el control de versiones, `Publicabilidad_Usar_Cache`, `DatosEnLocal`, `EnDesarrollo`. Confirmado que los objetos existen.

## Pendientes

- Las cuatro capacidades marcadas **pendiente de ficha** no se pueden afirmar hasta que se escriba la ficha del formulario que las sostiene.
- La clasificación «modificada» de las nueve capacidades con forma pendiente de comparar objeto a objeto con el snapshot: hasta que cada ficha contraste las líneas, la diferencia es de inventario, no de comportamiento.
- La capacidad añadida de concurrencia necesita un criterio de aceptación propio, que todavía no existe.