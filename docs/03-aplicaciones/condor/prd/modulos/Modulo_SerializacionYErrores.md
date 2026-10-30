# Módulos de serialización, presentación de errores y utilidades de desarrollo

> **Commit fijado:** `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · **Fuente:** `src/modules/JsonConverter.bas` (1447 líneas), `ErrorPresenter.bas` (186), `modDevTools.bas` (117), `modSimuladorNotificaciones.bas` (107), `ErrorLogger.bas` (53), `Pruebas.bas` (55) · **PRD:** [README](../README.md)

El último bloque de módulos de soporte: un conversor JSON propio con manejo de fechas UTC, la presentación y el registro de errores de usuario, y tres módulos que solo existen para desarrollo y pruebas —redimensionado de paneles, simulador de notificaciones y consultas de prueba. Los tres últimos no tienen equivalente en la versión web y su inventario debe reflejarlo.

Abreviaturas: `JC`, `modules/JsonConverter.bas`; `EP`, `modules/ErrorPresenter.bas`; `EL`, `modules/ErrorLogger.bas`; `DEV`, `modules/modDevTools.bas`; `SIM`, `modules/modSimuladorNotificaciones.bas`; `PRU`, `modules/Pruebas.bas`; `CE`, `classes/CondorError.cls`; `VG`, `modules/Variables Globales.bas`; `FU`, `modules/FUNCIONES UTILES.bas`.

## Superficie pública

| Objeto | Tipo | Procedimientos públicos | Qué resuelve |
|---|---|---|---|
| `JsonConverter` | módulo | `ParseJson`, `ConvertToJson`, `ParseUtc`, `ConvertToUtc`, `ParseIso`, `ConvertToIso`, `TextoParsedoParaJSON`, `TextoParsedoParaTxt`, `getJSonDeTabla`, `getJSonDeSQL`, `GenerarTextoJSON`, `SubItem` (`JC:174-1447`) | Conversión entre JSON y objetos de VBA, conversión de fechas y volcado de tablas y consultas |
| `ErrorPresenter` | módulo | `MostrarError`, `HayErrorPendiente`, `LimpiarError` (`EP:13-186`) | Decide si el último error se muestra como validación o como fallo y lo presenta |
| `ErrorLogger` | módulo | `LogToTable`, `LogFromGlobal` (`EL:11-53`) | Vuelca un `CondorError` en el log de errores |
| `modDevTools` | módulo | `AjustarPanelesBloqueo`, `Helper_PanelesBloqueo_AlFrente`, `Helper_PanelesBloqueo_AlFondo` (`DEV:9-117`) | Herramienta de diseño que reordena la etiqueta de bloqueo en todos los formularios |
| `modSimuladorNotificaciones` | módulo | `SimularTodasLasNotificaciones`, seis escenarios `Test_Notif_*`, `DispararNotificacionTest` (`SIM:6-107`) | Simulación de las notificaciones del flujo |
| `Pruebas` | módulo | `PruebaUsuariosCalidad`, `PruebaUsuariosTecnicos` (`PRU:5-55`) | Consultas de prueba de usuarios por rol |

## Procedimientos y llamadas

| Objeto | Procedimiento | Qué hace | Llama a | Datos y efectos | Evidencia |
|---|---|---|---|---|---|
| `JC` | `ParseJson`, `ConvertToJson` | Analiza y serializa JSON a `Dictionary` y `Collection` | funciones privadas `json_*` | Sin efectos | `JC:174-200`, `JC:201-462`; gramática en `JC:463-666` |
| `JC` | `ParseUtc`, `ConvertToUtc`, `ParseIso`, `ConvertToIso` | Traduce fechas entre hora local, UTC e ISO 8601 | `utc_ExecuteInShell`, `utc_DateToSystemTime` | Sin efectos | `JC:895-1047`; ejecución de proceso en `JC:1079-1109` |
| `JC` | `TextoParsedoParaJSON`, `TextoParsedoParaTxt` | Prepara texto para incrustarlo en JSON o en un archivo | — | Sin efectos | `JC:1186-1227` |
| `JC` | `getJSonDeTabla`, `getJSonDeSQL` | Serializa el contenido de una tabla o de una consulta | — | Lectura | `JC:1228-1367` |
| `EP` | `MostrarError` | Muestra el último error como advertencia o como fallo, según su origen | `EsValidacionFuncional`, `FU.ManejadorErroresFormulario` | Muestra un `MsgBox` | `EP:13-57`; nunca propaga error en `EP:14`, clasificación en `EP:27-29` |
| `EP` | `EsValidacionFuncional` | Decide si el error es de validación por número y por origen | — | Sin efectos | `EP:58-115`; número 513 en `EP:65`, comparación en `EP:67-68` |
| `EL` | `LogToTable` | Inserta el error en el log con usuario, módulo y procedimiento | `getdb()`, `OpenRecordset` | `INSERT` en `tbLogErrores` | `EL:11-47`; usuario por defecto en `EL:23`, corte de origen en `EL:24-25` |
| `EL` | `LogFromGlobal` | Registra el error pendiente global | `LogToTable` | `INSERT` en `tbLogErrores` | `EL:48-53` |
| `DEV` | `AjustarPanelesBloqueo` | Recorre los formularios del proyecto y reordena la etiqueta de bloqueo | `Application.Forms` | Modifica el diseño de los formularios | `DEV:9-107`; propósito declarado en `DEV:10-12` |
| `SIM` | `SimularTodasLasNotificaciones` y escenarios | Dispara las notificaciones del flujo con datos de ejemplo | `DispararNotificacionTest` | Envía notificaciones reales | `SIM:6-107` |
| `PRU` | `PruebaUsuarios*` | Devuelve usuarios de ejemplo por rol | — | Sin efectos | `PRU:5-55` |

## Flujos de trabajo

- **Normal:** el manejador de errores de los formularios captura la excepción, la pasa al presentador (`EP:13-57`), que decide si es validación o fallo técnico, y al registrador (`EL:11-47`), que la deja en la base. La decisión se toma por el número 513 y por el texto del origen (`EP:65-68`).
- **Alternativos:** `ErrorPresenter` nunca falla: silencia sus propios errores (`EP:14`) y, si no hay error pendiente, muestra un mensaje genérico (`EP:16-19`).
- **Errores:** `ErrorLogger` captura el error para poder recurrirse a una impresión de depuración (`EL:12`) y, si el usuario activo no existe, escribe «Sistema» (`EL:23`).

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| El error 513 de la aplicación es de validación, no técnico | `EsValidacionFuncional` | `EP:65` |
| La clasificación compara también el origen, no solo el número | comparación en mayúsculas | `EP:67-68` |
| Sin usuario activo, la traza se atribuye al sistema | `LogToTable` | `EL:23` |
| La clave de la transacción de cambios es un campo JSON | `JC:1228-1367` | — |

## Lecturas y escrituras

| Operación | Recurso | Tipo | Transacción | Evidencia |
|---|---|---|---|---|
| Registro de errores | `tbLogErrores` | `Recordset` | Ninguna | `EL:19` |
| Volcado de tablas y consultas a JSON | tablas y consultas de la base | lectura por `Recordset` | Ninguna | `JC:1228-1367` |
| Reajuste de paneles | formularios del proyecto de Access | diseño | — | `DEV:9-107` |

## Coexistencia

Cuatro de los seis módulos no tienen equivalente y deben quedar registrados como retirables: `modDevTools` manipula el diseño de los formularios en modo diseño (`DEV:9-12`), algo que no existe en la versión web; `modSimuladorNotificaciones` dispara notificaciones reales con datos de ejemplo (`SIM:6-107`); `Pruebas` devuelve usuarios de ejemplo (`PRU:5-55`); y la parte de `JsonConverter` que lanza un proceso del sistema para convertir fechas (`JC:1079-1109`) no puede trasladarse tal cual. Lo que sí se traslada es el contrato de JSON: el formato que produce `ConvertToJson` y la representación de fechas `ParseIso` y `ConvertToIso` (`JC:201-462`, `JC:953-1047`) deben reproducirse en el destino, porque los datos ya guardados se leen con este mismo formato. La presentación de errores, en cambio, se convierte en la respuesta de la API: el número 513 y su clasificación (`EP:58-115`) pasan a ser un contrato de validación, no un `MsgBox`.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual.** `JsonConverter` implementa un analizador y un serializador JSON completos, con búfer propio y soporte de UTC. `ErrorPresenter` y `ErrorLogger` separan mostrar de registrar. Los tres módulos restantes son herramientas de desarrollo y pruebas que viven en el mismo ejecutable que la aplicación.

**Posibles defectos, a verificar.**
1. `utc_ExecuteInShell` lanza un proceso del sistema para convertir fechas (`JC:1091`): depender de un proceso externo para una operación aritmética.
2. `modDevTools` modifica formularios en modo diseño y los guarda (`DEV:10-12`): si se ejecuta por accidente, altera la base de la aplicación.
3. `modSimuladorNotificaciones` envía notificaciones reales, no simuladas en un entorno de pruebas (`SIM:6-107`).
4. `ErrorPresenter` silencia cualquier error propio (`EP:14`), lo que puede ocultar un fallo del propio presentador.
5. `Pruebas` y los escenarios de simulación son código de pruebas en producción (`PRU:5-55`, `SIM:15-68`).

**Propuestas web pendientes de decisión.** Que la conversión de fechas se haga con la biblioteca estándar del destino y se documente el formato. Que la clasificación de errores sea un contrato compartido por las ocho aplicaciones del roadmap. Que las herramientas de desarrollo y los escenarios de simulación vivan fuera del entregable.

## Criterios de aceptación

1. Un JSON generado por la aplicación se lee con la misma estructura en el destino, incluidas las fechas ISO. Trazado a `JC:201-462` y `JC:953-1047`.
2. Un error de validación se presenta como advertencia y uno técnico como fallo. Trazado a `EP:27-29` y `EP:65-68`.
3. Todo error presentado queda registrado con usuario, módulo y procedimiento. Trazado a `EL:19-25`.
4. Un error sin usuario activo se atribuye al sistema. Trazado a `EL:23`.

## Pendientes

- Confirmar contra el binario si `modDevTools`, `modSimuladorNotificaciones` y `Pruebas` se despliegan con la aplicación o solo en el entorno de desarrollo.
- Confirmar qué componente reemplaza la conversión de fechas por proceso externo (`JC:1079-1109`) y si su resultado se usa en alguna decisión de negocio.
- Confirmar si el JSON que produce `getJSonDeTabla` y `getJSonDeSQL` se guarda en alguna tabla o solo se usa para exportar (`JC:1228-1367`).
- Confirmar el formato exacto de fecha que almacenan las columnas fecha de las tablas de auditoría, para no romper comparaciones al migrar.