# Coexistencia Access y web: No Conformidades

Medición estática del código de `main` (producción) en el commit `1d7a2f826a1e3c582874e4e29886a6183efb9067`, sobre `src/forms`, `src/modules` y `src/classes`. No se ha abierto ningún `.accdb` ni se ha leído el backend: el ERD de origen sigue pendiente del propietario (véase el apartado 9). Todo destino es propuesta hasta que se apruebe el ERD global. Las fichas de [formularios](README.md) aportan los apartados «Coexistencia» y «Posibles defectos» de cada pantalla.

Los recuentos son de líneas que casan con el patrón (`grep -rE` sobre las tres carpetas) y los ficheros que las contienen; no son recuentos de ejecuciones.

## 1. Vínculos y conexiones

- Todo el acceso a datos pasa por `getdb()`, que devuelve un `DAO.Database` cacheado en `m_dbCached` (`Variables Globales.bas:157-189`). Hay 309 llamadas en 40 ficheros.
- Si el objeto cacheado ya no es válido (se fuerza con `.Name`, error 3420), se descarta y se reabre (`Variables Globales.bas:161-176`).
- La apertura es `getWorkspace().OpenDatabase(...)` sobre el workspace global, con la ruta del backend activo y la contraseña en la cadena de conexión (`Variables Globales.bas:184`, `Variables Globales.bas:144-151`). `CheckAndReconnect` repite la apertura (`Variables Globales.bas:191-222`). No se documentan aquí ni la ruta ni la contraseña.
- La configuración se lee de la tabla local `TbConfiguracionBackends` con `CurrentDb.OpenRecordset` (`Variables Globales.bas:540-541`): backend activo (PROD, SANDBOX o TEST) y contraseña, volcados en `TempVars` (`Variables Globales.bas:546-553`).
- Medición: `OpenDatabase` 2 líneas en 1 fichero; `CurrentDb` 7 líneas en 2 ficheros (`Funciones Generales.bas:3187-3230` ejecuta SQL contra la base local); `DBEngine` 1 línea. `ADODB` aparece 4 veces en 2 ficheros, solo como `Stream` para ficheros (`Funciones Generales.bas:1695-1710`, `HTML.bas:72-87`): no hay ADO contra datos.
- El backend externo de riesgos se alcanza con el mismo `getdb()` (`RiesgoRepositorio.bas:91-92`); no hay un segundo vínculo.
- Implicación para la web:
  - Hay un único punto de apertura sustituible, pero la cadena de conexión es específica de DAO.
  - El vínculo a PostgreSQL (ODBC) cambiaría el comportamiento de `dbOpenDynaset`, `Edit`/`Update` y `AddNew` (24 `AddNew` en el código; ficha [FormNCProyectoAC](formularios/Form_FormNCProyectoAC.md)).
  - Las tablas locales (`TbConfiguracionBackends`) no tienen equivalente web: su contenido sería configuración de despliegue.

## 2. Adaptaciones de SQL Jet/ACE a PostgreSQL vinculado

El SQL se construye por concatenación en `constructor.bas`, `Funciones Generales.bas` y en las clases `*Operaciones.cls`.

| Patrón | Recuento | Ficheros principales |
| --- | --- | --- |
| Literal de fecha `#mm/dd/yyyy#` | 24 líneas en 10 ficheros | `IndicadorRepositorio.bas:15,56-57`; `Instalador.bas:838`; `ACProyecto.cls:562,597,631`; `ReplanificacionesProyectoOperaciones.cls:183` |
| `Like '*...*'` (comodín Jet) | 11 líneas en 3 ficheros (una es el `Like` de VBA, `Funciones Generales.bas:4065`) | `constructor.bas:3166-3169, 5845-5848, 4611`; `NCProyecto.cls:1620` |
| `DELETE *` | 16 líneas en 14 ficheros | `NCProyectoOperaciones.cls:944`; `ACProyectoOperaciones.cls:331`; `ARProyectoOperaciones.cls:351`; `Funciones Generales.bas:2939, 3185-3190` |
| `Nz(` | 1002 líneas en 92 ficheros (casi todas en VBA, pocas dentro de SQL) | `Instalador.bas:317` |
| `IIf(` | 12 líneas en 9 ficheros | `Funciones Generales.bas:1033-1035`; `CacheNCProyecto.bas:1302` |
| `Format(` | 110 líneas en 35 ficheros; `Format(Year(...),'0000')` en 10 líneas | `Variables Globales.bas:465`; `constructor.bas:4151` y consultas siguientes |
| `DCount` / `DMax` | 4 y 3 líneas, en 1 fichero cada uno | `InicializadorCache.bas:700-710, 977-985` |
| `DLookup` | 0 | |
| `Max(` para identificadores | 14 líneas en 3 ficheros | `Funciones Generales.bas:2036-2049`; `constructor.bas:1165` |
| Booleanos Sí/No como texto | `'Sí'` literal en SQL: `RiesgoRepositorio.bas:106`; propiedades `EsAdministrador = "Sí"` (`Usuario.cls:65`) | |
| Parámetros booleanos `True`/`False` escritos en SQL | `CacheNCProyecto.bas:1302` | |

Notas por patrón:

- **Fechas.** `Format(d, "mm/dd/yyyy")` rodeado de `#` no existe en PostgreSQL; hay que sustituirlo por parámetros tipados o `'yyyy-mm-dd'::date`. `RiesgoRepositorio.bas:97` y `106` insertan además la misma fecha en dos columnas.
- **Comodín.** `Like '*texto*'` pasa a `ILIKE '%texto%'`; el texto del usuario se concatena sin parametrizar (`constructor.bas:3166-3169`), lo que además es un riesgo de inyección (a verificar con la ficha [FormExpedientesBusqueda](formularios/Form_FormExpedientesBusqueda.md)).
- **`DELETE *`** no es válido en PostgreSQL: debe ser `DELETE FROM`.
- **Identificadores `Max+1`.** `DameID` lee `Max(campo)` y devuelve `+1` (`Funciones Generales.bas:2036-2049`); `constructor.getID` hace lo mismo (`constructor.bas:1165`). Dos usuarios simultáneos pueden obtener el mismo valor. En la web debe haber secuencia o identidad en servidor; mientras coexistan, ambos lados deben compartir el mismo contador (decisión pendiente, apartado 4).
- **Autonumérico.** No hay `COUNTER` ni lectura del último identificador generado en el código medido: todos los identificadores se calculan en cliente (ficha [FormNCProyectoGeneral](formularios/Form_FormNCProyectoGeneral.md)). El tipo del identificador en origen depende del ERD pendiente.
- **Concatenación de SQL.** 67 líneas con `.Execute` sin `dbFailOnError` frente a 10 con él; el escape es `Replace(texto, "'", "''")` en unos puntos (`CacheNCProyecto.bas:1301`) y la eliminación de `'` y `;` en otros (ficha [FormNCProyectoAC](formularios/Form_FormNCProyectoAC.md)).
- **Consultas guardadas.** Las definiciones de `src/queries` usan el mismo dialecto; su traducción se trata en el ERD y no aquí.

## 3. Transacciones

- `BeginTrans`: 0 líneas; `CommitTrans`: 0 líneas de DAO. Los 7 resultados de `CommitTrans` y los 7 de `Rollback` son funciones de la caché propia, no del motor: `IniciarTransaccionCache` (`CacheNCProyecto.bas:608-637`), `CommitTransaccionCache` (`CacheNCProyecto.bas:646-680`) y `RollbackTransaccionCache` (`CacheNCProyecto.bas:689-730`). El comentario de `getWorkspace` anticipa transacciones «futuras» sobre el workspace global (`Variables Globales.bas:146-147`).
- Conclusión medida: ninguna escritura de negocio está dentro de una transacción de base de datos.
- Escrituras multi-paso sin transacción (de las fichas):
  - Alta de acción correctiva: alta, log y recálculo sin límite transaccional (ficha [FormNCProyectoAC](formularios/Form_FormNCProyectoAC.md), apartado «Coexistencia»).
  - Alta de NC: identificador y código de NC calculados en cliente, `Edit`/`Update` sobre `TbNoConformidades` (ficha [FormNCProyectoGeneral](formularios/Form_FormNCProyectoGeneral.md)).
  - Vínculo NC y riesgos: `DELETE FROM TbRiesgosNC` seguido de un `INSERT` por riesgo en bucle (`RiesgoRepositorio.bas:96-106`); si falla un `INSERT`, el vínculo queda a medias.
  - Escrituras de log con `On Error Resume Next` (`CacheNCProyecto.bas:1292-1302`) que pierden el error.
- Propuesta (pendiente de aprobación): una unidad de trabajo por operación de negocio (alta o edición más log más recálculo) con una transacción del lado del servidor.

## 4. Reglas compartidas

Reglas que hoy solo aplica VBA y que la web no puede omitir mientras coexistan (todas por confirmar contra el ERD):

- Identificador y código de NC calculados en cliente (`Funciones Generales.bas:2036-2049`; `NCProyecto.cls:1620` busca por prefijo del código).
- Saneado de texto (`'` y `;`) en alta de acciones, no aplicado a las notas (`ACProyectoOperaciones.cls:133-134`, según la ficha [FormNCProyectoAC](formularios/Form_FormNCProyectoAC.md)).
- Recálculo de fechas y estado de acciones (`ACProyecto.cls:562-631`) y de máximos de fechas (`Funciones Generales.bas:3680-3787`).
- Exclusividad de los grupos de usuario: administrador anula calidad y técnico (`Variables Globales.bas:391-401`).

Restricciones compartidas propuestas, pendientes de aprobación: unicidad del código de NC, no nulidad de los campos obligatorios que hoy valida el formulario, claves foráneas de las relaciones «inferidas por código» y secuencia común para los identificadores. Ninguna está aprobada.

## 5. Concurrencia y política de conflictos

- Ninguna apertura declara bloqueo explícito: no hay `dbPessimistic` ni `LockType` en las tres carpetas; el único modo explícito es `dbOpenDynaset` (`CacheNCProyecto.bas:312`). Se aplica el bloqueo por defecto de DAO/ACE.
- Refresco: `Requery` aparece en 10 líneas de 4 ficheros (por ejemplo `Form_FormNCProyectoGestion.form.txt:6693,6750`); no hay `Refresh`. Las listas no se refrescan solas ante cambios de otro usuario.
- Pérdida de conexión: `getdb()` y `CheckAndReconnect` reabren al fallar el objeto, y devuelven el error en `p_Error` (`Variables Globales.bas:189-190`); no hay reintento de la operación que falló.
- Condición de carrera conocida: `Max+1` (apartado 2).
- Propuesta a decidir: control optimista por versión o marca de tiempo en la web, con aviso al usuario en caso de conflicto; durante la coexistencia, la regla ha de ser la misma en ambos lados. No hay decisión tomada.

## 6. Seguridad y auditoría

- Roles: `EsAdministrador`, `EsCalidad` y `EsTecnico` (`Variables Globales.bas:97-99`). Se calculan al arrancar: administrador por la propiedad del usuario o por los permisos de la aplicación (`Usuario.cls:59-77`; `UsuarioAplicacionPermisos.cls:68-76`); si lo es, los otros dos son No; después se evalúa calidad (`Variables Globales.bas:391-401`); en otro caso es técnico (`Entorno.cls:410-415`).
- Identidad: el usuario se determina por el entorno del equipo (`Variables Globales.bas:380-384`); la web necesitará su propio proveedor de identidad y una correspondencia de roles. Pendiente de decisión.
- Actor en logs: la caché registra el usuario de Windows (`CacheNCProyecto.bas:1283`) o `Sistema` (`CacheNCProyecto.bas:725`); `TbLogCache` guarda `Usuario`, `Exito`, `DuracionMs` y `FechaOperacion` (`CacheNCProyecto.bas:1296-1302`). Los logs de NC se escriben con `AddNew` (`LogNCProyecto.cls:223`, `LogNCAuditoria.cls:223`). Ninguno registra el cliente (Access o web): propuesta, añadir el origen a cada entrada de log.
- Secretos: la contraseña del backend viaja en la cadena de conexión (`Variables Globales.bas:184`). La web debe leerla de un almacén de secretos y no replicarla. Hay además nombres reales en el diseño de formularios (ficha [FormNCProyectoGeneral](formularios/Form_FormNCProyectoGeneral.md)), que no se reproducen aquí.

## 7. Pruebas bidireccionales

A ejecutar después, en un entorno aislado y con datos ficticios. Ninguna se ha ejecutado en esta medición.

- Crear en Access y leer en la web, y a la inversa, cada entidad principal (NC de proyecto, NC de auditoría, acciones, seguimiento).
- Dos altas simultáneas (una por lado) para comprobar identificadores y códigos.
- Edición simultánea de la misma NC desde ambos lados.
- Búsqueda de texto con acentos, mayúsculas y caracteres `'` y `;`.
- Fechas en el límite de día y zona horaria.
- Vínculo NC y riesgos: fallo a mitad de la inserción.
- Caída y recuperación de la conexión.
- Roles: un usuario de cada grupo en ambos lados.

## 8. Corte en la Lanzadera y dependencias

- Corte: la Lanzadera abre hoy la aplicación Access; el corte consistiría en redirigir su entrada a la web. La reversión está por definir (qué se conserva, quién la decide y en cuánto tiempo). El código de NC solo conoce la Lanzadera por la preferencia `MantenerLanzaderaAbierta` (`Usuario.cls:31`).
- Gestión de Riesgos: lee y escribe `TbRiesgosNC` y consulta `TbRiesgos` y `TbProyectosEdiciones` (`RiesgoRepositorio.bas:17-20`; `constructor.bas:2740`). Cualquier cambio de esquema afecta a ambas aplicaciones.
- Expedientes: la búsqueda de expedientes se hace desde `FormExpedientesBusqueda` (ficha [FormExpedientesBusqueda](formularios/Form_FormExpedientesBusqueda.md)), con un `SELECT *` de `TbExpedientes` filtrado en VBA (`constructor.bas:2415-2436`).
- Brass: en el código solo se encuentra una referencia de texto a `BRASS_20` en el diseño de un formulario (`Form_FormNCProyectoGestion.form.txt:685`); la relación con el catálogo de motivos queda pendiente (ficha [FormMotivosNoRequiereControlEficacia](formularios/Form_FormMotivosNoRequiereControlEficacia.md)).

## 9. Pendientes

- ERD de origen real, pendiente del propietario (extracción en su máquina Windows). Sin él, ni tipos de clave, ni nulabilidad, ni relaciones reales.
- Decisión de la política de conflictos (apartado 5) y de la secuencia de identificadores (apartado 4).
- Definición de la reversión del corte y del proveedor de identidad de la web.
- Contrastar la traducción de las consultas guardadas de `src/queries`.
- Aprobación de las restricciones compartidas propuestas.
