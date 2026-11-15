# Fase 0: No Conformidades (`NoConformidades`)

Épica [`docs/07-migracion/epica-coexistencia-postgresql.md`](../epica-coexistencia-postgresql.md), [fase 0. Inventario](../epica-coexistencia-postgresql.md#fase-0-inventario). No consta issue propia para esta ficha: la fase 0 de la épica deja «pendiente: el resto de backends». Fecha: 2026-11-04.
Base de análisis: rama `release/2026-009.rc1`, SHA `f281b4e3ad33e5411ce6e0e6f59ad388698a10b1`, worktree `00_NO_CONFORMIDADES-worktrees/2026-009.rc1`.
Fuentes: ERD del front y del backend con [`scripts/dysflow-erd/extract-erd.ps1`](../../../scripts/dysflow-erd/extract-erd.ps1) (solo lectura, sobre copia) y trazado del código con CodeGraph y `read`/`grep` sobre `src/`. CodeGraph estuvo operativo en este entorno (índice del worktree de la RC1, sin cambios pendientes); no hizo falta declararlo como fallback.

## Base de análisis

| Hecho | Evidencia |
|---|---|
| La rama base es `release/2026-009.rc1` | `docs/no_conformidades/roadmap.md:5` (mapa del ciclo de `release-2026-009`), `:64` (orden propuesto: «Crear la RC desde main con el baseline `179b56c`») |
| La UAT está pendiente, no superada | `docs/no_conformidades/roadmap.md:3` («Estado: `uat_handoff` … pendiente firma de las acceptance webs»), `:25` (NC-009-20: «Pendiente: firma UAT-31 a UAT-33»), `:77-83` (compuerta `uat_handoff`) |
| La RC1 no se ha promovido a `main` | `docs/no_conformidades/roadmap.md:83` («STOP: no promover a `main` hasta que ambas firmas estén recogidas») |
| `main` es `d834997e16ae1129fb70f12ad4181bda2e353a6b` y su `src` es otra aplicación | `main` tiene 16 módulos en `src/modules` y 41 clases; la RC1 tiene 90 módulos y 41 clases. En `main` no existe `modOrigenExpedientes.bas` ni ninguna llamada a `getdb(p_Origen:=…)` |
| El binario de la RC1 coincide con su `src` | `docs/no_conformidades/roadmap.md:25` («binario igual que `src/` (`verify_code` sin accionables)») |

Consecuencia: el inventario de abajo describe la RC1, no `main`. La diferencia no es ruido: `main` y la RC1 no abren los mismos backends (ver [Aperturas directas](#aperturas-directas)).

## Estructura del front

| Hecho | Valor |
|---|---|
| Tablas del front | 2, ambas propias: `TbConfiguracionBackends` y `TbTipologiaAux` |
| Vínculos del front | 0. El front no tiene tablas vinculadas |
| Formularios | 49 (`src/forms/*.form.txt`); 12 cargan en `Form_Timer` (`Form_FormNCProyectoGestion.cls:874` y otros) |
| Formularios con `RecordSource` | 0 |
| Controles con `ControlSource` | 0 |
| Controles con `RowSource` | 76, todos con `RowSourceType = "Value List"` y contenido literal; ninguno cita una tabla ni empieza por `SELECT` |
| Consultas guardadas del front | 3: `Consulta1`, `Consulta2`, `Consulta3`, con su SQL en `src/queries/` |

**La excepción de Brass se confirma en su forma exacta para esta aplicación: los formularios están desconectados.** No hay `RecordSource`, ni `ControlSource`, ni un solo `RowSource` que apunte a una tabla: los 76 `RowSource` del diseño son listas literales que el código sobrescribe en tiempo de ejecución con listas de cabeceras (`Form_FormExpedientesBusqueda.cls:181, 295`, `Form_FormAuditoriasGestion.cls:287`), y ninguna asignación en tiempo de ejecución nombra una tabla. Los datos los carga el código, con `Form_Timer` o al abrir el formulario (por ejemplo, `Form_FormNCProyectoGestion.cls:874-886` llama a `EstablecerDatos`).

Por eso el recuento de referencias no se apoya en el diseño: se apoya en el código, que es donde esta aplicación lee y escribe.

## Aperturas directas

| Apertura | Código | Observación |
|---|---|---|
| Backend propio | `src/modules/Variables Globales.bas:230` (`getdb`), apertura en `:254` (sandbox) y `:290` (backend activo) | La ruta y la clave salen de `TbConfiguracionBackends` del propio front (`Variables Globales.bas:683-684`), no están escritas en el código. La conexión es de lectura y escritura |
| Expedientes (`Expedientes_datos.accdb`) | `src/modules/modOrigenExpedientes.bas:99` (`OpenDatabase(m_Ruta, False, True, …)`, solo lectura), ruta y clave tomadas del `Connect` del vínculo (`:90`) | Se abre por `getdb(p_Origen:="Expedientes")` (`constructor.bas:2233, 2296, 2381, 2531, 2600, 5125`). Si no se puede abrir, el lector cae al vínculo |
| Riesgos (`Gestion_Riesgos_Datos.accdb`) | El mismo abridor (`modOrigenExpedientes.bas:99`), disparado por `getdb(p_Origen:="Riesgos")` (`constructor.bas:2976, 5843`) | Mismo contrato: solo lectura y vuelta al vínculo |
| Lanzadera (`Lanzadera_Datos.accdb`) | El mismo abridor, disparado por `getdb(p_Origen:="Lanzadera")` (`constructor.bas:407, 5126`; `modCatalogoUsuarios.bas:140`; `modOrigenExpedientes.bas:132` para el arranque) | Solo lectura. Es la apertura directa de `Lanzadera_Datos` que la RC1 añade y `main` no tiene |
| `CurrentDb()` | `Variables Globales.bas:684` (configuración del front) y `Funciones Generales.bas:3113, 3116, 3125, 3156` | Las tres últimas apuntan a `TbACParaLista` y `TbARParaLista`, tablas que no existen ni en el front ni en el backend (ver N5) |
| `DoCmd.OpenTable`, `TransferDatabase`, `DoCmd.OpenDatabase` | Ninguno | No hay apertura de tabla por la interfaz ni transferencia de datos |
| `Option Compare Database` | 178 de los 180 ficheros `.bas` y `.cls` de `src/` | La comparación no distingue mayúsculas en toda la aplicación (ver N10) |

La ruta de la conexión directa no está escrita en el código: se deduce del `Connect` del vínculo del backend y se corrige según el origen de datos del backend activo (`modOrigenExpedientes.bas:90-99`, `modOrigenDatos.bas:135-159`). En producción apunta a `\\datoste\APLICACIONES_DYS\Aplicaciones PpD\No Conformidades` y a las rutas de los vínculos; en pruebas, al mismo nombre de fichero dentro de la carpeta de pruebas (`modOrigenDatos.bas:17-18`). Ninguna credencial aparece en el código: la del backend se lee del campo `PasswordBackend` de la tabla de configuración del front.

**El caso de `Lanzadera_Datos` justifica la base de análisis.** En `main` esta aplicación no abre `Lanzadera_Datos`: la búsqueda de `OpenDatabase`, `DBEngine` y `Workspaces` en `src/` de `main` solo devuelve el `getdb()` de su propio backend (`src/modules/Variables Globales.bas:184, 221`), y las lecturas de usuarios van por los vínculos. La RC1 añade una apertura directa de solo lectura (`modOrigenExpedientes.bas:99`) con la ruta y la clave tomadas del vínculo (`:90`). Un inventario hecho sobre `main` habría descrito un mecanismo de acceso que la rama que va a producción ya no tiene. El detalle por front está en [aperturas directas de `Lanzadera_Datos`](aperturas-directas-lanzadera.md#no-conformidades); las dos alternativas de adaptación ya están planteadas allí y no se repiten aquí.

## Tablas de otros backends que consume el front

El front no tiene vínculos: las tablas de otras aplicaciones están en su backend y las lee a través de él o por conexión directa.

| Tabla | Backend de origen | Cómo la usa el front |
|---|---|---|
| `TbExpedientes` | `Expedientes_datos.accdb` | Lectura por conexión directa (`constructor.bas:2296, 2381, 2600`) y por vínculo |
| `TbExpedientesJuridicas`, `TbJuridicas` | `Expedientes_datos.accdb` | Lectura por conexión directa (`constructor.bas:2233`) |
| `TbExpedientesResponsables` | `Expedientes_datos.accdb` | Lectura por conexión directa (`constructor.bas:2531, 5125`) |
| `TbProyectos`, `TbProyectosEdiciones` | `Gestion_Riesgos_Datos.accdb` | Lectura por conexión directa (`constructor.bas:2976`) |
| `TbRiesgos` | `Gestion_Riesgos_Datos.accdb` | Lectura por conexión directa (`constructor.bas:5843`) y por vínculo (`CacheNCProyecto.bas:995-999`) |
| `TbRiesgosNC` | `Gestion_Riesgos_Datos.accdb` | **Escritura** por vínculo (`RiesgoRepositorio.bas:96-105`, `NCProyectoOperaciones.cls:1554-1559`); lectura por vínculo (`CacheNCProyecto.bas:995-999`) y por conexión directa (`constructor.bas:5843-5852`) |
| `TbUsuariosAplicaciones` | `Lanzadera_Datos.accdb` | Lectura por conexión directa (`constructor.bas:407, 5126`, `modCatalogoUsuarios.bas:140`) |
| `TbUsuariosAplicacionesPermisos` | `Lanzadera_Datos.accdb` | Lectura por conexión directa (`constructor.bas:407`) |
| `TbCorreosEnviados` | `Correos_datos.accdb` | **Escritura** por vínculo (`Correo.cls:120-136`, `Funciones Generales.bas:1322-1338`) |

Dos consecuencias para la [regla R-2](../epica-coexistencia-postgresql.md#reglas-del-legacy):

- `TbRiesgosNC` y `TbCorreosEnviados` las escribe esta aplicación y viven en el backend de otra. Con la regla «cada tabla vive en el backend de la aplicación que la usa», el dueño real hay que registrarlo; el dueño no se puede cerrar aquí sin las fichas de Gestión de Riesgos y de Correos.
- La tabla de correos que escribe esta aplicación se rellena con datos personales, y además hay destinatarios literales en el código: la copia oculta por defecto (`Correo.cls:116-118`) y la lista de destinatarios del aviso de error (`Funciones Generales.bas:1319`). Quien reciba el trasvase a PostgreSQL hereda ese dato.

## Backend `NoConformidades_Datos`

Copia de desarrollo: `C:\00repos\datos\NoConformidades_Datos.accdb`. Los vínculos apuntan a las copias locales de los otros backends, no a `\\datoste\…`; las rutas de producción no se pudieron abrir en este entorno.

| Hecho | Valor |
|---|---|
| Tablas | 47: 36 propias y 11 vinculadas |
| Destino de los 11 vínculos | 4 a `Expedientes_datos.accdb`, 4 a `Gestion_Riesgos_Datos.accdb`, 2 a `Lanzadera_Datos.accdb`, 1 a `Correos_datos.accdb` |
| Relaciones | 14, de las que 13 son de la aplicación y una es de sistema (`MSysNavPaneGroupCategories` → `MSysNavPaneGroups`) |
| Consultas guardadas | 4: `ColumnaFalta`, `Consulta1`, `NC_Estudio_Replanificacion_2026`, `Replanificaciones_2026_Max_Fecha_Fin_Prevista`. Ninguna se referencia desde `src/` |
| Tablas sin clave primaria | 3: `TbCacheListadoNC`, `TbCorreosEnviados`, `TbNoConformidadesIngresoPorLotesDetalle` |
| Campos multivalor, adjuntos, OLE o hipervínculo | 0. Los 578 campos del backend tienen `isComplex` falso y ninguno es de tipo OLE |
| Reglas de validación | 0, ni de tabla ni de campo |
| Tipos con equivalente directo en PostgreSQL | 7 campos `Integer` de 16 bits y 2 `Currency` (`TbExpedientes.ImporteLicitacion`, `ImporteContratacion`), los tres primeros en tablas de otros backends |
| Valores por defecto | 31, todos elementales (`now()`, `date()`, `0`, `Yes`, `No`) |

Sobre el atributo de campo: el ERD marca con el bit 16 dos claves primarias `Long`, `TbLogCache.IDLog` y `TbCacheIndicadoresProyectoDetalle.IDCacheDetalle`. La inserción del registro de caché omite `IDLog` (`CacheNCProyecto.bas:1305-1310`), de modo que son claves autogeneradas y no campos multivalor; además, ningún campo del backend es complejo. La equivalencia exacta de ese bit con las constantes de DAO no se comprobó contra la biblioteca.

## Veredicto de propiedad

Clasificación según la [épica](../epica-coexistencia-postgresql.md#clasificación-de-tablas). La columna de referencias es el recuento por nombre exacto de tabla en `src/` de la RC1 (`.bas`, `.cls`, `.form.txt`), separando código de producción de código de test.

**Se quedan en el backend y las usa el front**: 19 tablas propias. `TbNoConformidades` (194 referencias), `TbNoConformidadesAuditoria` (105), `TbNCAccionCorrectivas` (123), `TbNCAccionesRealizadas` (121), `TbNCAuditoriaAccionesRealizadas` (77), `TbNCAuditoriaAccionCorrectivas` (73), `TbCacheNCProyecto` (25), `TbDocumentosAuditorias` (18), `TbNCDocumentos` (16), `TbReplanificacionesProyecto` (16), `TbAuditorias` (33), `TbTiposNCProyectos` (11), `TbLogCache` (7), `TbLog` (6), `TbLogAuditoria` (5), `TbReplanificacionesAuditoria` (6), `TbTipologia` (4), `TbAuxPuntoNorma` (2) y `TbHerramientaDocAyuda` (1). Los anexos que la aplicación usa hoy no están en `TbAnexos`, sino en `TbNCDocumentos` (proyecto) y `TbDocumentosAuditorias` (auditoría).

**Solo las lee el instalador**: `TbAnexosAuditoria` (`Instalador.bas:126`) y `TbAnexosNCAuditorias` (`Instalador.bas:170`). `Instalador.bas` es un módulo de migración de datos que rellena `TbDocumentosAuditorias` a partir de ellas. No son huérfanas mientras la migración no se haya ejecutado en producción, pero su único consumidor es un procedimiento de una sola vez.

**No las usa el front**: 15 tablas propias, con cero referencias en `src/` de la RC1 y sin ninguna forma de acceso por diseño (no hay `RecordSource`, `ControlSource` ni `RowSource` que las nombre):

| Tabla | Filas | Observación |
|---|---|---|
| `TbAnexos` | 228 | La tabla de anexos anterior a `TbNCDocumentos`; la relación `TbNoConformidadesTbAnexos` sigue declarada |
| `TbNCARAvisos` | 432 | Avisos por acción correctiva, con referencias a ids de correo |
| `TbNoConformidadesIngresoPorLotesTemporal` | 166 | Entrada masiva por Excel |
| `TbNoConformidadesIngresoPorLotesDetalle` | 166 | Entrada masiva por Excel, sin clave primaria |
| `TbNoConformidadesIngresoPorLotesPrincipal` | 1 | Entrada masiva por Excel |
| `TbTareasExplicaciones` | 10 | Explicación de tareas |
| `TbConsultasPorFechas` | 8 | Catálogo de descripciones |
| `TbConfiguracion` | 1 | Interruptor de la caché, que NC-009-11 dejó apagada |
| `TbNCInformacionRAC` | 2 | Enlazada con `TbNoConformidades` |
| `TbNCInformacionRACAnexos` | 1 | Enlazada con `TbNCInformacionRAC` |
| `TbAuditoriaLog` | 0 | Tabla vacía |
| `TbCacheIndicadoresProyectoHeader` | 0 | Tabla vacía |
| `TbCacheIndicadoresProyectoDetalle` | 0 | Tabla vacía |
| `TbCacheListadoNC` | 0 | Tabla vacía, sin clave primaria |
| `TbConexiones` | 0 | Tabla vacía |

La función de entrada masiva por Excel (las tres tablas `IngresoPorLotes`) y la de anexos RAC no tienen código en la RC1: no hay ninguna referencia a `TbNoConformidadesIngresoPorLotes*`, `TbTareasExplicaciones` ni `TbNCInformacionRAC` en `src/`, ni por nombre completo ni por SQL dinámico (no hay construcción de nombres de tabla a partir de cadenas).

**Del front**: `TbConfiguracionBackends` es la tabla de arranque (ruta del backend activo, clave y parámetros de entorno) y la lee el propio front con `CurrentDb` (`Variables Globales.bas:683-684`); las otras siete apariciones de su nombre en `src/` son comentarios o mensajes de error. `TbTipologiaAux` (23 filas) no la usa nadie en la RC1.

**Tablas de otros backends**: las 11 vinculadas están bien ubicadas respecto de sus dueños, salvo las dos que esta aplicación escribe (`TbRiesgosNC` y `TbCorreosEnviados`), cuyo dueño real hay que registrar con las fichas de los otros aplicativos.

**Referencias a tablas que no existen**: `TbACParaLista` y `TbARParaLista` (`Funciones Generales.bas:3112-3116, 3125, 3156`) no están ni en el front ni en el backend; solo las alcanza `RellenarTbProyectoParaLista` (`Funciones Generales.bas:3096`), que no tiene llamadores. La consulta guardada `Consulta3` del front (`src/queries/Consulta3.sql`) cita `TbACParaLista`, y `Consulta1` cita `TbNCDocumentosAux`, que tampoco existe.

## Propuestas para después del descubrimiento global

Ninguna se aplica: la puerta de descubrimiento global sigue cerrada (#1576).

| Id | Propuesta | Grupo (fase 1) | Comprobación previa | Reversión |
|---|---|---|---|---|
| N1 | Poner en cuarentena las 15 tablas propias sin referencias del backend, renombrándolas a `zzBAJA_<tabla>`. Para `TbNCInformacionRAC` y `TbNCInformacionRACAnexos` hay que quitar antes las relaciones `TbNoConformidadesTbNCInformacionRAC` y `TbNCInformacionRACTbNCInformacionRACAnexos`. | Imprescindible | Confirmar que ningún otro front ni la web las leen (grafo de vínculos de todos los backends) y que no hay consumidores fuera de Access: informes, Excel o tareas programadas. Antes, copia de seguridad validada y recuento de filas en producción, porque 10 de las 15 tienen datos. | Quitar el prefijo y volver a declarar las relaciones. |
| N2 | Revisar `TbAnexosAuditoria` y `TbAnexosNCAuditorias`: su único lector es `Instalador.bas:126` y `:170`, un módulo de migración de una sola vez. Propuesta: ejecutarla en producción y retirarlas después. | Imprescindible | Comprobar que la copia de `TbDocumentosAuditorias` de producción tiene los anexos migrados: las filas de `TbDocumentosAuditorias` deben coincidir con las de `TbAnexosAuditoria` unidas por `IDAuditoria` y nombre de archivo. | Reponer las tablas desde la copia. |
| N3 | Declarar clave primaria en `TbCorreosEnviados`, la única tabla viva sin clave que escribe esta aplicación. Sin clave no es actualizable por ODBC. | Imprescindible | La tabla vive en `Correos_datos.accdb`: acordar el cambio con su dueño. Verificar duplicados: `SELECT IDCorreo FROM TbCorreosEnviados GROUP BY IDCorreo HAVING COUNT(*) > 1` debe devolver 0 filas. | Quitar la clave. |
| N4 | Sustituir el generador `MAX + 1` (`constructor.getID`, `src/modules/constructor.bas:1196-1209`) por una identidad. Afecta a 18 llamadas sobre 16 tablas: `TbNCAccionCorrectivas`, `TbNCAccionesRealizadas`, `TbNCAuditoriaAccionCorrectivas`, `TbNCAuditoriaAccionesRealizadas`, `TbAuditorias`, `TbDocumentosAuditorias`, `TbNCDocumentos`, `TbLog`, `TbLogAuditoria`, `TbNoConformidades`, `TbNoConformidadesAuditoria`, `TbReplanificacionesProyecto`, `TbReplanificacionesAuditoria`, `TbTiposNCProyectos`, `TbCorreosEnviados` y `TbRiesgosNC`. | Imprescindible | La lectura de `MAX` no reserva el identificador. Probar dos inserciones simultáneas sin colisiones y medir el hueco de ids por tabla: si el hueco es grande, el cambio de tipo de columna no es trivial. | Volver al generador actual. |
| N5 | Revisar las referencias a tablas que no existen: `TbACParaLista` y `TbARParaLista` (`Funciones Generales.bas:3112-3116, 3125, 3156`) y las consultas guardadas `Consulta1` y `Consulta3` del front (`src/queries/`). Propuesta: retirar el procedimiento `RellenarTbProyectoParaLista` (`:3096`), sin llamadores, y las dos consultas, o recuperar las tablas. | Imprescindible | Reproducir el recorrido afectado en runtime: la ejecución de `RellenarTbProyectoParaLista` falla hoy porque `TbACParaLista` no existe, y las dos consultas guardadas fallan al ejecutarlas. | Restaurar el procedimiento y las consultas. |
| N6 | Retirar `NCProyectoOperaciones.ActualizarRiesgosNC` (`src/classes/NCProyectoOperaciones.cls:1545-1568`): escribe `IDNoConformidad` en `TbRiesgosNC`, columna que no existe en el backend de Riesgos (sus campos son `ID`, `IDRiesgo`, `IDNC`, `FechaRegistro`, `ParaNC`, `FechaDecison`), y no tiene llamadores. El camino vivo es `RiesgoServicio.ActualizarRiesgosNC` → `RiesgoRepositorio.GuardarAsociaciones` (`RiesgoRepositorio.bas:96-105`), que sí usa `IDNC`. | Imprescindible | Confirmar que la asociación de riesgos a una NC se guarda bien por el camino vivo: abrir una NC con riesgos asociados, guardarlos y comprobar las filas de `TbRiesgosNC`. | Restaurar el procedimiento corregido. |
| N7 | Poner en cuarentena `TbTipologiaAux` del front (23 filas, cero referencias). Su equivalente en el backend es `TbTipologia` (23 filas), que sí usa el instalador (`Instalador.bas:667, 727-732`). | Imprescindible | Comparar el contenido de las dos tablas (código y descripción de cada tipología) para confirmar que `TbTipologiaAux` es la copia antigua antes de retirarla. | Reponer la tabla desde el binario del front. |
| N8 | Registrar el dueño real de `TbRiesgosNC` (backend de Gestión de Riesgos) y de `TbCorreosEnviados` (backend de Correos): las escribe esta aplicación. En la migración, la tabla se crea en el esquema PostgreSQL de su dueño. No se mueve nada en el legacy. | Imprescindible | Cruzar con las fichas de Gestión de Riesgos y de Correos: si el backend de origen también las usa, son propias y compartidas; si no, quedan mal ubicadas y el dueño es esta aplicación. | Sin cambios en el legacy: es un registro de propiedad. |
| N9 | Revisar las 4 consultas guardadas del backend (`ColumnaFalta`, `Consulta1`, `NC_Estudio_Replanificacion_2026`, `Replanificaciones_2026_Max_Fecha_Fin_Prevista`): ninguna se referencia desde `src/`. Propuesta: retirarlas o documentarlas como consultas de análisis. | Recomendable si es barato | Ningún front las referencia (0 apariciones en `src/` de la RC1 y de `main`). Antes hay que leer su SQL, que el ERD no exporta. | Recrearlas desde este inventario. |
| N10 | Decidir cómo se reproduce la comparación sin distinguir mayúsculas: 178 de los 180 módulos y clases de `src/` llevan `Option Compare Database`, y el SQL construido lo hereda. | Aplazar a la refactorización final | Medir en PostgreSQL, con los datos reales, cuántas consultas devuelven un resultado distinto por la diferencia de mayúsculas y minúsculas en la comparación de texto. | Sin cambios en el legacy: es una decisión de la fase 2. |
| N11 | Decidir dónde vive `TbConfiguracionBackends` (tabla de arranque del front, con la ruta y la clave del backend en claro). Por ahora no puede pasar al backend sin romper el arranque: es la tabla que dice dónde está el backend. | Aplazar a la refactorización final | Medir el coste: una tabla con una fila y un lector (`Variables Globales.bas:683-684`). La alternativa es un fichero de configuración fuera del binario, con su propio mecanismo de credenciales. | Sin cambios en el legacy: es una decisión de la fase 2. |

Lo que **no** se toca ahora: los 11 vínculos del backend, las 19 tablas propias vivas, las 2 tablas de identidad de la Lanzadera (compartidas con 7 backends), las 13 relaciones de aplicación del backend y las 3 conexiones directas de solo lectura a otros backends, cuyo mecanismo de adaptación ya está analizado en [aperturas directas de `Lanzadera_Datos`](aperturas-directas-lanzadera.md#propuesta-de-adaptación).

## Ficha por tabla

Referencias: número de apariciones del nombre exacto de la tabla en `src/` de la RC1. «Prod» es código de producción; «Test» es código de prueba (`Test_*.bas`, `Bench_*.bas` y `TestHelper.bas`).

### Tablas propias del backend

| Tabla | Clave primaria | Filas | Refs. Prod | Refs. Test | Clasificación |
|---|---|---|---|---|---|
| `TbAnexos` | IDAnexo | 228 | 0 | 0 | candidata a huérfana (N1) |
| `TbAnexosAuditoria` | IDAnexo | 8 | 1 | 0 | propia, la lee solo el instalador (N2) |
| `TbAnexosNCAuditorias` | IDAnexo | 0 | 1 | 0 | propia, la lee solo el instalador (N2) |
| `TbAuditoriaLog` | IDLog | 0 | 0 | 0 | candidata a huérfana (N1) |
| `TbAuditorias` | IDAuditoria | 8 | 33 | 10 | propia |
| `TbAuxPuntoNorma` | PuntoNorma | 28 | 2 | 2 | propia |
| `TbCacheIndicadoresProyectoDetalle` | IDCacheDetalle | 0 | 0 | 0 | candidata a huérfana (N1) |
| `TbCacheIndicadoresProyectoHeader` | IDCacheIndicadorProyecto | 0 | 0 | 0 | candidata a huérfana (N1) |
| `TbCacheListadoNC` | — | 0 | 0 | 0 | candidata a huérfana (N1) |
| `TbCacheNCProyecto` | IDNoConformidad | 0 | 25 | 1 | propia |
| `TbConexiones` | Usuario | 0 | 0 | 0 | candidata a huérfana (N1) |
| `TbConfiguracion` | ID | 1 | 0 | 0 | candidata a huérfana (N1) |
| `TbConsultasPorFechas` | Descripcion | 8 | 0 | 0 | candidata a huérfana (N1) |
| `TbDocumentosAuditorias` | IDDocumento | 281 | 18 | 0 | propia |
| `TbHerramientaDocAyuda` | NombreFormulario, NombreArchivoAyuda | 10 | 1 | 0 | propia |
| `TbLog` | IDLog | 4745 | 6 | 17 | propia |
| `TbLogAuditoria` | IDLog | 689 | 5 | 11 | propia |
| `TbLogCache` | IDLog | 5542 | 7 | 7 | propia |
| `TbNCARAvisos` | ID | 432 | 0 | 0 | candidata a huérfana (N1) |
| `TbNCAccionCorrectivas` | IDAccionCorrectiva | 656 | 123 | 47 | propia |
| `TbNCAccionesRealizadas` | IDAccionRealizada | 1684 | 121 | 42 | propia |
| `TbNCAuditoriaAccionCorrectivas` | IDAccionCorrectiva | 88 | 73 | 33 | propia |
| `TbNCAuditoriaAccionesRealizadas` | IDAccionRealizada | 115 | 77 | 27 | propia |
| `TbNCDocumentos` | IDDocumento | 1123 | 16 | 0 | propia |
| `TbNCInformacionRAC` | IDInformacionRAC | 2 | 0 | 0 | candidata a huérfana (N1) |
| `TbNCInformacionRACAnexos` | IDAnexoInformacionRAC | 1 | 0 | 0 | candidata a huérfana (N1) |
| `TbNoConformidades` | IDNoConformidad | 473 | 194 | 104 | propia |
| `TbNoConformidadesAuditoria` | ID | 77 | 105 | 62 | propia |
| `TbNoConformidadesIngresoPorLotesDetalle` | — | 166 | 0 | 0 | candidata a huérfana (N1) |
| `TbNoConformidadesIngresoPorLotesPrincipal` | IDLoteExcel | 1 | 0 | 0 | candidata a huérfana (N1) |
| `TbNoConformidadesIngresoPorLotesTemporal` | IDTmp | 166 | 0 | 0 | candidata a huérfana (N1) |
| `TbReplanificacionesAuditoria` | IDReplanificacion | 13 | 6 | 0 | propia |
| `TbReplanificacionesProyecto` | IDReplanificacion | 179 | 16 | 18 | propia |
| `TbTareasExplicaciones` | NodoTarea | 10 | 0 | 0 | candidata a huérfana (N1) |
| `TbTipologia` | CodTipologia | 23 | 4 | 0 | propia |
| `TbTiposNCProyectos` | IDTipo | 13 | 11 | 0 | propia |

### Tablas vinculadas del backend

| Tabla | Clave primaria | Refs. Prod | Refs. Test | Backend de origen | Clasificación |
|---|---|---|---|---|---|
| `TbCorreosEnviados` | — | 4 | 0 | `Correos_datos.accdb` | la escribe esta aplicación; dueño por registrar (N3, N8) |
| `TbExpedientes` | IDExpediente | 38 | 15 | `Expedientes_datos.accdb` | propia y compartida de Expedientes |
| `TbExpedientesJuridicas` | IDExpedienteJuridica | 4 | 0 | `Expedientes_datos.accdb` | propia y compartida de Expedientes |
| `TbExpedientesResponsables` | IDExpedienteResponsable | 12 | 0 | `Expedientes_datos.accdb` | propia y compartida de Expedientes |
| `TbJuridicas` | IDJuridica | 5 | 0 | `Expedientes_datos.accdb` | propia y compartida de Expedientes |
| `TbProyectos` | IDProyecto | 5 | 2 | `Gestion_Riesgos_Datos.accdb` | propia y compartida de Gestión de Riesgos |
| `TbProyectosEdiciones` | IDEdicion | 7 | 1 | `Gestion_Riesgos_Datos.accdb` | propia y compartida de Gestión de Riesgos |
| `TbRiesgos` | IDRiesgo | 17 | 4 | `Gestion_Riesgos_Datos.accdb` | propia y compartida de Gestión de Riesgos |
| `TbRiesgosNC` | ID | 18 | 2 | `Gestion_Riesgos_Datos.accdb` | la escribe esta aplicación; dueño por registrar (N8) |
| `TbUsuariosAplicaciones` | CorreoUsuario | 52 | 15 | `Lanzadera_Datos.accdb` | propia y compartida de la Lanzadera (identidad) |
| `TbUsuariosAplicacionesPermisos` | CorreoUsuario, IDAplicacion | 9 | 0 | `Lanzadera_Datos.accdb` | propia y compartida de la Lanzadera (identidad) |

### Tablas del front

| Tabla | Clave primaria | Filas | Refs. Prod | Refs. Test | Clasificación |
|---|---|---|---|---|---|
| `TbConfiguracionBackends` | ID | 1 | 8 (1 lectura; el resto, comentarios o mensajes) | 44 | tabla de arranque del front; destino por decidir (N11) |
| `TbTipologiaAux` | CodTipologia | 23 | 0 | 0 | candidata a huérfana en el front (N7) |

## Lo que no se pudo verificar

1. El SQL de las 4 consultas guardadas del backend: el ERD exporta nombre y tipo, no la sentencia. El SQL de las 3 del front sí se conoce, por `src/queries/`.
2. Si las 15 tablas sin referencias las consume otra aplicación: exige el grafo de vínculos de todos los backends, la tarea global de la fase 0.
3. Si esas mismas tablas tienen consumidores fuera de Access: informes en Excel, tareas programadas o consultas externas. El [protocolo de huérfanas](../epica-coexistencia-postgresql.md#protocolo-de-huérfanas) lo exige antes de borrar nada.
4. El dueño real de `TbRiesgosNC` y `TbCorreosEnviados`: requiere las fichas de Gestión de Riesgos y de Correos.
5. La paridad entre `src/` y el binario se apoya en el `verify_code` que registra el roadmap (`docs/no_conformidades/roadmap.md:25`), no en una comprobación hecha en esta sesión. Las expresiones de propiedad de formulario que no se exportan a `src` quedan sin revisar.
6. La ausencia de macros y de `AutoExec` es la regla del propietario, no una comprobación de esta sesión: no se inspeccionó el proyecto VBA del binario más allá del `src` exportado.
7. Las rutas de producción: los vínculos del backend apuntan a las copias locales de desarrollo (`C:\00repos\datos\…`) y no se pudieron abrir las rutas de `\\datoste\…`. Los recuentos de filas son de desarrollo.
8. El estado de UAT en el tracker: la evidencia es el roadmap del repositorio, no la issue de UAT.
9. La equivalencia entre el bit de atributo 16 del ERD y las constantes de DAO: los dos campos marcados son claves `Long` autogeneradas, no campos multivalor, pero el mapeo del bit no se contrastó con la biblioteca.
