# Coexistencia — Gestión de Riesgos

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Issue:** #1079 · **Fase:** 4 · **PRD:** [README](README.md)

Cómo convive hoy esta aplicación con el resto y con la web que se va a construir, medido sobre el código del commit fijado. Todo lo que aparece aquí es **medición estática**: no se ha ejecutado la aplicación ni se ha abierto el binario.

## 1. Los ocho backends que abre

`src/modules/Variables Globales.bas` es el punto único de apertura. Hay **ocho** funciones, no una:

| Función | Línea | Backend que abre | Llamadas en producción | Ficheros que la usan |
|---|---|---|---|---|
| `getdb` | `:485` | `Gestion_Riesgos_Datos.accdb` (el propio) | **362** | 43 |
| `getdbExpedientes` | `:518` | backend de Expedientes | 16 | `Constructor.bas`, `Variables Globales.bas` |
| `getdbNC` | `:539` | backend de No Conformidades | 12 | `Constructor.bas`, `NC.cls`, `Form_FormRiesgoNC.cls` |
| `getdbLanzadera` | `:457` | backend de Lanzadera (identidad) | 10 | `Constructor.bas`, `Usuario.cls`, `UsuarioAplicacionPermisos.cls` |
| `getdbControlCambios` | `:653` | backend de control de cambios | 7 | `Constructor.bas` |
| `getdbAGEDO` | `:566` | backend AGEDO 2.0 | 5 | `Constructor.bas`, `Documento.cls` |
| `getdbCorreo` | `:624` | base de correo | 3 | `Correo.cls` |
| `getdbAgedys` | `:594` | backend AGEDYS | 2 | `Pedido.cls` |
| | | **Total** | **417** | |

El nombre del backend propio se asigna en un único sitio, `Variables Globales.bas:490` (`m_NombreDatos = "Gestion_Riesgos_Datos.accdb"`); **no hay ruta en el repositorio**, se resuelve en ejecución. Para cualquier extracción hace falta la ruta que usa el propietario.

**Consecuencia para la migración:** la aplicación no es una unidad desplegable. Siete de sus ocho accesos son a datos de otras aplicaciones. Una migración que quiera hacer el corte por aplicación se lleva por delante identidad, expedientes, no conformidades, control de cambios, AGEDO, AGEDYS y correo.

### Tablas vinculadas

El ERD del propio legacy (`ERD/Estructura_Datos.md`) declara **15 tablas vinculadas**, que no son de esta aplicación:

| Backend de origen | Tablas |
|---|---|
| `Lanzadera_Datos.accdb` | `TbAplicaciones`, `TbUsuariosAplicaciones`, `TbUsuariosAplicacionesPermisos` |
| `Expedientes_datos.accdb` | `TbExpedientes1`, `TbExpedientesResponsables`, `TbExpedientesSuministradores`, `TbSuministradores` |
| `AGEDYS_DATOS.accdb` | `TbDPD`, `TbExpedientes`, `TbResponsablesExpedientes`, `TbSuministradoresSAP` |
| `AGEDO20_Datos.accdb` | `TbDocumentos`, `TbDocumentosID` |
| `NoConformidades_Datos.accdb` | `TbNoConformidades` |
| `Gestion_Riesgos_Datos.accdb` | `TbProyectos1` — un vínculo a sí mismo, que es **D-GR-5** |

Las tres primeras de Lanzadera son la identidad y los permisos, y **las usan también otras aplicaciones**: el PRD de Lanzadera las señala como propiedad suya. Cualquier cambio de estructura o de claves de esas tres tablas afecta a las ocho aplicaciones del roadmap y debe coordinarse en el ERD global.

## 2. Adaptaciones de Jet/ACE

Lo que el código da por hecho del motor y PostgreSQL no da:

| Construcción | Medición | En PostgreSQL |
|---|---|---|
| `OpenRecordset` | 283 llamadas | Sustituir por el acceso parametrizado del ORM |
| `Nz(` | 1 321 llamadas | `COALESCE` |
| `Err.Raise` | 3 752 llamadas | Excepciones tipadas; el código propaga por cadena |
| `LIKE '*texto*'` | 5 llamadas | `ILIKE '%texto%'` |
| `dbFailOnError` | 4 llamadas | Sin equivalente directo; hay que decidir si el error aborta la transacción |
| `Max + 1` para identificadores | 8 llamadas | Ver § 4 |
| `fso.FileExists` | 74 llamadas | El fichero vive fuera de la base; hay que decidir si pasa a almacenamiento de objetos |
| `TempVars` | 126 llamadas | Estado de sesión, no tabla |
| `DoCmd.OpenForm` | 128 llamadas | Navegación de la web |

Los cinco `LIKE '*…*'` son además un problema de corrección, no solo de dialecto: el texto se concatena en el SQL sin escapar. Queda como hallazgo de ficha, no como conclusión de este documento.

## 3. Reglas compartidas

- **El frontend no tiene tablas propias.** Es la regla del propietario y aquí se cumple en lo medido: las 68 tablas del ERD propio son del backend. Las excepciones son las de caché (§ 6) y las tablas temporales `~TMP*` que Access crea al ejecutar consultas de unión.
- **Las consultas guardadas son cero.** No hay carpeta `src/queries/` ni `queries.json`. Las tres llamadas a `CreateQueryDef` usan **nombre vacío**: son consultas efímeras, no guardadas. Todo el SQL de la aplicación está embebido en el código, lo que significa que **no hay un punto único donde revisar las consultas**: hay que buscarlas en los 195 ficheros de producción.
- **`getdb()` devuelve una base abierta por ruta**, no una conexión gestionada. La consecuencia está medida en § 4: una sola transacción en toda la aplicación y ninguna reversa.

## 4. Concurrencia: el punto más frágil

**No hay transacciones.** En todo `src` hay **un solo** `BeginTrans`, **un solo** `CommitTrans` y **ningún** `RollbackTrans`:

| Fichero | Transacciones |
|---|---|
| `src/modules/PublicabilidadEdicion.bas` | 1 `BeginTrans`, 1 `CommitTrans` |

Un único par, sin reversa, en una aplicación que escribe en 85 tablas. El resto de las escrituras va acia Access sin atomicidad: o se aplica todo o queda a medias según dónde se corte. En una migración esto no es un detalle heredable, es un requisito de la web.

**Los identificadores se calculan con `Max + 1`.** `DameID` (`src/modules/Funciones Generales.bas:2191`, cálculo en `:2223`) lee el máximo de la columna y suma uno, sobre el nombre de tabla que le pasan:

```
Funciones Generales.bas:3949   m_ID = DameID("TbProyectosEdicionesSuministradores", "ID", getdb(), p_Error)
Funciones Generales.bas:4033   m_ID = DameID("TbProyectosEdiciones", "IDEdicion", , p_Error)
Funciones Generales.bas:4088    m_ID = DameID("TbRiesgosPlanMitigacionPpal", "IDMitigacion", , p_Error)
```

Con dos usuarios guardando a la vez, ambos pueden leer el mismo máximo y escribir el mismo identificador. No hay restricción de unicidad que lo impida mientras el ERD de origen no confirme las claves (D-GR-4), así que **la conclusión queda condicionada**: es un riesgo a verificar contra el binario, no un defecto probado.

**Los errores se silencian mucho.** Hay **263** apariciones de `On Error Resume Next` en ficheros de producción, frente a 16 de `On Error GoTo 0`. La proporción indica que la mayoría de los fallos se propagan como «sigo como si nada». Consecuencia práctica: **el comportamiento de error real de la aplicación es en buena parte desconocido a partir del código**, y eso hay que resolverlo en UAT antes de fijar criterios de aceptación web.

## 5. Seguridad

- **Hay credenciales en el código.** `src/classes/Usuario.cls` contiene literales de tipo contraseña o clave. No se reproducen aquí. Es un hallazgo de ficha y, por su naturaleza, uno que se comunica al propietario y **no se publica en el repositorio del legacy**.
- **No hay autenticación federada ni LDAP.** No aparece ninguna llamada a LDAP, Active Directory ni `GetObject`. La identidad se resuelve leyendo las tablas de Lanzadera (`getdbLanzadera`, 10 llamadas), es decir, **confiando en el usuario de Windows y en las tablas de la otra aplicación**. La web debe sustituirlo por el adaptador de autenticación unificado; es la misma decisión que el PRD de Lanzadera deja abierta (P-LAN-3).
- **El correo no usa SMTP.** `Correo.cls` lee y escribe en una base de datos (`getdbCorreo`, 3 llamadas), no en un servidor de correo. No hay `CDO.`, `MAPI.` ni `SMTP` en el código. El equivalente en la web es una cola de salida, como la `mail_outbox` que Lanzadera ya tiene; conviene decidirlo antes, no durante el corte.
- **Rutas de red y ficheros locales.** 74 llamadas a `fso.FileExists` y 5 a `Environ(`. La aplicación depende de rutas de fichero que no están en el repositorio. Ninguna se reproduce en este PRD.

## 6. Auditoría y caché

**Auditoría.** Dos mecanismos distintos: `TbLog` y `TbLogPublicaciones` (11 columnas, con evidencia en 2 ficheros de `src`) y `PublicacionLog`. `TbLog` **no tiene ninguna cita en el código de producción**: o se escribe por SQL construido que no la nombra, o está heredada. Hay que verificarlo antes de asumir que existe auditoría efectiva.

**Caché.** Seis tablas materializan en disco el árbol de riesgos y el control de cambios: `TbArbolCacheMeta`, `TbCacheArbolRiesgosMeta`, `TbCacheArbolRiesgosNodo`, `TbCacheControlCambiosMeta`, `TbCacheControlCambiosRow` y `TbCachePublicabilidadEdicion`. El estudio previo ya señalaba la lentitud del árbol (D88) y la existencia del flag `Publicabilidad_Usar_Cache` (hoy «No»).

Estas tablas están en tensión directa con la solución que la web necesita: en PostgreSQL el árbol debería salir de una **CTE recursiva sobre los datos**, no de una tabla que se reconstruye a mano. Migrarlas perpetuaría el problema que las motivó. **Es una decisión abierta y probablemente de las más importantes del PRD**, no un hecho: la alternativa sería una tabla de presentación mantenida por disparador, que tiene su propio coste.

## 7. Reglas compartidas que la migración tiene que respetar

1. El frontend no tiene tablas propias. Las de caché son la excepción a decidir, no a extender.
2. La identidad y los permisos son de Lanzadera y los usan ocho aplicaciones: no se redefinen aquí.
3. Los siete backends externos son de otras aplicaciones: su corte se coordina, no se decide desde este PRD.
4. `Err.Raise` propaga por cadena 3 752 veces: el mapa de errores es parte del comportamiento observable y hay que migrarlo, no inventarlo.

## 8. Pruebas bidireccionales

Durante la coexistencia hacen falta dos direcciones, y ninguna está resuelta hoy:

- **Legacy → web.** Las consultas que el legacy hace contra tablas que la web va a sustituir. Con quince tablas vinculadas y siete backends externos, la web tiene que seguir sirviendo exactamente lo que el legacy lee. Hoy no hay contrato ni prueba que lo fije.
- **Web → legacy.** Lo que la web escriba y el legacy deba leer. No aplica todavía: no existe módulo `app/src/modules/gestion_riesgos`.

**Prueba mínima imprescindible antes de cualquier corte:** extraer el conjunto real de consultas que cada formulario ejecuta contra cada tabla vinculada y comprobarlo contra la API de la web. Con 195 ficheros de producción y el SQL embebido, esa extracción es trabajo, no un trámite. Es candidata a issue propuesta.

## 9. Corte en la Lanzadera y reversión

El corte de Gestión de Riesgos **no puede ser de esta aplicación**. Su secuencia real está atada a:

1. Que Lanzadera tenga su identidad y sus permisos migrados, porque `TbUsuariosAplicaciones` y `TbUsuariosAplicacionesPermisos` son suyas y esta aplicación las lee en cada entrada.
2. Que los cuatro backends externos (Expedientes, No Conformidades, AGEDO, AGEDYS) sigan respondiendo con el mismo contrato, porque esta aplicación los consulta en sus formularios (`Form_FormRiesgoNC`, `Form_FormExpedientesBusqueda`, `Documento.cls`, `Pedido.cls`).
3. Que el correo siga saliendo, por base de datos y no por SMTP.

**La reversión no está definida y no puede definirse aquí.** Un corte de esta aplicación deja en producción un legacy que, si hay que volver atrás, depende de siete backends que pueden haber cambiado mientras tanto. La reversión solo es definible cuando exista un ERD global que fije qué cambió en cada backend y una prueba bidireccional que lo detecte. **Queda como decisión del propietario**, junto con P-EXP-3 (arquitectura de datos de coexistencia), que sigue abierta y escalada.

## 10. Resumen de lo medido

| Magnitud | Valor | Lectura |
|---|---|---|
| Backends que abre | 8 (1 propio, 7 ajenos) | La aplicación no es desplegable sola |
| Tablas vinculadas | 15, de 6 backends | De las cuales 3 son propiedad de Lanzadera |
| Llamadas a `getdb*` en producción | 417 | El acceso a datos está en clases y módulos, no en la UI |
| Transacciones | 1 `BeginTrans`, 1 `CommitTrans`, 0 `RollbackTrans` | Sin atomicidad |
| `On Error Resume Next` | 263 | Comportamiento de error mayoritariamente silenciado |
| `Err.Raise` | 3 752 | El mapa de errores es parte del comportamiento |
| Identificadores por `Max + 1` | 8 llamadas | Riesgo de colisión, condicionado al ERD de origen |
| `OpenRecordset` | 283 | Todo el DAO es procedural |
| Consultas guardadas | 0 | El SQL está embebido en el código |
| Tablas de caché | 6 | En tensión con la solución web del árbol |

## Pendientes

- ERD de origen (D-GR-4): claves, índices, relaciones y nulabilidad, que confirmarían o refutarían el riesgo de `Max + 1` y darían la cardinalidad real de los vínculos.
- Resolver D-GR-5 (`TbProyectos1` vinculado a sí mismo) antes de usar `Estructura_Datos.md` como evidencia.
- Decidir si las seis tablas de caché se migran.
- Definir el contrato de lectura que la web debe servir a este legacy mientras conviven.
- Definir la reversión, que depende de lo anterior.