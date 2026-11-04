# Fase 0: aperturas directas de `Lanzadera_Datos` desde otros fronts

> **Issue:** #1570 · **Épica:** [coexistencia progresiva](../epica-coexistencia-postgresql.md), fase 2 · **Inventario relacionado:** [fase 0 de la Lanzadera](lanzadera-datos.md#verificación-de-r-1-y-r-2) · **Fecha:** 2026-11-04

## Propósito

Inventario de solo lectura de los fronts que abren `Lanzadera_Datos.accdb` sin pasar por un vínculo de su backend. Es el cuarto mecanismo de acceso de la fase 2: tiene que adaptarse a ODBC en el mismo corte que `Lanzadera_Datos`. Este documento propone alternativas; la decisión es del propietario.

## Reglas del propietario

Confirmadas expresamente el 2026-11-04: no hay macros ni `AutoExec`; los formularios son desconectados y sus controles se rellenan en tiempo de ejecución; el front no tiene tablas vinculadas del backend. Los vínculos inventariados en este documento están **dentro de los backends**, no en los fronts. El análisis de código se basa en CodeGraph, trazando los procedimientos que cargan formularios y controles. No se deducen usos de datos a partir de propiedades de diseño vacías.

Estas reglas retiran el anterior pendiente de inspeccionar macros y `AutoExec`: era una suposición del análisis, no una carencia del legacy.

## Fuente y método

- **Código:** `src` de cada front en su rama `main` local, con el SHA de la tabla de resumen. No se hizo `fetch`: el SHA fija lo inspeccionado, no lo desplegado en producción.
- **Rutas:** relativas al `src` de cada front. `apertura` es la línea donde empieza la función; `OpenDatabase` es la línea de la llamada real.
- **Búsqueda inicial:** `Lanzadera`, `OpenDatabase`, `DBEngine` y `Workspaces` en `src`; el explorador usó lectura directa al no disponer del acceso multirrepositorio a CodeGraph. **Contraste posterior del coordinador:** `codegraph_explore` con `projectPath` funcionó en los doce repositorios, consultando `getdbLanzadera`, `getUsuario` y `getAplicacionesPermisos`. En Condor traza también `CheckAndReconnect` y `OpenDatabase`. Se preservan las aristas marcadas como dinámicas: el grafo estático no es prueba de ejecución ni de ausencia de invocaciones.
- **Vínculos:** extracción DAO de solo lectura de las copias de desarrollo de los backends (la misma de la [fase 0 de la Lanzadera](lanzadera-datos.md#fuente-y-método)).
- **Recuentos:** «llamadas» son las líneas de `src` que invocan la función de apertura, y «procedimientos», los distintos que las contienen. Miden la superficie de código que toca la adaptación, no el esfuerzo de implementarla.
- **Claves:** diez fronts escriben la clave del `.accdb` como literal en el código; Condor la obtiene de `GetPasswordDB()`. Los literales no se reproducen aquí.

## Mecanismo común

Los 11 fronts con apertura repiten el mismo patrón:

- La ruta depende de `TempVars("DatosEnLocal")`. Con `Sí`, `m_URLRutaAplicacionesLocal` + `000datoslocal\Lanzadera_Datos.accdb`; con `No`, `m_URLRutaAplicacionesRemotas` + `0Lanzadera\Lanzadera_Datos.accdb`; con otro valor, error 1000. Condor trata cualquier valor distinto de `Sí` como remoto (`FUNCIONES UTILES.bas:234-235`).
- `OpenDatabase(ruta, False, False, …)`: modo compartido y de lectura y escritura.
- Sin caché, salvo Condor, que guarda la conexión en `g_dbLanzadera`. Expedientes guarda la conexión en una variable local dentro de `getUsuario`; no se ha demostrado una caché global.
- El origen de `TempVars("DatosEnLocal")` no se ha trazado más allá de su lectura.

## Resumen por front

| Front | SHA de `main` | Apertura → `OpenDatabase` | Llamadas / procedimientos | Identidad vinculada en su backend | Escritura observada |
|---|---|---|---|---|---|
| AGEDYS | `350896d1300c68818f3dfa345bb8b5e4173f7d39` | `Variables Globales.bas:439` → `:458` | 12 / 10 | Sí (`AGEDYS_DATOS`) | `TbAplicacionesAperturas`; llamador comentado |
| AGEDYS ECO | `a263ead7c2e6bab3aae85fdb07ce2193d5f91dba` | `Variables Globales.bas:439` → `:458` | 12 / 10 | Sí (`AGEDYS_DATOS`, compartido con AGEDYS) | `TbAplicacionesAperturas` (`AddNew`); llamador comentado |
| Brass | `3f77d586814096c04f24bd2b6c229831cfbfcd34` | `Variables Globales.bas:520` → `:539` | 6 / 5 | No | `TbAplicacionesAperturas`; llamadores comentados |
| Contratos | `effb968c08ae0862517cbc3838e17bdaa161a060` | `Variables Entorno.bas:249` → `:268` | 8 / 7 | No | `TbAplicacionesAperturas`; alcanzable |
| Registros | `44f36d3774a4870e7f70d8c8aa71710965f61411` | `Variables de Entorno.bas:372` → `:391` | 8 / 7 | No | `TbAplicacionesAperturas`; alcanzable |
| Gestión de Riesgos | `36fcf60161fa5ec84c3956b1308537ddf2579256` | `Variables Globales.bas:457` → `:476` | 9 / 8 | Sí (`Gestion_Riesgos_Datos`) | Ninguna |
| HPS | `dc194f4cacf0c7db9388c6d01e0aecc7a8547663` | `VariablesEntorno.bas:545` → `:564` | 5 / 5 | Sí (`HPST`) | Ninguna |
| Solicitudes HPS | `cfa1937ab05a9e40150ea9236243898de96f84da` | `Variables Globales.bas:430` → `:449` | 6 / 5 | No (vincula Expedientes y HPS) | Ninguna |
| Seguridad | `7445049de98ed0bfff245a084ce4b43ef3dbe00f` | `Variables de Entorno.bas:403` → `:422` | 5 / 5 | No | Ninguna |
| Expedientes | `495b96596fa29b647df3959a8638abe785b44a11` | `Variables Globales.bas:487` → `:506` | 6 / 5 | Sí (`Expedientes_datos`) | Ninguna |
| Condor | `97b98732501545727c91e5807d55af8613c65800` | `FUNCIONES UTILES.bas:64` → `CheckAndReconnect:211` → `:252` | 4 / 4 | Sí (`condor_datos`) | Ninguna |
| No Conformidades | `d834997e16ae1129fb70f12ad4181bda2e353a6b` | Sin apertura en `main` | 0 / 0 | Sí (`NoConformidades_Datos`) | Ninguna |
| Lanzadera (propietaria) | `d01d5037f90041f3bebc55104f0c35a71e90cae1` | `getdb()`, `Variables Globales.bas:253` → `:310` | — | Backend propio | — |

Total de los 11 fronts con apertura: 81 llamadas en 71 procedimientos.

«Identidad vinculada» significa que el backend vincula `TbUsuariosAplicaciones` y `TbUsuariosAplicacionesPermisos`. En esos fronts hay doble camino: la misma tabla llega por vínculo y por apertura directa. `Tareas_datos1` también las vincula, pero Tareas no tiene repositorio inspeccionado.

## Tablas por la conexión directa

Los 11 fronts leen `TbUsuariosAplicaciones` y `TbUsuariosAplicacionesPermisos`. La única escritura observada es en `TbAplicacionesAperturas` (`RegistroApertura` con `AddNew`, `RegistroCierre` con `UPDATE`). Ningún front escribe `TbConexionesRegistro` en su `src`.

### AGEDYS

- Lecturas: `constructor.bas:1252` (`getUsuario`), `:1294` y `:1329` (`getUsuariosAdministradores`), `:1384` (`getUsuariosCalidad`), `:1436` (`getUsuariosEconomia`), `:1488` (`getUsuariosSecretaria`), `:3648` (`getAplicacionesPermisos`), `:3723` (`getListaUsuarios`); `UsuarioAplicacionPermisos.cls:263` (`TableDefs`).
- Escrituras: `Funciones Generales.bas:8420` y `:8425` (`RegistroApertura`, `DameID1` y `AddNew`), `:8469` (`RegistroCierre`, `UPDATE`).
- Entrada: `Variables Globales.bas:654` → `getUsuario` (`:719`, `:728`), desde `Form_Form0BDOpciones.cls` (`Form_Load:131`, `Detalle_DblClick:85`). Las llamadas a las escrituras están comentadas en `Variables Globales.bas:824` y `Form_Form0BDOpciones.cls:16`.

### AGEDYS ECO

- Lecturas: `constructor.bas:793` (`getUsuario`), `:832` (`getAplicacionesPermisos`), `:906` (`getListaUsuarios`), `:1509` y `:1544` (`getUsuariosAdministradores`), `:1599` (`getUsuariosCalidad`), `:1651` (`getUsuariosEconomia`), `:1706` (`getUsuariosSecretaria`); `Usuario.cls:282`; `UsuarioAplicacionPermisos.cls:263`.
- Escrituras: `FUNCIONES UTILES.bas:966` y `:971` (`RegistroApertura`, `AddNew`). Su llamador está comentado en `Variables Globales.bas:802`.
- Entrada: `Variables Globales.bas:652`, desde `Form_Form0BDOpciones.cls:38`, `:379` y `:421`.

### Brass

- Lecturas: `Constructor.bas:71` (`getUsuario`), `:110` (`getUsuarios`, recordset de tipo tabla), `:165` (`getAplicacionesPermisos`).
- Escrituras: `Funciones Generales.bas:7623`, `:7628` y `:7672`. Sus llamadores están comentados en `Form_Form0BDOpciones.cls:16` y `:73`.
- Entrada: `Variables Globales.bas:264`, desde `Form_Form0BDOpciones.cls:57`.

### Contratos

- Lecturas: `Constructor.bas:71` (`getUsuario`), `:113` (`getAplicacionesPermisos`); `Funciones Generales.bas:156` (`getEsUsuarioAutorizado`, llamadores sin trazar); `Usuario.cls:183`; `UsuarioAplicacionPermisos.cls:263`.
- Escrituras: `Funciones Generales.bas:510` y `:515` (`RegistroApertura`, `DameID` y `AddNew`), `:558` (`RegistroCierre`).
- Entrada: `Form_Inicio.cls` `Form_Load:63` (tras la carga de entorno en `:59`) y `BotonSalir_Click:21`. Las escrituras son alcanzables desde estos eventos.

### Registros

- Lecturas: `constructor.bas:74` (`getUsuario`), `:114` (`getAplicacionesPermisos`), `:182` (`getListaUsuarios`); `Usuario.cls:178`; `UsuarioAplicacionPermisos.cls:263`.
- Escrituras: `Funciones Generales.bas:2767`, `:2772` y `:2834`.
- Entrada: `Form_Form0BDOpciones.cls` `Form_Open:99` y `cmdSalir_Click:15`. Las escrituras son alcanzables desde estos eventos.
- La consulta guardada `queries/Consulta3.sql` une `TbUsuariosAplicaciones` con `TbConexionesRegistro` (solo lectura, con un SSID fijo). No se ha trazado de dónde obtiene esas tablas.

### Gestión de Riesgos

- Lecturas: `Constructor.bas:3018` (`getUsuario`), `:3057` y `:3091` (`getUsuariosAdministradores`), `:3145` (`getUsuariosCalidad`), `:3198` (`getUsuariosCalidadAvisos`), `:3264` (`getUsuariosTecnicos`), `:3421` (`getAplicacionesPermisos`); `Usuario.cls:314`; `UsuarioAplicacionPermisos.cls:268`.
- Entrada: `Variables Globales.bas:204`, desde `Form_frmSplash.cls:61` y `Form_Form0BDOpciones.cls:82`.
- `TbAplicacionesVersiones` se lee por `getdbControlCambios`, no por la conexión a `Lanzadera_Datos`.

### HPS

- Lecturas: `Constructor.bas:1430` (`getUsuario`), `:1468` (`getUsuarios`), `:1625` (`getAplicacionesPermisos`); `Usuario.cls:186`; `UsuarioAplicacionPermisos.cls:263`.
- Entrada: `VariablesEntorno.bas:301`, desde `Form_FormInicial.cls:28`.

### Solicitudes HPS

- Lecturas: `constructor.bas:90` y `:125` (`getUsuariosAdministradores`), `:180` (`getUsuariosTecnicos`), `:235` (`getUsuariosTramitadores`), `:358` (`getUsuario`), `:401` (`getAplicacionesPermisos`).
- Entrada: `Form_frmSplash.cls` `Form_Timer:37`.

### Seguridad

- Lecturas: `Constructor.bas:72` (`getUsuario`), `:111` (`getUsuarios`, recordset de tipo tabla), `:167` (`getAplicacionesPermisos`); `Usuario.cls:282`; `UsuarioAplicacionPermisos.cls:263`.
- Entrada: `Form_Form0BDOpciones.cls` `Form_Load:109` y `Detalle_DblClick:87`.

### Expedientes

- Lecturas: `constructor.bas:76` (`getUsuario`), `:138` (`getAplicacionesPermisos`), `:193` y `:228` (`getUsuariosAdministradores`), `:294` (`getUsuariosTecnicos`); `modResponsablePorRolHelper.bas:107` (`GetUsuariosPorRol`) → `LookupUsuarioPorID:204`.
- Entrada: `Variables Globales.bas:257`, desde `frmSplash.cls:38`, `Form_Form0BDOpciones.cls:177` y `Form_Form0BDOpcionesTecnicos.cls:122`.

### Condor (confirmado)

- Apertura: `getdbLanzadera` (`FUNCIONES UTILES.bas:64`) llama a `CheckAndReconnect(g_dbLanzadera, "Lanzadera")` en `:68`; `CheckAndReconnect` (`:211`) elige `Lanzadera` en `:234`, construye la ruta en `:235` y abre con `getWorkspace().OpenDatabase` en `:252`.
- Lecturas: `AplicacionRepositorio.bas:27` (permisos), `UsuarioRepositorio.bas:42` (usuarios), `UsuarioServicio.cls:141` (`getResponsablesCalidad`, une las dos tablas de identidad).
- Entrada: `Variables Globales.bas:56` → `getdbLanzadera` en `:94`, desde `frmSplash.cls:49`.

### No Conformidades

- **`main`:** sin apertura. La búsqueda de `Lanzadera`, `OpenDatabase`, `DBEngine` y `Workspaces` solo devuelve el `getdb()` de su propio backend (`Variables Globales.bas:149`, `:184` y `:221`). Las lecturas de usuarios de `constructor.bas:167`, `:210`, `:323` y `:383` van por `getdb()`, es decir, por los vínculos.
- **`release/2026-009.rc1`** (commit `f281b4e3ad33e5411ce6e0e6f59ad388698a10b1`; su `src` es idéntico al de `4810fa225b9995956cf474893d8c90d460092c58`): añade una apertura directa **de solo lectura**. `modOrigenExpedientes.bas:99` llama a `OpenDatabase(m_Ruta, False, True, …)`, con la ruta y la clave tomadas del `Connect` del vínculo de `TbUsuariosAplicaciones` (`:90`), no escritas en el código. La conexión de Lanzadera está en `:166` y la usan `constructor.bas:407` y `:5126` con `getdb(p_Origen:="Lanzadera")`. Si esa rama llega a producción, No Conformidades pasa a tener doble camino.

## Propuesta de adaptación

El coste de cada front se expresa como superficie acotada: una función de apertura más sus llamadas observadas. Es lo que hay que revisar y probar, no una estimación de esfuerzo. Ninguna alternativa está elegida.

| Alternativa | Qué cambia | Superficie por front | Riesgos |
|---|---|---|---|
| A. Vínculo ODBC del propio backend (R-1) | Las lecturas pasan al abridor `getdb()` del backend propio; ese backend contiene los vínculos necesarios y el front permanece sin tablas vinculadas | Las llamadas de la tabla de resumen, más los vínculos nuevos | Hay que crear vínculos en los backends sin identidad vinculada (Brass, Contratos, Registros, Solicitudes HPS, Seguridad) y en todos los que escriben `TbAplicacionesAperturas`, que hoy no la vinculan |
| B. Conexión ODBC centralizada | La función de apertura de cada front cambia la ruta del `.accdb` por una cadena ODBC sin DSN; las llamadas no cambian | Una función por front (11 en total) | Sobre ODBC, DAO no admite recordsets de tipo tabla (Brass `Constructor.bas:110`, Seguridad `Constructor.bas:111`) y `TableDefs` (`UsuarioAplicacionPermisos.cls`) devuelve otra colección: hay que probarlo por front |

Por front:

| Front | A: llamadas que cambian | A: vínculos que faltan | B: funciones que cambian |
|---|---|---|---|
| AGEDYS / AGEDYS ECO | 12 / 12 | `TbAplicacionesAperturas` en `AGEDYS_DATOS` (solo si se reactivan las escrituras) | 1 / 1 |
| Brass | 6 | Identidad y, si se reactivan las escrituras, `TbAplicacionesAperturas` | 1 |
| Contratos | 8 | Identidad y `TbAplicacionesAperturas` | 1 |
| Registros | 8 | Identidad y `TbAplicacionesAperturas` | 1 |
| Gestión de Riesgos | 9 | Ninguno | 1 |
| HPS | 5 | Ninguno | 1 |
| Solicitudes HPS | 6 | Identidad | 1 |
| Seguridad | 5 | Identidad | 1 |
| Expedientes | 6 | Ninguno | 1 |
| Condor | 4 | Ninguno | 1 |
| No Conformidades | 0 en `main` | Ninguno | 0 en `main`; 1 si entra la rama `rc1` |

La alternativa B deja la clave fuera del `.accdb` solo si la cadena ODBC no repite el patrón de literal en el código; eso forma parte de la decisión.

## Superficie medida y requisitos previos al corte

Ampliación de #1570 (2026-11-04), sobre las revisiones fijadas arriba. A cuenta líneas de llamada existentes, no comentarios ni definiciones. B cuenta líneas no vacías del abridor y de la selección de ruta más apertura. Son superficies de revisión: no presupuestos de implementación ni líneas de diff garantizadas.

| Front | Función existente (inicio–fin) | B: función no vacía / bloque de ruta y apertura | A: líneas de llamada |
|---|---|---|---|
| AGEDYS | `Variables Globales.bas:439–466` | 22 / 10 | 12 |
| AGEDYS ECO | `Variables Globales.bas:439–466` | 22 / 10 | 12 |
| Brass | `Variables Globales.bas:520–547` | 22 / 10 | 6 |
| Contratos | `Variables Entorno.bas:249–276` | 22 / 10 | 8 |
| Registros | `Variables de Entorno.bas:372–399` | 22 / 10 | 8 |
| Gestión de Riesgos | `Variables Globales.bas:457–484` | 22 / 10 | 9 |
| HPS | `VariablesEntorno.bas:545–572` | 22 / 10 | 5 |
| Solicitudes HPS | `Variables Globales.bas:430–457` | 22 / 10 | 6 |
| Seguridad | `Variables de Entorno.bas:403–430` | 22 / 10 | 5 |
| Expedientes | `Variables Globales.bas:487–514` | 22 / 10 | 6 |
| Condor | `FUNCIONES UTILES.bas:234–235,252` | 2 líneas de selección propias y 1 apertura compartida | 4 |
| No Conformidades `main` | Sin apertura directa identificada | 0 | 0 |
| No Conformidades RC | `modOrigenExpedientes.bas:166–168`; abridor compartido `:64–107` | 3 líneas propias, más adaptación del abridor compartido | 4 más 1 camino de arranque |

En los diez abridores iguales: 220 líneas no vacías de función, 100 de bloque de ruta y apertura, y 77 llamadas. Con Condor, A suma las 81 llamadas del inventario inicial. La RC de No Conformidades no entra en ese total: además de `constructor.bas:407,5126`, lee mediante `OrigenLanzadera_Abrir` en `constructor.bas:336` y `modCatalogoUsuarios.bas:140`; arranca por `Variables Globales.bas:441` y `modOrigenExpedientes.bas:132`. Si no consigue abrir, existe retorno al vínculo (`:170–183`). Los ficheros citados no presentan diferencias locales frente al HEAD RC comprobado.

No se puede sustituir indiscriminadamente el abridor compartido: Condor lo usa también para sus otras cuatro bases y la RC de No Conformidades para Expedientes y Riesgos. Tampoco está demostrado que una cadena ODBC pueda reemplazar sin más el contrato DAO actual: requiere ensayo del modo de apertura, SQL y recordsets.

### Acciones de preparación derivadas

| Id | Acción propuesta, sin ejecutarla en el legacy | Evidencia / aceptación antes del corte |
|---|---|---|
| P1 | Sustituir la asignación `MAX+1` por una fuente de ID atómica, coordinando todos los escritores de aperturas | `DameID1` en AGEDYS (`Funciones Generales.bas:1261`), Brass (`:332`) y Registros (`:495`); `DameID` en AGEDYS ECO (`FUNCIONES UTILES.bas:519`) y Contratos (`Funciones Generales.bas:213`). Solo leen `MAX` y devuelven el siguiente valor; no reservan el ID. Probar dos inserciones simultáneas sin colisiones. No es un cambio solo del backend: adaptar también los fronts escritores. |
| P2 | Extraer credenciales del código y definir cómo se proporcionan en Citrix | Diez abridores contienen un literal, sin reproducirlo aquí; Condor usa `GetPasswordDB`. Acordar almacenamiento, permisos y rotación antes de implementar. |
| P3 | Revisar la vida de las conexiones antes de ODBC | AGEDYS declara `wks` y `db1` a nivel de módulo (`Variables Globales.bas:375,377`) y reasigna una apertura por llamada. Medir sesiones y cierre; no se ha probado una fuga ni su impacto en PostgreSQL. |
| P4 | Decidir qué hacer con `getEsUsuarioAutorizado` de Contratos antes de conservarla o retirarla | `Funciones Generales.bas:143–207`: sin llamadores encontrados en `src`, solo lee permisos pero llama a `Desconectar` (`:128`), que cierra la base global propia. Tiene lógica de autorización pendiente de revisión; no declararla muerta ni corregirla sin revisar el grafo de llamadores y los requisitos. |

Los auxiliares de ID citados no escriben registros ni cierran la base recibida; si reciben `Nothing` eligen otra base por defecto. La lectura de `MAX` seguida de `AddNew` no constituye una operación atómica.

### Corrección de la incertidumbre sobre macros y arranque

Se consultaron los contratos de lectura de dysflow y una muestra de `FormLogin`, pero esa comprobación no era el camino apropiado para este ecosistema. El propietario confirmó después que no existen macros ni `AutoExec` y que los formularios son desconectados. Se retira ese bloqueo y se usa CodeGraph sobre los procedimientos de carga y eventos; no se ejecuta el arranque contra producción.

El contraste posterior con `codegraph_explore(projectPath, query)` funcionó en los doce fronts: el fallo anterior pertenecía al acceso del explorador, no al servidor. No convierte las aristas dinámicas en ejecución demostrada. La adaptación concreta y sus pruebas funcionales siguen pendientes de elegir e implementar; esta issue solo entrega el inventario y su superficie medida.

## Riesgo residual

- No se inspeccionan macros ni `AutoExec`: no existen por regla del propietario. La revisión se centra en los eventos y procedimientos de carga en ejecución mediante CodeGraph.
- Las llamadas comentadas a las escrituras (AGEDYS, AGEDYS ECO, Brass) no prueban ejecución; hay que conservar las limitaciones de resolución dinámica del grafo.
- Las rutas construidas en tiempo de ejecución se buscaron por `OpenDatabase`, `DBEngine` y `Workspaces`, no por todo SQL dinámico.
- Los auxiliares de ID y `getEsUsuarioAutorizado` ya se inspeccionaron en `src` (ampliación anterior); sus posibles invocaciones externas al código siguen sin verificarse.
- No hay pruebas en ejecución: todo lo anterior es lectura de código y de vínculos.
