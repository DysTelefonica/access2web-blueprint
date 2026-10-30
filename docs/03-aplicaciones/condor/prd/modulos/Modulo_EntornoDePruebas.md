# Módulos de entorno de pruebas y batería de pruebas

> **Commit fijado:** `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · **Fuente:** `src/modules/SandboxBuilder.bas` (618 líneas), `SandboxProvisioner.bas` (1003), `TestSandbox.bas` (591), `modBattery_Canonical.bas` (7813) · **PRD:** [README](../README.md)

Cuatro módulos que no forman parte del comportamiento de negocio: clonan la base de datos en un entorno de pruebas, copian los backends externos a directorios laterales, y ejecutan una batería de pruebas unitarias del flujo de trabajo. Están en el mismo ejecutable que la aplicación y los inventaría el estudio previo como código Ordinary, pero su función es de andamiaje. Esta hoja los clasifica y los marca como **no trasladables a la versión web**, que es lo que corresponde al criterio de inventario por estado.

Abreviaturas: `SB`, `modules/SandboxBuilder.bas`; `SP`, `modules/SandboxProvisioner.bas`; `TS`, `modules/TestSandbox.bas`; `BAT`, `modules/modBattery_Canonical.bas`; `FU`, `modules/FUNCIONES UTILES.bas`.

## Superficie pública

| Objeto | Tipo | Procedimientos públicos | Qué resuelve |
|---|---|---|---|
| `SandboxBuilder` | módulo | `BuildSandbox`, `OpenSandbox`, `CloseSandbox` (`SB:51-171`) | Clonar la base de Condor, localizar sus tablas vinculadas y abrir la copia |
| `SandboxProvisioner` | módulo | `ProvisionSandbox_Fase2_Localizar`, `Sandbox_HasLinkedTables` y quince procedimientos privados (`SP:50-1003`) | Aprovisionar el entorno: localizar rutas, detectar vínculos, copiar backends externos y crear laterales por compactación |
| `TestSandbox` | módulo | `Sandbox_Ruta`, `Sandbox_Existe`, `Sandbox_Clonar`, `Sandbox_AbrirInstancia`, `Sandbox_CerrarInstancia`, `Sandbox_Eliminar`, `Test_StartTransaction` y otros escenarios (`TS:43-591`) | Arnés del entorno de pruebas |
| `modBattery_Canonical` | módulo | `Canonical_Setup`, `Canonical_TearDown`, `Canonical_GetTestList`, `Canonical_RunAll`, `Canonical_Log_Trace` y 161 pruebas unitarias `*_UT_*` (`BAT:66-7813`) | Batería de pruebas del flujo de trabajo, de los repositorios y de los datos |

## Procedimientos y llamadas

| Objeto | Procedimiento | Qué hace | Datos y efectos | Evidencia |
|---|---|---|---|---|
| `SB` | `BuildSandbox` | Crea la copia de la base y registra los ficheros laterales para su limpieza | Copia ficheros | `SB:51-117`; traza de entrada en `SB:54`, colección de laterales en `SB:57` |
| `SB` | `OpenSandbox`, `CloseSandbox` | Abre y cierra la instancia del entorno | Abre y cierra una base | `SB:118-171` |
| `SB` | `ResolveSandboxPaths`, `FileCopyCondorBackend`, `ResolveCondorBackendPath` | Resuelven rutas y copian el backend | Copian ficheros | `SB:172-305` |
| `SB` | `DetectLinkedTablesAndSidecars`, `ExtractBackendPathFromConnect` | Detectan las tablas vinculadas y extraen el backend de la cadena de conexión | Sin escritura | `SB:306-425` |
| `SB` | `CopyExternalBackendsToSidecars` | Copia los backends externos a directorios laterales junto a la copia | Copia ficheros | `SB:426-618` |
| `SP` | `ProvisionSandbox_Fase2_Localizar` | Punto de entrada del aprovisionamiento | Localiza el entorno | `SP:71-227` |
| `SP` | `CreateSidecarViaCompactDatabase` | Crea un lateral por compactación de la base externa | Crea ficheros | `SP:502-734` |
| `TS` | `Sandbox_*` | Ruta, existencia, clonación, apertura, cierre y borrado del entorno | Copia y borra ficheros | `TS:58-204` |
| `TS` | `Test_StartTransaction` y escenarios | Ejercitan transacciones sobre el entorno | Sin escritura sobre producción | `TS:205-591` |
| `BAT` | `Canonical_Setup`, `Canonical_TearDown` | Prepara y recoge la corrida de pruebas | Escribe un fichero de registro | `BAT:162-248`; registro en `BAT:66-68` |
| `BAT` | `Canonical_GetTestList`, `Canonical_RunAll` | Devuelven el catálogo de pruebas y ejecutan la batería | Sin efectos sobre producción | `BAT:249-475` |
| `BAT` | 161 pruebas `*_UT_*` | Ejercitan transiciones, códigos de error y datos de los repositorios | Sin efectos sobre producción | `BAT:476-7813`; ejemplos en `BAT:476`, `BAT:520` |

## Flujos de trabajo

- **Normal:** para ejecutar pruebas se prepara el entorno (`SB:51-117`), se localizan las tablas vinculadas (`SB:306-425`), se copian los backends externos a laterales (`SB:426-618`), se abre la instancia (`SB:118-149`) y se lanza la batería (`BAT:446-475`), que deja su registro en un fichero (`BAT:66-68`).
- **Alternativos:** `SandboxProvisioner` repite el mismo trabajo con otro punto de entrada y su propia fase (`SP:71-227`), lo que sugiere una migración del módulo antiguo al nuevo a medio hacer.
- **Errores:** el arnés registra el fallo y continúa; la batería no detiene la corrida ante el fallo de una prueba.

## Reglas, validaciones y cálculos

| Regla | Dónde | Evidencia |
|---|---|---|
| Ningún módulo de este grupo escribe en el backend de producción | operan solo sobre copias | `SB:221-305`, `SP:363-501`, `TS:107-138` |
| La batería se ejecuta sobre la instancia del entorno | `Canonical_RunAll` | `BAT:446-475` |
| Los laterales se registran para poder limpiarlos | `m_colSidecarFiles` | `SB:56-57` |

## Lecturas y escrituras

| Operación | Recurso | Efecto | Evidencia |
|---|---|---|---|
| Copia de la base y de sus laterales | ficheros de la aplicación | escritura en disco | `SB:221-305`, `SB:426-618`, `SP:363-501` |
| Ejecución de pruebas | instancia del entorno | sin escritura en producción | `BAT:446-475` |

## Coexistencia

Ninguno de los cuatro módulos se traslada a la versión web, y no por decisión de diseño sino porque su objeto es el fichero `.accdb`: clonar un fichero, compactarlo y resolver sus vínculos son operaciones de Access. Lo que sí debe sobrevivir es **lo que la batería demuestra**. Las 161 pruebas unitarias (`BAT:476-7813`) son la especificación ejecutable del flujo de trabajo: describen qué transiciones existen, qué errores se lanzan y con qué número, y qué datos se guardan. Ese catálogo es el insumo más valioso que este grupo aporta a la migración, porque permite validar el comportamiento del destino contra el del origen sin releer el código de negocio. La recomendación que este PRD deja sobre la mesa, sin decidirla, es reutilizar esos nombres y expectativas como casos de prueba de la API.

## Comportamiento actual · Posibles defects · Propuestas web pendientes de decisión

**Comportamiento actual.** Cuatro módulos de andamiaje que clonan, aprovisionan y prueban la aplicación sobre copias locales, con registro en fichero.

**Posibles defectos, a verificar.**
1. `SandboxBuilder` y `SandboxProvisioner` resuelven lo mismo con dos conjuntos de procedimientos (`SB:51-618` y `SP:50-1003`): hay una migración interna sin cerrar.
2. El código de pruebas y el de andamiaje se despliegan con la aplicación: 10025 líneas de los 93 módulos de `src/modules` son de pruebas.
3. Los procedimientos de la batería son públicos (`BAT:476-7813`): cualquiera que abra el proyecto puede ejecutarlos contra el entorno equivocado.

**Propuestas web pendientes de decisión.** Que las 161 pruebas se conviertan en la batería de la API y conserva su nombre como identificador de caso. Que el entorno de pruebas del destino sea una base propia con sus propias credenciales, en lugar de una copia de ficheros locales.

## Criterios de aceptación

1. Ninguna operación de estos cuatro módulos alcanza la base de producción. Trazado a `SB:221-305` y `TS:107-138`.
2. El catálogo de pruebas es reproducible y su registro queda en un destino conocido. Trazado a `BAT:249-475` y `BAT:66-68`.
3. La batería cubre transiciones de estado, códigos de error y escritura de datos. Trazado a `BAT:476-7813`.

## Pendientes

- Confirmar contra el binario si `SandboxBuilder` está retirado en favor de `SandboxProvisioner` o si ambos siguen activos (`SB:51-618`, `SP:50-1003`).
- Confirmar si los 161 procedimientos de la batería se despliegan con el ejecutable de producción (`BAT:476-7813`).
- Confirmar qué parte de la batería cubre cada repositorio y cada servicio, para poderla reutilizar como especificación de la API.