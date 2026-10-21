# Coexistencia de Condor

> **Commit fijado:** `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · **Issue:** #1075 · **Fase:** 3 del [roadmap](../../../07-migracion/roadmap-descubrimiento-prd.md) · **PRD:** [README](README.md)

Resumen de lo que Condor necesita para convivir con la base heredada y con la versión web durante la transición, y para poder retirarse sin dejar datos ni reglas a medias. Las cifras de este documento son mediciones estáticas sobre el clon de solo lectura del commit fijado; las decisiones que aparecen marcadas como abiertas son del propietario y este documento no las toma.

Abreviaturas de cita: `FU`, `modules/FUNCIONES UTILES.bas`; `USR`, `modules/UsuarioRepositorio.bas`; `SOL`, `modules/SolicitudRepositorio.bas`; `ADJ`, `modules/AdjuntoRepositorio.bas`; `LOG`, `modules/LogCambioRepositorio.bas`; `ERR`, `modules/LogErrorRepositorio.bas`; `COR`, `modules/CorreoRepositorio.bas`; `RECH`, `modules/RechazoRepositorio.bas`; `PC`, `modules/DatosPCRepositorio.bas`; `RC`, `modules/RepositorioComun.bas`; `CE`, `classes/CondorError.cls`; `LC`, `classes/LogCambio.cls`; `LE`, `classes/LogError.cls`; `DOC`, `classes/DocumentoServicio.cls`; `SH`, `modules/SnapshotHelper.bas`; `BAT`, `modules/modBattery_Canonical.bas`; `U`, `classes/Usuario.cls`; `US`, `classes/UsuarioServicio.cls`; `ES`, `classes/ExpedienteServicio.cls`; `WF`, `classes/WorkflowServicio.cls`.

## 1. Vínculos y conexiones

Condor abre **cinco** bases por ruta y las mantiene en globales de aplicación.

| Conexión | Uso | Obtención | Nº de llamadas en `src` |
|---|---|---|---|
| Condor | backend propio | `getdb()` (`FU:84-124`) | 155 |
| Lanzadera | identidad y expediente | `getdbLanzadera()` (`FU:64-83`) | 7 |
| Expedientes | expediente | `getdbExpedientes()` (`FU:156-190`) | 11 |
| No Conformidades | vínculo de la concesión | `getdbNoConformidades()` (`FU:191-210`) | 14 |
| Correos | registro de notificación | `getdbCorreo()` (`FU:136-155`) | 7 |

`CheckAndReconnect` decide la ruta de cada una y la abre con `Workspace.OpenDatabase`, componiendo la contraseña en la cadena de conexión (`FU:233-252`). Hay dos conmutadores de modo: si `Application.TempVars("DatosEnLocal")` vale «Sí» se abre la copia local de cada backend (`FU:235`, `FU:238`, `FU:240`, `FU:242`, `FU:244`) y si `EnPruebas` vale «Sí» se abre la base de pruebas de Condor (`FU:237`). `FUNCIONES UTILES` conserva además una routine que borra el contenido de la carpeta de datos locales (`FU:518-537`).

**En el destino.** Cada una de esas cinco conexiones se convierte en un enlace. La decisión no es solo técnica: el expediente, la identidad, las no conformidades y el registro de correo pertenecen a otras aplicaciones, así que el esquema global tiene que decidir si se replican, se consultan por servicio o se sustituyen. Este PRD no recomienda; deja la decisión en [tablas-enlazadas](matriz-campos/tablas-enlazadas.md).

**Sin equivalente en el destino:** los conmutadores de modo local y de pruebas. En el destino el destino de los datos es único y el entorno de pruebas tiene credenciales propias.

## 2. Adaptaciones de SQL y transacciones

| Rasgo de Access | Medida | Dónde | Adaptación |
|---|---|---|---|
| `Max(id) + 1` como clave primaria | 7 apariciones | hoja de módulos de datos | secuencia o identidad del destino |
| SQL concatenada en la construcción | 77 apariciones | hojas de módulos | parámetros ligados |
| `IIf` en SQL o en expresión | 31 apariciones | hojas de formularios y módulos | `CASE` o expresión del lenguaje de destino |
| `FIRST(...)` | 9 apariciones | hojas de módulos | `LIMIT` o agregación |
| `TOP 1` | en listados de rechazo, adjunto y revisión | hojas de módulos | `LIMIT 1` conservando el orden |
| `Nz(valor, "")` | 742 apariciones | todo el código | decisión explícita entre `NULL` y `''` |
| `InStrRev`, `Environ` | 7 y 5 apariciones | hojas de módulos | funciones del lenguaje de destino |
| `CreateQueryDef` efímero | 9 apariciones | hoja de módulos | consulta parametrizada real |
| `Recordset` devuelto al llamador | 94 firmas | hojas de módulos y clases | lista de objetos o recurso paginado |

**Transacciones.** Trece ficheros de producción abren transacción propia: repositorio de rechazo, servicios de datos de los tres bloques, servicio de adjuntos, servicio de solicitud, servicio de revisión y servicio de flujo, más cinco formularios. Hay 38 llamadas a `BeginTrans` y el mismo número a `CommitTrans`; la reversión se hace con `ws.Rollback`, nunca con `RollbackTrans`. El patrón consistente es «si el llamador no pasó conexión, la abre y la cierra yo», y su traducción al destino es una regla única: la unidad transaccional la decide el servidor.

**Consistencia de tipos.** El volcado del ERD mezcla `TEXT` con indicadores «Sí»/«No», `MEMO` para todo lo largo y `LONG` para claves que el código calcula a mano. La matriz de campos deja la asignación de columnas abierta a la espera del ERD de origen ([matriz-campos](matriz-campos.md)).

## 3. Reglas compartidas

Tres reglas cruzan aplicaciones y ninguna está consolidada:

1. **Identidad y rol.** `UsuarioServicio.DeterminarRol` traduce los indicadores de `TbUsuariosAplicacionesPermisos` a un rol con jerarquía administrador, calidad y técnico (`US:78-111`), y `PermisoSuficiente` aplica esa misma jerarquía a las transiciones (`WF:1410-1418`). Los indicadores son cadenas de texto en una tabla de Lanzadera, no banderas.
2. **Catálogo de estados y transiciones.** `tbEstados` y `tbTransiciones` son de Condor, pero el visor de estado y las precondiciones dependen de ellas (`WF:820-873`, `WF:1701-1733`). Si el ERD global los coloca en un esquema común, la regla pasa a ser compartida de facto.
3. **Expediente y RAC.** Condor lee el expediente de otra base y compone a partir de él el código de la solicitud (`ES:106-161`), y de ahí sale el RAC que se muestra en el visor.

## 4. Concurrencia y política de conflictos

La aplicación no tiene ninguna política de conflictos: no hay columna de versión, ni bloqueo optimista, ni marca de tiempo de modificación en la escritura de los bloques de datos. Las reglas de escritura que sí existen son:

- **Claves calculadas.** `Max(id) + 1` en adjuntos, logs, correos y bloques (`PC:80-86`, `ADJ:45-47`, `LOG:39`, `ERR:49`, `COR:18`, `SOL:107-114`): dos sesiones simultáneas pueden obtener la misma clave.
- **Transición de estado.** `EjecutarTransicion` comprueba precondiciones y luego escribe (`WF:1211-1409`), de modo que dos usuarios que actúen a la vez sobre la misma solicitud pueden pasar las precondiciones y transicionar ambas.
- **Rechazo y decisión de revisión.** El rechazo se desactiva y se inserta el nuevo en secuencia (`RECH:146`), lo que admite dos rechazos activos si las dos sesiones se cruzan.
- **Adjuntos.** Una etapa del flujo admite un adjunto, elegido por la fecha de subida más reciente (`ADJ:99`): dos subidas simultáneas dejan el último gana.

**Política propuesta, no decidida.** Que el estado de la solicitud lleve control de versión y que toda transición condicione la escritura a la versión leída; que la clave la asigne el destino; y que la unicidad de «un rechazo activo» y «un adjunto por etapa» se imponga con una restricción, no con código.

## 5. Seguridad y auditoría

- **Contraseñas en el código.** `GetPasswordDB` las compone al abrir la conexión (`FU:126-135`, `FU:252`) y la entidad `Usuario` transporta un campo de contraseña (`U:19`).
- **SQL con identificadores concatenados.** `RegistroExiste` construye la sentencia con el nombre de tabla y de campo (`FU:832`); `UsuarioRepositorio.getUsuario` arma el `WHERE` por unión de condiciones (`USR:41`).
- **Auditoría.** Tres tablas registran cambios, errores y transiciones de estado. `CondorError` se autoregistra al lanzarse (`CE:64-75`), de modo que cualquier excepción de la aplicación deja traza; la suplantación se guarda en dos de las tres tablas (`LC:21`, `LE:20`).
- **Datos personales.** La matriz marca por nombre las columnas personales y de credencial; es una heurística para priorizar, no una clasificación.
- **Salida de la aplicación.** Los documentos, los PDF y las instantáneas de depuración se escriben en rutas locales (`DOC:291-471`, `SH:176-259`), fuera de todo control de acceso.

## 6. Pruebas bidireccionales

Dos direcciones, y ninguna automatizada hoy:

**Del legacy a la web.** La batería de `modBattery_Canonical` contiene 161 pruebas unitarias del flujo, los repositorios y los datos (`BAT:476-7813`). Es la especificación ejecutable del comportamiento y el mejor insumo para validar el destino. Requiere un arnés que hoy no existe en la web.

**De la web al legacy.** Ninguna. Para una transición con dos sistemas vivos hace falta un contrato de lectura bidireccional sobre las entidades compartidas: expediente, identidad, no conformidad y registro de correo.

## 7. Plan de corte y reversión

**Corte propuesto:** por aplicación y por ruta, en este orden, que sale de la dependencia real entre ellas:

1. **Lecturas enlazadas primero.** Migrar la lectura de expediente e identidad a la fuente aprobada, dejando las escrituras propias intactas. Es el paso de menor riesgo porque solo lee.
2. **Bloques de datos.** Migrar `tbSolicitudes` y sus tres bloques con el servicio de flujo como consumidor único, manteniendo el legacy en lectura.
3. **Escritura entre bases.** Resolver el vínculo con No Conformidades, que hoy es una cadena con el código de la concesión.
4. **Transición de estado.** Migrar `WorkflowServicio` cuando los bloques ya estén en el destino, porque es el consumidor de todo lo anterior.
5. **Auditoría y adjuntos.** Por último, porque ninguna otra parte depende de ellos para operar.

**Reversión: pendiente, y es la decisión más urgente del documento.** Hoy no existe reversión. Si el destino falla a mitad de una transición, el estado de la solicitud queda en un backend y el log en el otro, y no hay operación que los vuelva a alinear. La reversión propuesta, que el propietario debe validar, tiene tres piezas: una bandera de corte por aplicación que devuelva el tráfico al legacy, una cola de operaciones replicadas desde el destino al legacy durante la ventana de convivencia, y un procedimiento de reconciliación que detecte y vuelva a alinear estados divergentes. Ninguna de las tres está implementada.

## 8. Riesgos que este documento deja abiertos

| ID | Riesgo | Por qué importa |
|---|---|---|
| P-CON-1 | Modelo de identidad compartido | Condor lee usuarios y permisos de Lanzadera; afecta a las siete aplicaciones del roadmap |
| P-CON-2 | Tratamiento del expediente | Condor compone el código de la solicitud a partir de un texto del expediente; cambiar ese formato invalida códigos emitidos |
| P-CON-3 | Vínculo con No Conformidades | Hoy es una cadena con el código de la concesión, escrito desde Condor en otra base |
| P-CON-4 | Reversión de corte | No existe; sin ella, un corte a mitad de una transición deja datos repartidos entre dos backends |
| P-CON-5 | Registro de correo | Condor registra y no envía: la frontera entre Correos y Condor está en el registro |
| P-CON-6 | Conservación del hash de versión | Los ciclos de validación de revisión compare hashes; el algoritmo debe sobrevivir al cambio |
| P-CON-7 | Modo local de datos | La copia local por aplicación y su routine de borrado no tienen equivalente y son datos fuera del control del repositorio |

Los siete se proponen como propuestas de issue en issues-propuestas (tanda siguiente) cuando esa hoja llegue.