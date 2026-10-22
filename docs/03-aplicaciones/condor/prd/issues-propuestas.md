# Issues propuestas del PRD de Condor

> **Commit fijado:** `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · **Issue:** #1075 · **Fase:** 6 del [roadmap](../../../07-migracion/roadmap-descubrimiento-prd.md) · **PRD:** [README](README.md)

Unidades implementables que se deducen de este PRD. **No se publican como issues hasta que el propietario cierre el PRD** y se decidas las tres decisiones de frontera (modelo de identidad, esquema global y reversión). Ninguna duplica una capacidad que ya cubra otra aplicación: donde la capacidad es de otro PRD, se enlaza a su issue.

Todas llevarían la etiqueta `app/condor` y `type:tech`, además de la de la aplicación afectada cuando la decisión sea compartida.

Abreviaturas de cita: `FU`, `modules/FUNCIONES UTILES.bas`; `ADJ`, `modules/AdjuntoRepositorio.bas`; `RECH`, `modules/RechazoRepositorio.bas`; `PC`, `modules/DatosPCRepositorio.bas`; `LOG`, `modules/LogCambioRepositorio.bas`; `ERR`, `modules/LogErrorRepositorio.bas`; `COR`, `modules/CorreoRepositorio.bas`; `SOL`, `modules/SolicitudRepositorio.bas`; `BAT`, `modules/modBattery_Canonical.bas`; `DOC`, `classes/DocumentoServicio.cls`; `SH`, `modules/SnapshotHelper.bas`; `AS`, `classes/AdjuntosServicio.cls`; `US`, `classes/UsuarioServicio.cls`; `WF`, `classes/WorkflowServicio.cls`; `SS`, `classes/SolicitudServicio.cls`; `ES`, `classes/ExpedienteServicio.cls`; `NC`, `classes/NoConformidad.cls`; `NCS`, `classes/NoConformidadServicio.cls`; `SIM`, `modules/modSimuladorNotificaciones.bas`; `VR`, `classes/ValidacionRevision.cls`.

| Propuesta | Capacidad | Evidencia | Depende de | Criterios de aceptación | Etiquetas |
|---|---|---|---|---|---|
| P-CON-1 Modelo de identidad compartido | C-CON-2 | `US:78-111`, `WF:1410-1418`, `tbUsuariosAplicacionesPermisos` de Lanzadera | Decisión de identidad del roadmap; PRD de Lanzadera (#1074) | Un usuario tiene una identidad única en el roadmap; el rol se resuelve en el servidor y no desde indicadores de texto | `app/condor`, `app/lanzadera` |
| P-CON-2 Clave estable del vínculo con no conformidades | C-CON-15 | `NC:19`, `NCS:71-105` | PRD de No Conformidades (#1080) | El vínculo entre una concesión y su no conformidad no depende de una cadena de código, y sobrevive a un cambio de código | `app/condor`, `app/no-conformidades` |
| P-CON-3 Clave de expediente estable | C-CON-14 | `ES:106-161`, `SS:397-425` | Decisión de esquema global; PRD de Expedientes (#1073) | El código de una solicitud se compone desde una clave de expediente estable, no desde un campo de texto cuyo formato pueda cambiar | `app/condor` |
| P-CON-4 Reversión del corte | — | [coexistencia](coexistencia.md) §7 | Decisión de reversión del roadmap | Existe una bandera de corte por aplicación, una cola de operaciones replicadas y un procedimiento de reconciliación que devuelve a alinear estados divergentes | `app/condor` |
| P-CON-5 Control de versión en la solicitud | C-CON-13 | `WF:1211-1409` | P-CON-4 | Dos transiciones simultáneas sobre la misma solicitud producen una sola transición efectiva y la otra se rechaza con mensaje | `app/condor` |
| P-CON-6 Claves primarias asignadas por el destino | todas | `PC:80-86`, `ADJ:45-47`, `LOG:39`, `ERR:49`, `COR:18`, `SOL:107-114` | Decisión de esquema global | Ninguna tabla de Condor calcula su clave con `Max(id) + 1`; dos inserciones simultáneas obtienen claves distintas | `app/condor` |
| P-CON-7 Unicidad del rechazo activo y del adjunto por etapa | C-CON-10, C-CON-11 | `RECH:146`, `RECH:179`, `ADJ:99` | Decisión de esquema global | La unicidad la impone una restricción del destino, no una secuencia de código | `app/condor` |
| P-CON-8 Almacenamiento de los ficheros adjuntos | C-CON-11 | `AS:17-131`, `AS:484-541` | Decisión de almacenamiento | Un adjunto es una fila más un fichero con clave estable; guardar y eliminar son idempotentes y no dejan huérfanos | `app/condor` |
| P-CON-9 Generación de documentos en el destino | C-CON-17 | `DOC:44`, `DOC:155`, `DOC:347`, `DOC:557`, `DOC:672` | Decisión de formato de documento | La plantilla deja de ser un fichero de Word del cliente y el documento y el PDF se generan en el destino con el mismo contenido | `app/condor` |
| P-CON-10 Conservación del hash de versión de la revisión | C-CON-9, C-CON-21 | `VR:23`, `SH:154-175`, `SH:664-696` | Decisión de algoritmo | Los ciclos abiertos antes de la migración siguen siendo verificables contra el destino con el mismo algoritmo de hash | `app/condor` |
| P-CON-11 Doble explícito de notificaciones para pruebas | C-CON-18 | `SIM:6-107` | — | El entorno de pruebas no envía correo real y las transiciones que notifican se pueden ejercitar de forma segura | `app/condor` |
| P-CON-12 Retirada del modo local de datos | C-CON-4 | `FU:235-244`, `FU:518-537` | P-CON-4 | Ningún conmutador de copia local existe en el destino y no queda ninguna carpeta de datos locales que limpiar | `app/condor` |
| P-CON-13 Batería de pruebas del legacy como contrato | C-CON-22 | `BAT:476-7813` | — | Las 161 pruebas unitarias del legacy tienen equivalente ejecutable contra la API del destino, con el mismo nombre de caso | `app/condor` |
| P-CON-14 Contrastar con el ERD de origen | — | `ERD/schema.sql` anterior al commit fijado; D-CON-5 | Ejecución del workflow del propietario en su máquina Windows | La matriz de campos deja de ser «candidata a nivel de tabla» y pasa a columna asignada, con claves, nulabilidad y relaciones confirmadas | `app/condor` |

## Orden de ejecución propuesto

1. **P-CON-14** primero, porque cambia la matriz de campos y con ella el resto de las decisiones de datos.
2. Después **P-CON-1**, **P-CON-2** y **P-CON-3**, que son decisiones compartidas y condicionan el esquema global.
3. A continuación **P-CON-4**, que sin reversión impide hacer el resto del trabajo con confianza.
4. Y luego el resto, agrupadas por bloque de datos, flujo y trazabilidad.

## Propuestas que este PRD **no** hace

- **No recomienda** el modelo de identidad ni el de reversión: son decisiones del propietario y este estudio las deja abiertas.
- **No propone** un tipo de solicitud nuevo. Si el producto necesita recuperar el tipo `PC_SUB` retirado (C-CON-7), es una capacidad añadida que hay que decidir antes de publicar nada.
- **No duplica** las capacidades de otros PRD: expediente (#1073), identidad (#1074), no conformidades (#1080).