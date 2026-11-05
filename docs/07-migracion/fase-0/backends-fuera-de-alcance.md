# Fase 0: backends de aplicaciones fuera del alcance

Issue #1638 (documento gemelo de [`backends-compartidos.md`](backends-compartidos.md)). Épica `docs/07-migracion/epica-coexistencia-postgresql.md` (fase 0).
Alcance: SICA, Control de cambios, AGEDO y Tareas. Sus aplicaciones **no entran** en la migración (decisión 6 de la épica y decisión del propietario sobre el catálogo), pero sus backends sostienen tablas que consumen aplicaciones que sí entran. Se inventarían como **dependencias**: para saber qué sobrevive, quién es su dueño real y qué está mal ubicado.
Método: ERD con `scripts/dysflow-erd/extract-erd.ps1` sobre copia temporal, en solo lectura. Recuentos de desarrollo.

## `SICA_datos`

| Tabla | Filas | Clave primaria | Índices |
|---|---|---|---|
| `TbAperturas` | 7 800 | **Ninguna** | Ninguno |
| `TbCierres` | 54 | **Ninguna** | Ninguno |
| `TbUltimaImportacion` | 1 | **Ninguna** | Ninguno |
| `TbArchivos` | 0 | **Ninguna** | Ninguno |

Cuatro tablas, ninguna con clave primaria ni un solo índice, y `TbAperturas` con 7 800 filas. Cero relaciones de aplicación. Ninguna aplicación dentro del alcance las consume, según el grafo de dependencias.

## `Control_Cambios_datos`

| Tabla | Filas | Clave primaria |
|---|---|---|
| `TbAplicacionesVersiones` | 1 | Sí |
| `TbVersionCambios` | 34 | Sí |
| `TbCambioDocumentos` | 19 | Sí |

Tres tablas con clave y **tres relaciones encadenadas**: `TbAplicacionesVersiones` → `TbVersionCambios` → `TbCambioDocumentos`. Es un control de versiones de cambios documentales, con la tabla de versiones de aplicación como raíz. Ninguna aplicación dentro del alcance lo consume, según el grafo.

## `AGEDO20_Datos`

66 tablas propias, 2 vínculos a `Lanzadera_Datos`, 12 relaciones, 1 consulta guardada. Ocho tablas **sin clave primaria**:

| Tabla | Filas | Índices |
|---|---|---|
| `Entradas` | 14 125 | ninguno |
| `Salidas` | 11 609 | ninguno |
| `TbDocExpProcesoRegistroTArchivos` | 810 | `IDProceso` |
| `TbCorreoAcciones` | 910 | `IDCorreo`, `IDIntegracion`, `IDProceso` |
| `TbAuxCodRepetidos` | 170 | `IDRepetido` |
| `TbMatClasificadoEntregaDetalleAux` | 22 | `IDEntrega`, `IDEntregaDetalle` |
| `TbCartas` | 20 | `IDCarta`, `PrimaryKey`, `SegundoKey` |
| `TbVersionCliente` | 1 | ninguno |

**Consumidores dentro del alcance:** Gestión de Riesgos vincula `TbDocumentos` (6 510 filas) y `TbDocumentosID` (3 280). AGEDYS Técnico vincula cuatro tablas más. Los dos vínculos a Lanzadera son `tbUsuarios` y `TbUsuariosAplicaciones`: **el primero es un homónimo de `TbUsuarios` de Lanzadera**, con la `b` en minúscula. Dos tablas de nombre parecido en backends distintos, con consumidores distintos.

## `Tareas_datos1`

| Propia | Filas | Clave primaria |
|---|---|---|
| `TbCorreosEnviados` | 1 620 | **Ninguna** |
| `TbTareas` | 10 | Sí |
| `TbTipoTareas` | 9 | Sí |
| `TbHTMLPorUsuario` | 0 | Sí |

Y **13 vínculos**: 8 a `AGEDYS_DATOS` (`TbDPDOfertaSuministrador`, `TbFacturasDetalle`, `TbNPedido`, `TbProyectos`, `TbSolicitudesOfertasPrevias`, `TbSuministradoresSAP`, `TbVisadoFacturas_Nueva`, `TbVisadosGenerales`), 2 a `Expedientes_datos` (`TbExpedientes`, `TbExpedientesResponsables`) y 3 a `Lanzadera_Datos` (`TbUsuariosAplicaciones`, `TbUsuariosAplicacionesPermisos`, `TbUsuariosAplicacionesTareas`).

Dos hallazgos:

- **`TbUsuariosAplicacionesTareas` está mal ubicada**: vive en `Lanzadera_Datos` y su único consumidor es Tareas. Por la regla de propiedad por uso del front propietario, la tabla pertenece a Tareas; lo que sobra es el vínculo en Lanzadera.
- **`TbCorreosEnviados` está duplicada**: Tareas tiene su propia tabla con 1 620 filas, sin clave, mientras `Correos_datos` tiene la suya con 35 394 filas. Dos tablas con el mismo nombre y el mismo propósito en backends distintos.

**No se ha localizado ningún front de Tareas.** El backend existe y enlaza 13 tablas de cuatro backends, pero no hay checkout de la aplicación ni entrada en el catálogo de la lanzadera. Es un backend sin aplicación conocida.

## Propuestas

Ninguna se aplica: la puerta de descubrimiento global sigue cerrada (#1576).

| Id | Propuesta | Grupo (fase 1) | Comprobación previa | Reversión |
|---|---|---|---|---|
| S1 | Localizar el front de Tareas o declarar el backend como material archivado. Sin esa respuesta, sus cuatro tablas propias no se pueden clasificar. | Imprescindible | Es una pregunta para el propietario: ¿existe la aplicación Tareas y por qué no está en el catálogo? | Sin cambios mientras no se sepa. |
| S2 | Fijar clave primaria en `TbAperturas` (7 800 filas) y en las otras tres tablas de SICA, que no tienen ninguna. | Imprescindible | Contar duplicados por la columna que identifique cada registro, sobre copia reciente de producción. | Quitar la clave. |
| S3 | Declarar `TbUsuariosAplicacionesTareas` como tabla de Tareas: retirar el vínculo de `Lanzadera_Datos` y crearla en el backend de Tareas, o documentar la excepción. | Imprescindible | Confirmar que ningún otro front la usa: la ficha de Lanzadera la marcó mal ubicada por el mismo motivo. | Restaurar el vínculo. |
| S4 | Resolver la duplicidad de `TbCorreosEnviados`: Tareas tiene su propia (1 620 filas, sin clave) y Correos tiene otra (35 394 filas, sin clave). | Imprescindible | Determinar si ambas registran el mismo hecho con destinatarios distintos o si una está abandonada. Comparar destinatarios y fechas. | Sin cambios mientras no se decida. |
| S5 | Fijar clave primaria en las ocho tablas sin clave de AGEDO, empezando por `Entradas` (14 125) y `Salidas` (11 609). | Imprescindible | Recuento de duplicados en la copia de producción. AGEDO está fuera del alcance de la aplicación, pero sus tablas las consume Gestión de Riesgos y AGEDYS Técnico, que sí están dentro. | Quitar la clave. |
| S6 | Renombrar el homónimo `tbUsuarios` de AGEDO para que no se confunda con `TbUsuarios` de Lanzadera, o documentar la diferencia. | Recomendable si es barato | Confirmar que ninguna consulta las usa indistintamente. | Renombrar de vuelta. |
| S7 | Declarar las relaciones de `Control_Cambios_datos` con integridad referencial: la cadena de tres tablas ya las cumple. | Recomendable si es barato | Consultas de huérfanos con 0 filas. | Quitar la relación. |
| S8 | Revisar si SICA y Control de cambios siguen teniendo consumidor. Ninguna aplicación dentro del alcance los usa, según el grafo. | Recomendable si es barato | Confirmar que ninguna aplicación del catálogo los abre. | Sin cambios. |

## Preguntas que solo el propietario puede responder

1. **Tareas**: ¿existe la aplicación? Si no existe, ¿sus cuatro tablas propias se archivan con su backend y sus 13 vínculos se pueden retirar?
2. **Correos**: ¿`TbCorreosEnviados` duplicada en Tareas y en Correos son dos hechos distintos o uno abandonado?
3. **AGEDO**: está fuera del alcance de la aplicación, pero `TbDocumentos` y `TbDocumentosID` las consume Gestión de Riesgos, que sí migra. ¿Se migran con AGEDO o se traspasan a Gestión de Riesgos?

## Lo que no se pudo verificar

1. Los vínculos en producción (`\\datoste\…`) no son alcanzables: los campos se leyeron del ERD, no del origen.
2. Los recuentos son de la copia de desarrollo.
3. No hay front de SICA, Control de cambios ni Tareas en `C:\00repos\codigo`, así que no se ha podido medir su uso en código: el veredicto de propiedad se apoya solo en el grafo de vínculos.
4. El SQL de la única consulta guardada de AGEDO no se ha leído.
