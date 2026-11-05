# Catálogo único de accionables: conflictos, bloqueos y decisiones

Documento gemelo de [`catalogo-accionables.md`](catalogo-accionables.md). Issue #1581, tarea T14.

## Conflictos entre propuestas

Una tabla tocada por propuestas de dos o más aplicaciones no se puede tratar en su ficha: o se resuelve una vez para todas, o se declaran incompatibles. Las tablas con más consumidores van primero.

| Tabla | Aplicaciones que la tocan | Consumidores (grafo) | Lectura |
|---|---|---|---|
| `TbUsuariosAplicaciones` | 2: Expedientes, Solicitudes HPS | 8 | 2 propuestas sobre la misma tabla |
| `TbCorreosEnviados` | 7: Condor, Correos y Registro de salida, Expedientes, Gestion de Riesgos, No Conformidades, SICA, Control de cambios, AGEDO y Tareas, Solicitudes HPS | 7 | 7 propuestas sobre la misma tabla |
| `TbExpedientes` | 3: Expedientes, Registros, Solicitudes HPS | 6 | 3 propuestas sobre la misma tabla |
| `TbNoConformidades` | 2: Expedientes, No Conformidades | 3 | 2 propuestas sobre la misma tabla |
| `TbAplicaciones` | 2: Gestion de Riesgos, Lanzadera | 2 | 2 propuestas sobre la misma tabla |
| `TbResponsablesExpedientes` | 2: AGEDYS Tecnico, Gestion de Riesgos | 2 | 2 propuestas sobre la misma tabla |
| `TbExpedientesSuministradores` | 2: HPS, Solicitudes HPS | 2 | 2 propuestas sobre la misma tabla |
| `TbExpedientesLugaresEjecucion` | 2: HPS, Solicitudes HPS | 2 | 2 propuestas sobre la misma tabla |
| `TbLog` | 4: AGEDYS Tecnico, Correos y Registro de salida, No Conformidades, Seguridad | no compartida | 4 propuestas sobre la misma tabla |
| `TbConfiguracionBackends` | 3: Condor, Gestion de Riesgos, No Conformidades | no compartida | 3 propuestas sobre la misma tabla |
| `TbExpedientes1` | 2: Expedientes, Gestion de Riesgos | no compartida | 2 propuestas sobre la misma tabla |
| `TbFacturasPrincipal` | 2: AGEDYS Tecnico, Expedientes | no compartida | 2 propuestas sobre la misma tabla |
| `TbExpedientesConEntidades` | 2: Expedientes, Gestion de Riesgos | no compartida | 2 propuestas sobre la misma tabla |
| `TbTipologia` | 2: Gestion de Riesgos, No Conformidades | no compartida | 2 propuestas sobre la misma tabla |
| `TbSalidas_RES` | 2: AGEDYS ECO, Correos y Registro de salida | no compartida | 2 propuestas sobre la misma tabla |
| `TbHerramientaDocAyuda` | 2: AGEDYS Tecnico, Lanzadera | no compartida | 2 propuestas sobre la misma tabla |
| `TbHPSEquivalencia` | 2: HPS, Solicitudes HPS | no compartida | 2 propuestas sobre la misma tabla |
| `TbUsuariosEntidades` | 2: HPS, Solicitudes HPS | no compartida | 2 propuestas sobre la misma tabla |
| `tbUsuarios` | 2: Registros, SICA, Control de cambios, AGEDO y Tareas | no compartida | 2 propuestas sobre la misma tabla |
| `TbConfiguracion` | 2: Registros, Solicitudes HPS | no compartida | 2 propuestas sobre la misma tabla |
| `TbUsuarios` | 2: SICA, Control de cambios, AGEDO y Tareas, Solicitudes HPS | no compartida | 2 propuestas sobre la misma tabla |
| `TbAuxExpPrado` | 2: AGEDYS ECO, AGEDYS Tecnico | no compartida | 2 propuestas sobre la misma tabla |

### Conflictos que no son de tabla sino de decisión

- **Cuarentena contra cambios solo aditivos.** Catorce propuestas proponen renombrar tablas a `zzBAJA_<tabla>`. La regla de la coexistencia dice que los cambios de esquema son solo aditivos. Renombrar es una excepción a esa regla y necesita una decisión expresa del propietario, no una propuesta de limpieza.
- **Vinculado y retirado.** Una propuesta retira un vínculo y otra repone la tabla en el mismo backend. Retirar un vínculo no retira la tabla de origen: renombra el acceso, no el dato.
- **Orden de migración contra catálogo.** El catálogo de la lanzadera y el grafo de dependencias dan órdenes distintos. El grafo manda: una tabla compartida obliga a coordinar a sus consumidores.

## Pruebas que bloquean

| Prueba pendiente | Qué desbloquea | Por qué falta |
|---|---|---|
| Recuento de filas y duplicados sobre una copia reciente de producción | Toda propuesta de clave primaria y toda cuarentena | Los recuentos disponibles son de copias de desarrollo y divergen entre sí: dos copias de `AGEDYS_DATOS` tienen 16 tablas con distinto recuento |
| Exportar el SQL de las consultas guardadas | Las propuestas sobre consultas: unas 64 consultas en seis backends | `MSysQueries` no es legible con el rol disponible; el SQL guardado en tablas (`TbConsultasDetalle`, `TbConsultasPrincipal`) sí se ha leído y se ejecuta en runtime |
| Volcar el catálogo `TbAplicaciones` | Saber qué aplicaciones existen de verdad | Declara 19 aplicaciones y la matriz de vínculos está desfasada: 6 tablas declaradas que ya no existen y 72 existentes sin declarar |
| Confirmar si existe la aplicación Tareas | Las propuestas sobre `Tareas_datos1` y `TbUsuariosAplicacionesTareas` | Existe el backend con 13 vínculos y no hay front ni entrada en el catálogo |
| Reproducir en runtime los recorridos afectados | Toda retirada de tabla o desplazamiento | Varias referencias ya probadas en runtime fallan hoy (`TbACParaLista`, `TbDPDDocAnexada`, `TbLog` sobre `CurrentDb()`) |

## Qué necesita decidir el propietario

Ordenado por lo que bloquea más cosas.

1. **Dueño de `TbCorreosEnviados`.** Una tabla, sin clave primaria, 35 394 filas, escrita desde al menos cinco aplicaciones y leída desde siete backends. Si Correos pasa a ser servicio de la plataforma, hay que decidir su esquema y su contrato antes de migrar nada que dependa de él; si sigue siendo un backend más, hay que declarar quién lo mantiene. *Recomendación*: tratarlo como servicio de la plataforma, porque es el único punto de concentración real del ecosistema y ninguna mejora estructural suya sobrevive a un modelo de tabla compartida si sigue siendo un `.accdb` más.
2. **Reparto de las 102 tablas de `AGEDYS_DATOS` entre las dos mitades de APGEDYS.** El recuento único da 31 compartidas, 12 solo de la parte técnica, 4 solo de la económica y 55 que no usa ninguno de los dos fronts. De ahí depende si APGEDYS necesita un backend común o dos, y qué tablas se retiran por no tener consumidor. *Recomendación*: un backend común, con las 55 sin consumidor tratadas por el protocolo de cuarentena tras el contraste con el grafo.
3. **Tareas queda fuera del alcance** (decisión del propietario, 2026-11-05). Sus cuatro tablas propias no las consume ninguna aplicación del alcance, así que no migran. Sigue abierto un efecto: `TbUsuariosAplicacionesTareas` vive en `Lanzadera_Datos`, que sí migra, y su único consumidor es Tareas.
4. **Duplicidad de `TbCorreosEnviados`**, con 1 620 filas en Tareas y 35 394 en Correos. Con Tareas fuera del alcance, la copia de Tareas no migra y deja de ser un conflicto de la migración.
5. **Tablas de AGEDO que consumen aplicaciones del alcance.** AGEDO queda fuera del alcance pero sigue en uso en Access (decisión 14). `TbDocumentos`, `TbDocumentosID` y `TbDocExpProceso` las consumen Gestión de Riesgos y AGEDYS. *Recomendación*: migrar esas tablas a PostgreSQL antes de que sus consumidores pasen a la web, con AGEDO vinculándolas por ODBC, como cualquier otro front Access.
6. **Cuarentena contra cambios solo aditivos.** Catorce propuestas renombran tablas. *Recomendación*: aprobar la excepción de forma explícita para la cuarentena, que es reversible y no cambia el significado, y mantener la regla para todo lo demás.
7. **Orden de migración.** El grafo de dependencias obliga a un orden distinto del catálogo de la lanzadera: una tabla compartida obliga a migrar a la vez a su dueño y a sus consumidores. *Recomendación*: derivar el orden del grafo y no del catálogo.
8. **Copia de producción.** Ninguna acción del grupo (a) puede ejecutarse sin una copia reciente de producción verificada, porque los recuentos de desarrollo divergen entre sí. *Recomendación*: fijar una copia de referencia con fecha y hash antes de la primera acción.
9. **AGEDO y Tareas como consumidores permanentes.** Siguen en uso en Access (decisión 14): cada corte de Lanzadera, Expedientes o AGEDYS debe reapuntar sus vínculos, ninguna tabla que consuman puede ponerse en cuarentena y la refactorización final no las alcanza mientras sigan vivas. Primer conflicto ya detectado: `LAN-A4` pone en cuarentena `tbUsuarios`, que AGEDO consume. *Recomendación*: sacar `tbUsuarios` de `LAN-A4`.

## Resumen por aplicación o backend

| Aplicación o backend | Propuestas | (a) | (b) | (c) |
|---|---|---|---|---|
| AGEDYS ECO | 18 | 11 | 5 | 2 |
| AGEDYS Tecnico | 16 | 9 | 6 | 1 |
| Brass | 6 | 6 | 0 | 0 |
| Condor | 9 | 5 | 4 | 0 |
| Correos y Registro de salida | 7 | 4 | 3 | 0 |
| Expedientes | 11 | 8 | 3 | 0 |
| Gestion de Riesgos | 10 | 8 | 2 | 0 |
| HPS | 12 | 6 | 6 | 0 |
| Lanzadera | 7 | 5 | 2 | 0 |
| Lanzadera (aperturas) | 4 | 4 | 0 | 0 |
| No Conformidades | 11 | 8 | 1 | 2 |
| Registros | 14 | 6 | 6 | 2 |
| SICA, Control de cambios, AGEDO y Tareas | 8 | 5 | 3 | 0 |
| Seguridad | 12 | 9 | 1 | 2 |
| Solicitudes HPS | 14 | 6 | 8 | 0 |
| **Total** | **159** | 100 | 50 | 9 |

## Lo que no se pudo verificar

1. El SQL de las consultas guardadas de seis backends, por permisos sobre `MSysQueries`.
2. El SQL efectivo de cinco consultas de Brass cuyo filtro se compone en variables.
3. Los recuentos de filas, que son de copias de desarrollo y divergen entre copias del mismo backend.
4. Si la aplicación Tareas existe, y si existen front de SICA o de Contratos fuera de los repositorios conocidos.
