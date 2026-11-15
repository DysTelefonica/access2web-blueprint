# Orden de paso de las aplicaciones a la web

Decisión 24 de la [épica](../epica-coexistencia-postgresql.md), aprobada por el propietario el 2026-11-05. Resuelve el punto 7 del [catálogo de conflictos](catalogo-accionables-conflictos.md#qué-necesita-decidir-el-propietario).

## Por qué los datos ya no marcan el orden

Con el traslado único (decisión 19), todos los backends pasan a PostgreSQL a la vez, antes del primer corte. Desde ese momento, una aplicación que pasa a la web no arrastra a ninguna otra: las que siguen en Access leen y escriben las mismas tablas por ODBC. El grafo de dependencias ya no obliga a un orden; lo marcan lo que se aprende en cada corte, el riesgo y el acoplamiento entre aplicaciones que se escriben tablas mutuamente.

## Orden

| Orden | Aplicación | Motivo |
|---|---|---|
| 0 | Lanzadera web | Es la plataforma: usuarios y permisos de todas las aplicaciones. Existe un MVP y la decisión 12 la pone primero. |
| 1 | Expedientes | Primera aplicación de negocio (fase 5), con plan de coexistencia propio. Tiene muchos consumidores, pero con los datos compartidos no la bloquean. |
| 2 | Condor | Sin consumidores documentados y sin controles conectados en diseño; solo escribe en la cola de correo. Consolida el método con poco riesgo. |
| 3 | Seguridad | Ninguna otra aplicación consume sus tablas; solo escribe en la cola de correo. |
| 4 | HPS | Va seguida de Solicitudes HPS, que escribe cuatro de sus tablas. |
| 5 | Solicitudes HPS | Junto a HPS, las escrituras sobre `TbUsuarios`, `TbUsuariosEntidades`, `TbUsuariosHistoricos` y `TbMotivoHPS` se rediseñan una sola vez. |
| 6 | Gestión de Riesgos | Va seguida de No Conformidades: se escriben tablas mutuamente (`TbRiesgosNC`, `TbNoConformidades`). |
| 7 | No Conformidades | Ver Gestión de Riesgos. |
| 8 | Registros | Front pequeño, pero con muchos controles conectados. Va antes que APGEDYS porque AGEDYS ECO escribe en su `TbSalidas_RES`. |
| 9 | Brass | Aislada en datos, pero es la de más interfaz acoplada a los datos: la más cara de rehacer. |
| 10 | APGEDYS | La más grande, y además funde AGEDYS Técnico y AGEDYS ECO en una sola aplicación (decisión 13). |

## Cifras de partida

Proceden de las fichas de la fase 0; «no consta» significa que la ficha no da el dato, no que sea cero.

| Aplicación | Formularios | Tablas propias | Escribe en tablas de | Controles conectados | Propuestas del grupo (a) |
|---|---:|---:|---|---|---:|
| Lanzadera | no consta | 35 | Correos | no consta | 9 |
| Expedientes | 45 | 47 | Gestión de Riesgos, No Conformidades, Correos | 1 | 8 |
| Condor | 45 | 15 | Correos | 0 en diseño (SQL asignado al ejecutar) | 5 |
| Seguridad | 54 | 36 | Correos | 5 formularios | 9 |
| HPS | 36 | 18 | Correos | 3 formularios | 6 |
| Solicitudes HPS | 34 | 11 | HPS (4 tablas), Correos | 0 | 6 |
| Gestión de Riesgos | 61 | 64 | No Conformidades, AGEDO, Correos | 1 | 8 |
| No Conformidades | 49 | 36 | Gestión de Riesgos, Correos | 0 | 8 |
| Registros | 11 | 11 | Lanzadera, Correos | 37 en 7 formularios | 6 |
| Brass | no consta | 57 | Lanzadera (alcance no demostrado) | 106 en 41 formularios | 6 |
| APGEDYS | 45 + 40 | 102 (compartidas) | Registros, AGEDO, Correos | 2 + 0 | 9 + 11 |

## Límites

- Varias fichas no dan el número de módulos ni de clases; el orden se basa en los datos disponibles y se revisa si un corte destapa un tamaño distinto.
- El grafo mide vínculos, no todos los accesos directos; una escritura cruzada no documentada puede obligar a agrupar dos cortes más.
- Las propuestas del grupo (a) se aplican en el saneamiento (fase 1), antes del traslado único; su número no es trabajo pendiente de cada corte.
