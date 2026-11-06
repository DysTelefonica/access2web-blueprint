# Cierre de la fase 0

La fase 0 de la [épica](../epica-coexistencia-postgresql.md#fase-0-inventario) queda **hecha** el 2026-11-06: su criterio de terminado se cumple con la evidencia que enlaza este documento. Pasa a **revisada** cuando el propietario valide esa evidencia; hasta entonces no se aplica nada en el legacy (decisión 12).

## Criterio de terminado

| Criterio | Evidencia |
|---|---|
| Todo backend dentro del alcance tiene su lista de vínculos | [Mapa de propiedad](mapa-propiedad.md): 535 tablas propias y 127 vínculos de los 18 backends de la [copia de referencia](copia-referencia.md); [grafo de dependencias](grafo-dependencias.md) entre backends. |
| Toda tabla tiene clasificación | [Clasificación de las tablas](clasificacion-tablas.md): las 535 tablas de los backends y las 50 de los fronts, con los siete puntos para el propietario resueltos salvo la cola de correo de Solicitudes HPS (abajo). |
| Toda tabla tiene ficha | Una ficha por aplicación o backend en `fase-0/`; los datos de cada tabla en las [comprobaciones del grupo (a)](comprobaciones-grupo-a.md) (clave primaria, huérfanos, duplicados) y en el [coste de los cambios](coste-cambios.md) (columnas, tipos, reglas y valores por defecto). |
| Todo front tiene su inventario de patrones de riesgo | [Patrones de riesgo del VBA](patrones-riesgo.md): las trece aplicaciones con un solo método. |
| R-1 y R-2 constan como confirmadas o con sus excepciones | [Grafo de dependencias](grafo-dependencias.md): R-1 se confirma en los backends y se refuta en los fronts, con las excepciones medidas. |
| Coste de cada cambio estructural | [Coste de los cambios](coste-cambios.md): líneas de código por tabla y por columna, sumando todos los fronts. |

El trabajo de la fase 1 está en el [catálogo único de accionables](catalogo-accionables.md): 164 propuestas, cada una con la issue de su repositorio cuando ya existe.

## Lo que queda abierto

Nada de esto impide cerrar la fase 0, pero cada punto tiene que estar resuelto antes de la fase en que se usa.

| Punto | Qué falta | Antes de |
|---|---|---|
| Cola de correo de Solicitudes HPS | El propietario elige la cola común, pero la local tiene emisor propio en producción y recordatorios que la común no soporta: falta decidir si se pierden o si la cola común se amplía ([clasificación, punto 3](clasificacion-tablas.md#para-decidir-por-el-propietario)). | Fase 1 de Solicitudes HPS |
| Tablas de AGEDO que leen otras aplicaciones | AGEDO ya no se usa (decisión 27), así que nadie alimenta `TbDocumentos`, `TbDocumentosID` ni `TbDocExpProceso`: Gestión de Riesgos y AGEDYS leen datos congelados. Falta decidir si se conservan como histórico o si esas aplicaciones dejan de leerlas. | Fase 1 de Gestión de Riesgos y AGEDYS |
| Botón de AGEDO en la Lanzadera | `ComandoAGEDO` de `FormAplicacionesOficina` sigue abriendo AGEDO; con la decisión 27 sobra. Falta su issue. | Fase 1 de la Lanzadera |
| Aviso de Telegram | El aviso que lee `TbConexionesRegistro` solo existe en la rama `feat/telegram-notifier` de Automatización, sin desplegar. Si se publica, choca con la retirada del registro de conexiones ([Lanzadera#25](https://github.com/DysTelefonica/Lanzadera/issues/25)). | Fase 1 de la Lanzadera |
| Issues de fase 1 sin crear | `DEC-21` (`TbExpedientesAnualidades` única en Expedientes), `COR-C1` (`IDCorreo` autonumérico, decisión 22), `CLS-H1` (cuarentena de las 164 huérfanas, una issue por backend) y las propuestas del catálogo con «—». Se crean al empezar cada aplicación, en el [orden de la decisión 24](orden-cortes-web.md). | Fase 1 de cada aplicación |
| Hipótesis H-1 a H-4 y macros de datos | Se ensayan en Citrix ([hipótesis abiertas](../epica-coexistencia-postgresql.md#hipótesis-abiertas)); las macros de datos de las tablas no las expone DAO y se comprueban al exportar el esquema. | Fase 3 |

## Cómo se hizo

Todo en solo lectura: los backends se leyeron con DAO sobre copias temporales de la copia de referencia, verificada por SHA-256, y el código, sobre la rama base de cada aplicación (su RC1 cuando aún no ha pasado UAT). La historia de cada paso está en el registro de la issue #1581 (`odd/tasks/1581-fase0-descubrimiento-global.md`), y todo se reproduce con los scripts de `scripts/dysflow-erd/`.
