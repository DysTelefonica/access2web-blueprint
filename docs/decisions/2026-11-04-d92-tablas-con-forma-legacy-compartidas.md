# D92 — Durante la coexistencia, la web trabaja sobre las tablas con forma legacy que Access vincula por ODBC

- estado: adoptada
- fecha: 2026-11-04
- decide: usuario (dueño del producto)
- motivo: Con la sustitución de las aplicaciones de una en una y la reversión por aplicación, una sola fuente de verdad garantiza que el legacy siga funcionando y que la reversión sea posible.
- origen: Issue #1545, decisión 10 de la épica de coexistencia (`docs/07-migracion/epica-coexistencia-postgresql.md`); resuelve la P-EXP-3

## Contexto

La épica de coexistencia de la issue #1545 sustituye las aplicaciones legacy de una en una, y durante la transición Access y la web conviven sobre un único PostgreSQL. La P-EXP-3 preguntaba si la web trabaja sobre las tablas con forma legacy o sobre un esquema propio sincronizado.

Las opciones A, B y C, con sus costes, están analizadas en la [coexistencia de Expedientes](../03-aplicaciones/expedientes/prd/coexistencia.md), apartados 11 y 12.

## Decisión

Durante la coexistencia, la web trabaja directamente sobre las tablas con forma legacy en PostgreSQL, que Access vincula por ODBC (opción A). Hay una sola fuente de verdad y no existe esquema web propio con sincronización.

Cuando el legacy esté apagado por completo, el modelo se refactoriza a un esquema moderno para dejar de arrastrar la deuda técnica de las tablas legacy.

## Alternativas descartadas

- **B. Vistas con triggers `INSTEAD OF`.** Tablas modernas para la web y un esquema legacy de vistas editables para Access. Descartada.
- **C. Dos familias de esquemas con triggers bidireccionales.** Esquemas modernos para la web y un esquema legacy para Access, sincronizados en ambos sentidos. Era la dirección preferida del propietario; el 2026-11-04 confirmó de forma expresa que la opción A la sustituye.

## Consecuencias

- Durante la coexistencia, los cambios de esquema son solo aditivos: columnas nuevas que admiten nulos y tablas nuevas. Nunca cambia el significado de una columna o tabla existente.
- Las mejoras estructurales se hacen en el backend Access antes de trasladar cada backend a PostgreSQL.
- Los adaptadores de PostgreSQL y las migraciones de DDL requieren ajuste para apuntar a las tablas con forma legacy: #803, #804, #813 y las migraciones de identidad de la Lanzadera web `0001` y `0008`.
- El dominio, los casos de uso, la interfaz y la capa HTTP siguen siendo válidos gracias a la arquitectura hexagonal: el cambio queda en los adaptadores.
- La reversión de una release web nunca puede borrar tablas con forma legacy: `DROP SCHEMA <módulo> CASCADE` (DA-7 y D82 en `docs/architecture.md`) destruiría datos que Access sigue usando. Mientras dure la coexistencia, la reversión de la web se limita a lo que la web añadió, y DA-7 y D82 deben revisarse en una issue aparte.
