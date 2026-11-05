# D93 — Dos esquemas por aplicación

- estado: adoptada
- fecha: 2026-11-05
- decide: usuario (dueño del producto)
- motivo: La reversión de una release web debe preservar las tablas compartidas que Access sigue utilizando.
- origen: Issue #1559

## Contexto

D92 establece que Access y la web comparten en PostgreSQL las tablas con forma legacy. Tras el corte irreversible de los backends a PostgreSQL, Access las vincula por ODBC. Borrar el esquema web no debe destruir esos datos ni los objetos de los que dependen.

## Decisión

Cada aplicación tiene dos esquemas en el PostgreSQL compartido: `<app>_legacy` para las tablas con forma legacy y `<app>` para lo que añade la web. Por ejemplo, `expedientes_legacy` y `expedientes`. La web lee y escribe directamente las tablas compartidas, sin sincronización entre copias.

D93 precisa D14, que establece PostgreSQL compartido y esquemas por aplicación, y D82, que exige migraciones compatibles mediante expand and contract. Mantiene la fuente de verdad compartida de D92.

## Alternativas descartadas

- **Un solo esquema por aplicación.** Una reversión con `DROP SCHEMA` podría borrar las tablas que Access sigue utilizando.

## Consecuencias

- Los vínculos ODBC de Access apuntan a las tablas de `<app>_legacy`.
- Las migraciones web solo pueden hacer cambios aditivos en `<app>_legacy`; nunca ejecutan `DROP` ni `ALTER` destructivos allí.
- La reversión de una release web es `DROP SCHEMA <app> CASCADE;` únicamente sobre el esquema web, nunca sobre `<app>_legacy`. Los objetos legacy no deben depender de objetos del esquema web: `CASCADE` también elimina dependencias fuera del esquema borrado.
- La refactorización final, tras el apagado completo del legacy, es el momento de retirar los esquemas `_legacy`.
