# ERD de origen: No Conformidades

> **Estado:** **pendiente del propietario.** No se ha extraído el ERD de origen; este documento registra qué se sabe, qué falta y cómo obtenerlo. Referencia: D-NC-3 del [README](README.md#discrepancias-abiertas).

## Qué hay y qué no

| Fuente | Contenido | Límite |
|---|---|---|
| [ERD/schema.sql](../ERD/schema.sql) (estudio previo, snapshot de 2026-08-06) | Volcado de `NoConformidades_Datos.accdb`: tablas, columnas y tipos con longitud | Sin claves primarias ni candidatas, índices, relaciones, nulabilidad ni valores por defecto. Declara 49 tablas; el DDL trae 42 |
| `src` de `main` (commit fijado) | Uso de tablas por formularios, módulos y clases | Solo da relaciones *inferidas por código* |
| Clon de solo lectura del legacy | Frontend `NoConformidades.accdb` | El backend de datos no está en el repositorio y su ruta no consta en el clon; el fichero de configuración del repositorio contiene un secreto y no se cita |

## Tablas con evidencia

Las 42 tablas del DDL, sus 428 columnas, su uso en `main` y las claves y relaciones inferidas por código están en la matriz campo a campo: el fichero `matriz-campos.md` y sus seis partes, que llegan en la quinta tanda de la matriz del PRD (#1284). Resumen: 181 filas con destino propuesto, 247 pendientes (sobre todo por tablas sin uso en `main`) y 3 columnas ausentes del DDL.

## Relaciones

Ninguna está **declarada** (no hay metadatos). Todas las de la matriz son **inferidas por código**, cada una con su cita. No se inventa ninguna otra.

## Extracción que falta

Ejecutar en la máquina Windows del propietario, en solo lectura y con `dysflow`, sobre `NoConformidades_Datos.accdb`:

1. Lista de tablas (reales, vinculadas y de sistema) para resolver la diferencia 49 frente a 42.
2. Por tabla: claves primarias y candidatas, índices, nulabilidad, valores por defecto, longitud y precisión.
3. Relaciones declaradas con reglas de actualización y borrado en cascada.
4. Confirmar las columnas que `src` usa y el DDL no trae (`IDNoConformidadCalculada` en `TbNCDocumentos`; `FechaUltimoUso`, `HitsConsultas` y `TamanioBytes` en la caché por NC).

El workflow de extracción `workflow_dispatch` sobre el runner Windows con etiqueta `dysflow-erd` está en la PR #1144 (abierta). Cuando esté en `main`, se usará con la ruta del backend, que falta y queda anotada como pendiente en [#1080](https://github.com/DysTelefonica/access2web-blueprint/issues/1080).

## Dependencias

Tablas compartidas con otras aplicaciones: `TbRiesgos*` (Gestión de Riesgos) y usuarios y permisos (Lanzadera). El ERD objetivo global no puede cerrarse sin este documento y los de las otras siete aplicaciones.
