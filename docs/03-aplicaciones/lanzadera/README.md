# 03 · Lanzadera

## Propósito

Carpeta de descubrimiento de **Lanzadera**: origen de usuarios, permisos y catálogo de aplicaciones (`TbAplicaciones`). Aquí aterrizará el contenido del Lote 1 de `exploration.md`.

## Estado

- **Mapeo:** resuelto.
- **Dysflow:** `valid`; frontend y backend resueltos.
- **Inventario `list_objects`:** 87 elementos resumidos (cifra preliminar, no capacidad).
- **Backend observado:** 35 tablas; frontend contiene `TbConfiguracionBackends`.

## Lote asociado

Lote 1 de `exploration.md` — Lanzadera.

## Entregable previsto

1. Catálogo real de `TbAplicaciones` con `IDAplicacion`, `NombreCorto`, ejecutable, backend y comando.
2. Usuarios, roles y permisos (`TbUsuariosAplicaciones`, `TbUsuariosAplicacionesPermisos`).
3. Arranque, comandos VBA y aperturas hacia otras aplicaciones.
4. Formularios de administración, informes y tareas.
5. Dependencias hacia el resto del ecosistema y excepciones.

## Fuentes de autoridad

1. `C:\00repos\documentacion\OPENSPEC\00_LANZADERA`.
2. `C:\00repos\codigo\00_LANZADERA\00_main`.
3. Inspección Dysflow solo lectura.
4. Engram solo como contexto histórico.

## Reglas de evidencia

- Toda fila del catálogo cita módulo y nombre exacto.
- Cada permiso describe el rol y la tabla origen.
- Sin evidencia Dysflow, una afirmación no asciende a `Verified-runtime`.

## Checklist

- [ ] Carpeta poblada solo cuando se ejecute el Lote 1.
- [ ] No se mezclan permisos de entornos distintos sin marcarlo.
- [ ] No se nombra a APAP ni APAP_WEB.

## Siguiente paso

Esperar la aprobación del usuario para iniciar el Lote 1; poblar capacidades en `05-capacidades/` y hallazgos en `08-decisiones-y-preguntas-abiertas.md`.
