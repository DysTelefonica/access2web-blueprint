# 03 · No_Conformidades

## Propósito

Carpeta de descubrimiento de **No_Conformidades**: no conformidades de proyectos y auditorías, con su excepción documentada respecto a `Nemotecnico`.

## Estado

- **Mapeo:** resuelto.
- **Dysflow:** `valid` en `get_capabilities`, pero `list_objects` y `list_tables` **fallaron** con error de runtime no desglosado. Inventario runtime pendiente de diagnóstico.
- **Backend observado:** `NoConformidades_Datos.accdb`.
- **Frontend observado:** `NoConformidades.accdb`.
- **Documentación local ya disponible:** `00_No_Conformidades/docs/ERD/` y cambios OpenSpec previos.

## Lote asociado

Lote 5 de `exploration.md` — No Conformidades.

## Entregable previsto

1. Capacidades de NC de proyectos y de NC de auditorías, separadas.
2. Cachés, formularios, indicadores y documentos asociados.
3. Permisos y consultas a Expedientes; delimitación exacta de la **excepción `Nemotecnico`**.
4. Dependencias con Lanzadera (parciales) y Expedientes (parciales con excepción).

## Fuentes de autoridad

1. `C:\00repos\documentacion\OPENSPEC\00_No_Conformidades`.
2. `C:\00repos\codigo\00_NO_CONFORMIDADES\00_main`.
3. Inspección Dysflow solo lectura, **una vez diagnosticado el error de inventario**.
4. Engram solo como contexto histórico.

## Reglas de evidencia

- La **excepción `Nemotecnico`** se documenta como `Excepción` con su flujo exacto (formulario, evento, módulo); no se generaliza.
- No se afirma recuento de objetos hasta que Dysflow devuelva inventario normalizado.
- Las rutas de proyecto y de auditoría se mantienen como capítulos separados.

## Checklist

- [ ] Diagnóstico del fallo de inventario Dysflow registrado en `08-decisiones-y-preguntas-abiertas.md`.
- [ ] Excepción `Nemotecnico` delimitada por flujo, no por aplicación.
- [ ] APAP y APAP_WEB no aparecen.

## Siguiente paso

Resolver el fallo de inventario Dysflow antes de poblar esta carpeta; consultar con el usuario la puerta de acceso de solo lectura.
