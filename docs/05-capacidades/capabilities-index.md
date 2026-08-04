# 05 · Capacidades — Índice

## Propósito

Índice maestro de las capacidades de negocio del ecosistema legacy. Sigue `access-vba-capability-docs`: cada capacidad se describe como contrato de comportamiento, con datos, dependencias, integraciones, excepciones, confianza y trazabilidad.

## Convenciones

- **Tier** de cada ficha: `critical` | `standard` | `minimal`.
- **Confianza** por hecho: `Verified-runtime` / `Verified-static` / `Intended` / `Likely` / `Divergent`, con fecha y evidencia.
- **Trazabilidad**: enlace a GH issue/PR cuando exista; nunca como fuente de verdad.
- **Plantilla**: `access-vba-capability-docs/assets/capability-doc-template.md`.
- **Listado de capacidades** (formato del índice): `apps · capacidad · tier · estado · fuente · dependencia principal · excepción`.

## Estado del contenido

**Placeholder.** Aún no se ha poblado ninguna ficha; la estructura se cierra con los Lotes 1 a 8 y se reconcilia en el Lote 9.

## Reglas del índice

- Una ficha por capacidad de negocio, no por formulario o módulo.
- Las fichas se enlazan desde `03-aplicaciones/<app>/` y desde la matriz de dependencias.
- APAP y APAP_WEB no aparecen en el índice ni en las fichas.

## Plantilla de fila (placeholder)

| Capacidad | Aplicación | Tier | Estado | Fuente | Dependencia principal | Excepciones |
|---|---|---|---|---|---|---|

## Checklist

- [ ] Cada capacidad tiene su ficha con contrato de comportamiento.
- [ ] Cada regla de negocio sin test se etiqueta `Verified-static` con su test pendiente.
- [ ] Las excepciones se referencian desde la matriz global y desde la ficha.

## Siguiente paso

Empezar el índice durante el Lote 1; cada lote posterior añade filas y fichas correspondientes.
