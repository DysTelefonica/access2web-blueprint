# 04 · Integraciones y operación — Tablas vinculadas y backends

## Propósito

Inventario de los enlaces entre frontend y backend: `TbConfiguracionBackends`, `TbTablasAVincular`, helpers `getdb()` / `getdbLanzadera` / `getdbExpedientes` y rutas físicas observadas. Sirve de base para `07-migracion/` y para el descubrimiento de cada aplicación.

## Qué va aquí

- Por aplicación: tablas declaradas como vinculadas y su backend origen.
- Helpers de apertura, manejo de errores y rutas alternativas.
- Rutas observadas en checkout; nunca rutas de red inferidas.

## Estado del contenido

**Placeholder.** Se poblará con la información de los Lotes 1 a 8 de `exploration.md`.

## Fuentes de autoridad

1. `TbConfiguracionBackends` y `TbTablasAVincular` en cada frontend.
2. `C:\00repos\codigo\<app>\00_main` por aplicación.
3. Inspección Dysflow solo lectura.
4. Engram solo como contexto histórico.

## Reglas de evidencia

- Una tabla se declara `vinculada` solo si la fuente lo afirma; un enlace manual no es una FK.
- Las rutas se citan textualmente; no se normalizan en esta fase.
- Las configuraciones Dysflow no se modifican aquí.

## Checklist

- [ ] Cada tabla vinculada lleva su backend origen y su helper de apertura.
- [ ] Las rutas observadas se distinguen de las rutas resueltas en runtime.
- [ ] APAP y APAP_WEB no aparecen.

## Siguiente paso

Cruzar con `02-topologia-ecosistema/topologia-frontends-backends.md` para mantener una única fuente de verdad de rutas.
