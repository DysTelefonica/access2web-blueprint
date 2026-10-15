[← Volver al README](../README.md) · [← DOCS](../DOCS.md) · [← AGENTS](../AGENTS.md)

# access2web-blueprint — Codebase Guide

> **Frase que organiza todo el repo**: «Lanzadera es la madre: ahí nacen usuarios, aplicativos y permisos. Las otras 7 apps son consumidoras.»

Esta guía es el **índice corto** que orienta al mantenedor y al contribuidor nuevo hacia la página de `docs/codebase/` que necesitan. No replica el detalle; cada página siguiente lo posee.

## Para quién es esta guía

| Lector | Use esta guía para |
|---|---|
| Mantenedor nuevo | Construir el modelo mental del monorepo antes de tocar código. |
| Contribuidor de una app | Encontrar el módulo que owns el comportamiento que va a cambiar. |
| Revisor de PR | Comprobar que el cambio toca la capa correcta y respeta los invariantes. |
| IA (Claude, OpenCode, Gemini) | Localizar el paquete que owns el cambio antes de proponer código. |

## Modelo mental en 90 segundos

```text
Apps legadas (Access/VBA)         Este monorepo                          Código hexagonal por app
┌────────────────────┐         ┌──────────────────────────┐         ┌──────────────────────────┐
│ Lanzadera ★        │         │   docs/                   │         │  app/src/modules/         │
│ Expedientes        │         │     03-aplicaciones/      │         │   ├─ lanzadera/   (MVP)  │
│ HPS, HPS_Sol.      │  ──►    │     <app>/{epic,walk-     │  ──►    │   ├─ expedientes/ (plan) │
│ Brass              │ walk-   │     through,capabilities} │  impl.  │   └─ <futuras>           │
│ GR, NC, Condor     │ through │   CODEBASE-GUIDE +       │         └──────────────────────────┘
│                    │         │   docs/codebase/*.md      │                    ▲
└────────────────────┘         │   app/ (código web)       │                    │ (corte por módulo)
                               │   scripts/check_*.py      │         ┌──────────────────────────┐
                               │   .github/workflows/      │         │ PostgreSQL + S3 + (F2)   │
                               └──────────────────────────┘         │ Mistica + HTMX + Alpine  │
                                                                         └──────────────────────────┘
```

★ = Lanzadera es la «madre»: ahí nacen usuarios, aplicativos y permisos que las otras 7 apps consumen.

## Tabla de páginas

| Página | Trabajo |
|---|---|
| [Modelo mental](codebase/mental-model.md) | Qué es, qué no es y qué debe preservar este repo. |
| [Mapa del repositorio](codebase/repository-map.md) | Propiedad por paquete y regla de ubicación del código nuevo. |
| [Interfaces](codebase/interfaces.md) | Rutas HTTP, plantillas, CLI y scripts. |
| [Testing y CI](codebase/testing-and-ci.md) | Niveles PR / nightly / release y qué corre en cada uno. |
| [Lo que no está aquí](codebase/what-is-not-here.md) | Lo que NO existe en este repo y cómo no inventarlo. |
| [Manual del mantenedor](codebase/maintainer-playbook.md) | Listas de comprobación por tipo de cambio. |

## Camino de lectura

1. [Modelo mental](codebase/mental-model.md) — lea esto primero.
2. [Mapa del repositorio](codebase/repository-map.md) — antes de editar código.
3. [Interfaces](codebase/interfaces.md) — si el cambio es visible para el usuario.
4. [Testing y CI](codebase/testing-and-ci.md) — si el cambio toca CI, migraciones o código testeable.
5. [Lo que no está aquí](codebase/what-is-not-here.md) — si el cambio promete algo que no ve en el repo.
6. [Manual del mantenedor](codebase/maintainer-playbook.md) — antes de abrir o revisar un PR.

## Siguiente paso

Lea [Modelo mental](codebase/mental-model.md).
