---
name: architecture-guardrails
description: 'Trigger: before implementing structural changes in `app/`, `tests/` o `openspec/`; when proposing a new architectural decision (D-<n>); when the IA arrives at this repository and before touching anything cross-cutting; when extending or contradicting an existing D-<n>. Companion skill de `documentation-alan-style`: esa prescribe cómo se escribe, esta prescribe qué se debe saber antes de escribir.'
license: Apache-2.0
metadata:
  author: ardelperal
  version: "0.1"
---

## Activation Contract

Load this skill **before** any of:

- Implementing a change that touches `app/`, `tests/`, `openspec/`, `docs/architecture.md`, `docs/03-aplicaciones/` o las epic por app.
- Proposing a new architectural decision (D-<n>).
- Extending, contradiciendo o reemplazando una D-<n> vigente.
- Creando un nuevo módulo en `app/src/modules/<nuevo>/`.
- Primera visita a este repositorio.

Do **not** load this skill for:

- Escribir prosa, plantillas o nombres de archivo (eso es `documentation-alan-style`).
- Preparar un PR que ya cumple la convention multi-app (`app/<slug>` + commit `fix(platform): ...`) — el gate ya pasó.
- Decisiones de producto no arquitectónicas (qué feature, qué UI) que viven en `openspec/changes/<change>/specs/`.
- Trabajo sobre el material de entrada Access en `data/` o `inputs/`: su contrato es «read-only replacement», gobernado por `dysflow-*`, no por esta skill.

## Hard Rules

1. **Front-door antes de tocar código.** Lea, en este orden: `AGENTS.md` → `CODEBASE-GUIDE.md` → `docs/architecture.md` → `CONTRIBUTING.md` → `docs/calidad-de-codigo-y-ci.md` → `openspec/changes/<change>/design.md` (si el cambio pertenece a un change vivo) → `docs/03-aplicaciones/<app>/epic.md` (si toca una app específica). Saltarse cualquiera de estos pasos deja a la IA operando contra arquitectura obsoleta.
2. **Toda D-<n> tiene estado.** `vigente` u `OBSOLETO`. Las obsoletas se reemplazan, no se duplican. Buscar en `docs/architecture.md` §Decisiones arquitectónicas D-<n> cross-cutting vigentes antes de proponer una nueva.
3. **Capas enforced por gate.** `scripts/check_layers.py` rechaza imports que violen `ROOT_PACKAGE = "app.src.modules"` con `ALLOWED_IMPORTS` y `PURE_LAYERS = {domain, ports, application}` (DA-1). Mover un adapter a la capa equivocada es un gate failure, no un estilo.
4. **Cross-cutting se promote, no se filtra.** Si una decisión afecta a más de una app o a la plataforma entera, no se queda en `app/src/modules/<app>/`: aparece en `docs/architecture.md` §Decisiones y se etiqueta `cross-cutting` en su issue. Las apps restantes la leen desde allí.
5. **Las D-<n> no se inventan.** Antes de crear una nueva, busque la vigente en `architecture.md`, proponga extensión con heredada visible, o marque el cambio como `##ABIERTO##` en `openspec/changes/<change>/design.md`. Las D obsoletas pasan a la sección `##OBSOLETO##` con el reemplazo visible.
6. **El walkthrough no se edita a mano sin codegraph.** `docs/03-aplicaciones/<app>/walkthrough-*.json` se regenera con dysflow + codegraph-vba. Cambios «manuales» quedan desincronizados con la realidad del binario.
7. **PR con PR.** Antes de commit, link a un issue aprobado que cite la D-<n> que implementa o la que extiende. Sin issue aprobado, el CI rechaza el PR (ver `CONTRIBUTING.md` §Issue-First).

## Decision Gates

| Situation | Action |
|---|---|
| Cambio toca una capa hexagonal | Abrir el change con `ROOT_PACKAGE` + `ALLOWED_IMPORTS` visibles en el issue; citar la DA-<n> que aplica. |
| Surge una decisión cross-cutting desde una app | Promover primero a `docs/architecture.md` §Decisiones con estado `vigente`; después extender la DA-<n> en el change; etiquetar la issue con `cross-cutting`. |
| Módulo nuevo en `app/src/modules/<nuevo>/` | Abrir issue, citar la D-<n> que justifica su existencia (D5 plataforma modular, D68 monolito modular), crear su `epic.md` y `walkthrough-*.json`. |
| Una D-<n> contradice el código | Issue con `label cross-cutting`, decision.md abierto en `##ABIERTO##`, no parchear código antes de documentar. |
| Walkthrough mismatches con codegraph | Regenerar con `dysflow` + `codegraph`; no editar el JSON a mano sin un `codegraph_summary` que lo respalde. |
| Auth, secret manager, secrets, cache cross-cutting | Consultar `architecture.md` §Auth, §Secret manager, §Caché antes de implementar; heredar DA-2..DA-13 si el change es Lanzadera MVP, o proponer nueva DA-<n> si es app nueva. |
| Conventional commit con `<app>` placeholder | Cerrar con el slug real (`docs(expedientes): ...`); el placeholder `(app)` se considera violación de convention. |

## Execution Steps

1. **Al llegar al repo (primera visita):** lea los 7 ficheros del front-door en orden. Construya mental-model de 30 segundos desde `architecture.md` §90-second mental model. Identifique la «sentence que organiza» y la «What this is NOT» tabla. Si el `architecture.md` no existe, **no cree código todavía**: abra issue para crearlo antes.
2. **Antes de implementar:** identifique la D-<n> (o DA-<n>) que aplica. Si no la encuentra en `architecture.md`, abra issue `##ABIERTO##` y bloquee el cambio hasta que el design.md la declare.
3. **Al proponer una decisión nueva:** redacte con el formato del catálogo — contexto (por qué), decisión (qué), trade-offs, reversibilidad. No abrevie.
4. **Al abrir PR:** el body cita la D-<n> que implementa o la que hereda. Etiqueta `app/<slug>` para app-scoped, `cross-cutting` si afecta a varias. Conecta con issue aprobado.
5. **Al cruzar un gate:** `scripts/check_*.py` falla por una razón concreta. Arregle la causa raíz; no esquive el gate (`|| true` está prohibido por `tests/test_ci_workflow.py`).
6. **Al cerrar:** si la decisión es nueva y cross-cutting, abra PR a `architecture.md` §Decisiones en el mismo change o en uno independiente.

## Output Contract

Antes de considerar cerrado cualquier trabajo que toque `app/`, `openspec/` o `docs/architecture.md`:

- [ ] El issue aprobado cita la D-<n> (o pide crear una).
- [ ] La PR abierta etiqueta `app/<slug>` o `cross-cutting` y referencia el issue con `Closes #N`.
- [ ] `scripts/check_layers.py` corre limpio sobre `app/src/`.
- [ ] Si la decisión es cross-cutting, está en `docs/architecture.md` §Decisiones D-<n> con estado `vigente` u `OBSOLETO`.
- [ ] Si promueve un cambio en la convention multi-app, `CONTRIBUTING.md` §Convención multi-app también lo refleja.
- [ ] El walkthrough JSON de la app (si aplica) se regenera con dysflow + codegraph; no está editado a mano.

## References

- `docs/architecture.md` — fuente de verdad única de la arquitectura del monorepo. Léalo primero.
- `CODEBASE-GUIDE.md` — overview, ownership, reading path raíz.
- `CONTRIBUTING.md` — workflow de contribución + convention multi-app + label system.
- `docs/calidad-de-codigo-y-ci.md` — gates de calidad (los 12 `check_*.py` + los 4 workflows).
- `openspec/changes/lanzadera-mvp/design.md` — DA-1..DA-13 vigentes mientras el change esté vivo.
- `docs/03-aplicaciones/<app>/epic.md` — por-app; una epic por cada una de las 8 apps legadas.
- `AGENTS.md` §Skills — tabla de skills obligatorias y cross-cutting.
- Skill hermana `documentation-alan-style` — tono y plantillas; esta skill no la sustituye.
