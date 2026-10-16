[← Volver al Codebase Guide](../CODEBASE-GUIDE.md)

# Lo que no está aquí

> Esta página rastrea `main`, que puede contener cambios sin liberar. Para la última versión liberada, véase la etiqueta correspondiente.

Esta página enumera lo que **NO existe** en este repo y cómo no simularlo. Si una capability no aparece en la tabla, considérela ausente: la ausencia documentada es un invariante, no un hueco que rellenar.

## Capacidades ausentes

| Capability | Estado en este repo | Cómo no inventarla |
|---|---|---|
| Réplica exacta de los binarios `.accdb` legados | No commiteada. `data/staging/<app>/` se R2-pullea con `scripts/setup-staging.ps1`; el staging no vive en el repo. | No añada `.accdb` a `data/`; el staging se mantiene reproducible desde R2. |
| Suite E2E ejecutada en PR | No corre en CI. La suite Playwright en `e2e/` corre en `.github/workflows/release.yml::e2e` contra la imagen publicada. | No promueva la suite `e2e` al job de PR; un fallo verde alargaría cada PR minutos. |
| Memoria persistente entre sesiones | No presente. No hay `engram`, `memory.json` ni estado de runtime commiteado. | No añada `engram`, `memory.json` ni `runtime/` al repo; el runbook cita fuentes externas. |
| Cola de merge | No presente. La rama `main` se protege con el agregador `required` (`scripts/check_required_jobs.py`); `merge_group` desapareció con el modelo gentle-ai (ver `docs/gates.md` §3 y `.github/workflows/ci.yml`). | No reintroduzca `merge_group` ni workflows de queue. |
| Migración one-shot por app | Sólo Lanzadera (MVP) y Expedientes (F04 scaffold) tienen código hexagonal en `app/src/modules/<app>/`. Las otras 6 apps viven sólo en `docs/03-aplicaciones/<app>/`. | No cree `app/src/modules/<otra-app>/` sin abrir una epic en `openspec/changes/<change>/`. |
| Dashboard SPA separado | No presente. No existe bundle JS de UI ni servidor HTTP en puerto separado. La superficie web es la app FastAPI única en `/auth`, `/admin`, `/health`. | No añada `dashboard/`, `ui/`, ni un servidor en otro puerto. Use `templates/` de `delivery/http/`. |
| gRPC / WebSocket persistente | No presente. Sólo hay SSE puntual en `/admin/presence/stream`. | No introduzca `grpc` ni un servidor WS separado; abra una epic. |
| Bindings legacy Access/VBA | No presente. Los bindings se generan con dysflow + codegraph-vba, viven en `scripts/prompt-dysflow-form-walkthrough.md` y se R2-pullean. | No commitee bindings nuevos en el repo; el contrato está en `docs/prompts/`. |

## Sources externos válidos

| Si necesita... | Vaya a |
|---|---|
| Comandos de Engram | `docs/engram.md` no existe en este repo; el mantenedor externo es Gentleman-Programming/engram. |
| Binarios legacy Access | `data/staging/<app>/` R2-pulled; el script `scripts/setup-staging.ps1` documenta el flujo. |
| Skills reutilizables | `skills/README.md` (copia internalizada); las originales viven en Gentleman-Programming. |
| Contratos de dysflow | `docs/prompts/prompt-ia-mantenedora-dysflow-round-*.md` son work-in-progress hasta merge upstream. |

## Fronteradocumental explícita

- **No describa capabilities ausentes como si estuvieran implementadas.** Esta página las marca ausentes a propósito; prometerlas en otra parte del repo es una inconsistencia que rompe la confianza del lector.
- **No copie endpoints, rutas ni nombres de archivos de repos externos.** Si necesita uno, declare su ausencia y apunte al owner externo.
- **No añada un fichero `memory.json`, `engram.db`, ni equivalentes.** La frontera está pinada por la guard `scripts/check_legacy_hashes.py` (DA-13).

## Lista de comprobación del contribuidor

- [ ] Antes de proponer una capability nueva, buscó en este repo y confirmó que no existe.
- [ ] Si la capability está en esta tabla como ausente, no la añadió sin abrir una epic que la retire de la tabla.
- [ ] Si la capability es externa, apuntó al owner correcto (Engram, dysflow, Gentleman-Programming) sin replicar el contrato aquí.

## Siguiente paso

Continúe con [Manual del mantenedor](maintainer-playbook.md) para usar las listas de comprobación por tipo de cambio antes de abrir o revisar una PR.