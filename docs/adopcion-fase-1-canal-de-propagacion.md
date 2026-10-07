# Fase 1 — El canal de propagación, auditado en vivo

> **Qué es.** La auditoría del mecanismo real que propaga el catálogo de `DysTelefonica/team-skills` hacia este repositorio, con la salida literal de las dos ejecuciones intentadas hoy y el estado medido del destino.

> **Qué no es.** El cierre de la fase. La evidencia de la corrida vive en `docs/adopcion-fase-1-evidencia-propagacion.md`, y es lo que declara `phases."1".evidence.path`; este documento es la auditoría del mecanismo. Lo que sigue abierto es el artefacto del destino, no la fase: el PR #770 no puede mergear con ese tamaño.

## El mecanismo, leído del repositorio que lo aloja

| Fichero | Rol | Lee | Escribe |
|---|---|---|---|
| `fleet/registry.json` | Declaración del consumidor | — | — (entrada de este repositorio: `active_branch: main`, `bot_branch: skill-fleet/access2web-blueprint`, `primary_type: web`, sin `governance` declarado) |
| `scripts/propagate-team-skills.ps1` | La corrida real | El partial del tipo primario, el fragmento del consumidor, el partial de gobernanza (solo si el consumidor declara `governance`), el catálogo `skills/`, `fleet/governance-tiers.json` | El bloque de `AGENTS.md` entre marcadores, `.agents/skills/<n>/`, `.team-skills.yaml`, y el borrado de `skills/<n>/` solo para las skills que salen del catálogo; después commit y `git push --force-with-lease` a la rama de flota, y apertura o actualización del PR |
| `scripts/doctor-team-skills.ps1` | Comprobación de solo lectura | Registry, manifiesto, marcadores, catálogo | Nada. Códigos: 0 sincronizado, 1 deriva, 2 inconsistente, 3 error de entorno |
| `fleet/governance-tiers.json` | Canal de gobernanza | — | — (declara qué partial y qué skills llegan solo a un consumidor con `governance`; este repositorio no lo declara, así que no le llegan) |

El alcance real no incluye hooks: no existen `.githooks` ni `.husky`, y los espejos de runtime (`.claude/`, `.opencode/`, `.codex/`) los escribe el consumidor con `scripts/install-skills.sh`, no el propagador.

## Qué hay hoy en el destino, medido

- **El slice está en `main`.** `AGENTS.md` lleva el bloque entre `<!-- personal-skills:slice:access2web-blueprint @ v0092651 -->` y su cierre, con tres commits de flota (`@ vae17553` #755, `@ 4675cbe` #761, `@ f0e46d3` #767). El canal del slice funciona y está mergeado.
- **v1, mergeada, escribió `skills/`:** 50 entradas, de las cuales 4 son del consumidor (`README.md`, `architecture-guardrails`, `lanzadera-testing-strategy`, `maintainer-prompt-drafter`) y el resto vienen del catálogo.
- **v2, pendiente, escribe `.agents/skills/`:** la rama de flota `skill-fleet/access2web-blueprint` (PR #770, abierto desde el 2026-10-02) cambió 278 ficheros con 39.368 líneas en `f5b584b` y, tras las corridas de la evidencia, 265 ficheros con 37.715 líneas en `7be912a`: las skills de `.agents/skills/` (43), `AGENTS.md` y el `.team-skills.yaml`, que en `f5b584b` salía de 0 bytes y en `7be912a` sale escrito (1.419 bytes).
- **Migración incompleta.** Los dos destinos no contienen el mismo conjunto: de las skills del canal de gobernanza, `skills/` tiene 4 y `.agents/skills/` tiene 2. Mientras la migración no termine, el repositorio sostiene dos copias divergentes del catálogo.

## Las dos ejecuciones intentadas, con salida literal

```text
$ pwsh -NoProfile -File scripts/doctor-team-skills.ps1 -Json
Join-Path: ~/repos/team-skills/scripts/doctor-team-skills.ps1:55
  Cannot find drive. A drive with the name 'C' does not exist.
EXIT=1

$ pwsh -NoProfile -File scripts/propagate-team-skills.ps1 \
        -ConsumerId DysTelefonica/access2web-blueprint -DryRun
Catalogo: ~/repos/team-skills
Registry: ~/repos/team-skills/fleet/registry.json (commit 37765f4 @ rama main)
Join-Path: ~/repos/team-skills/scripts/propagate-team-skills.ps1:101
  Cannot find drive. A drive with the name 'C' does not exist.
EXIT=1
```

Causa: `fleet/registry.json` declara `local_layout.local_root: "C:\\00repos\\codigo"`, la entrada de este consumidor no declara `local_root` propio y el script no admitía sustitución por variable de entorno ni por parámetro. Tanto el doctor como el propagador resuelven el consumidor bajo esa raíz declarada y abortan. **La corrida real no era ejecutable desde este host con el script de la rama por defecto.**

Resuelto en `DysTelefonica/team-skills#372` (PR #374): los dos scripts aceptan `-LocalRoot <ruta>`, que sustituye la raíz del registro solo en esa corrida. Con ese parámetro, el modo seco y la corrida real terminan en 0 y dejan su salida en `docs/adopcion-fase-1-evidencia-propagacion.md`.

## Hallazgos

1. **El canal del slice es real, medido y mergeado.** Tres commits de flota sobre `AGENTS.md` y el bloque de marcadores en `main` con su versión de slice.
2. **La corrida real no resolvía en este host.** Sin `local_root` por consumidor no había forma declarada de apuntarla a un clon de Linux; la salida de arriba es el obstáculo completo, y no era un fallo del consumidor sino un dato de configuración del catálogo. Resuelto con el parámetro `-LocalRoot` (`team-skills#372`, PR #374), sin tocar el registro de la flota.
3. **El modo seco no es de solo lectura.** Antes de imprimir el plan, el script hace `git fetch`, `git checkout` de la rama de flota y `git add`/`git rm` en el índice (líneas 345-397 y 658-686); el commit y el push quedan después del corte. Un `-DryRun` sobre un clon de trabajo lo deja en la rama de flota con el índice tocado, así que la revisión previa que pide el plan de validación de #777 no es inocua.
4. **El manifiesto de la última corrida real sale vacío.** `.team-skills.yaml` mide 0 bytes en la rama de flota; el script escribe el manifiesto con `[void]$sb.ToString() | Out-File …`, donde `[void]` anula la tubería. Está reportado en `DysTelefonica/team-skills#366`.
5. **La última ejecución real no puede mergear tal cual.** PR #770: `review-budget` en rojo (39.368 líneas en `f5b584b` y 37.715 tras la corrida de la evidencia, contra un presupuesto de 400, sin `size-exception-reason:` en el cuerpo), `quality` en rojo y `required` en rojo. La causa del `quality` no consta en el log disponible; no se atribuye.
6. **El canal de gobernanza no llega a este repositorio.** Sin `governance` declarado, ni el partial de gobernanza ni sus skills se propagan. Es coherente con la decisión del operador de no declararlo en esta adopción, pero significa que la fase 5 no puede enlazar su documento generado desde un partial que hoy no existe aquí.

## Cómo se cerró la fase 1

La corrida real se ejecutó desde este host con el script de `DysTelefonica/team-skills#374`, que aporta `-LocalRoot <ruta>`: el parámetro sustituye la raíz del registro solo en esa corrida, así que el layout que comparten las corridas del mantenedor no cambia y este consumidor no declara `local_root` propio. La salida literal de las dos corridas, lo que dejó en el destino y los seguimientos abiertos están en `docs/adopcion-fase-1-evidencia-propagacion.md`, que es lo que declara `phases."1".evidence.path`.

## Ficheros y comandos de esta auditoría

```bash
# lo que se leyó (solo lectura)
sed -n '80,110p;340,400p;580,700p' scripts/propagate-team-skills.ps1
pwsh -NoProfile -File scripts/doctor-team-skills.ps1 -Json          # aborta: línea 55
pwsh -NoProfile -File scripts/propagate-team-skills.ps1 -DryRun     # aborta: línea 101
git show origin/skill-fleet/access2web-blueprint:.team-skills.yaml | wc -c   # 0
git ls-tree --name-only origin/skill-fleet/access2web-blueprint:.agents/skills | wc -l   # 46
gh pr checks 770
```

La corrida real dejó su evidencia en `docs/adopcion-fase-1-evidencia-propagacion.md`; este documento es la auditoría del mecanismo y aquel, el cierre de la fase.
