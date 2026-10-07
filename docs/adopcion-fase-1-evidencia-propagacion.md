# Fase 1 — Evidencia de la corrida real del canal de propagación

Corrida real del propagador del catálogo `personal-skills` sobre este
repositorio, y su modo seco previo. Es la evidencia que declara
`phases."1".evidence.path` en `.github/ci-pattern-adoption.json`.

La auditoría del canal está en `docs/adopcion-fase-1-canal-de-propagacion.md`;
este documento registra la ejecución que faltaba para cerrar la fase.

## Procedencia de la herramienta que ejecutó la corrida

| Dato | Valor |
|---|---|
| PR de la herramienta | `DysTelefonica/team-skills#374` (rama `feat/372-localroot-run-scoped`, commit `c97cb26`) |
| Issue de la herramienta | `DysTelefonica/team-skills#372` |
| Cambio que habilita la corrida | `scripts/propagate-team-skills.ps1` y `scripts/doctor-team-skills.ps1` aceptan `-LocalRoot <ruta>`, que sustituye `local_layout.local_root` del registro **solo en esa corrida**; sin el parámetro el comportamiento no cambia y el `local_root` de la entrada del consumidor conserva la prioridad |
| Sincronización con `main` previa a la corrida final | `DysTelefonica/team-skills#368`: la escritura del manifiesto pierde el `[void]` que anulaba su tubería. La corrida anterior a esa sincronización dejó un `.team-skills.yaml` de 0 bytes |
| CI del PR | `smoke` y `unit suite` en verde sobre `c97cb26` |
| Pruebas de la herramienta | `testing/suites/propagate-contract/test-propagate-contract.sh` 123/123 (los cuatro casos de `-LocalRoot` y el del manifiesto de #368) |

La corrida se lanzó desde la rama del PR porque el parámetro todavía no está en
la rama por defecto de `team-skills`. El artefacto que produce (rama de flota y
PR en este repositorio) no depende de qué copia del script lo lanza más allá de
la resolución de la raíz local.

## Qué raíz resolvió

`Resolve-ConsumerPath` resolvió la rama «canónico» del layout: la entrada de
este consumidor declara `active_branch: main`, así que la rama
`<canónico>-worktrees/<rama activa>` no aplica.

```text
===> DysTelefonica/access2web-blueprint @ <raíz local>/access2web-blueprint
```

Las rutas del host que ejecutó la corrida se sustituyen aquí por
`<raíz local>` y `<clon del catálogo>`: son rutas de una máquina, no del
repositorio, y este documento no debe fijarlas.

## Las corridas

Se hicieron dos corridas completas (modo seco y corrida real) con el script en
`c6bf1ca` y se repitieron con `c97cb26`, que ya trae la sincronización con `main`
y la corrección del manifiesto (`#368`). Lo que sigue es la salida de la corrida
final; la anterior solo difiere en el manifiesto, que salió de 0 bytes.

### 1. Modo seco, antes de la corrida real

```text
$ pwsh -NoProfile -File scripts/propagate-team-skills.ps1 \
        -ConsumerId DysTelefonica/access2web-blueprint \
        -LocalRoot <raíz local> -DryRun
Catalogo: <clon del catálogo>
Registry: <clon del catálogo>/fleet/registry.json (commit c97cb26 @ rama feat/372-localroot-run-scoped)

===> DysTelefonica/access2web-blueprint @ <raíz local>/access2web-blueprint
WARNING:   refresh script no existe en <raíz local>/access2web-blueprint/scripts/
  skill-fleet/access2web-blueprint regenerada desde origin/main (descartados 1 commit(s) previos de la flota)
  [43 líneas «DRY: would materialize .agents/skills/<skill>/», una por skill]
  DRY: slice would change = True; skills would materialize = 43; legacy skills/ to remove = 0

=== Resumen ===

id                                 action  pr_url
--                                 ------  ------
DysTelefonica/access2web-blueprint dry_run slice_changed=True
EXIT=0
```

### 2. Corrida real

```text
$ pwsh -NoProfile -File scripts/propagate-team-skills.ps1 \
        -ConsumerId DysTelefonica/access2web-blueprint -LocalRoot <raíz local>
Catalogo: <clon del catálogo>
Registry: <clon del catálogo>/fleet/registry.json (commit c97cb26 @ rama feat/372-localroot-run-scoped)

===> DysTelefonica/access2web-blueprint @ <raíz local>/access2web-blueprint
WARNING:   refresh script no existe en <raíz local>/access2web-blueprint/scripts/
  PR ya existente, actualizada por el push: https://github.com/DysTelefonica/access2web-blueprint/pull/770

=== Resumen ===

id                                 action            pr_url
--                                 ------            ------
DysTelefonica/access2web-blueprint pushed_pr_updated https://github.com/DysTele
                                                     fonica/access2web-blueprin
                                                     t/pull/770
EXIT=0
```

Las 43 líneas de materialización omitidas en el bloque del modo seco son
idénticas entre sí salvo el nombre de la skill; el conteo del resumen las
cierra. La última línea de cada bloque es el código de salida del proceso, no
salida del script.

## Qué dejó la corrida en el destino, medido después

| Medida | Antes (`f5b584b`) | Después (`7be912a`) |
|---|---|---|
| Rama de flota `skill-fleet/access2web-blueprint` | `f5b584b` | `7be912a` (`chore(fleet): apply personal-skills @ c97cb26 …`) |
| PR del destino | #770 abierto | #770 ampliado por el push, sigue abierto |
| Diff contra `main` | 278 ficheros, +39.368 / −1 | 265 ficheros, +37.715 / −5 |
| `.agents/skills/` | 46 skills | 43 skills |
| `.team-skills.yaml` | 0 bytes | 1.419 bytes y 55 líneas, con `include_skills` y `slice_block_sha256` |
| `AGENTS.md`, bloque de slice | `@ v0092651` | `@ veaf74c8`, con el partial `web.md` vigente del catálogo |
| `skills/` legacy | 50 entradas | 50 entradas, intacto |

Lo que la corrida **no** hizo: no mergea, no escribió en `main`, no retiró el
árbol legacy `skills/` y no tocó ningún repositorio de `team-skills`.

## Cierre de la fase

```text
$ python3 <skill de team-skills>/assets/bin/ci-pattern adoption check --phase 1 .
adoption check: fases 0..1 en verde sobre .
EXIT=0
```

`phases."1".evidence.path` apunta a este documento, que existe y no está vacío,
que es lo que el gate de la fase 1 ejecuta.

## Defectos y seguimientos que esta evidencia no cierra

1. **El manifiesto de 0 bytes queda cerrado por la corrida final.** La corrida
   con `c6bf1ca`, anterior a la sincronización con `main`, reprodujo el defecto
   de `DysTelefonica/team-skills#366`: `.team-skills.yaml` de 0 bytes, porque el
   script lo escribía con un `[void]` que anulaba la tubería. Con `#368` en la
   rama, la corrida final lo escribe completo (1.419 bytes, 55 líneas, con
   `include_skills` y `slice_block_sha256`). El artefacto intermedio (`2a3b1d7`)
   conserva el manifiesto vacío; era un defecto de la herramienta, no del
   consumidor.
2. **El PR #770 no puede mergear tal cual.** 37.715 líneas contra un presupuesto
   de revisión de 400, sin `size-exception-reason:` en el cuerpo. La corrida
   amplía el problema en lugar de resolverlo: el tamaño es del artefacto, no de
   la ejecución.
3. **La migración v1→v2 sigue incompleta.** `.agents/skills/` tiene 43 skills y
   `skills/` 50; la corrida no retira el árbol legacy.
4. **El modo seco no es de solo lectura.** La corrida seca hizo `git fetch` y
   dejó el clon en la rama de flota; el clon se devolvió a `main` después de la
   corrida real. La revisión previa que pide el plan de validación de #777 no
   puede ser un `-DryRun` inocuo.
5. **`refresh-team-skills.ps1` no existe en este consumidor** (aviso del
   script, no error): la comparación de hashes previa a la materialización no
   corre aquí.
6. **La herramienta que produjo esta evidencia vive en un PR sin mergear** de
   `team-skills` (#374). Si ese PR cambia antes de mergear, la corrida debería
   repetirse.

## Ficheros y comandos de esta corrida

```bash
# el modo seco y la corrida real, tal cual, desde la rama del PR #374
pwsh -NoProfile -File scripts/propagate-team-skills.ps1 \
  -ConsumerId DysTelefonica/access2web-blueprint -LocalRoot <raíz local> -DryRun
pwsh -NoProfile -File scripts/propagate-team-skills.ps1 \
  -ConsumerId DysTelefonica/access2web-blueprint -LocalRoot <raíz local>

# lo medido después (solo lectura)
git log --oneline -1 origin/skill-fleet/access2web-blueprint      # 7be912a
git show origin/skill-fleet/access2web-blueprint:.team-skills.yaml | wc -c   # 1419
git ls-tree --name-only origin/skill-fleet/access2web-blueprint:.agents/skills | wc -l   # 43
git diff --shortstat main...origin/skill-fleet/access2web-blueprint          # 265 ficheros, +37.715/-5
```
