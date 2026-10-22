# XCUT #705 — docs: resolver drift entre release-uat-rollback-traceability.md y el mecanismo real de release.yml

## Outcome

Reescribir `data/staging/no-conformidades/docs/capabilities/release-uat-rollback-traceability.md` para reflejar el mecanismo real vigente de release/UAT/rollback, documentado en `release.yml` y `docs/11-releases.md`. Eliminar los conceptos de la era Access/Dysflow (tags `PRUEBAS-###`, `openspec/REGRESSION-ANCHOR.md`, refs a `C:\00repos\documentacion\OPENSPEC\00_No_Conformidades`).

## Branch

`docs/705-release-uat-rollback-drift` desde `main @ 8c21d82`.

## Budget

Cambio de un solo archivo docs. ≤50 líneas netas esperadas.

## Estrategia

Opción (a) del issue: reescribir el contenido del doc preservando la metadata de "Capacidad: trazabilidad UAT, release y rollback" pero reemplazando las reglas documentales (`Intended`, `Verified-static`) con referencia a las reglas vigentes.

## Plantilla de reescritura

El doc reescrito debe:
- Preservar §0 Identidad (ID de capacidad, Tier) — son metadata de capacidad.
- Reemplazar §1 Intención de negocio → versión web: trazabilidad desde PR → merge a main → tag `v*` → imagen firmada → release preflight.
- Reemplazar §2 Contrato de comportamiento → referenciar `docs/11-releases.md` ciclo de publicación.
- Reemplazar §3 Reglas de negocio → referenciar `docs/11-releases.md` Gates (release-preflight, require-ci-success, cosign verify, mutation gate).
- Eliminar refs a `openspec/REGRESSION-ANCHOR.md`, `PRUEBAS-###`, `C:\00repos\documentacion\OPENSPEC\00_No_Conformidades`.
- Mantener el formato de secciones (markdown) compatible con `data/staging/no-conformidades/docs/capabilities/_template.md`.
- Idioma: castellano peninsular formal (per `documentation-alan-style`).

## Verificación local pre-PR

```bash
python3 -c "import yaml; doc = yaml.safe_load(open('.github/dependabot.yml'))"  # sanity
python3 scripts/check_branch_name.py
python3 scripts/check_pr_size.py --base-ref main  # 1 archivo ≤ 50 lineas
```

(Este PR no toca app/ ni tests/, así que pytest/ruff/mypy/layers/dry no aplican.)

## Cierre

Commit único con mensaje Conventional, push, abrir PR. Closes #705.

## Estado

- [ ] worktree creado desde main @ 8c21d82
- [ ] doc reescrito
- [ ] tamaño ≤ 50 líneas netas
- [ ] commit + push + PR
