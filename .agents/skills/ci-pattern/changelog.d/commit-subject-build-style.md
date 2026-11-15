type: patch

**Qué.** HR-14 rechazaba los commits `build:` y `style:`, aunque el nombre de rama (`check_branch_name.py`) sí admite `build/` y `style/`.
**Causa.** El patrón por defecto de `check_pr_contract.py` y el de `pr-contract.policy.example.json`, que copian los consumidores, se quedaron con nueve tipos, mientras que el de rama tenía once.
**Arreglo.** Los dos patrones admiten `build` y `style`. Test `test_hr14_accepts_build_and_style_like_the_branch_name`, con el patrón por defecto y con el de la política de ejemplo.
**A partir de ahora.** Un tipo nuevo de commit entra a la vez en el patrón de rama, en el de commits y en la política de ejemplo. Los consumidores con su propio `conventional_commit_pattern` (cadete, APAP_WEB) lo actualizan en su `.github/pr-contract.policy.json`. Fase 0 de CI/CD, punto 0.1.
