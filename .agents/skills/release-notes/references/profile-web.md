# Perfil web/contenedor

Repositorios que se despliegan como imagen de contenedor y entregan por PR (`cadete`, `access2web-blueprint`, `APAP_WEB`). El ejemplo de referencia es `DysTelefonica/cadete` `docs/releases/v1.2.1.md`.

## Fuentes de datos

| Dato | Fuente | Comando |
|---|---|---|
| PR del rango, cadenas y commits sin PR | `release_coverage.py collect` | `collect --repo-dir <clon> --from <anterior> --to <nuevo> --github <owner/name> --out range.json` |
| Qué cambia en cada PR | Cuerpo de la PR y sus commits | `gh pr view <N> --json title,body,commits` |
| Issue que cierra cada PR | `closingIssuesReferences` (ya en `range.json`) y la línea `Closes`/`cierra` del cuerpo | `gh issue view <N> --json title,body` |
| Antes y ahora | Diff de la PR y la prueba que falla sin el cambio | `gh pr diff <N>` |
| Migraciones | Ficheros nuevos de migraciones del rango | `git diff --name-status <anterior> <nuevo> -- <dir de migraciones>` |
| Configuración y dependencias | Ficheros de configuración, lockfiles | `git diff --stat <anterior> <nuevo> -- config/ composer.lock package-lock.json` |
| Pendiente | Issues abiertas citadas en las PR del rango | `gh issue list --state open --search "<términos>"` |

`collect` toma las PR de `gh pr list --search "updated:>=<fecha del tag anterior>"`. Una PR entra en el rango cuando su commit de cabeza o su merge están en `git rev-list <anterior>..<nuevo>`. Es directa si su merge está en la línea de primeros padres del tag nuevo; si no, `integrated_by` dice qué punta la integró.

## Procedencia

| Concepto | Valor |
|---|---|
| Release anterior | `<tag anterior>` → `<sha completo>` |
| Objetivo | `<tag nuevo>`, tag anotado en `main`; el commit del tag figura en la GitHub Release |
| Imagen | `<registro>/<imagen>:<tag>`, construida desde un clon limpio del tag; `org.opencontainers.image.revision` = commit del tag; el digest figura en la GitHub Release |
| Manifiesto de procedencia | Asset de la GitHub Release (`image-provenance-manifest.json` con `ci-pattern`) |
| CI de `main` y gate del tag | Run de CI sobre el SHA y salida literal de `check_release_tag.py` |
| Imagen de marcha atrás | `<registro>/<imagen>:<tag anterior>` |
| Migración de producción | Ruta del script y apartado del runbook |

## Con `ci-pattern`

La plantilla vive vendorizada en `.agents/skills/ci-pattern/assets/templates/RELEASE-NOTES.md.tmpl`; el workflow de release lee el fichero materializado (`--notes-file`). Conserve sus secciones y su orden. Los cuatro grupos de la skill ocupan el lugar de «Qué se corrige»; cada entrada lleva «En <anterior>», «En <nuevo>» y «Dónde», y la prueba que lo demuestra va en «Qué lo prueba». Añada «PR integradas» antes de «Registro de cierres» y «Pendiente tras esta versión» como subsección de este.

## Cómo desplegar

Siga el procedimiento del repositorio (en cadete, la skill `cadete-devops`). Pasos mínimos, cada uno con su comprobación: construir la imagen desde el tag; pasar el gate de procedencia contra la imagen en ejecución; anotar la imagen actual; cambiar la imagen; copia verificada de la base de datos antes de la migración; migración y su comprobación; pruebas tras el despliegue (salud, inicio de sesión, un flujo por grupo de cambios); marcha atrás con la imagen anterior y qué pasa con la migración.

## Exclusiones de la comprobación

La PR que añade las notas no puede citarse a sí misma antes de existir: `--exclude-pr <N> --exclude-issue <issue de las notas>`. En cadete v1.2.1 son `--exclude-pr 867 --exclude-issue 866`.
