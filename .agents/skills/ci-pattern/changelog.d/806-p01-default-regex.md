type: patch

El valor por defecto del parámetro P01 (`parameters.schema.json` y `ci-pattern.example.yaml`) deja de ser la regex anterior a #518, que exigía número de issue (`<tipo>/<número>-<slug>`), y pasa a la convención vigente `<tipo>/<slug>` sin número, la misma de APAP y BP. El documento operativo que genera `ci-pattern adoption generate-doc` deja de decir que el nombre lo construye un generador (HR-35, retirada) y dice quién aplica la regex: el ruleset del host o, sin rulesets, el gate de nombre de rama del patrón (#806).
