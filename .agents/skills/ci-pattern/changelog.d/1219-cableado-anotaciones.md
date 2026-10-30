type: patch

El cableado documentado del job `required` en `personal/ardelperal/ci-pattern/assets/required-jobs/README.md` pasa a ejecutar el generador de anotaciones antes del gate y a pasar `--annotations` cuando el fichero existe (#1219). El paso necesita el token del workflow (`github.token`) y `permissions: checks: read`, sin PAT ni secretos nuevos. Sin el fichero, la causa dice `no annotation` y el veredicto no cambia. El test `tests/test_readme_wiring.py` ejecuta el paso tal como está escrito.
