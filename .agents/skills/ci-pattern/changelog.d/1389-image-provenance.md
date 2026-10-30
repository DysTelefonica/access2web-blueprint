type: minor

HR-83 (#1389, 5/5): en un consumer que construye y despliega la imagen fuera del CI, producción se despliega solo desde una imagen cuyo contenido casa con el manifiesto de un tag de release en la rama por defecto. El gate del tag (HR-19, HR-20) mide el tag y `deploy-evidence` (HR-10) los jobs de deploy del CI; ninguno medía que la imagen desplegada saliera de ese tag. En `DysTelefonica/cadete` producción ejecutó `wcdy/cadete:2026.07.29B`, construida desde un commit nunca subido mientras su tag apuntaba a otro commit de `main` (cadete#748, cadete#755).

La plantilla `release.yml.tmpl` gana el job `provenance`, que sube el manifiesto como asset de la GitHub Release solo si existe `.github/image-provenance.json`. El partial web recoge la invariante y el asset documenta parámetros, códigos de salida y la normalización nula y explícita en su `README.md`.

Tests: el job nuevo en `assets/templates/tests/test_release_workflow_template.py`. Aceptación real: el gate rechaza el tar de `2026.07.29B` (revisión nunca subida).
