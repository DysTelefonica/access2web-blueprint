# Perfil VBA/Access

Proyectos Access de la flota (`Expedientes`, `AGEDYS_ECO`, `seguridad`, `GESTION_RIESGOS`, `Lanzadera`, `HPS`, `No_conformidades`). Se distribuye un fichero, no una imagen: el front-end `.accdb` (o `.accde`) versionado en el repositorio, con el código exportado en `src/`. Parte del trabajo entra por PR a la rama `release/<versión>.rc<n>` y parte se empuja como commits directos; ninguno de los dos caminos puede quedar fuera de las notas.

## Fuente de verdad

1. **Commits del rango** — `release_coverage.py collect` lista las PR y, en `direct_commits`, los commits de la línea de primeros padres que no son el merge de una PR. Cada uno se cuenta en las notas por su issue o por su SHA corto.
2. **Issues** — las que cierran las PR (`closingIssuesReferences`) y las citadas en los mensajes de los commits directos.
3. **Registro de cambios del front-end** — `CHANGELOG.md`, `docs/release-*-novedades*.html` o el documento que el repositorio use. Si no tiene la entrada de la versión, dígalo en «Pendiente».
4. **Ficheros por módulo y formulario** — el diff de `src/` agrupa los cambios por objeto Access.

## Fuentes de datos

| Dato | Fuente | Comando |
|---|---|---|
| PR, cadenas y commits directos | `release_coverage.py collect` | `collect --repo-dir <clon> --from <anterior> --to <nuevo> --github <owner/name> --out range.json` |
| Objetos tocados | `src/modules/*.bas`, `src/classes/*.cls`, `src/forms/*.cls` y `*.form.txt`, `src/queries/` | `git diff --stat <anterior> <nuevo> -- src/` |
| Cambio real frente a exportación | Diff de cada objeto; descarte las líneas que solo regenera la exportación (`Attribute VB_Name`, `AutoResize`, línea en blanco final) | `git diff <anterior> <nuevo> -- src/forms/<Form>.cls` |
| Versión del front-end | `VersionAplicacion` del `.ini` del proyecto, o la constante de versión del código | `git show <nuevo>:<Proyecto>.ini` |
| Fichero distribuido y su hash | Front-end en el tag | `git show <nuevo>:<Proyecto>.accdb \| sha256sum` y `git cat-file -s <nuevo>:<Proyecto>.accdb` |
| Coherencia binario/fuente | `verify_code` de dysflow sobre el binario del tag | herramienta `verify_code` de dysflow |
| Migraciones del back-end | `db/migrations/*.sql` o scripts DDL del rango | `git diff --name-status <anterior> <nuevo> -- db/` |
| Pruebas | Manifiestos `tests/*.json` y átomos `Test_*` nuevos o cambiados; evidencia UAT en `docs/uat/` | `git diff --name-status <anterior> <nuevo> -- tests/ docs/uat/` |

## Agrupación

Los cuatro grupos de la skill se mantienen. «Dónde» nombra el formulario (`Form_<Nombre>`), el control o el módulo y el procedimiento (`modulo.Procedimiento`). Las regeneraciones de `src/` desde el binario, la configuración de dysflow y las skills propagadas van a «Repositorio, pruebas y CI». Un commit de depuración que se revierte dentro del rango se cuenta junto al commit que lo revierte.

## Procedencia

El esqueleto `assets/RELEASE-NOTES-vba.md.tmpl` la fija: versión anterior y su commit, commit del tag, rama integrada y UAT, front-end con su SHA-256 y tamaño, versión declarada por el front-end, coherencia binario/fuente, front-end de marcha atrás (el del tag anterior, con su SHA-256) y migraciones del back-end.

## Cómo desplegar

Pasos, cada uno con su comprobación: copia del front-end en uso y, si hay migración, copia verificada del back-end; aplicar la migración y comprobarla; distribuir el front-end del tag (por la Lanzadera o la ruta de distribución del proyecto) y comprobar su SHA-256 y la versión que muestra; pruebas tras el despliegue (abrir los formularios tocados, un flujo por grupo de cambios); marcha atrás con el front-end del tag anterior y qué pasa con la migración.
