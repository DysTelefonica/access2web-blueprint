# image-provenance — la imagen desplegada es la de su release, byte a byte

Asset de `ci-pattern` (HR-83, DysTelefonica/team-skills#1389). Obligatorio en los
consumers web que **construyen y despliegan la imagen fuera del CI** porque el
registro (Quay tras VPN con doble factor) no es alcanzable desde el CI.

El gate del tag (`assets/release-gate/`, HR-19 y HR-20) mide el **tag**; el gate
`deploy-evidence` (HR-10) mide los jobs de deploy **del CI**. Ninguno mide que la
imagen que llega a producción se construyó desde ese tag. En `DysTelefonica/cadete`
producción ejecutó durante tres meses `wcdy/cadete:2026.07.29B`, construida desde
una copia local con un commit nunca subido (`30981a52`), mientras su tag apuntaba a
un commit de `main` que solo traía el Dockerfile y las notas (cadete#748,
investigación cadete#755). Cuando un hotfix posterior se construyó desde `main`, las
funciones de esa imagen desaparecieron de producción.

## Las tres piezas

| Pieza | Dónde corre | Qué hace |
|---|---|---|
| `image_provenance.py manifest` | workflow de release (CI), tras los gates del tag | sha256 de cada fichero del árbol desplegable **en el commit del tag**, leído de los objetos git; se sube como asset de la GitHub Release |
| `build_release_image.py` | máquina de quien construye | clon limpio del tag desde el remoto; rechaza tag ausente, fuera de `main` o árbol que no es el contenido git del tag; etiqueta la imagen con `org.opencontainers.image.revision` = commit del tag y `org.opencontainers.image.version` = tag |
| `image_provenance.py gate` | antes de `oc set image` | lee el árbol de la imagen, lo compara con el manifiesto de su release y comprueba la procedencia; imprime además el diff frente a la imagen que corre en el clúster |

Códigos de salida comunes: `0` correcto · `1` hallazgos (cada fichero nombrado) ·
`2` no medible (falla cerrado: archivo de imagen ilegible, sin árbol desplegable, sin
tag de release, manifiesto no disponible sin otros hallazgos, herramienta caída).

## Normalización: ninguna, y explícita

El manifiesto guarda el sha256 de los **bytes del blob git** en el commit del tag
(`git ls-tree` + `git cat-file`), nunca de un checkout. La imagen tiene que llevar
exactamente esos bytes. El script de build clona con `core.autocrlf=false` por
invocación (nunca lo escribe en ninguna configuración) y rechaza el build si el
checkout del árbol desplegable difiere del contenido git, por ejemplo por una regla
`eol=crlf` de `.gitattributes`.

Un fichero que solo difiere en finales de línea **es un hallazgo**: sale como
`changed in line endings only (CRLF in image)` para que la causa se vea, sin ocultar
que la imagen no es la release. La 29B de cadete tiene 4991 ficheros así (una copia
de Windows con `core.autocrlf=true`) además de 249 cambios reales. No se comparan
permisos ni propietario: el build puede hacer `chmod`/`chown`, no cambiar contenido.
Los enlaces simbólicos se comparan por destino.

## Qué rechaza el gate

| Hallazgo | Qué pasó |
|---|---|
| `provenance: image has no …revision label` | la imagen no salió del script de build |
| `provenance: image revision … is not a commit of the remote` | se construyó desde un commit nunca subido (el caso 29B) |
| `provenance: image revision … is not the commit of tag …` | el tag apunta a otro commit |
| `provenance: release tag … does not exist on the remote` | no hay release para esa imagen |
| `provenance: tag … is not on 'main'` | release fuera de la rama por defecto (HR-20) |
| `content: missing in image` / `extra in image` / `changed` | el árbol de la imagen no es el del manifiesto |

La procedencia se mide sobre un clon `--bare` fresco del remoto, no sobre una copia
local. El manifiesto se descarga de la release con `gh release download`; la
exclusión aplicada es la que registra el propio manifiesto.

## Parámetros por consumer

Un único fichero, p. ej. `.github/image-provenance.json` (ejemplo completo en
`image-provenance.example.json`):

| Clave | Obligatoria | Uso |
|---|---|---|
| `tree` | sí | árbol desplegable en el repositorio (`www`) |
| `image_path` | sí | dónde vive ese árbol en la imagen (`/var/www/html`) |
| `image` | sí | registro y nombre de la imagen, sin tag |
| `github_repo` | sí | `owner/name`: remoto git y releases |
| `exclude` | no | directorios instalados en el build (`vendor/`), globs relativos a `tree` |
| `default_branch` | no | por defecto `main` |
| `containerfile` | no | por defecto `Dockerfile` |
| `openshift` | no | `context`, `namespace`, `deployment`, `container` para `--current-from-cluster` |
| `engine_cmd`, `oc_cmd`, `gh_cmd` | no | ejecutables como lista (por defecto `podman`, `oc`, `gh`) |

## Uso

```bash
# CI (plantilla release.yml.tmpl, job `provenance`): lo hace el workflow.
python3 image_provenance.py manifest --config .github/image-provenance.json --tag v1.0.3 --out image-provenance-manifest.json

# Build: siempre desde el tag publicado, nunca desde la copia de trabajo.
python3 build_release_image.py --config .github/image-provenance.json --tag v1.0.3 --push

# Antes de `oc set image`:
python3 image_provenance.py gate --config .github/image-provenance.json \
  --image quay.example/wcdy/cadete:v1.0.3 --current-from-cluster
```

`--image-archive <tar>` lee un `podman save`/`docker save` (docker-archive u
oci-archive) sin motor de contenedores; `--manifest <fichero>` sustituye la
descarga; `--current-image` o `--current-archive` fijan la imagen de comparación.

## Riesgo residual

Quien tenga permisos de `oc` puede saltarse el script. El gate convierte el salto en
un acto explícito, no en un accidente; la política de admisión de OpenShift (solo
imágenes firmadas) es el seguimiento de infraestructura.

## Tests

`tests/test_manifest.py`, `tests/test_image_tree.py`, `tests/test_gate.py` y
`tests/test_build.py`: repositorios git y tarballs de imagen sintéticos, sin red ni
motor de contenedores.
