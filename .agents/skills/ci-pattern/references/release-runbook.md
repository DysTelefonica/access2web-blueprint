# Camino de release y de hotfix (HR-19, HR-20, HR-21)

Procedimiento del consumer paso a paso. El gobierno que lo acompaña está en
`assets/release-gate/` (los dos gates), `assets/templates/` (el workflow de
release y las notas) y este documento. Nada de esto vive en instrucciones
sueltas a los agentes.

## Quién lanza la release y con qué comandos

**Lanza un mantenedor del consumer, con su propia credencial.** Ni un agente ni
un workflow inventan el tag: el workflow publica, el mantenedor etiqueta.

```bash
# 1. Desde un checkout limpio y sincronizado con origin/main:
git fetch origin && git merge origin/main

# 2. Las notas del release, en el commit que se va a etiquetar (HR-19).
#    Plantilla: assets/templates/RELEASE-NOTES.md.tmpl → RELEASE-<TAG>.md

# 3. El tag anotado, sobre el commit de la rama por defecto. Nunca ligero.
git tag -a vX.Y.Z <sha> -m "vX.Y.Z"
git push origin refs/tags/vX.Y.Z
```

El `push` del tag dispara el workflow `release.yml`
(`assets/templates/.github/workflows/release.yml.tmpl`), que ejecuta en este
orden:

1. **`preflight`** — `check_release_tag.py` (semver estricto, anotado, versión
   mayor que la anterior, commit alcanzable desde la rama por defecto) y
   `check_release_ci.py` (el último run del workflow de CI para ese mismo
   commit terminó en `success`; sin runs para ese SHA el gate sale 2).
2. **`release`** — publica en el entorno protegido `release` la GitHub Release
   del tag, marcada *Latest*, leyendo las notas de `RELEASE-<TAG>.md`.
3. **`verify`** — comprueba que la release publicada es la del tag, no es un
   borrador y está marcada *Latest*, y que lo que quedó desplegado sirve el SHA
   del tag. Es comprobación de salida, no una batería de tests.

Si `preflight` falla, el tag se queda donde está: no hay publicación. El remedy
es siempre el mismo —arreglar en la rama por defecto, dejar el CI en verde
sobre el commit exacto que se quiere publicar y mover el tag— y volver a
lanzar. Un tag mal cortado no se borra y se rehace: se corta el siguiente.

## Procedimiento de hotfix (HR-20, HR-21)

1. **Issue.** Abra la issue canónica con `type:bug` y la evidencia del
   incidente. Sin issue aprobada no hay hotfix.
2. **Rama.** Rama con el prefijo que admite el gate de nombre de rama: hoy
   `fix/` (parámetro 1). **`hotfix/` lo rechaza** —no es una preferencia de
   estilo, el gate lo pone en rojo—; si su host aplica la regex por ruleset, el
   prefijo es el mismo.
3. **PR suelto.** PR contra la rama por defecto declarando **Chain Context
   1/1** (`chain:`, `position: 1/1`, `base:`, `depends-on: none`,
   `follow-up: none`, `starts-at:`, `ends-with:`, `review-budget:` y un
   diagrama con un único `📍`). HR-54 exige la sección también en un PR suelto,
   no solo en los encadenados.
4. **Merge con CI en verde.** El PR se mergea por el procedimiento normal, sin
   rutas paralelas ni un fix que viva solo en una rama.
5. **Gate del tag.** `check_release_tag.py` y `check_release_ci.py` corren en el
   `preflight` del workflow; para comprobarlo antes de cortar el tag, los dos
   gates se ejecutan en local contra el repositorio.
6. **Tag anotado y release.** La persona mantenedora corta el tag semver con el
   bump **PATCH** y lo empuja; el workflow publica.
7. **Post-mortem blameless** en la ruta de post-mortems del consumer
   (parámetro 9), con Timeline (UTC) / Impact / Root cause / What worked /
   What failed, causas de sistema y nunca personas. Cada action item se abre
   como issue con owner y se verifica que no queda ninguno solo en el documento.

## Política de suites: dónde corre cada una

Es una regla del patrón, no una preferencia:

- **Cada PR y cada push a `main`** corren las suites caras (E2E, integración,
  multiplataforma) que el consumer declara. Un PR no se mergea con la suite
  que cubre su área sin ejecutar.
- **Las más caras pueden reservarse** a la rama por defecto y a una ejecución
  nocturna (`schedule`), no a cada PR. Entonces su fallo se ve en `main` o en
  el nightly, no en el PR: es una decisión conscious y se escribe en la
  adopción del consumer, con lo que se ejecuta en el PR y lo que se ejecuta en
  `main`.
- **La release no repite ninguna suite.** El gate `check_release_ci.py` exige el
  CI en verde del commit exacto del tag; repetir ahí las suites costaría tiempo
  sin añadir señal, porque el CI ya midió esos mismos bytes. Si el CI del tag
  no está en verde, la release no sale (punto 1).

## Notas de release (HR-19, HR-21)

Se materializan como `RELEASE-<TAG>.md` en el commit etiquetado, desde
`assets/templates/RELEASE-NOTES.md.tmpl`, y el workflow las publica tal cual.
Las secciones son fijas: resumen (diciendo si es hotfix), Procedencia (versión
y SHA anterior y nuevo, y el enlace al run), Qué cambia, Cambios
incompatibles, Cómo actualizar y Qué se arregla (causa en una frase y enlace al
post-mortem o a la issue).

La causa raíz **no** vive en las notas: vive en su documento dedicado y las
notas solo la enlazan (HR-21).