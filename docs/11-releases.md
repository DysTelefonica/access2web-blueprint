# Releases

Contrato de publicación del blueprint. Define qué se publica, quién lo autoriza
y qué comprobaciones lo impiden cuando no se cumple.

> «Los canales son versiones, no ramas. La rama es dónde se trabaja; el tag es
> qué se publica.»

## Qué es / qué no es

| Es | No es |
|---|---|
| El contrato de publicación y promoción de este repo. | Una guía de despliegue en producción. |
| La descripción de tres canales identificados por tag. | Un modelo de ramas de larga vida. |
| Un conjunto de gates ejecutables en `scripts/`. | Una recomendación revisable caso a caso. |

## Canales

`main` es el tronco único: siempre publicable, no siempre publicado. Producción
es un tag, nunca una rama.

| Canal | Identidad | Quién lo consume |
|---|---|---|
| Producción | `v1.3.0` | Usuarios finales. La imagen desplegada. |
| Candidata | `v1.4.0-rc.2` | Calidad y desarrollo durante la aceptación. |
| Desarrollo | `main` | El equipo. No se despliega fuera de entornos de desarrollo. |

Ramas vigentes: `feat/*`, `fix/*`, `ci/*`, `chore/*` y `docs/*`, todas de vida
corta y fusionadas a `main` mediante PR. Y `release/vX.Y.x`, que existe
únicamente mientras hay un hotfix en curso.

## Ciclo de publicación

1. **El trabajo entra en `main`.** Cada cambio por su PR. El trabajo incompleto
   entra desactivado, nunca acumulado en una rama paralela.
2. **Se corta la candidata.** `git tag -a v1.4.0-rc.1` sobre el commit exacto de
   `main`. Ese es el congelamiento. `main` continúa; la candidata no.
3. **Se construye una vez.** El workflow `release` construye la imagen desde el
   tag y la publica en el registro. El digest resultante es la identidad del
   artefacto.
4. **Aceptación sobre esa imagen.** El alcance se obtiene con
   `git log v1.3.0..v1.4.0-rc.1`: un rango entre dos puntos inmutables. Las
   actas de aceptación se generan de ese rango.
5. **Promoción.** Con la aceptación firmada, se etiqueta el **mismo commit**
   como `v1.4.0` y se despliega el **mismo digest**. Sin fusión y sin recompilar.

Un caso de aceptación fallido no se parchea sobre la candidata. La corrección
entra en `main` y se corta `rc.2`. Los números de candidata son gratuitos.

### Retroceso

Volver atrás no es una operación de Git, es un redespliegue: `v1.3.0` es un tag
inmutable cuyo digest sigue publicado. Se redespliega y termina.

Esto exige migraciones compatibles hacia atrás. Añada columnas; no las renombre
ni las elimine en la misma release. La limpieza va una release después.

## Hotfix con una candidata en curso

Producción en `v1.3.0`, `v1.4.0-rc.2` en aceptación, y aparece un fallo crítico
en producción.

1. **La corrección entra primero en `main`**, por PR normal. Así la próxima
   release no puede regresar el fallo.
2. **Se abre `release/v1.3.x`** desde el tag `v1.3.0`.
3. **Cherry-pick** de la corrección, tag `v1.3.1`, y despliegue a producción.
4. **Se corta `v1.4.0-rc.3`** desde `main`. La candidata anterior se descarta.

El orden importa. Corregir primero la rama de release y aplazar el cherry-pick a
`main` es la vía por la que el mismo fallo se publica dos veces. La única
excepción es una incidencia grave en curso; en ese caso se invierte el orden y
el cherry-pick a `main` se hace el mismo día.

La nueva candidata no repite la aceptación completa: valida la corrección y una
prueba de humo del resto.

## Gates

La regla y su cumplimiento son el mismo archivo. Un documento que dice «no
etiquete nada que no sea `main`» lo incumple cualquiera; un script que termina
con código distinto de cero, no.

| Gate | Dónde | Rechaza |
|---|---|---|
| `scripts/release-preflight.sh` | Local y workflow `release` | Tag sin anotar, versión no canónica, o commit que no es el `origin/main` actual. |
| `scripts/require-ci-success.sh` | Workflow `release` | Tag sobre un commit cuyo workflow `ci` no está en verde. |
| `pre-push` (`gentleai.gatedTags`) | Máquina local | Publicación de un tag estable. Es una barrera, no un cerrojo. |

### La regla de identidad

Es la comprobación que sustituye a una rama de preproducción:

- Una **candidata** debe ser el `origin/main` exacto en el momento del corte.
- Un **tag estable** debe ser el `origin/main` exacto, o promover una candidata
  de su misma versión sobre el commit exacto de esa candidata.

El segundo caso es el que permite que un tag estable quede por detrás de `main`:
el tronco siguió avanzando mientras la candidata estaba en aceptación.

Verificación manual antes de publicar:

```bash
git tag -a v1.4.0-rc.1 -m "Candidata 1.4.0"
./scripts/release-preflight.sh v1.4.0-rc.1
git push origin v1.4.0-rc.1
```

## Autorización

Las candidatas las corta el equipo. **Los tags estables los corta únicamente la
persona responsable**, y solo después de que Calidad valide los casos de
aceptación.

Es la misma garantía que ofrecía el modelo anterior de ramas: nada llega a
producción sin validación previa. Cambia el punto de control, no el control.

## Referencias

- `scripts/release-preflight.sh` — la regla de identidad, ejecutable.
- `scripts/require-ci-success.sh` — el gate de CI verde.
- `.github/workflows/release.yml` — construcción, publicación y verificación.
- `docs/00-alcance-y-evidencia.md` — alcance del proyecto.
- `CONTRIBUTING.md` — convenciones de rama, commit y PR.
