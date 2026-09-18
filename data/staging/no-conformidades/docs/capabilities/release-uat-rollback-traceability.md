# Capacidad: trazabilidad UAT, release y rollback

> **Estado del documento — 2026-09-18.** Esta página reemplaza la versión
> previa (era Access/Dysflow) y queda alineada con el mecanismo vigente
> implementado en `.github/workflows/release.yml` y documentado en
> [`docs/11-releases.md`](../../../../../../docs/11-releases.md). Las
> referencias a tags `PRUEBAS-###`, a `openspec/REGRESSION-ANCHOR.md` y al
> repositorio externo en `C:\00repos\documentacion\OPENSPEC\` quedaron
> obsoletas y se han retirado. Si necesita el contrato canónico, consulte
> `docs/11-releases.md`; esta página documenta la *intención de la
> capacidad* (POR QUÉ existe y QUÉ trazabilidad aporta), no la mecánica
> del release (que vive en `release.yml`).

## §0 Identidad
- **ID de capacidad**: `CAP-RELEASE-UAT-ROLLBACK`
- **Tier**: standard
- **Estado**: active / capacidad documental-operativa
- **Source**: sdd (migrada desde la era Access/Dysflow)
- **Responsable / autoridad de producto**: Pendiente de confirmación — equipo de calidad / release
- **Última verificación**: 2026-09-18 mediante revisión documental; mecanismo vigente verificado contra `release.yml` y `docs/11-releases.md`
- **Confianza global**: `Verified-static` para la intención; la mecánica operativa la imponen `release.yml` y los gates definidos en `docs/11-releases.md`

## §1 Intención de negocio
- **Propósito**: Vincular cada capacidad con su flujo de aceptación (UAT), su publicación (release) y la posibilidad de revertirla (rollback) sin ambigüedad.
- **Usuarios / perfiles**: Equipo de calidad, release manager, revisores de PR y desarrolladores/agentes IA que cierran capacidades.
- **Problema que resuelve**: Evita afirmar «está hecho» con evidencias obsoletas o con commits no alcanzables desde `main`; ofrece la misma trazabilidad que la versión Access, pero sobre el release nativo de GitHub.
- **Valor de negocio / por qué existe**: Una release de calidad necesita trazabilidad desde petición → OpenSpec → código → prueba → UAT → producción → rollback. El portal Web hereda ese requisito; el mecanismo cambia, la intención no.
- **No-objetivos**: No ejecuta releases ni sustituye los gates de CI (`release-preflight.sh`, `require-ci-success.sh`, `cosign verify`, job `mutation`).
- **Origen de la intención**: Migración de la capacidad homónima desde el portal Access. El contrato vigente del release vive en [`docs/11-releases.md`](../../../../../../docs/11-releases.md).
- **Referencia de tracker de origen**: Issues fuente del grupo de release (etiqueta transversal `XCUT`) en el repositorio vivo.

## §2 Contrato de comportamiento

### Canales y trazabilidad
- **DADO** un cambio listo para UAT **CUANDO** se mergea por PR **ENTONCES** queda en `main` con un SHA único y queda enlazado a la issue que lo motiva.
- **DADO** un SHA de `main` **CUANDO** se taguea como `v1.4.0-rc.1` **ENTONCES** el workflow `release` ejecuta `preflight` y publica una imagen con digest inmutable asociada al tag.
- **DADO** un SHA candidato **CUANDO** se acepta con cosign **ENTONCES** se taguea el **mismo commit** como `v1.4.0` (sin recompilar) y se promueve el digest a producción.
- **DADO** un fallo en producción **CUANDO** se requiere rollback **ENTONCES** se redespliega el digest del tag anterior; `v1.4.0` queda substituida por `v1.3.0` sin operación de Git (el tag anterior es inmutable).

### Reglas de negocio
| ID regla | Enunciado vigente | Autoridad | Mecanismo |
|---|---|---|---|
| BR-REL-1 | Ninguna release se publica sin que el workflow `ci` esté en verde sobre el SHA exacto del tag. | `release.yml` → `preflight` → `scripts/require-ci-success.sh` | [`docs/11-releases.md`](../../../../../../docs/11-releases.md) §Gates |
| BR-REL-2 | El SHA candidato debe ser ancestro de `origin/main` en el momento del tag; el release no tolera commits no publicados. | `release.yml` → `preflight` → `scripts/release-preflight.sh` | `docs/11-releases.md` §Ciclo de publicación |
| BR-REL-3 | Cada tag `vX.Y.Z-rc.N` y `vX.Y.Z` es inmutable; no se reescribe. | Git tag annotation + `release.yml` publish step | `docs/11-releases.md` §Canales |
| BR-REL-4 | El job `mutation` debe haber corrido en verde sobre el SHA del tag antes del publish. | `ci.yml` job `mutation` (XCUT #703) | `docs/11-releases.md` §Gates (job mutation) |
| BR-REL-5 | Cada release queda firmada con cosign (identidad OIDC keyless de GitHub); el digest se verifica antes de aceptar. | `release.yml` → job `verify` | `docs/11-releases.md` §Verificación criptográfica |
| BR-REL-6 | El rollback no es una operación de Git: es un redespliegue del digest del tag anterior. | Operacional | `docs/11-releases.md` §Retroceso |

### Validaciones
- Toda capacidad cerrada debe tener enlace a su release (`v*` tag) y a su eventual rollback (`v*` tag anterior) trazable por SHA desde `main`.
