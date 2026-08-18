[← Back to repo](README.md)

# access2web-blueprint

> Estudio de cómo migrar los 8 aplicativos Access/VBA legados a una arquitectura web hexagonal (FastAPI + HTMX), ejecutado como blueprint abierto y MVP Lanzadera en este monorepo.

![build](https://img.shields.io/badge/build-passing-brightgreen) ![license](https://img.shields.io/badge/license-proprietary-blue) ![status](https://img.shields.io/badge/status-MVP%20Lanzadera-yellow) ![python](https://img.shields.io/badge/python-3.12%2B-blue)

## Quick start

```bash
# 1. Clonar y entrar
git clone https://github.com/DysTelefonica/access2web-blueprint && cd access2web-blueprint

# 2. Instalar las skills del proyecto (gentle-ai y engram ya deben estar)
bash scripts/install-skills.sh   # o install-skills.ps1 en Windows

# 3. Smoke test del runtime Access (vía Dysflow MCP)
#    Requiere un agente con Dysflow MCP conectado.
await tools.dysflow.get_capabilities({});
await tools.dysflow.list_objects({ outputMode: "summary" });
```

Si los tres comandos devuelven respuesta sin error, el repo está listo para trabajar.

## Documentation

| Doc | Audience | Read this when... |
|---|---|---|
| [AGENTS.md](AGENTS.md) | IAs | Necesita saber qué skill cargar antes de tocar código o docs |
| [DOCS.md](DOCS.md) | Humanos técnicos, IAs | Busca referencia técnica completa (8 apps, stack, decisiones) |
| [CODEBASE-GUIDE.md](CODEBASE-GUIDE.md) | Mantenedores | Necesita entender dónde vive la responsabilidad X |
| [CONTRIBUTING.md](CONTRIBUTING.md) | Contribuidores | Va a abrir un issue o un PR |
| [CHANGELOG.md](CHANGELOG.md) | Usuarios | Quiere saber qué cambió entre versiones |
| [docs/architecture.md](docs/architecture.md) | Arquitectos, IAs | Va a tocar código de plataforma o necesita una decisión D-<n> |
| [docs/calidad-de-codigo-y-ci.md](docs/calidad-de-codigo-y-ci.md) | Plataforma team | Arranca el MVP y necesita los 12 quality gates |
| [docs/03-aplicaciones/](docs/03-aplicaciones/) | Research, mantenedores | Necesita entender una app legacy concreta |
| [skills/README.md](skills/README.md) | Contribuidores nuevos | Necesita instalar las skills del proyecto |

## What this is

El repo es **self-contained** desde el 2026-08-06: las 8 apps en alcance tienen su frontend, backend, `src/` y `docs/` (cuando existe) copiados bajo `data/staging/<app>/`. Esto permite que cualquier IA o humano trabaje sobre el blueprint sin acceso a los repos originales `C:\00repos\codigo\00_<APP>\`.

| App | Snapshot |
|---|---|
| Condor | `data/staging/condor/` (default Dysflow) |
| HPS | `data/staging/hps/` |
| HPS Solicitudes | `data/staging/hps-solicitudes/` |
| BRASS | `data/staging/brass/` |
| Gestion de Riesgos | `data/staging/gestion-riesgos/` |
| No Conformidades | `data/staging/no-conformidades/` |
| Expedientes | `data/staging/expedientes/` |
| Lanzadera | `data/staging/lanzaderas/` |

Total: **404.45 MB** committeados. Detalle por app y herramientas en [`data/staging/README.md`](data/staging/README.md).

## Estructura del repo

```text
access2web-blueprint/
├── README.md                          # este archivo
├── AGENTS.md                          # índice de skills para IAs
├── CODEBASE-GUIDE.md                  # ownership + reading path para mantenedores
├── DOCS.md                            # technical reference raíz
├── CONTRIBUTING.md                    # workflow + label system + multi-app
├── CHANGELOG.md                       # cambios por versión
├── docs/
│   ├── 00-alcance-y-evidencia.md      # puerta de entrada — qué apps, qué fuentes
│   ├── 01-inventario-aplicaciones.md  # tabla de las 8 apps + snapshot
│   ├── 02-topologia-ecosistema/       # cómo se conectan
│   ├── 03-aplicaciones/<app>/         # discovery por app (epic + walkthroughs)
│   ├── 04-integraciones-y-operacion/  # batch, correo, rutas
│   ├── 05-capacidades/                # índice de capacidades
│   ├── 06-autorizacion-legacy-matriz.md
│   ├── 09-arquitectura-objetivo-y-principios.md
│   ├── AGENT-SETUP.md                 # setup por agente (Claude, OpenCode, …)
│   ├── architecture.md                # fuente de verdad única arquitectónica
│   ├── calidad-de-codigo-y-ci.md      # 12 quality gates + 4 workflows
│   ├── design/mockups/                # mockups Mistica autocontenidos
│   └── prompts/                       # reportes al mantenedor de dysflow
├── app/                                # MVP Lanzadera — Python hexagonal (desde 2026-08-09)
├── data/staging/<app>/                # snapshot self-contained (8 apps)
├── inputs/                             # evidencia cruda (no se commitea por defecto)
├── openspec/                           # SDD por app
├── scripts/                            # check_*.py + setup-staging + sync-to-r2
├── skills/                             # skills internalizadas del proyecto
├── .dysflow/project.json               # config Dysflow unificada
├── .codegraph-vba/                     # índice codegraph del repo raíz
└── .github/workflows/                  # ci + security + release
```

## Herramientas

- **Dysflow MCP** (`adapterVersion 2.36.x`): única vía canónica para source↔binary sync, SQL execution, test execution y form UI en Access. Cargue `dysflow-usage` + `dysflow-arnes` antes de tocar.
- **CodeGraph-VBA MCP**: code intelligence sobre los `.bas`/`.cls`/`.form.txt`/`.report.txt`. Cargue `codegraph-usage`. Cada app tiene su propio `.codegraph-vba/`.

## Next steps

1. **Lee [DOCS.md](DOCS.md)** si quiere el technical reference completo (8 apps, stack pinned, decisiones D1-D178).
2. **Lee [CODEBASE-GUIDE.md](CODEBASE-GUIDE.md)** si va a tocar código, abrir un PR o agregar una nueva app.
3. **Lee [CONTRIBUTING.md](CONTRIBUTING.md)** antes de abrir un issue; el workflow CI exige rama `<tipo>/<nº issue>-<slug>` y PR ≤ 400 líneas.
4. **Lee [docs/architecture.md](docs/architecture.md)** si va a implementar el MVP de plataforma; las decisiones D-<n> son vinculantes.
5. **Arranca con un walkthrough** de la app que le interese (`docs/03-aplicaciones/<app>/epic.md`).

## Estado del proyecto

- **Research**: cerrado para las 8 apps (épicas mergeadas).
- **Plataforma MVP**: Lanzadera en `app/src/modules/lanzadera/` (FastAPI + HTMX + Alembic + Argon2id).
- **Migración**: pending por app; Lanzadera es la primera, las otras 7 la consumen.
- **CI**: workflow `deterministic-quality-harness` v1.4 verde en `main`.

## License

Proprietary — Telefónica. El blueprint se publica para revisión interna; el código de Lanzadera está en fase MVP y aún no es estable para consumo externo.