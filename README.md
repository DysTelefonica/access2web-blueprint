# access2web-blueprint

Estudio de cómo migrar los aplicativos Access/VBA legacy a una arquitectura web.

## Foundation (2026-08-06)

El repo es ahora **self-contained**: las 8 apps en alcance tienen su frontend, backend, `src/` y `docs/` (cuando existe) copiados bajo `data/staging/<app>/`. Esto permite que cualquier IA o humano trabaje sobre el blueprint sin tener acceso a los repos originales `C:\00repos\codigo\00_<APP>\`.

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

Total: **404.45 MB** committeados a git. Detalle por app y herramientas en [`data/staging/README.md`](data/staging/README.md).

## Estructura del repo

```
access2web-blueprint/
├── README.md                          # este archivo
├── docs/                              # discovery, topología, capacidades, migración
│   ├── 00-alcance-y-evidencia.md      # puerta de entrada — qué apps, qué fuentes
│   ├── 01-inventario-aplicaciones.md  # tabla de las 8 apps + snapshot
│   ├── 02-topologia-ecosistema/       # cómo se conectan
│   ├── 03-aplicaciones/<app>/         # discovery por app
│   ├── 04-integraciones-y-operacion/  # batch, correo, rutas
│   ├── 05-capacidades/                # índice de capacidades
│   ├── 06-autorizacion-legacy-matriz.md
│   ├── 06-seguridad-y-trazabilidad.md
│   ├── 07-migracion/                  # legacy → web
│   ├── 08-decisiones-y-preguntas-abiertas.md
│   └── 09-arquitectura-objetivo-y-principios.md
├── data/
│   └── staging/<app>/                 # snapshot self-contained (8 apps)
├── inputs/
│   └── automatizaciones-legacy/       # evidencia cruda (no se commitea)
├── openspec/                          # SDD por app
├── .dysflow/project.json              # config Dysflow unificada (default = Condor)
├── .codegraph-vba/                    # índice codegraph del repo raíz
└── .gitignore
```

## Herramientas

- **Dysflow MCP** (`adapterVersion 2.36.0`): única vía canónica para source↔binary sync, SQL execution, test execution y form UI en Access. Cargar `dysflow-usage` + `dysflow-arnes` antes de tocar.
- **CodeGraph-VBA MCP**: code intelligence sobre los `.bas`/`.cls`/`.form.txt`/`.report.txt`. Cargar `codegraph-usage`. Cada app tiene su propio `.codegraph-vba/`.

## Smoke test

```js
// runtime responde
await tools.dysflow.get_capabilities({});
// objetos del frontend default (Condor)
await tools.dysflow.list_objects({ outputMode: "summary" });
// symbol en el codegraph de condor
await tools.codegraph.codegraph_explore({
  query: "LeeConfiguracionLocal",
  projectPath: "C:/00repos/codigo/access2web-blueprint/data/staging/condor"
});
```
