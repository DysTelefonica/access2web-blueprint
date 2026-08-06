# data/staging — Self-contained source-staging for the eight legacy Access apps

Este directorio contiene el **source tree** exportado y la **documentación** de cada aplicación legacy en scope. El repo es **self-contained** en la parte de texto: cualquier IA puede clonar el repo y trabajar solo con lo que está en git.

Los **binarios** (`.accdb` frontend/backend, `.codegraph-vba/`, `resources/`) **NO están en git**. Viven en **Cloudflare R2** y se descargan con `scripts/setup-staging.ps1`. Razón: pesan ~870 MB en total, demasiado para push a GitHub.

## Por qué existe

Los repos originales (`00_CONDOR`, `00_HPS`, `HPS_SOLICITUDES`, `00_BRASS`, `00_GESTION_RIESGOS`, `00_NO_CONFORMIDADES`, `00_EXPEDIENTES`, `00_LANZADERA`) viven en worktrees separados con su propio `.dysflow/project.json` cada uno y son READ-ONLY para esta tarea. Para que un agente pueda hacer un análisis cross-app (queries SQL reales, blast radius entre apps, smoke tests) necesita acceso local a:

- El **frontend** (`.accdb`) de cada app — en R2.
- El **backend** (`.accdb` con tablas vinculadas) — en R2.
- El árbol **`src/`** exportado de cada binario (`.bas`, `.cls`, `.form.txt`, `.report.txt`) — **en git**.
- La **`docs/`** original (input; puede estar outdated) — **en git**.

## Qué hay por app

Cada subdirectorio `data/staging/<app>/` sigue la misma estructura:

```
data/staging/<app>/
├── frontend/<app>_<frontend>.accdb            # ⚠️ DESCARGADO de R2 (no en git)
├── backend/<app>_datos.accdb                  # ⚠️ DESCARGADO de R2 (no en git)
├── src/                                        # ✓ EN GIT (text, diffable)
├── docs/                                       # ✓ EN GIT (input, puede estar outdated)
├── resources/                                  # ⚠️ DESCARGADO de R2 (no en git)
├── .codegraph-vba/                             # ⚠️ DESCARGADO de R2 (no en git, regenerable)
└── README.md                                   # ✓ EN GIT
```

Apps en scope (8):

| App | Frontend | Backend | Source base original | src files | docs |
|---|---|---|---|---|---|
| `condor` | `CONDOR.accdb` | `condor_datos.accdb` | `C:\00repos\codigo\00_CONDOR\staging` | 221 | 20 |
| `hps` | `HPS.accdb` | `HPST.accdb` | `C:\00repos\codigo\00_HPS\staging` | 139 | — |
| `hps-solicitudes` | `Solicitudes_HPS.accdb` | `Solicitudes_HPS_datos.accdb` | `C:\00repos\codigo\HPS_SOLICITUDES` | 108 | 2 |
| `brass` | `Gestion_Brass_Gestion.accdb` | (mismo archivo) | `C:\00repos\codigo\00_BRASS\00_main` | 218 | — |
| `gestion-riesgos` | `Gestion_Riesgos.accdb` | `Gestion_Riesgos_Datos.accdb` | `C:\00repos\codigo\00_GESTION_RIESGOS\staging` | 328 | 29 |
| `no-conformidades` | `NoConformidades.accdb` | `NoConformidades_Datos.accdb` | `C:\00repos\codigo\00_NO_CONFORMIDADES\staging` | 249 | 51 |
| `expedientes` | `Expedientes.accdb` | `Expedientes_datos.accdb` | `C:\00repos\codigo\00_EXPEDIENTES\staging` | 302 | 69 |
| `lanzaderas` | `Lanzadera.accdb` | `Lanzadera_Datos.accdb` | `C:\00repos\codigo\00_LANZADERA\staging` | 89 | — |

Notas puntuales:

- **BRASS** no tiene backend separado; el mismo `.accdb` actúa como frontend y backend.
- **HPS** llama a su backend `HPST.accdb` (no `_datos.accdb`).
- **HPS_SOLICITUDES** vive en `HPS_SOLICITUDES` (sin prefijo `00_`) y no tiene `staging/` ni `00_main/`.
- Para HPS, HPS_SOLICITUDES, BRASS y LANZADERAS el repo original **no tiene `docs/`** propio del repo (solo `AGENTS.md` / `README.md` / `CHANGELOG.md` que no se copian).

Total: 8 apps × (`frontend` + `backend?` + `src/` + `docs/?`) en R2 ≈ 870 MB en descarga. **Texto** en git (`src/`, `docs/`, READMEs, scripts) ≈ 50 MB.

## Cómo usar

### 1. En una máquina nueva, descargar los binarios desde R2

```powershell
# Verificar que rclone tiene el remote "cloudflare-r2"
rclone listremotes

# Descargar los binarios + .codegraph-vba/ para una app
pwsh -File scripts/setup-staging.ps1 -App condor

# O todas las 8 apps
pwsh -File scripts/setup-staging.ps1
```

Las carpetas `frontend/`, `backend/`, `resources/`, `.codegraph-vba/` se pueblan desde `cloudflare-r2:access2web-staging-binaries/<app>/`. Lo que ya está local se respeta.

### 2. Con Dysflow (MCP)

`get_capabilities` ya resuelve la app **condor** como default. Para apuntar a otra app, pasá `accessPath` y `backendPath` por llamada:

```js
// Ejemplo: leer objetos de HPS
await tools.dysflow.list_objects({
  accessPath: "C:/00repos/codigo/access2web-blueprint/data/staging/hps/frontend/HPS.accdb",
  backendPath: "C:/00repos/codigo/access2web-blueprint/data/staging/hps/backend/HPST.accdb",
  allowExternalAccessPath: true
});
```

El campo `backends[]` en `.dysflow/project.json` documenta los paths exactos de las 8 apps. Es metadata para humanos / agentes — el runtime de Dysflow no lo interpreta (su contrato exige un único frontend por worktree).

### 3. Con CodeGraph-VBA (MCP)

Cada app tiene su propio `.codegraph-vba/` independiente. Para explotar el índice de una app concreta:

```js
await tools.codegraph.codegraph_explore({
  query: "LeeConfiguracionLocal",
  projectPath: "C:/00repos/codigo/access2web-blueprint/data/staging/condor"
});
```

Cambiá el `projectPath` por `data/staging/<app>` para apuntar a otra app. Cada índice reconoce automáticamente `.bas`, `.cls`, `.form.txt`, `.report.txt` en su `src/`.

### 4. Cómo regenerar el índice de una app (opcional)

El índice en R2 es funcional. Si querés reconstruirlo desde cero:

```powershell
codegraph-vba index "C:/00repos/codigo/access2web-blueprint/data/staging/<app>"
```

`index` rebuilda desde cero; `sync` actualiza incrementalmente. El watcher auto-sync se ocupa mientras se trabaja; solo tocá `sync` si el watcher está deshabilitado.

## Limitaciones y caveats

- **NO** sincronizar cambios desde `data/staging/<app>/src/` hacia los repos originales (`00_<APP>\staging\src\`) — los snapshots son de solo lectura local.
- **NO** escribir en producción. Los `.accdb` de staging son **mutables** (writeReady=true para auditoría y tests) pero cualquier cambio debe vivir en el repo original, no acá.
- **Secretos legacy**: Los `.accdb` pueden contener contraseñas embebidas en linked tables (D93, D104, D109). Viven en R2 (no en git) por esta razón.
- **Drift**: Si los repos originales avanzan, los snapshots en R2 quedan stale. Para re-sincronizar: re-subir a R2 con `rclone copy` desde el staging del repo original.
- **`.dysflow/project.json` runtime contract**: el runtime solo soporta un frontend por worktree. Las otras 7 apps se acceden por override per-call.
- **Cambiar de ordenador**: simplemente `git clone` + `scripts/setup-staging.ps1`. Las credenciales de R2 tienen que estar en `rclone` config local.
