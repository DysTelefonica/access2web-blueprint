# data/staging/condor

Snapshot self-contained de **Condor**.

## Origen (READ-ONLY)

| Carpeta | Tipo | Comentario |
|---|---|---|
| `C:\00repos\codigo\00_CONDOR\staging` | fuente de verdad | staging activo; `.dysflow/project.json` presente |
| `C:\00repos\codigo\00_CONDOR\00_main` | comparación | `CONDOR.accdb` 11.12 MB; src/ 208 archivos |

Se eligió `staging/` por ser la baseline funcional con `src/` más completo (221 archivos vs 208) y por incluir `docs/` propia.

## Archivos locales

| Path | Tamaño | Fuente |
|---|---|---|
| `frontend/CONDOR.accdb` | 42.84 MB | `00_CONDOR\staging\CONDOR.accdb` |
| `backend/condor_datos.accdb` | 5 MB | `00_CONDOR\staging\condor_datos.accdb` |
| `src/` (221 archivos) | 10.17 MB | `00_CONDOR\staging\src\` |
| `docs/` (20 archivos) | 0.38 MB | `00_CONDOR\staging\docs\` |

## Notas

- **Default en Dysflow**: `CONDOR.accdb` y `condor_datos.accdb` están copiados también a la raíz del repo (`C:\00repos\codigo\access2web-blueprint\CONDOR.accdb`) para que `frontendFile: "CONDOR.accdb"` resuelva bajo el worktree root sin overrides.
- **src/** está organizado en `classes/`, `forms/`, `modules/`, `queries/`, `reports/`. Hay un `.disco/` con configuración del backend (`backends.json`).
- **docs/** tiene una `DISCOVERY_MAP.md` raíz más subcarpetas `capabilities/`, `epics/`, `ERD/`, `features/`, `testing/`. Es la discovery viva de la app.

## Acceso con Dysflow

Como es el default, no requiere overrides:

```js
await tools.dysflow.list_objects({ outputMode: "summary" });
```

Desde el MCP el `projectConfig.accessPath` resuelve a `C:/.../CONDOR.accdb` y `projectConfig.writeReady === true`.

## Acceso con CodeGraph

```js
await tools.codegraph.codegraph_explore({
  query: "LeeConfiguracionLocal",
  projectPath: "C:/00repos/codigo/access2web-blueprint/data/staging/condor"
});
```
