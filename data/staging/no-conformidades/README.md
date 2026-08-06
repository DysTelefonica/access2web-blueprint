# data/staging/no-conformidades

Snapshot self-contained de **No Conformidades**.

## Origen (READ-ONLY)

| Carpeta | Tipo | Comentario |
|---|---|---|
| `C:\00repos\codigo\00_NO_CONFORMIDADES\staging` | fuente de verdad | staging activo; `.dysflow/project.json` con `allowWrites=true` |
| `C:\00repos\codigo\00_NO_CONFORMIDADES\00_main` | comparación | binarios más livianos (41.25 MB vs 66.96 MB) |

Se eligió `staging/` por ser la baseline funcional con más historia.

## Archivos locales

| Path | Tamaño | Fuente |
|---|---|---|
| `frontend/NoConformidades.accdb` | 66.96 MB | `00_NO_CONFORMIDADES\staging\NoConformidades.accdb` |
| `backend/NoConformidades_Datos.accdb` | 32.18 MB | `00_NO_CONFORMIDADES\staging\NoConformidades_Datos.accdb` |
| `src/` (249 archivos) | 9.41 MB | `00_NO_CONFORMIDADES\staging\src\` |
| `docs/` (51 archivos) | 0.88 MB | `00_NO_CONFORMIDADES\staging\docs\` |

## Notas

- **Frontend pesa 66.96 MB** — el más grande del set.
- **docs/** es extensa (51 archivos). Tiene `DISCOVERY_MAP.md`, `uat/`, etc.
- El repo original tiene `database/` aparte del `src/` (con objetos de seed/migration); **no se copia** acá porque está fuera del scope pedido (solo `src/`, `docs/`, `frontend/`, `backend/`).
- También tiene `openspec/` en staging; **no se copia**.

## Acceso con Dysflow

```js
await tools.dysflow.list_objects({
  accessPath: "C:/00repos/codigo/access2web-blueprint/data/staging/no-conformidades/frontend/NoConformidades.accdb",
  backendPath: "C:/00repos/codigo/access2web-blueprint/data/staging/no-conformidades/backend/NoConformidades_Datos.accdb",
  allowExternalAccessPath: true
});
```

## Acceso con CodeGraph

```js
await tools.codegraph.codegraph_explore({
  query: "NoConformidad",
  projectPath: "C:/00repos/codigo/access2web-blueprint/data/staging/no-conformidades"
});
```
