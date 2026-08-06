# data/staging/brass

Snapshot self-contained de **BRASS**.

## Origen (READ-ONLY)

| Carpeta | Tipo | Comentario |
|---|---|---|
| `C:\00repos\codigo\00_BRASS\00_main` | fuente de verdad | **no existe** `staging/`; `00_main` es el branch más estable |
| `C:\00repos\codigo\00_BRASS\develop` | paralelo | mismo binario que `00_main` (24.12 MB) |
| `C:\00repos\codigo\00_BRASS\release_2026-001` | release | binario 24.32 MB (un poco más grande por compact) |

Se eligió `00_main` por ser la línea principal.

## Archivos locales

| Path | Tamaño | Fuente |
|---|---|---|
| `frontend/Gestion_Brass_Gestion.accdb` | 24.12 MB | `00_BRASS\00_main\Gestion_Brass_Gestion.accdb` |
| `backend/` | — | **no existe** — BRASS no tiene backend separado |
| `src/` (218 archivos) | 9.65 MB | `00_BRASS\00_main\src\` |
| `docs/` | — | **no existe** en el repo original |

## Notas

- **No hay backend separado**: el único `.accdb` (`Gestion_Brass_Gestion.accdb`) hace de frontend y backend simultáneamente. Para `dysflow` se lo apunta como `accessPath` Y `backendPath` al mismo archivo (es válido para apps single-file).
- **No hay `docs/` propia** del repo; la discovery vive en `access2web-blueprint/docs/03-aplicaciones/brass/`.

## Acceso con Dysflow

```js
await tools.dysflow.list_objects({
  accessPath: "C:/00repos/codigo/access2web-blueprint/data/staging/brass/frontend/Gestion_Brass_Gestion.accdb",
  backendPath: "C:/00repos/codigo/access2web-blueprint/data/staging/brass/frontend/Gestion_Brass_Gestion.accdb",
  allowExternalAccessPath: true
});
```

## Acceso con CodeGraph

```js
await tools.codegraph.codegraph_explore({
  query: "Brass",
  projectPath: "C:/00repos/codigo/access2web-blueprint/data/staging/brass"
});
```
