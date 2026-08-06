# data/staging/hps

Snapshot self-contained de **HPS** (Hojas de Problemas / Seguimiento).

## Origen (READ-ONLY)

| Carpeta | Tipo | Comentario |
|---|---|---|
| `C:\00repos\codigo\00_HPS\staging` | fuente de verdad | staging activo; `.dysflow/project.json` con allowlist amplia de tests |
| `C:\00repos\codigo\00_HPS\00_main` | comparación | mismo frontend/backend que staging (28.88 MB / 7.1 MB) |

`staging/` y `00_main/` tienen idénticos `.accdb`; se eligió `staging/` por su src/ canónico.

## Archivos locales

| Path | Tamaño | Fuente |
|---|---|---|
| `frontend/HPS.accdb` | 28.88 MB | `00_HPS\staging\HPS.accdb` |
| `backend/HPST.accdb` | 7.1 MB | `00_HPS\staging\HPST.accdb` |
| `src/` (139 archivos + `backends.json` + `tests.vba.json`) | 5.48 MB | `00_HPS\staging\src\` |
| `docs/` | — | **no existe** en el repo original |

## Notas

- **Backend se llama `HPST.accdb`**, no `HPS_datos.accdb`. Es el patrón histórico de esta app.
- **No tiene `docs/`** propia en el repo; la documentación de la app vive en `docs/03-aplicaciones/hps/` de `access2web-blueprint`.
- `src/` incluye `tests.vba.json` con la lista de tests TDD registrada — apunta a `dysflow test_vba`.
- El layout flat del staging original (`bin/`, `classes/`, `forms/`, `modules/`, `queries/`, `reports/`, `src/`) **no se preserva**; aquí solo se copió `src/`.

## Acceso con Dysflow

```js
await tools.dysflow.list_objects({
  accessPath: "C:/00repos/codigo/access2web-blueprint/data/staging/hps/frontend/HPS.accdb",
  backendPath: "C:/00repos/codigo/access2web-blueprint/data/staging/hps/backend/HPST.accdb",
  allowExternalAccessPath: true
});
```

## Acceso con CodeGraph

```js
await tools.codegraph.codegraph_explore({
  query: "RelinkToLocalData",
  projectPath: "C:/00repos/codigo/access2web-blueprint/data/staging/hps"
});
```
