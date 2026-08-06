# data/staging/hps-solicitudes

Snapshot self-contained de **HPS Solicitudes**.

## Origen (READ-ONLY)

| Carpeta | Tipo | Comentario |
|---|---|---|
| `C:\00repos\codigo\HPS_SOLICITUDES` | fuente de verdad | sin prefijo `00_`; no tiene subcarpetas `staging/` ni `00_main/` |

Es la única app del set sin layout `staging/00_main`. El checkout raíz ES el staging.

## Archivos locales

| Path | Tamaño | Fuente |
|---|---|---|
| `frontend/Solicitudes_HPS.accdb` | 21 MB | `HPS_SOLICITUDES\Solicitudes_HPS.accdb` |
| `backend/Solicitudes_HPS_datos.accdb` | 3.77 MB | `HPS_SOLICITUDES\Solicitudes_HPS_datos.accdb` |
| `src/` (108 archivos) | 4.51 MB | `HPS_SOLICITUDES\src\` |
| `docs/` (2 archivos) | ≈0 MB | `HPS_SOLICITUDES\docs\` |

## Notas

- **Solo 2 archivos en `docs/`**: `DISCOVERY_MAP.md` y `Informes/Informe_Correccion_ID256_2026-001.md`. El grueso de la documentación está en `access2web-blueprint/docs/03-aplicaciones/hps-solicitudes/`.
- `Solicitudes_HPS.accdb` se vincula contra `Solicitudes_HPS_datos.accdb` con la convención `_datos` estándar.

## Acceso con Dysflow

```js
await tools.dysflow.list_objects({
  accessPath: "C:/00repos/codigo/access2web-blueprint/data/staging/hps-solicitudes/frontend/Solicitudes_HPS.accdb",
  backendPath: "C:/00repos/codigo/access2web-blueprint/data/staging/hps-solicitudes/backend/Solicitudes_HPS_datos.accdb",
  allowExternalAccessPath: true
});
```

## Acceso con CodeGraph

```js
await tools.codegraph.codegraph_explore({
  query: "Solicitudes",
  projectPath: "C:/00repos/codigo/access2web-blueprint/data/staging/hps-solicitudes"
});
```
