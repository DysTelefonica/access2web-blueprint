# data/staging/gestion-riesgos

Snapshot self-contained de **Gestion de Riesgos**.

## Origen (READ-ONLY)

| Carpeta | Tipo | Comentario |
|---|---|---|
| `C:\00repos\codigo\00_GESTION_RIESGOS\staging` | fuente de verdad | staging activo; `.dysflow/project.json` con allowlist amplia |
| `C:\00repos\codigo\00_GESTION_RIESGOS\00_main` | comparación | binarios más livianos (19.36 MB vs 38.99 MB en staging) |

`staging/` es baseline funcional con más historia (front 38.99 MB, backend 16.19 MB vs `00_main` con 19.36 / 15.09). La queja histórica del árbol de ediciones/riesgos está caracterizada en la discovery (D88).

## Archivos locales

| Path | Tamaño | Fuente |
|---|---|---|
| `frontend/Gestion_Riesgos.accdb` | 38.99 MB | `00_GESTION_RIESGOS\staging\Gestion_Riesgos.accdb` |
| `backend/Gestion_Riesgos_Datos.accdb` | 16.19 MB | `00_GESTION_RIESGOS\staging\Gestion_Riesgos_Datos.accdb` |
| `src/` (328 archivos) | 11.39 MB | `00_GESTION_RIESGOS\staging\src\` |
| `docs/` (29 archivos) | 0.43 MB | `00_GESTION_RIESGOS\staging\docs\` |

## Notas

- **Acoplamiento declarado con Lanzadera** vía `getdbLanzadera()`. Está incorporado en D86/D87 como punto de migración.
- **src/** es el más grande del set (328 archivos) — refleja la densidad de la app (riesgos, ediciones, anexos, indicadores, correos).
- **docs/** tiene un layout disperso (`prompts/`, `uat/`, `ui-estilo-form-anexos.md`).
- Hay una queja histórica de rendimiento del árbol de ediciones caracterizada en D88 (HTMX + CTE recursivo + lazy expansion propuesto).

## Acceso con Dysflow

```js
await tools.dysflow.list_objects({
  accessPath: "C:/00repos/codigo/access2web-blueprint/data/staging/gestion-riesgos/frontend/Gestion_Riesgos.accdb",
  backendPath: "C:/00repos/codigo/access2web-blueprint/data/staging/gestion-riesgos/backend/Gestion_Riesgos_Datos.accdb",
  allowExternalAccessPath: true
});
```

## Acceso con CodeGraph

```js
await tools.codegraph.codegraph_explore({
  query: "Test_Helper_CacheHit_GetAnexosDeEdicion_LecturaRepetidaMismoID",
  projectPath: "C:/00repos/codigo/access2web-blueprint/data/staging/gestion-riesgos"
});
```
