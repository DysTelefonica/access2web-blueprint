# data/staging/lanzaderas

Snapshot self-contained de **Lanzadera**.

## Origen (READ-ONLY)

| Carpeta | Tipo | Comentario |
|---|---|---|
| `C:\00repos\codigo\00_LANZADERA\staging` | fuente de verdad | staging activo; scripts/, tests/ propios |
| `C:\00repos\codigo\00_LANZADERA\00_main` | comparación | binarios idénticos a staging |

`staging/` y `00_main/` tienen idénticos `.accdb`; se eligió `staging/` por tener scripts/ y tests/ (no copiados acá pero queda como referencia).

## Archivos locales

| Path | Tamaño | Fuente |
|---|---|---|
| `frontend/Lanzadera.accdb` | 16.94 MB | `00_LANZADERA\staging\Lanzadera.accdb` |
| `backend/Lanzadera_Datos.accdb` | 5.44 MB | `00_LANZADERA\staging\Lanzadera_Datos.accdb` |
| `src/` (89 archivos + `backends.json` + `tests.vba.json`) | 5.44 MB | `00_LANZADERA\staging\src\` |
| `docs/` | — | **no existe** en el repo original |

## Notas

- **Suite más chica del set**: 89 archivos en `src/`. Lanzadera es una app de identidad/permisos/sesión relativamente contenida.
- **No tiene `docs/`** propia del repo; la documentación está en `access2web-blueprint/docs/03-aplicaciones/lanzadera/` y `02-topologia-ecosistema/lanzadera-identidad-permisos.md`.
- Disposiciones finales consolidadas en `docs/03-aplicaciones/lanzadera/capabilities.md`.
- Acoplada con `gestion-riesgos` vía `getdbLanzadera()` desde la otra app.

## Acceso con Dysflow

```js
await tools.dysflow.list_objects({
  accessPath: "C:/00repos/codigo/access2web-blueprint/data/staging/lanzaderas/frontend/Lanzadera.accdb",
  backendPath: "C:/00repos/codigo/access2web-blueprint/data/staging/lanzaderas/backend/Lanzadera_Datos.accdb",
  allowExternalAccessPath: true
});
```

## Acceso con CodeGraph

```js
await tools.codegraph.codegraph_explore({
  query: "Lanzadera",
  projectPath: "C:/00repos/codigo/access2web-blueprint/data/staging/lanzaderas"
});
```
