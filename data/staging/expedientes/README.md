# data/staging/expedientes

Snapshot self-contained de **Expedientes**.

## Origen (READ-ONLY)

| Carpeta | Tipo | Comentario |
|---|---|---|
| `C:\00repos\codigo\00_EXPEDIENTES\staging` | fuente de verdad | staging activo; `e2e-integration-rest`, `fix-e2e-json-utf8` slices |
| `C:\00repos\codigo\00_EXPEDIENTES\00_main` | comparación | `Expedientes.accdb` 16.75 MB; sin `Expedientes_datos.accdb` separado |

Se eligió `staging/` por tener el par `frontend/backend` completo y más documentación.

## Archivos locales

| Path | Tamaño | Fuente |
|---|---|---|
| `frontend/Expedientes.accdb` | 23.33 MB | `00_EXPEDIENTES\staging\Expedientes.accdb` |
| `backend/Expedientes_datos.accdb` | 4.35 MB | `00_EXPEDIENTES\staging\Expedientes_datos.accdb` |
| `src/` (302 archivos) | 9.15 MB | `00_EXPEDIENTES\staging\src\` |
| `docs/` (69 archivos) | 0.46 MB | `00_EXPEDIENTES\staging\docs\` |

## Notas

- **docs/** es la más completa del set (69 archivos). Discovery funcional, técnico y de migración documentado.
- Hay un `.temp/` en el staging original con copias compact (`Expedientes_compact.accdb`, `Expedientes.compact.accdb`); **no se copian** — son artefactos runtime.
- Hay `exports/` y `openspec/` en el staging original; **no se copian**.
- Índice `.codegraph-vba/` regenerado (331 MB) — refleja la densidad documental + de código.
- El staging original tiene `expedientes_relacion.json` (metadata cross-app) — **no se copia** acá; vive solo en el repo original.

## Acceso con Dysflow

```js
await tools.dysflow.list_objects({
  accessPath: "C:/00repos/codigo/access2web-blueprint/data/staging/expedientes/frontend/Expedientes.accdb",
  backendPath: "C:/00repos/codigo/access2web-blueprint/data/staging/expedientes/backend/Expedientes_datos.accdb",
  allowExternalAccessPath: true
});
```

## Acceso con CodeGraph

```js
await tools.codegraph.codegraph_explore({
  query: "Expedientes",
  projectPath: "C:/00repos/codigo/access2web-blueprint/data/staging/expedientes"
});
```
