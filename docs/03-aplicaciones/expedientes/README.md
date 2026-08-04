# 03 · Expedientes

## Propósito

Carpeta de descubrimiento de **Expedientes**: aplicación que crea y mantiene el expediente del que dependen casi todas las demás.

## Estado

- **Mapeo:** resuelto, con aviso.
- **Dysflow:** `path-mismatch`; `accessPath` absoluto heredado. Proyecto **no write-ready**; solo lectura en esta fase.
- **Inventario `list_objects`:** 158 elementos resumidos (cifra preliminar).
- **Frontend observado:** `Expedientes.accdb` con `TbAuxEstadosMartina`.
- **Backend observado:** `Expedientes_datos.accdb` localizado en staging (no en `00_main`).
- **Rutas de entrada:** indicadas en `00_EXPEDIENTES/00_main/AGENTS.md`.

## Lote asociado

Lote 2 de `exploration.md` — Expedientes.

## Entregable previsto

1. Alta, edición y estados del expediente; entidades, suministradores y responsables.
2. Anexos, consultas, informes, E2E/JSON y batch.
3. Contratos de `IDExpediente` / `CodExp` / `CodExpLargo` / `Nemotecnico`.
4. Consumidores actuales (sin extrapolar).
5. Clases `UsuarioAplicacionPermisos` y operaciones de expediente.

## Fuentes de autoridad

1. `C:\00repos\documentacion\OPENSPEC\00_EXPEDIENTES`.
2. `C:\00repos\codigo\00_EXPEDIENTES\00_main` (incluye `AGENTS.md`).
3. Inspección Dysflow solo lectura; no se corrige el `path-mismatch` aquí.
4. CodeGraph-VBA ya presente en este `00_main`.
5. Engram solo como contexto histórico.

## Reglas de evidencia

- La configuración Dysflow no se modifica; el aviso se traslada a `08-decisiones-y-preguntas-abiertas.md`.
- Los consumidores se listan solo si la evidencia los sostiene; la matriz global se reconcilia en el Lote 9.
- Los identificadores se tratan como contrato de negocio, no como detalle.

## Checklist

- [ ] Aviso `path-mismatch` registrado para corrección en lote aprobado.
- [ ] Consumidores actuales documentados con su evidencia.
- [ ] APAP y APAP_WEB no aparecen.

## Siguiente paso

Poblar capacidades tras resolver si el backend definitivo vive en `00_main` o en staging.
