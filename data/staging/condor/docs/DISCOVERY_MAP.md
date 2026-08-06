# Mapa de descubrimiento de CONDOR

Este mapa orienta a una IA o persona antes de modificar CONDOR. La documentación de capacidades es la puerta de entrada funcional; el código y Dysflow siguen siendo la fuente de verdad de comportamiento ejecutable.

## Rutas canónicas

- Repo de código en este worktree: `C:\00repos\codigo\00_CONDOR_staging`
- Documentación externa / OpenSpec: `C:\00repos\documentacion\OPENSPEC\00_CONDOR`
- Índice local de capacidades: [`docs/capabilities/capabilities-index.md`](capabilities/capabilities-index.md)
- PCSUB: [`docs/capabilities/CAP-001-pcsub.md`](capabilities/CAP-001-pcsub.md)

## Orden recomendado de lectura

1. `README.md` del repo de código.
2. `docs/capabilities/capabilities-index.md`.
3. Documento de capacidad aplicable.
4. Artefactos SDD externos relacionados.
5. Código fuente relevante en `src/`.
6. Evidencia de pruebas y fixtures en `tests/` y `docs/testing/`.
7. Verificación con Dysflow antes de afirmar estado runtime.

## OpenSpec externo relevante

- `C:\00repos\documentacion\OPENSPEC\00_CONDOR\README.md`
- `openspec/changes/fixture-first-tdd-suite`
- `openspec/changes/backend-getdb-single-entrypoint`
- `openspec/changes/document-template-mapping-e2e`
- `openspec/changes/pcsub-nueva-solicitud`
- `openspec/changes/pcsub-staging-recovery`
- `openspec/changes/pcsub-twin-gap-analysis`
- `openspec/changes/pcsub-guardar-phase-advancement`
- `changes/pcsub-full-coverage-dpcdcf`

## Evidencia local útil

- `docs/ERD/condor_datos.md`: referencia de esquema exportado.
- `docs/testing/pcsub-fixture-graph.md`: grafo fixture-first PCSUB.
- `docs/testing/pcsub-coverage-map.md`: cobertura y deuda PCSUB.
- `tests/tests.vba.json`: manifiesto estricto actual de pruebas PCSUB exportadas.
- `tests/tests.pcsub.json`: smoke PCSUB agregado.

## Regla operativa

No confundir documentación con prueba. Si una regla está documentada pero no tiene ejecución Dysflow vigente, debe quedar como `Verified-static`, `Intended` o `Likely`, nunca como `Verified-runtime`.
