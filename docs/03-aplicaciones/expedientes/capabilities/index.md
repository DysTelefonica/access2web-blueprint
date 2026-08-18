[← Back to Expedientes README](../README.md)

# Expedientes — capacidades documentadas (índice)

Este índice enlaza las capacidades del Expedientes que ya están documentadas en `data/staging/expedientes/docs/capabilities/` (single source of truth del walkthrough de discovery). Cada CAP tiene un fichero estructurado con secciones §0 Identidad, §1 Intención, §2 Contrato de comportamiento, §3 Modelo, §4 Walkthrough, §5 Hallazgos, §6 Decisiones, §7 Anexos.

`data/staging/` es material de sólo lectura per `AGENTS.md` §«Alcance». No se modifica aquí; se referencia desde este índice.

## Capacidades ya documentadas

17 CAPs en `data/staging/expedientes/docs/capabilities/` a 2026-08-18:

| ID | Título corto | Ruta en staging | Dominio |
|---|---|---|---|
| CAP-001 | Alta de expediente | `data/staging/expedientes/docs/capabilities/CAP-001-alta.md` | Expediente (escritura) |
| CAP-002 | Edición de expediente | `data/staging/expedientes/docs/capabilities/CAP-002-edicion.md` | Expediente (escritura) |
| CAP-003 | Eliminación condicionada | `data/staging/expedientes/docs/capabilities/CAP-003-eliminacion.md` | Expediente (escritura) |
| CAP-006 | Comercial | `data/staging/expedientes/docs/capabilities/CAP-006-comercial.md` | Catálogos (master-detail) |
| CAP-007 | CPV | `data/staging/expedientes/docs/capabilities/CAP-007-cpv.md` | Catálogos (master-detail) |
| CAP-008 | Ejército | `data/staging/expedientes/docs/capabilities/CAP-008-ejercito.md` | Catálogos (master-detail) |
| CAP-009 | Suministrador | `data/staging/expedientes/docs/capabilities/CAP-009-suministrador.md` | Suministradores |
| CAP-010 | Lugar | `data/staging/expedientes/docs/capabilities/CAP-010-lugar.md` | Catálogos (master-detail) |
| CAP-011 | PECAL | `data/staging/expedientes/docs/capabilities/CAP-011-pecal.md` | Catálogos (master-detail) |
| CAP-012 | RAC | `data/staging/expedientes/docs/capabilities/CAP-012-rac.md` | Catálogos (master-detail) |
| CAP-013 | Grado | `data/staging/expedientes/docs/capabilities/CAP-013-grado.md` | Catálogos (master-detail) |
| CAP-014 | Órgano | `data/staging/expedientes/docs/capabilities/CAP-014-organo.md` | Catálogos (master-detail) |
| CAP-015 | Oficina | `data/staging/expedientes/docs/capabilities/CAP-015-oficina.md` | Catálogos (master-detail) |
| CAP-023 | Búsqueda avanzada | `data/staging/expedientes/docs/capabilities/CAP-023-busqueda-avanzada.md` | Consulta |
| CAP-025 | Excel | `data/staging/expedientes/docs/capabilities/CAP-025-excel.md` | Exportación |
| CAP-040 | Tareas (bandeja) | `data/staging/expedientes/docs/capabilities/CAP-040-tareas-bandeja.md` | Tareas |

## Capacidades pendientes de documentar

21 CAPs (CAP-016..022, 024, 026..038) no tienen fichero en staging ni en `docs/`. Véase el issue de gap creado en paralelo para el plan de cierre.

## Cómo se aplica a access2web-blueprint

Este índice es el hub de las capacidades del Expedientes a nivel repo. La fuente de verdad por capacidad vive en `data/staging/expedientes/docs/capabilities/` (read-only); este índice sólo expone qué hay y dónde buscarlo. El código destino de cada capacidad vivirá en `app/src/modules/expedientes/` siguiendo la forma hexagonal (DA-1). Las decisiones operativas D86 (forma hexagonal del legacy preservada), D87 (Test_* VBA preservados), D94 (FKs conceptuales), D102 (booleans Text(2) → BOOLEAN) siguen aplicando.

## Lista de comprobación final

- [ ] Las 17 entradas referencian paths staging que existen (verificables con `ls`).
- [ ] Las 21 capacidades pendientes quedan abiertas en el issue de gap, no se pierden.
- [ ] El path `data/staging/` NO se ha modificado (read-only).
- [ ] El índice se ajusta al contrato de `skills/documentation-alan-style/SKILL.md` (§3 + §8): castellano peninsular formal con usted, sin emojis decorativos.
- [ ] Cross-references desde DOCS.md y `docs/03-aplicaciones/expedientes/README.md` siguen resolviendo.

## Navigation

Previous: [Expedientes README](../README.md) | Next: [DOCS](../../../DOCS.md)
