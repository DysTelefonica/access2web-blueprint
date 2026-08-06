# 00 · Alcance y evidencia

## Propósito

Fija el alcance del descubrimiento del ecosistema legacy y declara la jerarquía de fuentes que debe respetarse en todo el árbol `docs/`. Este fichero es la **puerta de entrada**; ningún otro documento introduce alcance, aplicaciones o jerarquía de fuentes por su cuenta.

## Alcance

Solo se documentan estas ocho aplicaciones:

| # | Aplicación | Estado inicial (de `exploration.md`) |
|---|---|---|
| 1 | Lanzadera | Mapeo resuelto. Snapshot en `data/staging/lanzaderas/`. |
| 2 | Gestion_Riesgos | Mapeo resuelto. Snapshot en `data/staging/gestion-riesgos/`. |
| 3 | No_Conformidades | Mapeo resuelto; inventario Dysflow pendiente. Snapshot en `data/staging/no-conformidades/`. |
| 4 | Condor | Frontend localizado; backend y `.dysflow/project.json` pendientes. **Snapshot completo** en `data/staging/condor/` (default del runtime). |
| 5 | HPS_Solicitudes | Documentación localizada; checkout/binary no identificados. **Snapshot completo** en `data/staging/hps-solicitudes/`. |
| 6 | HPS | Mapeo resuelto. Snapshot en `data/staging/hps/`. |
| 7 | Brass | Frontend localizado; backend y `.dysflow/project.json` pendientes. **Snapshot completo** en `data/staging/brass/`. |
| 8 | Expedientes | Mapeo resuelto con aviso de `path-mismatch` en Dysflow. Snapshot en `data/staging/expedientes/`. |

**Fuera de alcance:** APAP y APAP_WEB. No se mencionan como nodos, evidencia ni comparación; esta exclusión es la única referencia admisible.

## Hechos de orientación del usuario (no sustituyen verificación)

1. Lanzadera es el origen de los usuarios y los permisos.
2. Expedientes crea el expediente del que dependen casi todas las aplicaciones.

Estos dos hechos guían la exploración, pero toda afirmación concreta debe verificarse contra código, documentación o Dysflow antes de ascender a hecho.

## Jerarquía de fuentes

Orden de autoridad, de mayor a menor. Un nivel inferior no puede elevar una hipótesis a hecho:

1. `C:\00repos\documentacion` — documentación y especificaciones existentes.
2. Checkout `00_main` de cada aplicación bajo `C:\00repos\codigo` — código, binarios y configuración local.
3. Snapshot self-contained en `data/staging/<app>/` de este repo — copia de los staging activos para que cualquier IA pueda trabajar sin acceso a los repos originales. **Idéntica autoridad** que el checkout original; se actualiza re-copiando desde el source.
4. Inspección Dysflow en vivo, **solo lectura** (`get_capabilities`, `list_objects`, `list_tables`).
5. CodeGraph y CodeGraph-VBA como apoyo de navegación por símbolos.
6. Engram — **contexto histórico únicamente**; nunca se eleva a hecho actual.

## Reglas de evidencia

- Código y documentación vigente son la verdad. El producto y Engram solo aportan intención o historia.
- Toda afirmación de comportamiento se etiqueta: `Verified-runtime`, `Verified-static`, `Intended`, `Likely` o `Divergent`, con fecha y evidencia.
- Sin prueba no se afirma; sin trazabilidad no se cierra.
- Las configuraciones Dysflow con `path-mismatch` o ausentes no se mutan en esta fase.

## Modo de uso

| Si vienes a… | Empieza por |
|---|---|
| Entender qué aplicaciones existen | `01-inventario-aplicaciones.md` |
| Ver cómo se conectan | `02-topologia-ecosistema/` |
| Estudiar una aplicación concreta | `03-aplicaciones/<aplicacion>/` |
| Revisar integraciones, batch, correo, rutas | `04-integraciones-y-operacion/` |
| Buscar una capacidad de negocio | `05-capacidades/capabilities-index.md` |
| Preparar la migración a web | `07-migracion/` |
| Cerrar dudas o divergencias | `08-decisiones-y-preguntas-abiertas.md` |
| Revisar decisiones de producto y arquitectura aprobadas | `09-arquitectura-objetivo-y-principios.md` |
| **Trabajar con binarios + código VBA localmente sin tocar los repos originales** | `data/staging/README.md` (snapshot self-contained) |

## Checklist

- [ ] Toda mención a una aplicación identifica el checkout y la fuente documental.
- [ ] Ningún documento nombra a APAP o APAP_WEB fuera de este fichero.
- [ ] Ningún dato aparece como `Verified-runtime` sin evidencia Dysflow o de prueba.
- [ ] Las configuraciones Dysflow se respetan tal cual; no se modifican en esta fase.

## Siguiente paso

Confirmar con el usuario las cinco preguntas de mayor valor de `08-decisiones-y-preguntas-abiertas.md` antes de iniciar el Lote 1 (Lanzadera).
