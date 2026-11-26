# Alcance de la migración: las trece aplicaciones con icono en la Lanzadera

- estado: adoptada
- fecha: 2026-10-28
- decide: usuario (dueño del producto)
- motivo: D85 dejaba en alcance solo ocho IDs de `TbAplicaciones`, pero la Lanzadera muestra más aplicaciones a los usuarios. Para que Calidad decida qué se conserva antes de migrar a web, el inventario tiene que cubrir todas las aplicaciones que la Lanzadera ofrece hoy.
- origen: issue #1494

## Decisión

Entran en el alcance todas las aplicaciones con botón en las pestañas «En oficina» (`Form_FormAplicacionesOficina`) y «Fuera de la oficina» (`Form_FormAplicacionesFueraOficina`) de la Lanzadera, más la propia Lanzadera. Quedan fuera las de la pestaña de pruebas y las que el dueño del producto retira.

| ID | Aplicación | Repositorio |
|---|---|---|
| 12 | Lanzadera | `DysTelefonica/Lanzadera` |
| 5 | Gestión de Riesgos | `DysTelefonica/GESTION_RIESGOS` |
| 8 | No Conformidades | `DysTelefonica/No_conformidades` |
| 23 | Condor | `DysTelefonica/condor` |
| 22 | Solicitudes HPS | `DysTelefonica/HPS_SOLICITUDES` |
| 17 | HPS | `DysTelefonica/HPS` |
| 6 | Brass | `DysTelefonica/BRASS` |
| 19 | Expedientes | `DysTelefonica/Expedientes` |
| 16 | Seguridad | `DysTelefonica/seguridad` |
| 18 | Gestión Contratos | `DysTelefonica/CONTRATOS` |
| 3 | AGEDYS técnico | `DysTelefonica/AGEDYS` |
| 2 | AGEDYS económico | `DysTelefonica/AGEDYS_ECO` |
| 10 | Registro | `DysTelefonica/REGISTROS` |

Se retiran por decisión del dueño del producto AGEDO (ID 4) y SICA (ID 20). No tienen botón en esas pestañas SIRDEE (1) ni AGEDYS (21), ni las copias de pruebas (81, 101, 191 y 231).

El código de referencia de cada aplicación es su rama rc1 si tiene trabajo que todavía no está en `main`; si no, `main`. Los backends siguen en `C:\00repos\datos\`, como fijaba D85.
