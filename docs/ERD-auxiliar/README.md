# ERD Auxiliar — Backends compartidos

Schemas de backends auxiliares que no pertenecen a una app específica sino que son compartidos por varias apps del ecosistema legacy.

| Archivo | Tablas locales | Tablas vinculadas | Uso |
|---|---|---|---|
| AGEDYS_DATOS.sql | 104 | 35 | Sistema AGEDYS (gestión de expedientes antiguos) |
| AGEDO20_Datos.sql | 66 | 2 | AGEDO 2.0 (registro de documentos) |
| Seguridad_datos.sql | 36 | 0 | Seguridad y control de acceso |
| GestionContratos_datos.sql | 20 | 0 | Gestión de contratos |
| Registro_Ent_Salida_Datos.sql | 11 | 1 | Registro de entradas/salidas |
| Registro_Datos.sql | 5 | 0 | Registro general |
| Tareas_datos1.sql | 4 | 13 | Tareas programadas |
| SICA_datos.sql | 4 | 0 | SICA (control de acceso) |
| Control_Cambios_datos.sql | 3 | 0 | Control de cambios |
| Correos_datos.sql | 1 | 0 | Sistema de correos |

## Origen

Extraídos con Jackcess 4.0.7 + jackcess-encrypt 4.0.2 desde `C:\00repos\datos\` (password legacy `dpddpd`). Las tablas vinculadas (linked tables) que apuntan a otros `.accdb` se marcan con `-- LINKED/SKIPPED` porque sus paths Windows no son resolubles desde Linux.

## Nota

Estos schemas son **read-only reference**. No se modifican. Si el backend cambia en producción, re-extraer con el mismo proceso.
