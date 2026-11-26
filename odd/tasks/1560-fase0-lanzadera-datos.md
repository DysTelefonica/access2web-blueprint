# 1560 — Fase 0 de coexistencia: inventario de Lanzadera_Datos

Issue #1560. Épica `docs/07-migracion/epica-coexistencia-postgresql.md` (fase 0, decisión 12).
Rama `docs/1560-fase0-inventario-lanzadera-datos`. Solo lectura sobre backends; nada se escribe en `.accdb`.

## Specs

- S1. "la fase 0" — inventario de solo lectura de `Lanzadera_Datos` según la fase 0 de la épica.
- S2. Script reproducible de solo lectura sobre una copia temporal: tablas, campos (tipo, tamaño, obligatoriedad, defecto, validación, tipos complejos), índices (primario, único, campos), relaciones, recuento de filas, vínculos con origen. Reutilizable para los demás backends.
- S3. Ficha por tabla de las 35 tablas: clave primaria o su ausencia, índices, tipos problemáticos, recuento.
- S4. Consumidores por tabla: `TbTablasAVincular` y vínculos de los backends disponibles en `C:\00repos\datos`.
- S5. Huérfanos en las relaciones.
- S6. Recuento de referencias en código por tabla (front de la Lanzadera y fronts consumidores disponibles).
- S7. Clasificación de cada tabla (propia, propia y compartida, mal ubicada, mal ubicada con dueño por decidir, huérfana) y verificación de R-1 y R-2.
- S8. Los patrones de VBA del front se enlazan desde `coexistencia.md`, no se repiten.
- S10. «de aquí tienen qué salir accionables para que yo aplique a los backends ya mismo»: lista de acciones con comprobación previa y reversión.
- S9. La contraseña del backend nunca se escribe en el repositorio ni en la salida.

## Tasks

| ID | Specs | Route | Commit |
|---|---|---|---|
| T1 | S2, S9 | inline | commit de `extract-erd.ps1` |
| T2 | S3, S4, S5 | inline | commit del inventario |
| T3 | S6 | inline | commit del inventario |
| T4 | S1, S7, S8, S10 | inline | commit del inventario |

## Log

- L1 (2026-11-04, usuario): «la fase 0» — en respuesta a «¿Empezamos con la fase 0 de la Lanzadera o con la #1559?».
- L2 (2026-11-04): `get_schema` no expone claves ni índices; se usa DAO de solo lectura sobre una copia.
- L3 (2026-11-04, usuario): «de aquí tienen qué salir accionables para que yo aplique a los backends ya mismo».
- L4 (2026-11-04): los vínculos con origen inaccesible (`\datoste`) hacían fallar el script al leer los índices; ahora se registran en `linkError`. Regresión: la salida sobre `Lanzadera_Datos` es idéntica a la anterior.
- L5 (2026-11-04): T3 y T4 se hicieron inline porque solo esta sesión tiene dysflow y la evidencia estaba en `%TEMP%`.
