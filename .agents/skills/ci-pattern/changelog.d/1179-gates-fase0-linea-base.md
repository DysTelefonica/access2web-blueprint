type: feature

Fase 0 de la propuesta de un job de gates por repo con un veredicto unico (#1179, aprobado por el dueno el 2026-10-24): el canon mide la linea base de friccion y la publica, para que el flip de la fase 2 sea una comparacion y no una opinion.

Nuevo asset `assets/gates/measure_gates.py`: sobre un export de check-runs (JSON, sin red) calcula los rojos secundarios (los checks rojos que no son el primero del commit), los gates rojos por commit, los commits con re-ejecucion y el tiempo hasta el verde — la cifra que la auditoria de #1093 dejo a medias. Excluye el agregador `required` (que repite el rojo de la hoja) y los `cancelled`/`skipped`, y mide la primera onda de la CI: una re-ejecucion posterior no es la espera del merge.

Tests: `assets/gates/tests/test_measure_gates.py` cubre las reglas de cada cifra con fixtures que violan una y el contrato de salida (export ilegible o sin commits -> exit 2, HR-3). Doce casos, sin red.
