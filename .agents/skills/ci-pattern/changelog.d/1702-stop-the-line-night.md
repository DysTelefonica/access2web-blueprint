type: patch

La parada de línea (`stop_the_line.py`, #1627) deja de contar cualquier `workflow_dispatch` como noche: una ejecución ligera en verde ya no limpia un rojo nocturno. Un dispatch cuenta solo si el consumer fija `--dispatch-marker <texto>` y el título del run lo contiene; sin esa opción ningún dispatch cuenta, y el comando de desbloqueo lo dice. Además, un workflow nocturno inexistente (404) retiene con salida 2 y un mensaje que nombra el workflow y `--nightly-workflow`, en vez de dejar pasar o retener sin explicación. `check_governed_merge.py` pasa la opción nueva. Refs #1702.
