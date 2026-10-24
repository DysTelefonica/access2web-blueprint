type: feature

Fase 1 de la propuesta de un job de gates por repo con un veredicto unico (#1179, aprobado por el dueno el 2026-10-24): el job en sombra.

Nuevo asset `assets/gates/run_gates.py`: ejecuta los gates declarados del consumer una vez, en paralelo, y agrega sus veredictos en un unico veredicto — cuantos gates estan en verde y, para cada uno que falla, su nombre y su causa. La declaracion es un fichero versionado en el repo del consumer (`nombre: comando` por linea). Exit: 0 todos en verde · 1 algun gate en rojo · 2 declaracion rota o un gate que no se ejecuto (timeout o crash): un gate que no midio nunca dice «OK» (HR-3).

El fragmento del porting-guide publica el job en sombra: su check lo publica y NADIE lo exige (no entra en `required_jobs` ni en el `needs:` del agregador), su nombre lleva «(sombra)» y su techo cubre el gate mas lento. La sombra duplica el trabajo de los gates: es el precio de medir antes de decidir, y es temporal.

Tests: `assets/gates/tests/test_run_gates.py` cubre el veredicto (todos en verde, alguno en rojo con su causa, un gate que no midio sale 2), la salida `--json` y las cinco formas de declaracion rota. Once casos, sin red.
