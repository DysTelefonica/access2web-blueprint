type: minor

La parada de línea del merge gobernado (`stop_the_line.py`, #1627) admite que el consumer
apague cada señal por separado desde su `.github/host-contract.json`: la clave opcional
`stop_the_line` con un miembro booleano por señal (`night_red_enabled` para
`prerelease:night-red`, `main_red_enabled` para `prerelease:main-red`). Sin la clave, o
sin un miembro, la señal sigue activa y el comportamiento es el de hoy. Una señal apagada
no se consulta (el gate no lee su etiqueta) ni bloquea el merge, y el merge lo dice en una
línea por `stderr`. Una forma que no sea un objeto con esos miembros hace fallar cerrado
el merge (exit 2): una declaración rota que se ignorase dejaría la puerta en verde sin
ella. La forma la valida el merge gobernado
(`assets/runner-controls/check_governed_merge.py`), no el readback del host. Refs #1784.
