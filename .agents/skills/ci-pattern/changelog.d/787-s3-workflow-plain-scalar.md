type: patch

`workflow-policy` (#787): un valor en escalar plano que continúa en una línea más indentada (`run: cmd \` seguido de sus argumentos) se rechaza, como antes, con exit 2, pero ahora con la línea, la clave y la causa («YAML pliega las líneas y deja el `\` literal; use un bloque `|`») en lugar de `stray content`. APAP_WEB tuvo 35 pushes con `runner-guard` y 33 con `runner-detective` en rojo sin evaluar nada por este patrón.
