type: minor

El gate `workflow-policy` rechaza ahora las claves duplicadas de un mapeo (salida 2: GitHub rechaza el fichero con `startup_failure` y el check desaparece del PR, mientras que un lector que se queda con el último valor daba OK) y exige `timeout-minutes` en todo job, requerido o no (el valor por defecto de GitHub son 360 minutos). Un job que llama a un workflow reutilizable queda exento. Los consumers cuyos jobs no declaren timeout verán un hallazgo nuevo (#806).
