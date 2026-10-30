type: patch

El gate issue-spec vuelve a aceptar `--required-sections` y `--host-contract` como flags obsoletas que se ignoran con un aviso por stderr y no cambian el veredicto. Esta corrección de compatibilidad de #1103 evita que la propagación rompa el pr-gate de los consumers.
