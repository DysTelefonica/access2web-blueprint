type: patch

Los fragmentos de `changelog.d/` de #1179 vuelven al contrato vigente del consolidador: `type: minor` para los dos assets nuevos de gates y `type: patch` para el informe (#1232). La suite del consolidador comprueba además que todos los fragmentos reales del skill respetan ese contrato antes de que una release vuelva a salir con código 2.
