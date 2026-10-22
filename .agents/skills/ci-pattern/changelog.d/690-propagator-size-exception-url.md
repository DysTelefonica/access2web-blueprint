type: patch

Propagador (#690): el trozo de un único fichero por encima de `-MaxChunkLines` lleva en su cuerpo la sección «Excepción de tamaño» (asset canónico no divisible; la etiqueta sigue siendo del mantenedor), y `New-PrIfMissing` extrae la URL del PR con una expresión regular en lugar de tomar la primera línea de `gh pr create`, que con el árbol sucio es un aviso.
