type: patch

Regla 3 de `governed-delivery`: la cadena cabe en tres líneas — eslabón: `Refs #N` en el cuerpo y en los commits, con la etiqueta `chain:partial`; punta contra `main`: `Closes #N` y sin `chain:partial`. El ejemplo va con la regla: `fix(x): y (Refs #N)` es un eslabón y `fix(x): y (Closes #N)` es la punta. Se añade que una palabra de cierre en el cuerpo **o en un commit** cierra la issue al fusionar, así que en un eslabón la cierra antes de tiempo, y que el merge gobernado bloquea y nombra el commit y la línea que la rompen (#1758).
