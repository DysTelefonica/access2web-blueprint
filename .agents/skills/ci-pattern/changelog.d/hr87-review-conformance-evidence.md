type: minor

HR-87: el coordinador revisa la conformidad con la issue, no la CI. Cada criterio de aceptación lleva su evidencia en la PR, una línea por criterio, y sin ella el criterio no está hecho. La excepción de tamaño se pide antes del push, la decide el coordinador y se aplica con la etiqueta `size:exception` del gate `pr-size`, el único mecanismo de la flota (principio del dueño, 2026-11-02). Quién pone la etiqueta `size:exception` sigue pendiente de decisión del dueño (hoy, el mantenedor).
