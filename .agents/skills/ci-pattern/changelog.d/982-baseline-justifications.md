type: minor

`assets/ratchet/check_baseline_relaxation.py` rechaza también el borrado de las justificaciones de una baseline en código: un comentario `#` dentro de una constante `*BASELINE*` que desaparece (o se reescribe) mientras su entrada sigue existiendo exige un campo `RELAXATION` nuevo, igual que subir la cifra. Retirar una entrada con su comentario, o mover el comentario sin cambiar su texto, pasa (#982, Refs #974).
