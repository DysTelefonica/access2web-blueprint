type: minor

Nuevo asset `assets/ratchet/check_baseline_relaxation.py`: compara con la rama base las baselines `*BASELINE*` declaradas como constantes en los scripts de gate del consumer y rechaza toda relajación (cifra que sube, entrada o elemento nuevo) que no lleve un campo `RELAXATION` nuevo con motivo, referencia e identidad (HR-15, HR-33). Bajar o quitar entradas siempre pasa; sin base legible falla cerrado (#974, Refs #429).
