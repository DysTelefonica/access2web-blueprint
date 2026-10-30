type: patch

R8 de #1093: HR-54 deja de ser un único párrafo de unos 6.000 bytes y se reparte en dieciséis entradas, una por requisito (HR-54.1 a HR-54.16). Cada entrada lleva su gate entre paréntesis, o la marca de juicio en revisión si no tiene gate ejecutable. Ningún requisito cambia de fondo, ninguno se retira en esta PR y el número `HR-54` se conserva, así que las citas del resto del canon siguen válidas. `build-governance-partial.py --check` sigue en verde: las 50 HR vivas no cambian.
