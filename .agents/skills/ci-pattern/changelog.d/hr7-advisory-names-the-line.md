type: patch

**Qué.** El aviso de HR-7 de `pr-contract` para un PR con `chain:partial` que lleva una palabra de cierre ahora cita la línea, y dice que el merge gobernado lo bloquea.
**Causa.** En cadete#1031, un `follow-up: … closes #1029` dentro de Chain Context hizo que el merge gobernado rechazara la PR, con razón: GitHub cierra la issue al fusionar en `main`, esté donde esté la palabra. `pr-contract` ya avisaba (HR-7 es aviso desde #1172), pero sin decir dónde, y el aviso pasó desapercibido hasta el merge.
**Arreglo.** El aviso lista las líneas del cuerpo con la palabra de cierre. Test: `test_chain_partial_con_closes_en_el_cuerpo_cita_la_linea`.
**A partir de ahora.** En un tramo intermedio, el tramo siguiente se describe sin palabras de cierre (`follow-up: tramo 2, Refs #N`). Pasar este aviso a bloqueo, para que coincida con el merge gobernado, queda como decisión del dueño, porque revierte #1172.
