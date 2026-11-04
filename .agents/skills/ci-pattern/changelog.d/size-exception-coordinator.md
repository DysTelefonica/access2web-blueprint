type: major

**Qué.** La etiqueta `size:exception` la pone el coordinador del proyecto (decisión del dueño, 2026-11-02 17:09Z): nunca el worker y sin esperar a una persona, tras verificar que el corte no es posible, registrarlo y con las cuatro respuestas en el cuerpo de la PR. Lo dicen HR-87, la sección de principios, la tabla de puertas, la frontera con `deterministic-quality-harness`, la plantilla `CONTRIBUTING.md.tmpl` y los mensajes de `pr-size` y `pr-contract`.
**Causa.** HR-87 lo dejaba pendiente y el resto del canon decía «solo el mantenedor», así que la cola esperaba a una persona por cada excepción.
**Arreglo.** Todos esos textos dicen ahora que la pone el coordinador del proyecto. El gate no cambia: sigue leyendo la etiqueta.
**A partir de ahora.** El coordinador decide y aplica la excepción en el momento.

Migración: el coordinador de cada consumidor pone `size:exception` él mismo en las PR que la necesiten, en lugar de pedirla al mantenedor.
