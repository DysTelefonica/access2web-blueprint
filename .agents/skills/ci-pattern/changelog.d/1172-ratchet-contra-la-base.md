type: patch

El ratchet compara contra la base del PR: mide lo que el PR AÑADE, no el total (#1172, decisión del dueño del 2026-10-24; auditoría de fricción de gates en team-skills#1093, comentario `6389037313`).

La auditoría clasificó 10 de los 170 job-runs rojos de la ventana de 48 h como fricción del carril de lint: el ratchet comparaba los hallazgos del árbol entero contra la baseline, así que un PR que no añadía ninguna violación quedaba rojo por la deuda preexistente (el caso medido: `check_test_classification` sumaba los `TRY003` que ya había al mover la política a datos).

`check_ratchet.py` acepta ahora `--base-findings <fichero>`: los hallazgos medidos en la base del PR, en el mismo formato que `--findings`.

- Un hallazgo que la base ya tenía **avisa y no bloquea**: este PR no lo añade.
- Un hallazgo que la base no tenía y la baseline no admite **sigue bloqueando** (HR-15): eso sí lo añade el PR.
- Una identidad de la baseline que el PR eliminó **sigue bloqueando**: hay que encoger la baseline para fijar la mejora. Una entrada que ni la base ni el PR tienen es una entrada vieja: avisa, no la quita este PR.
- La caducidad de una exención (`target_date` vencida) **sigue bloqueando**: ese control no se relaja.
- Sin `--base-findings` el gate se comporta como antes y declara la comparación NO realizada; con él, la declara realizada y dice cuántos hallazgos ya estaban en la base.

Tests: ocho casos nuevos sobre el gate (el hallazgo preexistente avisa, el que añade el PR bloquea, la comparación se declara, la identidad eliminada bloquea, la entrada vieja avisa, la caducidad sigue bloqueando y una base ilegible sale 2).
