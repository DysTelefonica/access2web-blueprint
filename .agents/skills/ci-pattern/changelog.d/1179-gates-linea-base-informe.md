type: patch

El informe de la linea base de fricción de gates que decide el flip de la fase 2 (#1179, fase 0).

`references/gates-baseline.md` publica, con el asset `assets/gates/measure_gates.py`, las cifras que la auditoria de #1093 dejo a medias: el metodo para producir el export de check-runs y repetir la medicion, los rojos secundarios, los gates rojos por commit, las re-ejecuciones y el tiempo hasta el verde de los dos repos medidos (team-skills: 605 commits, 9 rojos secundarios, mediana de 3,7 min; cadete: 480 commits, 97 rojos secundarios, mediana de 3,9 min), los tres criterios que decidiran el flip de la fase 2 (el veredicto de la sombra tiene que coincidir con el rollup, el tiempo hasta el verde no puede crecer, y los rojos secundarios tienen que bajar) y los limites de la medicion.

La cifra que decide el flip es la de cada repo: cada consumer mide su propia linea base con este mismo asset antes de su fase 2.
