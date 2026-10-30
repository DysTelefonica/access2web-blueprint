# Línea base de fricción de gates (#1179, fase 0)

La línea base que decide el flip de la fase 2 (un job de gates por repo con un
veredicto único). Sin esta cifra, la fase 2 sería una opinión: con ella, es
una comparación.

El asset que la calcula es `assets/gates/measure_gates.py`; este documento es
el informe con la medición, el método para repetirla y los criterios de
decisión.

## Qué mide

| Métrica | Qué es | Por qué decide |
|---|---|---|
| **Rojos secundarios** | Por commit, los checks rojos que no son el primero | La redundancia: el mismo defecto reportado N veces (65 de 129 episodios, el 50 %) |
| **Gates rojos por commit** | El peor caso | El techo de ruido que ve quien abre el PR (7 checks en `cadete#631`) |
| **Commits con re-ejecución** | Los que se pusieron verdes sin commit nuevo | El tiempo de espera que no es trabajo (67 de 129, el 52 %) |
| **Tiempo hasta el verde** | Latencia de la CI: del primer run al último verde de la primera onda | Lo que un merge espera; la auditoría **no** lo midió |

El asset NO cuenta el check agregador (`required`, que repite el rojo de la
hoja) ni los `cancelled`/`skipped` (concurrencia o filtro de paths), y mide la
**primera onda** de la CI: una re-ejecución posterior no es la espera del merge.

## Método (reproducible)

1. Los commits de los PRs actualizados en la ventana, con sus check-runs. La
   API de solo lectura, un `gh api` por commit:

   ```bash
   gh api "repos/<owner>/<repo>/pulls?state=all&sort=updated&direction=desc&per_page=100"
   gh api "repos/<owner>/<repo>/pulls/<n>/commits?per_page=100"
   gh api "repos/<owner>/<repo>/commits/<sha>/check-runs?per_page=100"
   ```

2. El export: `{"repo", "window", "commits": [...]}`; cada commit lleva `sha`,
   `pr`, `committed_at` y sus `checks` (nombre, conclusión, `started_at`,
   `completed_at`).

3. Las métricas: `python3 assets/gates/measure_gates.py --export <export.json>`.

## La medición

Ventana de siete días (2026-10-17 → 2026-10-24), con el asset:

| Métrica | team-skills | cadete |
|---|---:|---:|
| Commits medidos (con check-runs) | 605 (452) | 480 (332) |
| Commits con rojo | 93 (20,6 %) | 101 (30,4 %) |
| **Rojos secundarios** | **9** | **97** |
| Gates rojos por commit (peor) | 3 (`#1103`) | 5 (`#559`) |
| Commits con re-ejecución | 154 (34,1 %) | 118 (35,5 %) |
| Tiempo hasta el verde (mediana · p90 · máx) | 3,7 · 6,7 · 20,1 min | 3,9 · 8,4 · 15,7 min |

La diferencia entre los dos repos es el argumento del flip: team-skills
publica tres checks por PR y apenas tiene redundancia (9 rojos secundarios en
605 commits); cadete publica más gates —algunos con dos nombres— y la
redundancia se multiplica (97 en 480 commits: **uno de cada cinco commits con
rojo tiene más de un check en rojo**).

BP y APAP_WEB no entran: su export completo son miles de llamadas a la API.
**Cada consumer mide su propia línea base con este asset antes de su flip**:
la cifra que decide el flip es la de ese repo, no la de otro.

La auditoría de #1093 (305 PRs, cuatro repos, 48 h) ya midió la redundancia y
los gates rojos por commit: **143 de 170 job-runs rojos eran fricción (84 %)**,
11 defecto real (6 %) y 16 infraestructura (9 %).

## Criterios para el flip (fase 2)

La fase 2 (cambiar required checks y protección de rama) **no se decide con
esta cifra**: se decide con la sombra (fase 1). Estos son los umbrales que la
fase 0 deja fijados, para que la comparación sea de dos cosas medidas con la
misma herramienta:

1. **El veredicto de la sombra tiene que coincidir con el rollup actual**: en
   cada commit, el conjunto de gates que la sombra declara en rojo es el mismo
   que el de checks rojos del rollup (criterio C6 de #1179). Si no coincide, la
   sombra miente y no se llega a la fase 2.
2. **El tiempo hasta el verde no puede crecer**: la sombra añade un job; si la
   mediana o el p90 suben, el job único no está pagando su ahorro.
3. **Los rojos secundarios tienen que bajar** cuando la fase 2 los elimine: si
   tras el flip siguen siendo secundarios, el cambio de required checks no
   compró nada y se deshace.

El cambio de required checks y de protección de rama va en el **mismo PR** que
el flip (criterio C5): sin las dos cosas, el repo queda sin candado.

## Límites de la medición

- Siete días no son dos meses: la línea base cambia con el ritmo de los merges,
  así que hay que repetirla antes del flip.
- El tiempo hasta el verde se mide desde el **inicio del primer run**, no desde
  el push: el commit se crea antes de empujarse y la latencia humana no es CI.
  El tiempo en cola del runner no es observable desde la API.
- Los commits sin check-runs (jobs filtrados por paths, o commits anteriores
  al workflow) no entran en las métricas de rojo: sin veredicto no hay nada
  que atribuir.
- La auditoría tuvo que inferir el 81 % de las causas (los logs de tres repos
  devolvieron 403); estas cifras salen de los check-runs, que sí son legibles.
