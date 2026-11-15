type: patch

El porting-guide sitúa la fase 4 en el `docs-gates` del PR (`--only-phase 4 pr-head`), junto a la fase 5 (`--only-phase 5`), como ya exigía G4.2 (#1359): con solo `--only-phase 5`, ningún job de PR medía la fase 4, porque `governance-seals` llega a la 3 (DysTelefonica/team-skills#1396, cadete#799). El párrafo de #1399 sale de dentro de la tabla de fases, que volvía a mostrarse partida.
