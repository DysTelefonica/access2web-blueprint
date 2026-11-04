type: patch

`agent-operating-rules` 1.6: HR-13 prohíbe fusionar `main` en la rama o pulsar «Update branch» con `merge_freshness` en `off`, salvo conflicto (`DIRTY`) o solape de `.github/workflows/` (#783). En APAP_WEB, cada uno de esos merges costaba unos 25 minutos de CI por eslabón (#3358, #3371), y el merge gobernado ya fusionaba la PR atrasada (#3351).
