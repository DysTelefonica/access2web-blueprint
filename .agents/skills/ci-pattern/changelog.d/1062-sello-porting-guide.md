type: minor

#1062: el commit del catálogo deja de declararse dos veces. El sello `# Auto-generado por propagate-team-skills.ps1 (catalog @ <sha40>).` de `.team-skills.yaml`, que el propagador reescribe en cada corrida, es la fuente única; lo que se llevaba escrito a mano quedaba viejo en la propagación siguiente, con un `diff` de 75 líneas en el `ref:` de un checkout.

Este eslabón cierra la segunda puerta: `references/porting-guide.md` deja de enseñar `pin: <sha literal>` en el job `governance-seals`. El fragmento resuelve el sello y hace el checkout con el output del paso, lo que sustituye a la acción `pattern-at-pin` que el canon publicaba sin existir en ningún repo. El paso final recibe su directorio por `env`, que hasta ahora el fragmento usaba sin declarar.

Dos medidas nuevas en `assets/templates/tests/test_governance_seals_job.py`: el fragmento no declara ningún sha de 40 hex escrito a mano —con su caso negativo, que mete el `pin:` que el canon publicaba y comprueba que el detector lo ve— y su `ref:` es el output del paso que lee el sello. Las ocho aserciones preexistentes del fragmento siguen verdes después de reescribirlo.