type: patch

El sello de gobernanza desalineado dice el hash esperado y el comando para re-sellar (#1172, decisión del dueño del 2026-10-24; auditoría de fricción de gates en team-skills#1093, comentario `6389037313`).

La auditoría clasificó 13 de los 170 job-runs rojos de la ventana de 48 h como fricción del gate `governance-seals`, todos por la misma causa: un fichero gobernado cambió, el manifiesto no se re-selló y el mensaje decía «difiere del sha256 del manifiesto de gobierno» sin más. Quien tenía el rojo no tenía ni el hash que poner ni el comando que ejecutar, y el re-sellado acababa a mano (un commit titled «reseal» en cada repo).

Tres cambios, ninguno mueve el veredicto: el sello sigue bloqueando.

1. El mensaje del gate nombra el hash **esperado**, el hash **encontrado** y el comando exacto `ci-pattern manifest seal <repo>`. Lo dice la fase 3 (`adoption check --phase 3`, la que corre el job `governance-seals`) y también `verify`.
2. El comando `manifest seal <repo>` existe: re-sella lo que YA está declarado (el sha256 de cada fichero de `files` y de cada bloque de slice de `slice_blocks`). No es un auto-firmado: no añade entradas ni inventa contenido, y un fichero declarado que no existe en el árbol no se re-sella —sale 1 nombrándolo—. Con `--out <ruta>` escribe la copia re-sellada y no toca el árbol, que es lo que usa el job.
3. El fragmento del job `governance-seals` que publica el porting-guide deja el manifiesto re-sellado como artefacto `governance-manifest-resealed` cuando el gate falla (`if: failure()`, sin `|| true` ni `continue-on-error`, y sin push: descargarlo y commitearlo es de quien abre el PR).

Tests: la CLI cubre el re-sellado (actualiza el hash, `--out` no toca el árbol, un fichero declarado ausente no se re-sella, el re-sellado es idempotente, y el mensaje del drift nombra los dos hashes y el comando); el fragmento del job se mide sobre el porting-guide (re-sella sin push, los dos pasos solo corren con `if: failure()`, la acción va pineada a 40 hex y el veredicto del gate no se traga).
