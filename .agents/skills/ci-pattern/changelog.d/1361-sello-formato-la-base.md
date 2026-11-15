type: patch

El formato del sello de gobierno lo negocia ahora la BASE del consumer (DysTelefonica/team-skills#1361). La base mide el PR con su canon —el de su rama base, el del sello `catalog @`—: un manifiesto en un formato que la base no mide es un rojo (rc=2) en el gate de la base, aunque el canon de la cabeza lo entienda (DysTelefonica/cadete#733, ronda 15). Que la base juzgue al PR es deliberado y no cambia: el PR no puede elegir el canon que lo mide.

El canon vigente mide ahora los dos formatos vivos, cada uno con su algoritmo: el hash se calcula con el formato que declara el manifiesto, no con el del canon que lo lee. El formato 2 (workflows sellados con el `uses:` normalizado y la imagen en crudo) vuelve a ser medido; el formato 1 sigue sin medirse y `ci-pattern manifest seal` lo migra.

`ci-pattern manifest seal` acepta `--schema-version N` (2 o 3) y sella en ese formato, con el algoritmo de ese formato; un formato que el canon no mide es error de uso (exit 2) y el manifiesto no se toca.

La ronda de propagación lee el `SCHEMA_VERSION` del canon vendorizado en la rama base del consumer y sella en el mayor formato que miden la base y el catálogo: con una base en el formato 2 la ronda sella 2 y vendoriza el canon nuevo, que lee 2 y 3; en la ronda siguiente, con la base ya en el formato 3, la ronda sella 3. Un consumer que no vendoriza el patrón se sella con el formato del catálogo.

La migración sigue comprobando antes los hashes del formato viejo: una deriva no se esconde al cambiar de formato, y `--reeseal` sigue siendo la vía de quien declara que cambia ese fichero.
