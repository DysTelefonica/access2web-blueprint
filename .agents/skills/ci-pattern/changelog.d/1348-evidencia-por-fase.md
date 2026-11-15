type: patch

`--seal-baseline <dir>` guarda la ruta de evidencia de la base POR FASE (DysTelefonica/team-skills#1348). La fase 1 y la fase 3 declaran documentos distintos, y cada una se compara con la de su fase en la base: con una sola ruta para las dos, la fase 3 se comparaba contra la declaración de la fase 1 y el gate daba un hallazgo falso, incluso con la base y la cabeza idénticas. Un cambio real de la evidencia de cada fase sigue dando hallazgo, nombrando la ruta de su fase en la base. Sin la opción el gate se mide como siempre.
