type: patch

La fase 3 de `adoption check` puede medir además contra la BASE del PR (`--seal-baseline <dir>`): el conjunto que hay que sellar es la unión de los `pattern_paths` de la base y los de la cabeza, de modo que un PR no puede reducir su propia base de sellado mientras el fichero siga en el árbol (DysTelefonica/team-skills#1317). Retirar un workflow sí se permite: si el fichero ya no existe, no hay nada que sellar. La opción es opt-in —sin ella el gate se mide como siempre— y una base sin contrato legible falla cerrado con exit 2.
