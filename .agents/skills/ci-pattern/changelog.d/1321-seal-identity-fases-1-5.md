type: patch

`--seal-baseline <dir>` mide también las fases 1 y 5 contra la BASE del PR, con el operador de IDENTIDAD (DysTelefonica/team-skills#1321). La evidencia (`evidence.path`, fase 1) y el documento operativo con el fichero que lo enlaza (`operating_doc.doc` y `operating_doc.linked_from`, fase 5) tienen que ser los que mide la base: el PR no puede apartar el gate cambiando lo que el gate lee. La retirada real de una ruta sigue permitida, y una base sin contrato legible sigue fallando cerrado con exit 2. Sin la opción el gate se mide como siempre.
