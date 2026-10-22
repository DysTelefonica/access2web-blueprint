type: minor

HR-68 (#805): el bloque `architecture` con `paths` admite `overrides` (`[{"path", "layer"}]`): una ruta exacta decide su capa antes que los patrones, para consumers con ficheros planos que no siguen la estructura de su capa (APAP_WEB). `exempt` sigue ganando; una ruta inexistente, una capa no declarada o una entrada mal formada es rojo. Sin `overrides` nada cambia.
