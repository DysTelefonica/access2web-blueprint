type: patch

`assets/ratchet/check_baseline_relaxation.py` ya no salta en silencio una baseline `*BASELINE*` cuyo valor no es literal: la avisa por stderr y declara que el gate no está protegido. La comparación de justificaciones borradas pasa a usar la pareja (entrada, texto), de modo que borrar el comentario de una entrada se rechaza aunque otra entrada conserve el mismo texto (#1084).
