type: patch

Propagador (#960): el `- base:` del Chain Context de un PR ya abierto se recalcula contra la base real del PR (`base.ref` en vivo, consulta separada) antes del PATCH. Antes se congelaba al componer el cuerpo: tras fusionarse el tramo 1 y retargetear el tramo 2 a `main`, HR-54 rechazaba el cuerpo porque la base declarada quedaba desfasada. Sin retarget, o si la base coincide, el cuerpo no cambia.
