type: minor

Propagador (#679): nuevo parámetro `-MaxChunkLines N`, que parte el tramo vendorizado en PRs encadenados de como máximo N líneas (cada trozo desciende del anterior, posiciones `1/total` a `total/total`). Un fichero que por sí solo supera N va en su propio PR, con aviso. Con 0 (por defecto) el tramo 2 sigue siendo un solo PR.
