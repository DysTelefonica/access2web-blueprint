type: patch

`governed-delivery` 1.4: la regla 3 dice cómo se relanza el CI después de re-apuntar el PR siguiente de una cadena. Se cierra y se reabre el PR, sin commit, o se empuja un commit vacío; nunca se fusiona main en la rama (HR-13). Antes decía «empujón» sin precisar (cadete#1074, #1495).
