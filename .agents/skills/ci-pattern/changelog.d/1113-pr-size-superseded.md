type: patch

El gate `pr-size` deja de fallar cuando la PR recibe un push nuevo o su base se mueve entre la lectura del evento y la del desglose (#1113). El gate lee el head y el desglose alrededor del listado de ficheros, reintenta una vez si el desglose no cuadra con `changed_files` y, si el head ya no es el del evento, informa «superado por un push más reciente» y sale con 0: el evento más reciente mide el head nuevo. El `changed_files` del payload deja de ser la autoridad; un desglose que no cuadra de forma persistente, con el head igual, sigue fallando cerrado.
