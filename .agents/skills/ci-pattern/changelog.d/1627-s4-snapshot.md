type: patch

**Qué.** `assets/prerelease/snapshot.py` carga y valida la instantánea de la noche (#1627, slice 4). Una instantánea que no cumple el contrato sale con 2 y nombra la clave o el campo, sin traceback: falta una clave, una fecha no es válida o le falta la zona horaria, o el fichero no existe.
**Causa.** El informe nocturno caía con un traceback y salida 1 ante una instantánea incompleta.
**Arreglo.** Validación y carga en un módulo propio, con un test por caso del contrato.
**A partir de ahora.** El informe nocturno valida la instantánea con este módulo antes de renderizar.
