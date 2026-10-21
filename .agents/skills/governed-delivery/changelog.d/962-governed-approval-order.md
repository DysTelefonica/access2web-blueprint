type: patch

La regla 2 admite que un agente aplique `status:approved` o lance el merge gobernado por orden explícita de un mantenedor, siempre que la cuenta de `gh` tenga permiso de mantenedor y la orden quede registrada en un comentario de la issue o del PR antes de actuar. Sin orden registrada, la autoaprobación sigue prohibida. Sin cambios en `metadata.version` ni en ningún gate (#962).
