type: patch

Las plantillas y el contrato del host del patrón dejan de declarar burocracia que ningún gate exige (#1093, R1 y R2).

`assets/templates/docs/CONTRIBUTING.md.tmpl` deja de decir que una PR con dos `type:*` falla cerrado y que `status:approved` la aplica solo un mantenedor sin cuya etiqueta no hay rama ni PR: la etiqueta `type:*` es una convención de clasificación y `status:approved` la puede poner cualquiera, agentes incluidos (decisión del dueño del 2026-10-23, 17:24, en los comentarios de #1093). El flujo de PR bloquea por tres cosas: que la PR enlace una issue, el tamaño (~400 líneas, `size:exception`) y que los commits sean convencionales; HR-53 y HR-64 avisan.

`assets/templates/README.md` deja de llamar «campo de datos» a la excepción de tamaño —es la etiqueta `size:exception`— y «obligatorias» a las seis secciones de la issue, que desde R2 son la plantilla del repo y no una puerta.

`assets/host-readback/host-contract.schema.json` marca `approvers` como retirado: desde R1 el gate de `issue-spec` no lo lee y `status:approved` no la exige ningún gate. La clave se conserva para que los contratos de los consumers que la declaran sigan siendo válidos.
