type: patch

Registro (#1088): el presupuesto de revisión del gate `pr-size` ya no excluye `skills/`. En el canon `skills/` es el catálogo de autoría (64 skills escritas aquí) y en un consumer es su catálogo propio, así que sus líneas cuentan para el presupuesto; solo queda fuera el árbol vendorizado `.agents/skills/`. Antes, una PR que reescribía una skill de 569 líneas salía en verde con `changed_lines: 0` y `vendored_lines: 569` (auditoría de gobierno del 2026-10-23, hallazgo A2).
