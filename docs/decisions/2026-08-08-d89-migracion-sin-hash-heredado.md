# D89 — La migración descarta los hashes heredados

- estado: adoptada
- fecha: 2026-08-08
- decide: usuario (dueño del producto)
- motivo: Ningún usuario migrado conserva hash: empieza con `password_hash=NULL` y `status='password_reset_required'`, sin columna `legacy_hash`.
- origen: `product/lanzadera-no-legacy-hash`
