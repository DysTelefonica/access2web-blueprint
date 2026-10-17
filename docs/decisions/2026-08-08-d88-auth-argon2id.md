# D88 — Autenticación con Argon2id

- estado: adoptada
- fecha: 2026-08-08
- decide: usuario (dueño del producto)
- motivo: Sustituye el hash heredado sin sal por un esquema moderno: Argon2id vía `argon2-cffi==25.1.0`, perfil `RFC_9106_LOW_MEMORY` (64 MiB, 3 iteraciones, 4 hilos).
- origen: `product/lanzadera-auth-modern-crypto`
