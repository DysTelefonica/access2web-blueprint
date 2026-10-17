# Registro de decisiones

Una decisión por fichero `AAAA-MM-DD-<slug>.md`. Este índice lista exactamente las decisiones del directorio; lo comprueba `tests/test_decisions_registry.py`.

## Formato

Tras el título, una lista con estas claves:

| Clave | Obligatoria | Valor |
|---|---|---|
| `estado` | Sí | `adoptada`, `rechazada`, `sustituida` o `aplazada` |
| `fecha` | Sí | Fecha ISO `AAAA-MM-DD` |
| `decide` | Sí | Rol o persona que decide |
| `motivo` | Sí | Por qué; no puede estar vacío |
| `sustituida-por` | Si `sustituida` | Slugs separados por comas de decisiones existentes; no a sí misma ni en ciclo |
| `revision` | Si `aplazada` | Condición o fecha de revisión |
| `origen` | No | De dónde procede (topic_key, issue o PR) |

## Decisiones

| Decisión | Estado | Fecha | Decide |
|---|---|---|---|
| [D36](2026-08-08-d36-preservar-hashes-heredados.md) | sustituida | 2026-08-08 | usuario (dueño del producto) |
| [D37](2026-08-08-d37-rehash-transparente.md) | sustituida | 2026-08-08 | usuario (dueño del producto) |
| [D88](2026-08-08-d88-auth-argon2id.md) | adoptada | 2026-08-08 | usuario (dueño del producto) |
| [D89](2026-08-08-d89-migracion-sin-hash-heredado.md) | adoptada | 2026-08-08 | usuario (dueño del producto) |
| [PRs siempre contra main](2026-10-15-prs-siempre-contra-main.md) | adoptada | 2026-10-15 | dueño del repositorio |
| [El registro de decisiones debe ser completo, sin baseline](2026-10-17-registro-de-decisiones-completo.md) | adoptada | 2026-10-17 | dueño del repositorio |

### D1–D29 migradas desde 08

| Decisión | Estado | Fecha | Decide |
|---|---|---|---|
| [D1](2026-08-04-d1-el-alcance-del-descubrimiento-se-limita.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D2](2026-08-04-d2-la-jerarquia-de-fuentes-de-autoridad-es.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D3](2026-08-04-d3-la-fase-1-se-ejecuta-en-9-lotes-con.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D4](2026-08-04-d4-las-configuraciones-dysflow-con-o.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D5](2026-08-04-d5-plataforma-unica-modular-permission.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D6](2026-08-04-d6-navegacion-app-first-anidada.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D7](2026-08-04-d7-lanzadera-solo-administracion.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D8](2026-08-04-d8-hexagonal-global-en-plataforma-y-modulos.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D9](2026-08-04-d9-autenticacion-inicial-email-password.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D10](2026-08-04-d10-siteminder-ocp-jwt-unificado.md) | aplazada | 2026-08-04 | usuario (dueño del producto) |
| [D11](2026-08-04-d11-notificacion-unificada-como-servicio.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D12](2026-08-04-d12-notificacion-v1-solo-email.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D13](2026-08-04-d13-cola-de-correo-por-tabla-como-adaptador.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D14](2026-08-04-d14-postgresql-compartido-con-esquemas-por.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D15](2026-08-04-d15-rendimiento-y-cache-deliberada-como.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D16](2026-06-13-d16-adjuntos-en-object-storage-s3-compatible.md) | adoptada | 2026-06-13 | usuario (dueño del producto) |
| [D17](2026-08-04-d17-adjuntos-sin-versionado-de-contenido.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D18](2026-08-04-d18-adjuntos-con-papelera-de-retencion.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D19](2026-08-04-d19-retencion-de-papelera-30-dias.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D20](2026-08-04-d20-restauracion-de-papelera-por-borrado-o.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D21](2026-08-04-d21-roles-administrador-global-admin-de.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D22](2026-08-04-d22-capacidades-de-admin-de-aplicacion.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D23](2026-08-04-d23-solo-administrador-global-nombra-admins.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D24](2026-08-04-d24-cli-operable-por-ia-para-todas-las.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D25](2026-08-04-d25-cli-exclusivo-para-administrador-global.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D26](2026-08-04-d26-confirmacion-explicita-en-acciones.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D27](2026-08-04-d27-logs-estructurados-canonicos-con.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D28](2026-08-04-d28-retencion-por-niveles-configurable-con.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D29](2026-08-04-d29-retencion-provisional-90-dias-hot-1-ano.md) | aplazada | 2026-08-04 | usuario (dueño del producto) |

Las D-n anteriores a este registro viven en `../08-decisiones-y-preguntas-abiertas.md` y se migran aquí en la issue #985; hasta entonces, las cuatro primeras filas de arriba (D36, D37, D88, D89) son la única cadena de sustitución migrada.
