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

### D30–D60 migradas desde 08

| Decisión | Estado | Fecha | Decide |
|---|---|---|---|
| [D30](2026-08-04-d30-cli-soporta-diagnosticos-operativos.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D31](2026-08-04-d31-toda-remediacion-requiere-aprobacion-del.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D32](2026-08-04-d32-inspecciones-de-salud-bajo-demanda.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D33](2026-08-04-d33-destinatarios-de-anomalias-configurables.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D34](2026-08-04-d34-solo-administrador-global-configura.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D35](2026-08-04-d35-resultado-del-estudio-roadmap-global.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D38](2026-08-04-d38-umbral-de-lockout-configurable-default.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D39](2026-08-04-d39-bloqueo-durante-una-hora-por-defecto.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D40](2026-08-04-d40-notificacion-a-administradores-globales.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D41](2026-08-04-d41-caducidad-de-contrasena-configurable.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D42](2026-08-04-d42-activacion-baja-y-creacion-de-roles-solo.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D43](2026-08-04-d43-notificacion-de-cambios-de-permiso-solo.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D44](2026-08-05-d44-suplantacion-para-pruebas-restringida-al.md) | adoptada | 2026-08-05 | usuario (dueño del producto) |
| [D45](2026-08-04-d45-capabilities-declaradas-por-modulo.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D46](2026-08-04-d46-vista-unica-por-defecto-vistas.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D47](2026-08-04-d47-uat-asignacion-explicita-de.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D48](2026-08-04-d48-uat-gobernanza-reservada-al.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D49](2026-08-04-d49-uat-excepciones-auditadas-y-visibles-en.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D50](2026-08-04-d50-uat-excepciones-visibles-para-usuarios.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D51](2026-08-04-d51-lanzadera-retira-formacion-videos.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D52](2026-08-04-d52-lanzadera-retira-mecanismo-de.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D53](2026-08-04-d53-lanzadera-retira-segmentacion-oficina.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D54](2026-08-05-d54-lanzadera-retira-gestion-de-rutas-y.md) | adoptada | 2026-08-05 | usuario (dueño del producto) |
| [D55](2026-08-05-d55-lanzadera-moderniza-auditoria-retira.md) | adoptada | 2026-08-05 | usuario (dueño del producto) |
| [D56](2026-08-04-d56-lanzadera-preserva-y-moderniza-identidad.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D57](2026-08-04-d57-lanzadera-uat-detallado-y-diseno-de.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D58](2026-08-04-d58-registro-de-aplicaciones-hibrido.md) | adoptada | 2026-08-04 | usuario (dueño del producto) |
| [D59](2026-08-05-d59-scheduler-unificado-sustituye-flags-y.md) | adoptada | 2026-08-05 | usuario (dueño del producto) |
| [D60](2026-08-05-d60-configuracion-global-de-horarios-y.md) | adoptada | 2026-08-05 | usuario (dueño del producto) |

### D61–D91 migradas desde 08

| Decisión | Estado | Fecha | Decide |
|---|---|---|---|
| [D90](2026-08-08-d90-reset-flow-con-tokens-one-time-de-24-h.md) | adoptada | 2026-08-08 | usuario (dueño del producto) |
| [D91](2026-08-08-d91-primer-global-admin-se-aprovisiona-con.md) | adoptada | 2026-08-08 | usuario (dueño del producto) |
| [D61](2026-08-05-d61-ejecucion-manual-bajo-demanda-de.md) | adoptada | 2026-08-05 | usuario (dueño del producto) |
| [D62](2026-08-05-d62-vista-previa-de-informe-sin-envio.md) | adoptada | 2026-08-05 | usuario (dueño del producto) |
| [D63](2026-08-05-d63-generar-y-enviar-directamente-sin-vista.md) | adoptada | 2026-08-05 | usuario (dueño del producto) |
| [D64](2026-08-05-d64-dashboard-global-de-operaciones-de.md) | adoptada | 2026-08-05 | usuario (dueño del producto) |
| [D65](2026-08-05-d65-evidencia-legacy-cola-por-tabla.md) | adoptada | 2026-08-05 | usuario (dueño del producto) |
| [D66](2026-08-05-d66-stack-backend-python-3-12-con-fastapi-0.md) | adoptada | 2026-08-05 | usuario (dueño del producto) |
| [D67](2026-08-05-d67-stack-frontend-htmx-2-0-4-jinja2-3-1.md) | adoptada | 2026-08-05 | usuario (dueño del producto) |
| [D68](2026-08-05-d68-estructura-del-repositorio-monorepo-con.md) | adoptada | 2026-08-05 | usuario (dueño del producto) |
| [D69](2026-08-05-d69-actualizacion-de-contadores-pendientes.md) | adoptada | 2026-08-05 | usuario (dueño del producto) |
| [D70](2026-08-05-d70-cache-de-aplicacion-selectiva-y.md) | adoptada | 2026-08-05 | usuario (dueño del producto) |
| [D71](2026-08-05-d71-redis-queda-como-opcion-detras-del.md) | adoptada | 2026-08-05 | usuario (dueño del producto) |
| [D72](2026-08-05-d72-rendimiento-http-etag-304-not-modified.md) | adoptada | 2026-08-05 | usuario (dueño del producto) |
| [D73](2026-08-05-d73-no-usar-insforge-como-baas-el-backend.md) | adoptada | 2026-08-05 | usuario (dueño del producto) |
| [D74](2026-08-05-d74-no-introducir-kubernetes-ni-openshift.md) | adoptada | 2026-08-05 | usuario (dueño del producto) |
| [D75](2026-08-05-d75-topologia-de-despliegue-abierta.md) | adoptada | 2026-08-05 | usuario (dueño del producto) |
| [D76](2026-08-05-d76-postgresql-gestionado-preferido-sobre.md) | adoptada | 2026-08-05 | usuario (dueño del producto) |
| [D77](2026-08-05-d77-contenedores-docker-desde-el-dia-uno.md) | adoptada | 2026-08-05 | usuario (dueño del producto) |
| [D78](2026-08-05-d78-versionado-semantico-por-modulo-y.md) | adoptada | 2026-08-05 | usuario (dueño del producto) |
| [D79](2026-08-05-d79-trunk-based-development-conventional.md) | adoptada | 2026-08-05 | usuario (dueño del producto) |
| [D80](2026-08-05-d80-catalogo-de-versiones-compatibles-entre.md) | adoptada | 2026-08-05 | usuario (dueño del producto) |
| [D81](2026-08-05-d81-despliegue-coexistente-estable-rc-en-uat.md) | adoptada | 2026-08-05 | usuario (dueño del producto) |
| [D82](2026-08-05-d82-migraciones-de-bd-backward-compatibles.md) | adoptada | 2026-08-05 | usuario (dueño del producto) |
| [D83](2026-08-05-d83-hps-solicitudes-se-documenta-como.md) | adoptada | 2026-08-05 | usuario (dueño del producto) |
| [D84](2026-08-05-d84-baseline-de-release-de-condor-y-brass.md) | adoptada | 2026-08-05 | usuario (dueño del producto) |
| [D85](2026-08-05-d85-baseline-operativo-actual-catalogo-con-8.md) | adoptada | 2026-08-05 | usuario (dueño del producto) |
| [D86](2026-08-05-d86-la-forma-hexagonal-del-legacy-dominio-en.md) | adoptada | 2026-08-05 | usuario (dueño del producto) |
| [D87](2026-08-05-d87-los-tests-vba-existentes-en-staging-por.md) | adoptada | 2026-08-05 | usuario (dueño del producto) |

Las D1 a D91 de la antigua tabla de `../08-decisiones-y-preguntas-abiertas.md` están migradas completas; `tests/test_decisions_registry.py` falla si falta alguna o si un documento vuelve a repetir sus filas. Las decisiones que comparten `origen` (por ejemplo D70, D71 y D72) comparten también el motivo de esa resolución conjunta.
