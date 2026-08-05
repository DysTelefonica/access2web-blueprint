# 03 · Lanzadera

Lanzadera es el portal Access/VBA que concentra identidad, permisos, catálogo y lanzamiento de aplicaciones. Este índice resume el **Lote 1 completado**; el detalle está separado para facilitar revisión.

## Identidad y evidencia

| Elemento | Evidencia | Confianza |
|---|---|---|
| Baseline funcional | `C:\00repos\codigo\00_LANZADERA\staging\Lanzadera.accdb` · `63ba5e01617fdda857503d43151f06bb7bc11829` | Verified-static/runtime · Dysflow, 2026-08-04 |
| Referencia publicada | `C:\00repos\codigo\00_LANZADERA\00_main\Lanzadera.accdb` · `1474e8e8c2a8c352599ffa8b846223c6eb6e0f17` | Comparison-only |
| Backend de uso consultado | `C:\00repos\datos\Lanzadera_Datos.accdb` | Verified-runtime · selección explícita read-only |
| Configuración | `.dysflow/project.json`, `src/backends.json`, `TbConfiguracionBackends` | Verified-static/runtime |
| Versión declarada | `lanzadera.ini`: `2026-002` | Verified-static |
| Grafo CodeGraph-VBA | Índice independiente en `staging` y `00_main`; exploración estática de flujos principales, 2026-08-04 | Verified-static · límites dinámicos explícitos |

## Inventario normalizado

| Tipo | Cantidad | Evidencia |
|---|---:|---|
| Tablas backend | 35 | `list_tables(target=backend)` |
| Tablas frontend | 1 (`TbConfiguracionBackends`) | `list_tables(target=frontend)` |
| Formularios / módulos de documento | 28 | `list_objects`, `list_vba_modules` |
| Informes | 0 | `list_objects` |
| Módulos estándar | 17 | `list_objects` |
| Clases | 14 | `list_objects` |
| Consultas exportables | 0 en frontend | `export_queries` |
| Macros | No inventariado por la superficie usada | Evidence gap |

## Documentos del lote

- [Capacidades y disposición](capabilities.md)
- [Mapa de formularios y navegación](forms.md)
- [Modelo de datos y evidencia ERD](data-model.md)
- [Integraciones, seguridad y riesgos](integrations-security.md)

## Fuentes y límites

La autoridad funcional de esta pasada es `staging`, seguida por Dysflow y el backend compartido seleccionado explícitamente; `00_main` solo es referencia publicada. No se ejecutaron tests, SQL de escritura, imports, exports, sync, compilación ni cambios en Lanzadera. Las operaciones Dysflow devuelven `dysflow.result/v1`; se validó el envelope defensivamente. Los hechos de intención de OpenSpec se marcan `Intended` y cualquier divergencia se conserva como `Divergent`.

La pasada CodeGraph-VBA confirmó estáticamente login, permisos, administración, navegación y lanzamiento. Los saltos por DAO, `Application.TempVars`, `Shell`/filesystem y resolución dinámica siguen siendo límites, no evidencia runtime.

## Disposiciones finales (post-Lote 1)

Las decisiones posteriores al descubrimiento **cierran el debate abierto** para cada capacidad detectada. Resumen en [`capabilities.md` – Disposiciones finales](capabilities.md#disposiciones-finales-post-lote-1); los contratos objetivo que sustituyen los mecanismos legacy, en [`integrations-security.md`](integrations-security.md). Las decisiones de identidad (hashes, lockout, activación, notificación de cambios de permiso), UAT, registro de aplicaciones y suplantación viven en `docs/06-seguridad-y-trazabilidad.md` y `docs/09-arquitectura-objetivo-y-principios.md`.

Estado Lanzadera tras el lote:

- ✅ Identidad, permisos y catálogo: **preservar y modernizar**.
- ✅ Auditoría de autenticación y apertura: **preservar y modernizar**.
- ❌ Lanzador Access / ubicación oficina-fuera / vídeos / gestión técnica de backend / telemetría de ubicación: **retirar**.
- 🔄 Cola de correo, tasks flags, hash de credenciales: **reemplazar mecanismo** mediante servicios compartidos.

## Próximo paso

Revisar este lote y aclarar los huecos de catálogo efectivo, macros, formularios de vídeo, consultas guardadas y política de migración de credenciales antes de pasar al lote siguiente.

## Inventario real Dysflow (2026-08-05, segunda pasada)

**Volúmenes principales del backend autoritativo** (`C:\00repos\datos\Lanzadera_Datos.accdb`):

- **`tbUsuarios`**: **156 filas** (usuarios Lanzadera activos en producción).
- **`TbUsuariosAplicacionesPermisos`**: **622 filas** (permisos por aplicación, alto volumen).
- **35 tablas totales** (vs 22 de HPS, 11 de HPS_Solicitudes, 15 de Condor, 42 de NoConformidades, 49 de Expedientes, 71 de Gestion_Riesgos).

**Hallazgos del inventario real**:

1. **Datos personales explícitos en `tbUsuarios`**: `Matricula_DNI` (Text 50), `Nombre`, `DirCorreo`, `telfijo`, `telmovil`. Mismo riesgo que HPS (D92) y HPS_Solicitudes (D98). 156 filas con datos personales.

2. **Sistema de flags booleanos mixtos**:
   - `SeLogean`, `ParaTareasProgramadas`, `Autorizador` son `YesNo` (type 1) en `tbUsuarios` (consistente con booleanos reales).
   - `EmplazamientoExterno`, `UsuarioDeGestionRiesgos`, `UsuariosI3D` son `Text 2` (Sí/No) en `tbUsuarios` (inconsistencia — booleanos como texto).

3. **Inconsistencias de naming detectadas**:
   - `tbUsuarios` (minúscula) vs `TbAplicaciones` (mayúscula) en la misma app.
   - `TbCuestionaroRespuestas` (typo: debería ser `TbCuestionarioRespuestas`).

4. **Sistema de cuestionarios y vídeos** (formación): `TbCuestionarios`, `TbCuestionarioPreguntas`, `TbCuestionaroRespuestas` (typo), `TbVideos`, `TbVideosCategorias`, `TbVideosCuestionario`, `TbVideosVisionados`. Sistema de formación interno para usuarios.

5. **Cross-app flags**: `UsuarioDeGestionRiesgos`, `UsuariosI3D` (Text 2) en `tbUsuarios` — flags de acceso a Gestion_Riesgos e I3D. Cross-app.

6. **Día de envío de tareas** (`DiaEnvioTareas`, Integer 1) — sistema de tareas programadas.

7. **`.dysflow/project.json` creado en esta pasada** (con `setup_project` autorizado) en `00_LANZADERA/00_main/.dysflow/`. `projectId: 00-lanzadera-staging`, `frontendFile: Lanzadera.accdb`, `allowWrites: true`, `destinationRoot: src`.
