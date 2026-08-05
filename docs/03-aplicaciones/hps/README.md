# 03 · HPS

## Propósito

Evidencia de descubrimiento de **HPS** (Homologación Personal Servicios), aplicación Access/VBA que mantiene el catálogo de usuarios HPS, su historial, cursos y observaciones. HPS es la aplicación responsable de la **gestión de personal sujeto a homologación** — quién está obligado, qué cursos ha hecho, fechas de renovación, observaciones y motivos. Es **lectura intensiva sobre usuarios** con poco write-heavy en producción.

## Estado

- **Fase:** descubrimiento completo, Batch 5 + inventario Dysflow real.
- **Fecha de evidencia:** 2026-08-05.
- **Repositorio:** `C:\00repos\codigo\00_HPS\00_main`. CodeGraph-VBA inspeccionado; inventario backend real vía Dysflow MCP.
- **Staging:** rama `staging` y `refactor_PLAN-003-clean-start` disponibles; no inspeccionadas en esta pasada.
- **Frontend:** `HPS.accdb` en `00_main` (30 MB).
- **Backend autoritativo:** `C:\00repos\datos\HPST.accdb` (7 MB). **Duplicado en `00_main/HPST.accdb`** (legacy local copy que NO se debe usar como autoridad).
- **Dysflow:** `00_main/.dysflow/project.json` válido (`projectId: 00-hps-staging`). Inventario backend completo: **22 tablas**, **6 FKs**, **27 columnas en `TbUsuarios`**.
- **APAP y APAP_WEB** no aparecen (proyecto personal; regla del blueprint).

## Lote asociado

Lote 4 del plan de discovery. Elegido por ser la app de gestión de personal más estable del ecosistema y por contraste con NoConformidades (cache-heavy vs read-heavy).

## Entregables

1. [Capacidades](capabilities.md)
2. [Formularios y call paths](forms.md)
3. [Modelo físico y diccionario](data-model.md)
4. [Matriz de migración](migration-matrix.md)
5. [Integraciones y automatización](integrations-automation.md)
6. [Seguridad y reglas](security-rules.md)

## Fuentes de autoridad

1. `C:\00repos\codigo\00_HPS\00_main\src` (clases, forms, módulos) — codegraph-vba.
2. `C:\00repos\documentacion\OPENSPEC\00_HPS` (documentación previa).
3. Dysflow read-only sobre `HPS.accdb` (frontend) y `HPST.accdb` (backend autoritativo en `C:\00repos\datos\`).
4. Engram como contexto histórico.

## Reglas de evidencia

- No se han realizado imports, exports, sync, tests, compile, cleanup ni escrituras.
- Se excluyen valores personales (DNI, nombres, correos, teléfonos, fechas de nacimiento), credenciales, hashes, hosts y nombres de máquina. La columna `DNI`, `Nombre`, `Apellido_1`, `Apellido_2`, `Telefono`, `Correo_e`, `F_Nacimiento` están en la lista de campos sensibles para migración (ver D92).
- Las rutas UNC y hosts no se reproducen.
- Todo campo/registro del backend queda como `preservar hasta decisión` salvo los marcados `legacy copy` que requieren disposición explícita.

## Hallazgos críticos del lote

1. **107 callers de `getdb()`** — el más bajo de las 8 aplicaciones. HPS es **read-heavy, no write-heavy**: lee usuarios, observaciones, histórico; escribe solo al actualizar HPS o cargar nuevos datos.
2. **3 módulos de caché propios** (`cacheSuministrador.bas`, `cacheUsuario.bas`, `Mod_Cache_Core.bas`, `Mod_Sincronizacion_Historico.bas`, `Mod_StartupCacheInitialization.bas`, `CacheConsistencyAudit.bas`). HPS también tiene capa de caché selectivo maduro — confirma D91 cross-app.
3. **9 archivos `Test_*.bas`** con cobertura significativa (Test_AnexoSelectionTransaction, Test_CacheConsistencyAudit, Test_HistoricoAdjuntosTransactionWrapper, Test_HPSConfig, Test_HPSEntorno, Test_LocalReadAccessAuthorization, Test_PerAnexoMove, Test_RealTimeIndicatorCoherence, Test_StartupCacheInitialization). Confirma D87 (tests VBA como evidencia de comportamiento).
4. **4 tablas "Copia de..."** y 1 sentinel "Errores de pegado" — patrón legacy de copia antes de cambios masivos + sentinel de errores en operaciones de pegado masivo. Requieren disposición explícita en la matriz de migración.
5. **Volumen real**: 345 usuarios activos (`TbUsuarios`), 242 históricos (`TbUsuariosHistoricos`), 1280 relaciones HPS (`TbHPS`), 330 observaciones (`TbObservaciones`).
6. **⚠️ HALLAZGO CRÍTICO DE SEGURIDAD: HPS TIENE CACHÉ EN EL FRONTEND**. El frontend `HPS.accdb` (30 MB) contiene **12 tablas locales** con datos cacheados, incluyendo:
   - `TbDatosLocal` con **344 filas** y TODOS los datos personales (DNI, Nombre, Apellidos, Teléfono, Correo, Fecha de Nacimiento, Lugar de Nacimiento, Observaciones).
   - `TbUsuariosHistoricosLocal` con **235 filas** con datos personales históricos.
   - `TbCursosLocal`, `TbSuministradoresLocal`, `TbDatosLocalParaIndicadores`, `TbUsuariosSICALocal` (esta vacía, 0 filas), `TbUsuariosSICALocalParaIndicadores`, `TbVinculosTablas`.
   - `TbConfiguracionHPS` (clave-valor local: Clave, Valor, Activo, FechaModificacion, UsuarioModificacion).
   - `tblInfo`, `tblSettings` (info del frontend).
   - **Esto es un patrón raro** que NO aparece en Lanzadera/Expedientes/Gestion_Riesgos/NoConformidades. El frontend tiene un sistema de caché local completo.
7. **`.gitignore` del repo `00_HPS` NO excluye `*.accdb`** — solo `*.accde`, `*.mdb`, `*.mde`, `HPST.accdb` (específico). Riesgo: si `HPS.accdb` se versiona, los datos personales quedan en git.
8. **Duplicación frontend/backend**: `TbHPS` y `TbUsuariosHistoricos` existen en ambos `.accdb`. Esto genera ambigüedad (`ACCESS_TABLE_AMBIGUOUS`) que requiere `target: "backend"` explícito en Dysflow. La duplicación es **cache local sincronizada** (consistente con el hallazgo 6).
9. **FKs conceptuales sin FK física** (`IDExpediente`, `IDEmpresaUsuario`, `IDEmpresaHPS`, `IDJuridicaContrato`, `IDSolicitud`). Las relaciones con Expedientes/Empresas/Solicitudes son por ID sin constraint. Migración debe decidir si formalizar las FK o mantener como referencia conceptual.
10. **Booleanos como Text(2)**: `CursoEnVigor`, `Requiere_Curso`, `RequiereComunicacionConcesion`. Inconsistencia detectada — ya presente en NoConformidades. Estandarizar a `BOOLEAN` en PostgreSQL con migración explícita.
11. **29 clases** vs 44-47 de las otras apps. HPS tiene superficie de dominio menor.
12. **`clsTestDouble*`** (5 archivos): Test Doubles para tests. Indica disciplina TDD madura.

## Checklist

- [x] Inventario funcional, formularios, clases y módulos documentados vía codegraph-vba.
- [x] Inventario real Dysflow del backend (22 tablas, 493 filas principales, 6 FKs, 27 columnas en `TbUsuarios`).
- [x] Volumen real: 345 usuarios activos, 242 históricos, 1280 relaciones HPS, 330 observaciones.
- [x] Inconsistencias detectadas en FKs, tipos Sí/No, duplicación frontend/backend.
- [x] APAP y APAP_WEB no aparecen en esta evidencia.
- [x] D92 propuesto para disposition de campos sensibles y legacy copies.
- [ ] Épica + tickets + matriz de migración de datos para HPS.

## Siguiente paso

Cruzar el inventario con la documentación previa en `OPENSPEC/00_HPS`, resolver D92 (disposición de legacy copies + campos sensibles), y luego generar la épica + tickets para la migración de HPS. Continuar después con Condor (Lote 5), Brass (Lote 6) y HPS_Solicitudes (Lote 7) usando el mismo patrón codegraph + Dysflow.