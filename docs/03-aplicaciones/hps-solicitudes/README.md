[← Back to DOCS](../../../DOCS.md)

# 03 · HPS_Solicitudes

## Propósito

Evidencia de descubrimiento de **HPS_Solicitudes** (Solicitudes HPS), aplicación Access/VBA que gestiona el ciclo de vida de **solicitudes HPS**: alta, renovación, cambio de tipo, traspaso a ONS, justificación, responsables, adjuntos Excel, generación de plantillas HTML, correos automáticos, registro en HPS, exportación a Lanzadera.

## Estado

- **Fase:** descubrimiento completo, Batch 7 + inventario Dysflow real sobre backend autoritativo.
- **Fecha de evidencia:** 2026-08-05.
- **Repositorio:** `C:\00repos\codigo\HPS_SOLICITUDES` (sin prefijo `00_`, sin rama staging, sin rama develop, sin rama release; solo `main`). CodeGraph-VBA inspeccionado; inventario Dysflow completo.
- **Dysflow:** `.dysflow/project.json` creado en esta pasada con `setup_project` (autorizado por el user). `projectId: 00-hps-solicitudes-staging`, `frontendFile: Solicitudes_HPS.accdb`, `allowWrites: true`, `destinationRoot: src`.
- **Frontend:** `Solicitudes_HPS.accdb` en raíz (22 MB).
- **Backend autoritativo:** `C:\00repos\datos\Solicitudes_HPS_datos.accdb` (13 MB). **Duplicado en la raíz del repo** (`Solicitudes_HPS_datos.accdb` local).
- **Inventario Dysflow real**:
  - **11 tablas** en el backend autoritativo.
  - **245 solicitudes** en `TbSolicitudes`.
  - **27 responsables**, **7 justificaciones**, **13 grados HPS**, **9 correos enviados**.
  - **2058 filas en `TbLogsGeneral`** (log general, alto volumen).
  - **3 FKs** entre tablas de usuario (excluyendo MSysNavPane*).
  - **Sistema de gestión de traspasos a ONS** (Organismo Notificador de Seguridad) vía `Form_FormAdjuntaTraspasoONS.cls` y `URLAdjuntoEnvioONS`.
- **APAP y APAP_WEB** no aparecen (proyecto personal; regla del blueprint).

## Lote asociado

Lote 7 del plan de discovery. **La app más simple de las 8** en superficie de datos (11 tablas vs 22 de HPS, 15 de Condor, 42 de NoConformidades) pero **en uso activo** (245 solicitudes reales vs 1 de Condor).

## Entregables

1. [Capacidades](capabilities.md)
2. [Formularios y call paths](forms.md)
3. [Modelo físico y diccionario](data-model.md)
4. [Matriz de migración](migration-matrix.md)
5. [Integraciones y automatización](integrations-automation.md)
6. [Seguridad y reglas](security-rules.md)

## Fuentes de autoridad

1. `C:\00repos\codigo\HPS_SOLICITUDES\src` (clases, forms, módulos) — codegraph-vba.
2. `C:\00repos\documentacion\OPENSPEC\00_HPS_SOLICITUDES` (documentación previa).
3. Dysflow read-only sobre `Solicitudes_HPS.accdb` (frontend) y `Solicitudes_HPS_datos.accdb` (backend autoritativo en `C:\00repos\datos\`).
4. Engram como contexto histórico.

## Reglas de evidencia

- No se han realizado imports, exports, sync, tests, compile, cleanup ni escrituras.
- Se excluyen valores personales, correos, credenciales, hashes, hosts y nombres de máquina.
- Las rutas UNC y hosts no se reproducen (regla de evidencia transversal).
- `Solicitudes_HPS_datos.accdb` en la raíz del repo es **legacy local**; el backend autoritativo está en `C:\00repos\datos\`.
- APAP y APAP_WEB no se mencionan (proyecto personal del desarrollador).

## Hallazgos críticos del lote

1. **11 tablas en backend autoritativo** (vs 22 de HPS, 15 de Condor, 42 de NoConformidades). **La app más simple en superficie de datos** pero **en uso activo** (245 solicitudes reales).

2. **FK `TbResponsables.Correo → TbSolicitudes.emailResponsable`** ⚠️ — join por **texto email**, no por ID. Si el email cambia en TbResponsables, la FK lógica se rompe. Data integrity gap.

3. **FK `TbJustificaciones.idjustificacion → TbSolicitudes.idjustificacion`** ⚠️ — dirección unusual. La FK va de `TbJustificaciones` a `TbSolicitudes` (un Justificación referencia una Solicitud). Esto significa que la Solicitud tiene su propio `idjustificacion` y la Justificación la referencia de vuelta. **Confuso**.

4. **`TbSolicitudes` con 28 columnas y datos personales completos**: `DNI`, `Nombre`, `Apellido1`, `Apellido2`, `FNacimiento`, `LugarNacimiento`, `email`, `Telefono`. ⚠️ **245 filas con datos personales**. Mismo riesgo de seguridad que HPS (ver D92).

5. **Sistema de traspasos a ONS** (`Organismo Notificador de Seguridad`): `Form_FormAdjuntaTraspasoONS.cls` + `URLAdjuntoEnvioONS` (Memo en `TbSolicitudes`). Integración con sistema externo ONS.

6. **Sistema de plantillas HTML**: `Form_FormPlantillasHTML.cls` y `Form_FormWeb.cls`. Vistas web embebidas (similar a Condor con `WebVisorCacheServicio` pero en versión más simple).

7. **89 callers de `getdb()`** (intermedio entre Lanzadera/Expedientes y Gestion_Riesgos/NoConformidades).

8. **`IDAplicacion = "22"`** (producción). `EnPruebas` no implementado en `EVE` (comentado, no activo).

9. **`getdbLanzadera()` para identidad** — acoplamiento directo a Lanzadera (igual que el resto del ecosistema).

10. **2 archivos de tests VBA**: `Test.bas`, `TestParametrosParser.bas`. Cobertura básica (no tan maduro como HPS con 9 archivos o NoConformidades con 7).

11. **Flags `TempVars` en `EVE`**: `EnDesarrollo`, `DatosEnLocal`, `EnPruebas`, `ConCorreoCopiaGestor`, `ActivadoCorreoAutomatico`, `RegistroEnHPS`, `ExpedienteUnificado`. **Más flags que las otras apps** (configuración rica).

12. **Adjuntos Excel**: `Form_FormAdjuntarExcelSolicitante.cls` (solicitante adjunta Excel con datos de la solicitud). Integración con Excel.

## Checklist

- [x] Inventario funcional, formularios, clases y módulos documentados vía codegraph-vba.
- [x] Inventario real Dysflow del backend (11 tablas, 245 solicitudes, 2058 logs, 3 FKs).
- [x] **26 clases** + **~30+ forms** + **14 módulos** desglosados.
- [x] 28 columnas reales de `TbSolicitudes` con tipos y observaciones.
- [x] APAP y APAP_WEB no aparecen en esta evidencia.
- [ ] Épica + tickets + matriz de migración de datos para HPS_Solicitudes.
- [ ] Audit cross-cutting de D93 (password hardcoded) en los demás repos.

## Siguiente paso

Cruzar el inventario con la documentación previa en `OPENSPEC/00_HPS_SOLICITUDES`. Generar la **épica + tickets accionables + matriz de migración de datos** para HPS_Solicitudes (alcance expandido). Continuar después con Brass (Lote 6) y audit de D93 (password hardcoded) cross-cutting.

## Core invariants

- **245 solicitudes con datos personales completos (D92)**: `TbSolicitudes` contiene `DNI`, `Nombre`, `Apellido1`, `Apellido2`, `FNacimiento`, `LugarNacimiento`, `email`, `Telefono`. Riesgo de seguridad equivalente al de HPS. Cualquier artefacto del blueprint redacta con placeholders.
- **11 tablas (la app más simple del ecosistema)**: superficie de datos mínima pero en uso activo. La migración web preserva el ciclo de vida completo (alta, renovación, cambio de tipo, traspaso a ONS, justificación).
- **FK por texto email en `TbResponsables.Correo → TbSolicitudes.emailResponsable`**: data integrity gap. Si el email cambia en `TbResponsables`, la FK lógica se rompe. La nueva plataforma usa IDs numéricos para las FKs y mantiene el email como atributo.
- **Acoplamiento Lanzadera via `getdbLanzadera()` (D86)**: traduce a adaptadores con `IdentityPort`; nunca se mantiene acceso directo cross-app.
- **`Solicitudes_HPS_datos.accdb` en la raíz del repo es legacy local**: el backend autoritativo está en `C:\00repos\datos\Solicitudes_HPS_datos.accdb`. La nueva plataforma NO conserva el duplicado en el repo; el `.gitignore` debe excluirlo explícitamente.
- **Sistema de traspasos a ONS**: `Form_FormAdjuntaTraspasoONS.cls` + `URLAdjuntoEnvioONS` (Memo en `TbSolicitudes`). Integración externa preservada como contrato de puerto (`OnsIntegrationPort`).

## Contributor checklist

- [ ] El cambio respeta las 6 reglas de §Core invariants; el `ci / quality` check pasa verde.
- [ ] Si el PR introduce datos personales (DNI, Nombre, Apellidos, email, F_Nacimiento), sigue D92: redacción con placeholders y scrubbing antes de commit.
- [ ] Si el PR modifica el `.gitignore` raíz, se asegura que `Solicitudes_HPS_datos.accdb` (legacy local con datos reales) queda excluido.
- [ ] Si el cambio añade una migración Alembic, sigue `expand_and_contract` (D82): añadir columnas o tablas, sin `DROP` ni `ALTER` destructivos en la misma release.
- [ ] Si el cambio toca el sistema de traspasos a ONS, se conserva como contrato de puerto y se documenta en `integrations-automation.md`.
- [ ] El PR es ≤ 400 líneas (`additions + deletions`); si no, partir por unidad de trabajo o encadenar.

## Navigation

Previous: [hps](../hps/README.md) | Next: [condor](../condor/README.md)