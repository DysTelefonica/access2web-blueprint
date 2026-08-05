# NoConformidades — formularios, navegación y call paths

## Navegación principal

```text
Form_frmSplash
  -> EVE
  -> Form_Form0BDOpciones
      ├─> Form_Form0BDOpcionesParteProyectos
      │    -> Form_FormNCProyectoGestion (listado de NC de Proyecto)
      │        -> Form_FormNCProyectoGeneral | Form_FormNCProyectoGeneralConVinculoNC (alta/edición)
      │            -> Form_FormNCProyectoAC | Form_FormNCProyectoAR
      │            -> Form_FormNCProyectoControlEficacia | Form_FormNCProyectoControlEficaciaAlta
      │            -> Form_FormNCProyectoAcciones
      │            -> Form_FormNCProyectoDocumentos
      │            -> Form_FormNCProyectoNota
      │            -> Form_FormNCProyectoMotivoEliminado
      │            -> Form_FormNCProyectoReplanificaciones
      │            -> Form_FormNCProyectoSeguimiento
      │                -> Form_FormNCProyectoSeguimientoNC
      │                -> Form_FormNCProyectoSeguimientoTareas
      │        -> Form_FormNCProyectoTipologiaGestion -> Form_FormTipologiaNCProyecto
      │    -> Form_formRiesgosSeleccion (vinculación con Gestion_Riesgos)
      │    -> Form_FormExpedientesBusqueda (vinculación con Expedientes)
      │
      ├─> Form_Form0BDOpcionesAuditorias
      │    -> Form_FormAuditoriaSeleccion -> Form_FormAuditoria | Form_FormAuditoriasGestion
      │        -> Form_FormAuditoriaDocumentos
      │        -> Form_FormNCAuditoriaGeneral (alta/edición NC de Auditoría)
      │            -> Form_FormNCAuditoriaAC | Form_FormNCAuditoriaAR
      │            -> Form_FormNCAuditoriaControlEficacia | Form_FormNCAuditoriaControlEficaciaAlta
      │            -> Form_FormNCAuditoriaAcciones
      │            -> Form_FormNCAuditoriaDocumentos
      │            -> Form_FormNCAuditoriaNota
      │            -> Form_FormNCAuditoriaMotivoEliminado
      │            -> Form_FormNCAuditoriaReplanificaciones
      │            -> Form_FormNCAuditoriaSeguimiento
      │                -> Form_FormNCAuditoriaSeguimientoNC
      │                -> Form_FormNCAuditoriaSeguimientoTareas
      │
      └─> Form_Form0BDTecnicos
           -> Form_FormIndicadores
           -> Form_FormARAuditoriaDocumentos | Form_FormARProyectoDocumentos

Form_FormCorreo (modal de envío de correo)
```

## Call paths críticos

| Capacidad | Camino observado | Persistencia / efecto |
|---|---|---|
| Inicio | `frmSplash.Form_Timer → EVE → LeeConfiguracionLocal → TbConfiguracionBackends → getdb() → getUsuario()` | TempVars (incluye `BackendActivo`, `BackendPathConfigurado`, `EnPruebas`, IDAplicacion 8/81), sesión, cachés, tareas |
| Carga de NC | `Form_FormNCProyectoGestion → NCProyectoOperaciones.Listar → NCProyectoListItemVM[] → HTML` | vista flat optimizada para UI |
| Alta NC de Proyecto | `Form_FormNCProyectoGeneral.Alta → NCProyectoOperaciones.Alta → transacción DAO` | cabecera, ACs, ARs, replanificaciones, documentos, notas, logs |
| AC/AR | `Form_FormNCProyectoAC/AR → ACAuditoriaOperaciones/ACProyectoOperaciones.Registrar → transacción DAO` | AC/AR con responsables, fechas, motivos |
| Control de Eficacia | `Form_FormNCProyectoControlEficaciaAlta → ControlEficaciaOperaciones → transacción DAO` | estado y veredicto del control |
| Replanificación | `Form_FormNCProyectoReplanificaciones → ReplanificacionesProyectoOperaciones.Registrar` | replanificación con motivo |
| Vinculación con Riesgo | `Form_formRiesgosSeleccion → RiesgoServicio.Listar → selección → NCProyecto.RiesgoID` | FK conceptual |
| Caché | `getdb() → TbCacheNCProyecto (lookup by IDNoConformidad)` → si cache miss, regenerar via `CacheNCProyecto.GenerarCacheCompleto` | tabla con `CacheValida`, `HitsConsultas`, `TamanioBytes` |
| Kill switch | `Test_KillSwitch.bas → Test_KillSwitch_IsCacheEnabled_Atomic → BuildJsonFail/BuildJsonOk → AddLog → EndTestSession` | activa/desactiva la caché en runtime |
| Mantenimiento | `InicializadorCache.AyudaCache`, `InvalidarCachesObsoletos(30)`, `EliminarCachesInvalidos(60)`, `DiagnosticarIntegridad`, `RegenerarCachesInvalidos`, `PoblarCacheMasivo(soloFaltantes=True)`, `LimpiarLogsAntiguos(90,30)`, `MostrarEstadisticasUso`, `MostrarRendimiento` | comandos administrativos de la caché |
| Indicadores | `Form_FormIndicadores → IndicadorServicio → IndicadorRepositorio` (puede usar caché) | métricas agregadas |
| Informes | `InformeNCAuditorias.Generar → artefacto` | salida |
| Correos | `Form_FormCorreo → Correo.Enviar` | envío al administrador |

## Inventario normalizado

- **Formularios**: ~60 archivos `Form_*.cls` cada uno con su `.form.txt` compañero. Cada par representa código + layout exportado de un formulario Access.
- **Patrón `.cls + .form.txt`**: la presencia del `.form.txt` indica que el layout está exportado de forma legible; el `.cls` contiene la lógica VBA. La nueva plataforma no usa formularios Access pero puede inspeccionar el `.form.txt` para extraer estructura (controles, eventos, jerarquía).
- **Clases**: 44 clases en `src/classes/` (dominio, infraestructura, helpers, entidades externas compartidas con Lanzadera). El inventario CodeGraph incluye `NCAuditoria*`, `NCProyecto*`, `Auditoria*`, `AC*`, `AR*`, `Replanificaciones*`, `Documento*`, `CacheNCCacheRepositorio`, `Riesgo*`, `Log*`, `Seg*`, `Informe*`, `Indicador*`, `Entorno`, `Usuario*`, `Juridica`, `Correo`.
- **Módulos**: 25 archivos en `src/modules/` incluyendo bootstrap/factory/DAO (`Variables Globales.bas`, `constructor.bas`), helpers, cache (`InicializadorCache.bas`, `CacheNCProyecto.bas`, `CacheTrustDiagnostics.bas`, `ModuloCacheIndicadores.bas`, `mdlCursor.bas`), instalador, JSON/HTML helpers, y **7 archivos `Test_*.bas`** con cobertura significativa (Test_BackendConfigPaths, Test_CacheListadoEstados, Test_CacheTrustDiagnostics, Test_E2E_BateriaNC, Test_IndicadoresCaracterizacion, Test_Issue19_CEGating, Test_KillSwitch).
- **Reports/macros/queries**: no se inspeccionaron queries exportadas ni macros embebidas en esta pasada; requieren revisión del binario (bloqueada por D89).

## Nota de evidencia

CodeGraph-VBA se consultó primero sobre `00_main` y devolvió call paths dinámicos con 51 símbolos en 2 archivos para la query de caché. Dysflow `list_objects` **NO se pudo ejercitar** por el fallo documentado en [Seguridad § D89](security-rules.md#d89-diagnóstico-del-fallo-de-list_objects-de-dysflow). La inspección de UI se mantiene read-only y no se han alterado formularios.