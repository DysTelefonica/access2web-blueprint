[← Back to DOCS](../../../DOCS.md)

# 03 · Brass

## Propósito

Evidencia de descubrimiento de **Brass** (Gestión de Eventos con Equipos de Medida Calibrados), aplicación Access/VBA que gestiona el ciclo de vida de **eventos regulatorios**: alta, técnicos asignados, equipos de medida, calibración, facturación, materiales, planificación, partes de trabajo, libranzas, subcontratación, franqueo (cierre) con validación de SLA. Brass es la **app de gestión de eventos del dominio regulatorio de calidad/calibración**, con **5777 eventos en producción** y **25639 actividades**.

## Estado

- **Fase:** descubrimiento completo, Batch 8 + inventario Dysflow real sobre backend autoritativo.
- **Fecha de evidencia:** 2026-08-05.
- **Repositorio:** `C:\00repos\codigo\00_BRASS\00_main` (seleccionado por el user). CodeGraph-VBA inspeccionado; inventario Dysflow completo.
- **Otras ramas:** `develop`, `release_2026-001` (no inspeccionadas en esta pasada).
- **Frontend:** `Gestion_Brass_Gestion.accdb` (25 MB).
- **Backend autoritativo:** `C:\00repos\datos\Gestion_Brass_Gestion_Datos.accdb` (61 MB).
- **Dysflow:** `.dysflow/project.json` creado en esta pasada con `setup_project` (autorizado por el user). `projectId: 00-brass-00-main-clean`, `frontendFile: Gestion_Brass_Gestion.accdb`, `allowWrites: true`, `destinationRoot: src`.
- **Inventario Dysflow real**:
  - **60 tablas** en el backend autoritativo.
  - **27 FKs** entre tablas de usuario (excluyendo MSysNavPane*).
  - **5777 eventos** (`TbEventos`).
  - **25639 actividades** (`TbActividades`).
  - **7592 anexos** (`TbAnexos`).
  - **2351 materiales** (`TbMaterial`).
  - **806 equipos** (`TbEquipos`).
  - **14 equipos de medida** (`TbEquiposMedida`) con **28 calibraciones** (`TbEquiposMedidaCalibraciones`).
  - **115 técnicos** (`TbTecnicos`).
  - **624 planificaciones** (`TbPlanificacion`).
  - **519 gastos** (`TbGastos`).
  - **53 facturas principales** (`TbFacturaPrincipal`).
- **APAP y APAP_WEB** no aparecen (proyecto personal del desarrollador).

## Lote asociado

Lote 6 del plan de discovery. **Última app pendiente del study**. Posicionado por ser la app con dominio regulatorio más especializado (calibración) y volumen alto de datos (25639 actividades).

## Entregables

1. [Capacidades](capabilities.md)
2. [Formularios y call paths](forms.md)
3. [Modelo físico y diccionario](data-model.md)
4. [Matriz de migración](migration-matrix.md)
5. [Integraciones y automatización](integrations-automation.md)
6. [Seguridad y reglas](security-rules.md)

## Fuentes de autoridad

1. `C:\00repos\codigo\00_BRASS\00_main\src` (clases, forms, módulos) — codegraph-vba.
2. `C:\00repos\documentacion\OPENSPEC\00_BRASS` (documentación previa).
3. Dysflow read-only sobre `Gestion_Brass_Gestion.accdb` (frontend) y `Gestion_Brass_Gestion_Datos.accdb` (backend autoritativo en `C:\00repos\datos\`).
4. Engram como contexto histórico.

## Reglas de evidencia

- No se han realizado imports, exports, sync, tests, compile, cleanup ni escrituras.
- Se excluyen valores personales, correos, credenciales, hashes, hosts y nombres de máquina.
- ⚠️ **D104 — contraseña hardcodeada `"dpddpd"`** en `Variables Globales.bas:560` como contraseña REAL (no fallback). Documentada en [Seguridad § D104](security-rules.md#d104--contraseña-hardcodeada-dpddpd-como-real-en-variables-globalesbas560).
- Las rutas UNC y hosts no se reproducen.
- APAP y APAP_WEB no se mencionan (proyecto personal del desarrollador; regla transversal del blueprint).

## Hallazgos críticos del lote

1. **⚠️⚠️⚠️ D104 — contraseña hardcodeada `"dpddpd"` en `Variables Globales.bas:560`**. Es la contraseña **REAL** del backend, pasada directamente a `OpenDatabase("MS Access;PWD=dpddpd")`. **NO es un fallback**, es la contraseña directa. Si el repo `00_BRASS` se versiona, la contraseña queda en git. La rotación de la contraseña NO surte efecto sin cambio de código. Riesgo CRÍTICO.

2. **`getdb()` con 155 callers** en `Variables Globales.bas:535` (intermedio entre HPS_Solicitudes con 89 y Gestion_Riesgos con 308). Brass tiene bastante intensidad DAO.

3. **Sistema de eventos complejo** (`TbEventos` con 38 columnas, 5777 filas): incluye SLA (`TIEMPORESPUESTAEVENTO`), franqueo (`Franqueado` YesNo + `CAUSAFIN` + `FechaFinal` + `HORAFINALEVENTO`), workflow de notificación al cliente, 3 tipos de reparación mutuamente excluyentes, 14+ fechas.

4. **`IDEvento` es Text(50), no Long**: ⚠️ todas las FKs del sistema se hacen por texto. Decisión de diseño legacy. La nueva plataforma debe decidir si mantiene IDs como string (UUID) o los migra a numéricos.

5. **Sistema de BUI (Business Unit Identifier)**: `TbBUI`, `TbNodoBUI`, `TbSubsistemaBui`, `TbNodos` — jerarquía organizacional de 4 niveles.

6. **Sistema de equipos de medida con calibración** (dominio regulatorio): `TbEquiposMedida` (14), `TbEquiposMedidaCalibraciones` (28). Cada equipo tiene ~2 calibraciones. ⚠️ **Crítico**: la calibración tiene implicaciones de calidad/cumplimiento. Preservar como dominio regulatorio.

7. **Sistema de facturación completo** (8 tablas): 1 principal + 5 detalles + 1 perfiles + 1 conciliación. Modelo de "factura con múltiples involucrados" (actividades, eventos, gastos, materiales, subcontrataciones).

8. **Sistema de técnicos** (5 tablas): `TbTecnicos`, `TbTecnicosAusencias`, `TbTecnicosFiestas`, `TbTipoTecnico`, `TbTipoTecnicoPrecios`.

9. **`EquipoMedidaCalibracion` con 10 callers** (alta utilización): `EquipoMedida.cls`, `Evento.cls`, `EventoEquipoMedida.cls`, etc. La calibración es central al sistema.

10. **Sistema de planificación** (5 tablas): `TbPlanificacion`, `TbPlanificacionAnexos`, `TbPlanificacionEquipos`, `TbPlanificacionRegistrada`, `TbAuxPlanificacion`.

11. **`Evento.Franquear` workflow complejo**: valida SLA + MotivosNoFranqueable antes de cerrar el evento. Múltiples validaciones.

12. **`tbUsuarios` no listado en staging** (no hay inventario en `00_main/`; los usuarios están en Lanzadera). Brass **lee identidad directamente de Lanzadera** vía `getdbLanzadera()` (acoplamiento).

13. **28 clases** (vs 29-52 de las otras apps). Brass tiene una superficie de código media pero un modelo de datos MUY rico (60 tablas).

14. **TYPO en tablas**: `LIbranza` (debería ser "Libranza"), `TbAuxManteniminetosPreventivosCalendario` (debería ser "Mantenimientos"). Errores de ortografía en nombres de clases y tablas.

## Checklist

- [x] Inventario funcional, formularios, clases y módulos documentados vía codegraph-vba.
- [x] Inventario real Dysflow del backend (60 tablas, 27 FKs, 5777 eventos, 25639 actividades).
- [x] Schema de `TbEventos` (38 columnas) y volumen de las principales.
- [x] **D104** propuesto para contraseña hardcodeada.
- [x] **D102 cross-cutting**: booleanos Sí/No inconsistentes cross-app.
- [x] APAP y APAP_WEB no aparecen en esta evidencia.
- [ ] Épica + tickets + matriz de migración de datos para Brass.

## Siguiente paso

Cruzar el inventario con la documentación previa en `OPENSPEC/00_BRASS`. Generar la **épica + tickets accionables + matriz de migración de datos** para Brass. **Cerrar D104** (contraseña hardcodeada) y **D102** (booleanos cross-cutting) en una iteración posterior con todas las apps. Continuar con la fase de épicas por aplicación según el alcance expandido.

## Core invariants

- **⚠️ D104 contraseña hardcodeada REAL**: la cadena `"dpddpd"` en `Variables Globales.bas:560` se pasa directamente a `OpenDatabase("MS Access;PWD=dpddpd")`. NO es fallback: si la contraseña rota en backend, no surte efecto sin cambio de código. Riesgo CRÍTICO; remediación operativa separada (rotación + saneamiento del repo + cambio a `SecretManagerPort`).
- **`getdb()` con 155 callers (intermedio en DAO-intensity)**: Brass comparte el patrón `getdb()` con el resto del ecosistema; la migración web preserva la cobertura de queries y traduce a repositorios con `Optional ByRef db` cuando aplica.
- **`IDEvento` es `Text(50)`, no `Long`**: todas las FKs del sistema (`TbEventos`, `TbActividades`, `TbAnexos`, `TbGastos`, etc.) se hacen por texto. La nueva plataforma decide si mantiene IDs como string (UUID) o los migra a numéricos; la decisión se documenta en la épica correspondiente.
- **Calibración es dominio regulatorio crítico**: `TbEquiposMedida` + `TbEquiposMedidaCalibraciones` tienen implicaciones de calidad/cumplimiento. La nueva plataforma preserva el dominio regulatorio como caso de uso de primer orden con tests específicos.
- **60 tablas + 27 FKs (modelo MUY rico)**: los cambios cross-cutting pasan por la `migration-matrix.md` de Brass; cualquier decisión de normalización afecta al menos 4-5 tablas y exige cobertura de tests antes de merge.

## Contributor checklist

- [ ] El cambio respeta las 5 reglas de §Core invariants; el `ci / quality` check pasa verde.
- [ ] Si el PR toca `Variables Globales.bas`, NO se reintroduce ninguna contraseña hardcodeada (D104, D93); el acceso a backend se hace vía `SecretManagerPort` (D9-D10) o variable de entorno.
- [ ] Si el cambio añade una migración Alembic, sigue `expand_and_contract` (D82): añadir columnas o tablas, sin `DROP` ni `ALTER` destructivos en la misma release.
- [ ] Si el cambio afecta al dominio de calibración, el PR incluye tests específicos (QC-5 + DA-2) que cubren el ciclo de vida del equipo y sus calibraciones.
- [ ] El PR es ≤ 400 líneas (`additions + deletions`); si no, partir por unidad de trabajo o encadenar.

## Navigation

Previous: [gestion-riesgos](../gestion-riesgos/README.md) | Next: [hps](../hps/README.md)