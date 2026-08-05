# Brass — integraciones y automatización

## Contratos observados

| Sistema | Uso | Evidencia |
|---|---|---|
| **Lanzadera** (identidad y permisos) | identidad, permisos, entorno (vía `getdbLanzadera()`) | `src/modules/Variables Globales.bas` (en `getUsuario`) |
| **Lanzadera** (datos de tablas vinculadas) | ruta del backend via `getRutaBackendDesdeTablasVinculadas` | `src/modules/Variables Globales.bas:461-535` |
| **SICA** (presumido vía código compartido) | presumiblemente vía tablas compartidas | (no inspeccionado en detalle) |
| **AGEDYS** (presumido) | presumiblemente vía código compartido | (no inspeccionado) |
| **Microsoft Word / plantillas** (presumido) | plantillas de informes (no inspeccionado) | `Informe.cls`, `GestorInforme.cls` |
| **Microsoft Excel** (presumido) | export de informes (no inspeccionado) | (presumido) |
| Correos | envío de correos (presumido vía servicio de correo interno) | (no inspeccionado en detalle) |
| SharePoint/ficheros | enlaces a anexos externos | `Anexo.cls` |

## ⚠️⚠️⚠️ D104 — Contraseña hardcodeada `"dpddpd"` como REAL

**Estado**: PROPUESTO. **CRÍTICO**. Detalle completo en [Matriz de migración § D104](migration-matrix.md#d104--contraseña-hardcodeada-dpddpd-como-real-en-variables-globalesbas560) y [Seguridad § D104](security-rules.md#d104--contraseña-hardcodeada-dpddpd-como-real-en-variables-globalesbas560).

**Resumen**: `Variables Globales.bas:560`:

```vba
Set db = wks.OpenDatabase(m_URL, False, False, "MS Access;PWD=" & "dpddpd" & "")
```

La contraseña `"dpddpd"` se pasa **directamente** a `OpenDatabase`. NO es un fallback (a diferencia de Condor D93). Es la **contraseña real del backend**. Si el repo `00_BRASS` se versiona, esta contraseña queda en git.

**Riesgo de seguridad CRÍTICO**: rotación de la contraseña NO surte efecto sin cambio de código.

**Acciones inmediatas**:
1. **Rotar la contraseña** del backend `Gestion_Brass_Gestion_Datos.accdb` (asumiendo que la real es distinta).
2. **Mover la contraseña a env var**: `passwordEnv: "BRASS_BACKEND_PASSWORD"` o similar.
3. **Modificar el código** para usar `Environ$(passwordEnv)` en vez del literal.
4. **Auditar el historial de git** del repo `00_BRASS` para confirmar si la contraseña real fue distinta en algún momento (en `00_BRASS/00_main/AGENTS.md` o `CHANGELOG.md`).
5. **Aplicar HR-3 del arnés dysflow**: cero secretos en código fuente, cero secretos en repos.

## Sistema de calibración (dominio regulatorio)

Brass tiene un **sistema completo de calibración de equipos de medida** que es **dominio regulatorio**:

- `TbEquiposMedida` (14 filas) — equipos de medida calibrados.
- `TbEquiposMedidaCalibraciones` (28 filas) — calibraciones con fecha.
- `TbEquiposCalibrables` (volumen TBD) — equipos que requieren calibración.
- `TbEquiposCalibrablesFechas` (volumen TBD) — fechas de calibración.

**Implicaciones para la nueva plataforma**:
- El sistema de calibración tiene implicaciones de **calidad/cumplimiento normativo**.
- Las calibraciones tienen **fechas críticas** (`FechaCalibracion`, `FechaVencimientoCalibracion`, presumido).
- El estado de calibración (`EstadoCalibracion`) se calcula desde las fechas.
- La nueva plataforma debe **preservar este dominio** con su semántica regulatoria.

## Sistema de facturación con múltiples involucrados

Brass tiene un **modelo de facturación único** en el ecosistema: una factura puede involucrar **múltiples entidades** (actividades, eventos, gastos, materiales, subcontrataciones). Esto se traduce a un modelo de **líneas de factura** o **detalles N:M** entre `TbFacturaPrincipal` y las entidades involucradas.

- 1 tabla principal: `TbFacturaPrincipal` (53 filas).
- 5 tablas de detalle: `TbFacturaActividadesInvolucradas`, `TbFacturaEventosInvolucrados`, `TbFacturaGastosInvolucrados`, `TbFacturaMaterialesInvolucrados`, `TbFacturaSubcontratacionesInvolucradas`.
- 1 tabla de perfiles: `TbFacturaPrincipalPerfiles`.
- 1 tabla de conciliación: `TbGuiaConciliaciones`.

## Franqueo con SLA (workflow declarativo)

`Evento.Franquear(p_CausaFin)` valida:
- `MotivosNoFranqueable` (validaciones de negocio).
- `MotivosNoFranqueableSLA` (validaciones de SLA).
- `m_FechaFin`, `m_HoraFin` (fecha y hora de cierre, **hardcodeado a 15:00:00** en el código).
- `m_CausaFin` (causa de fin proporcionada por el usuario).

Si pasa las validaciones, ejecuta un `UPDATE` que setea `FechaFinal`, `HORAFINALEVENTO`, `CAUSAFIN`, `Franqueado=True`. Es un **workflow de cierre** con validaciones previas.

⚠️ **D108 propuesta**: el workflow de franqueo se traduce a un **workflow declarativo** en la nueva plataforma (similar a D96 de Condor), con transiciones y roles requeridos. La regla "no franqueable sin causa" debe preservarse.

## Sistema BUI / Nodos / Subsistemas (jerarquía organizacional)

Brass tiene un **sistema de jerarquía organizacional** de 4 niveles:

- `TbBUI` (Business Unit Identifier).
- `TbNodos` (nodos dentro de BUI).
- `TbSubsistemaBui` (subsistemas de BUI).
- `TbNodoBUI` (relación nodo-BUI).
- `TbBuiIDEvento` (relación BUI-evento, presumida).

En la nueva plataforma, esto se traduce a un **modelo recursivo** (CTE o `ltree` de PostgreSQL).

## Sistema de técnicos (5 tablas)

- `TbTecnicos` (115 filas) — técnicos activos.
- `TbTecnicosAusencias` — ausencias (1:N).
- `TbTecnicosFiestas` — fiestas (1:N).
- `TbTipoTecnico` — tipos de técnico.
- `TbTipoTecnicoPrecios` — precios por tipo (1:N).

⚠️ **FKs por texto** (D106): `ALIAS` en `TbTecnicos` se usa como FK a otras tablas. **Inconsistencias de nombre** (`ALIAS` vs `ALIASTECNICO` vs `Alias`).

## Sistema de planificación (5 tablas)

- `TbPlanificacion` (624 filas).
- `TbPlanificacionAnexos` (1:N).
- `TbPlanificacionEquipos` (N:M con `TbEquipos`).
- `TbPlanificacionRegistrada` (1:1 con `TbPlanificacion`).
- `TbAuxPlanificacion` (auxiliar, presumido legacy).

## Configuración y flags de operación

Flags activos hoy en producción (inferidos de `Variables Globales.bas`):

| Flag / Campo | Valor por defecto | Significado |
|---|---|---|
| `EnDesarrollo` | (TempVar) | Modo desarrollo |
| `DatosEnLocal` | (TempVar) | Modo local |
| `m_ActiveBackendURL` | ruta al backend activo | URL del backend |
| `m_BackendSandboxURL` | ruta al sandbox | URL del sandbox (testing) |
| `m_BackendSandboxPassword` | contraseña del sandbox | Contraseña del sandbox |
| `m_PasswordBackend` | contraseña del backend | Contraseña del backend (⚠️ D104) |
| `m_TestingMode` | `True/False` | Modo testing → sandbox |

En la nueva plataforma estos flags se mueven a **configuración del módulo** (no TempVars, no código).

- `EnDesarrollo`, `DatosEnLocal` → variables de entorno del runner.
- `m_ActiveBackendURL` → config del puerto de persistencia.
- `m_BackendSandboxURL`, `m_BackendSandboxPassword`, `m_PasswordBackend` → secret manager / env var (**NUNCA en código**, ver D104).
- `m_TestingMode` → flag de runtime del runner (no en código de aplicación).

Para rutas UNC, hosts y nombres de máquina concretos: **NO se reproducen en este artefacto** (regla de evidencia).