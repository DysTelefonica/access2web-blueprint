# Condor — matriz de migración (scaffold de preservación)

Esta matriz no decide el esquema PostgreSQL. Su regla es conservadora: todo campo/registro se conserva hasta que negocio apruebe una disposición. Disposiciones específicas de Condor en § D93-D95.

| Fuente | Significado/candidato de dominio | Transformación | Estado | Validación/rechazo | Decisión abierta |
|---|---|---|---|---|---|
| `tbSolicitudes` (1 fila en staging) | solicitud de calidad | separar identidad, tipo, estado, auditoría, vinculaciones | mapeado preliminar | volumen real en producción | **FK conceptual sin constraint (D94)** |
| `tbAdjuntos` (0 filas en staging) | adjuntos de una solicitud | FK a `tbSolicitudes` + metadatos + path externo | external reference | comprobar existencia de fichero/URL sin copiar datos | repositorio destino (D16) |
| `tbDatosCDCA` (volumen TBD) | datos de Solicitud tipo CD_CA | FK a `tbSolicitudes` + campos específicos del tipo | mapeado preliminar | cardinalidad; ⚠️ staging sin datos reales | cardinalidad vs Solicitud |
| `tbDatosCDCASUB` (volumen TBD) | datos de Solicitud tipo CD_CA_SUB | FK a `tbSolicitudes` + campos específicos del tipo | mapeado preliminar | cardinalidad | cardinalidad vs Solicitud |
| `tbDatosPC` (volumen TBD) | datos de Solicitud tipo PC | FK a `tbSolicitudes` + campos específicos del tipo | mapeado preliminar | cardinalidad | cardinalidad vs Solicitud |
| `tbDatosPCSUB` (volumen TBD) | datos de Solicitud tipo PC_SUB | FK a `tbSolicitudes` + campos específicos del tipo | mapeado preliminar | cardinalidad | cardinalidad vs Solicitud |
| `tbEstados` (9 filas) | catálogo de estados del workflow | entidad de catálogo | mapeado preliminar | unicidad; coherencia con `tbLogEstados` | versionado |
| `tbTransiciones` (volumen TBD) | transiciones de estado registradas | entidad de auditoría + FK a `tbEstados` (origen y destino) | mapeado preliminar | secuencialidad; **FK a `tbSolicitudes` no explícita** ⚠️ | migrar con FK explícita |
| `tbHistorialRechazos` (volumen TBD) | histórico de rechazos | FK a `tbSolicitudes` + FK a `tbRechazos` | mapeado preliminar | trazabilidad | retención |
| `tbRechazos` (0 filas en staging) | rechazos de solicitudes | entidad con motivo, fecha, autor | mapeado preliminar | cardinalidad | retención |
| `tbLogCambios` (volumen TBD) | log de cambios general | evento + timestamp + actor + tabla afectada | mapeado preliminar | trazabilidad | **traducir a logs estructurados canónicos (D27)** |
| `tbLogErrores` (3 filas en staging) | log de errores | evento + timestamp + tipo + mensaje | mapeado preliminar | trazabilidad | **traducir a logs estructurados canónicos (D27)** |
| `tbLogEstados` (volumen TBD) | log de cambios de estado | evento + FK a `tbSolicitudes` + estado anterior/nuevo + timestamp | mapeado preliminar | trazabilidad; **FK a `tbSolicitudes` y `tbEstados` no explícitas** ⚠️ | migrar con FK explícita |
| `tbMapeoCampos` (volumen TBD) | mapeo de columnas legacy ↔ modernas | entidad de configuración | mapeado preliminar | unicidad; uso real | deprecate o migrar a mapping config |
| `tbValidacionRevision` (volumen TBD) | validación de revisión de calidad | FK a `tbSolicitudes` + estado + comentarios + autor + fecha | mapeado preliminar | trazabilidad | retención |
| `tbConfiguracion` (no en staging; presumido en `00_main`) | configuración de la app | config + flags | mapeado preliminar | unicidad | migrar a config del módulo |
| `tbConfiguracionBackends` (no listado en staging; presumido) | config de backend PROD/SANDBOX/TEST | config + flag activo + IDAplicacion | mapeado preliminar | unicidad | migrar a config del puerto (D9-D10) |
| catálogos compartidos con Lanzadera (vía `getdbLanzadera`) | identidad / permisos / entorno | identidad vía adaptador unificado (D9-D10) | mapeado preliminar | coherencia | acoplamiento directo a romper (D86/D87) |
| integraciones externas (Lanzadera/Expedientes, NoConformidades, Gestion_Riesgos) | referencias externas | claves externas y snapshot contractual | external reference | reconciliar por identificador, sin asumir propiedad | ownership y sincronización |

## Ledger

- **Mapeado preliminar:** cabecera de Solicitudes, 4 tipos de datos, estados, adjuntos, rechazos, logs.
- **Necesita decisión:** FKs conceptuales a formalizar (D94), password hardcoded (D93), vinculación con NC (D95).
- **External reference:** adjuntos, identidad, Lanzadera, NoConformidades.
- **No hay campos/filas declarados obsoletos** (salvo los marcados en D93-D95).

## D93 · Password hardcoded como fallback en `GetPasswordDB`

**Estado**: PROPUESTO. Riesgo de seguridad **CRÍTICO** ⚠️⚠️⚠️.

### Hallazgo

`src/modules/FUNCIONES UTILES.bas:150` (y referencias en el código de staging):

```vba
Public Function GetPasswordDB() As String
    On Error GoTo Errores
    Dim objEntorno As New entorno
    GetPasswordDB = objEntorno.PasswordDB
    Exit Function
Errores:
    ' Fallback hardcoded legacy
    GetPasswordDB = "dpddpd"
End Function
```

Si la lectura del INI (`objEntorno.PasswordDB`) falla por cualquier razón, **la función cae al literal `"dpddpd"`**. Esta contraseña:

1. **Está hardcodeada en código fuente** — visible para cualquiera con acceso al repo.
2. **Es la misma para todos los entornos** (PROD, SANDBOX, TEST).
3. **Si fue rotada en algún momento, el fallback sigue siendo la versión vieja** — el código se desincroniza con el binario.
4. **Si fue rotada y la nueva está en el INI pero el INI falla por otro motivo** (permisos, encoding), el código cae a la versión vieja — **bypass de la rotación**.

### Recomendaciones inmediatas

1. **Eliminar el fallback hardcoded**. Si `GetPasswordDB` falla, **fallar explícitamente** con `Err.Raise` y mensaje claro.
2. **Cargar la contraseña solo desde el INI** (o env var en la nueva plataforma).
3. **Rotar la contraseña real del binario** (por si ya fue expuesta).
4. **Investigar el historial de git** para confirmar si la contraseña `"dpddpd"` es la real o una dummy.
5. **Aplicar la regla HR-3 del arnés dysflow**: cero secretos en código fuente, cero secretos en repos.

### Impacto cross-cutting

El mismo patrón puede existir en otras apps (Lanzadera, Gestion_Riesgos, NoConformidades, HPS). Acción: auditar `GetPasswordDB` / equivalentes en los 8 repos.

## D94 · FKs conceptuales sin constraint

**Estado**: PROPUESTO. Disposiciones pendientes.

| Tabla | Columna | FK conceptual | Recomendación |
|---|---|---|---|
| `tbSolicitudes` | `idEstadoInterno` | `tbEstados` | **formalizar FK** en PostgreSQL (migración de datos debe garantizar coherencia) |
| `tbSolicitudes` | `idExpediente` | Lanzadera `TbExpedientes` | mantener como referencia conceptual (D86/D87); FK cross-app requiere acuerdo de orden de migración |
| `tbSolicitudes` | `idNCAsociada` | NoConformidades `TbNoConformidades` | mantener como referencia conceptual; FK cross-app idem |
| `tbTransiciones` | (presumido: `idSolicitud`, `idEstadoOrigen`, `idEstadoDestino`) | `tbSolicitudes`, `tbEstados` | formalizar FK en PostgreSQL |
| `tbLogEstados` | (presumido: `idSolicitud`, `idEstadoAnterior`, `idEstadoNuevo`) | `tbSolicitudes`, `tbEstados` | formalizar FK en PostgreSQL |

## D95 · Vinculación con NoConformidades

**Estado**: PROPUESTO.

`Form_frmGestionSolicitud.cls:1908` verifica la NC vinculada antes de eliminar:

```vba
Dim ncServ As New NoConformidadServicio
Dim ncVinculada As NoConformidad
Dim mensajeConfirmacion As String
Set ncVinculada = ncServ.getNoConformidadPorCodigoCondor(m_ViewModel.Solicitud.codigoSolicitud)
```

La columna `idNCAsociada` en `tbSolicitudes` es la FK conceptual. La relación es:

- **Código de NC ↔ Código de Solicitud de Condor** (relación por código, no por ID directa).
- `NoConformidadServicio.getNoConformidadPorCodigoCondor` busca NC por código de Condor.

Recomendación:

- **Mantener como referencia conceptual** en PostgreSQL (FK cross-app requiere orden de migración estricto: NoConformidades antes que Condor).
- **Documentar la regla de negocio**: `idNCAsociada` solo se setea cuando hay NC explícitamente vinculada; no es FK directa.
- **Validar en la migración**: si hay NCs vinculadas, deben migrarse antes que las Solicitudes de Condor que las referencian.

## D96 · Workflow declarativo via `tbTransiciones`

**Estado**: PROPUESTO.

`tbTransiciones` (4 columnas: `idTransicion`, `idEstadoOrigen`, `idEstadoDestino`, `rolRequerido`) define el **workflow declarativo** de Condor: qué transiciones de estado son válidas y qué rol se requiere para ejecutarlas. Es la representación del workflow en datos, no en código.

Recomendación:

- **Preservar como tabla de datos** en PostgreSQL (`transiciones` con FKs a `estados(origen)` y `estados(destino)`, y `rol_requerido` como enum).
- **La columna `rolRequerido` se traduce a una verificación de capabilities** en la nueva plataforma (D45-D46). El workflow service (`WorkflowServicio.cls`) lee esta tabla y valida en runtime.
- **Disponer de un endpoint admin** para que el workflow pueda evolucionar sin deploys de código.
- **Disponer de un endpoint de "transiciones disponibles"** que devuelva, para un usuario dado y un estado actual, las transiciones que puede ejecutar.

## D97 · `tbMapeoCampos` con 183 filas — config que se preserva como datos

**Estado**: PROPUESTO.

`tbMapeoCampos` tiene **183 filas** en el backend autoritativo. Es **config de mapeo entre columnas legacy y modernas** que se preserva como datos, no como código. Esto significa que la nueva plataforma debe:

- **Migrar las 183 filas como `INSERT INTO tb_mapeo_campos VALUES (...)`** en la migración inicial.
- **Exponer un servicio de mapeo** (`MapeoServicio.cls` ya existe en staging) que la nueva plataforma use para resolver dinámicamente las equivalencias entre columnas legacy y modernas.
- **Disponer de un endpoint admin** para mantener el mapeo sin deploys.

Recomendación:

- **Preservar como tabla `mapeo_campos` en PostgreSQL** con PK + columnas equivalentes a la legacy.
- **NO migrar como código** (mapeos hardcodeados en `MapeoServicio.cls`). El mapeo es **datos de runtime**.
- **Disponer de UI admin** (CRUD sobre `mapeo_campos`) con control de capabilities (D45).