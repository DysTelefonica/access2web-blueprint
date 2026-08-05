# HPS — seguridad y reglas

## Autorización

La matriz común se mantiene en [06-autorizacion-legacy-matriz.md](../../06-autorizacion-legacy-matriz.md). Aquí solo se conserva el comportamiento específico de HPS:

- **Tres roles** (Administrador, Calidad, Técnico), análogo al resto del ecosistema.
- El acceso a funcionalidades específicas (renovación, baja, modificación de histórico) está condicionado por capacidades del rol.
- `Form_FormInicial02ConsultasPrincipal` + 8 formularios de consulta permiten lecturas filtradas por criterios (contratistas, datos, empresa, estado, grado, jurídica, motivo HPS, nombre, proyecto asignado).

## Permisos por aplicación

Los permisos efectivos se cargan desde `UsuarioAplicacionPermisos` (clase compartida con Lanzadera). El acceso a HPS está condicionado por `IDAplicacion = "17"` (producción) o `"51"` (pruebas). En la nueva plataforma esto se reemplaza por el **catálogo de capabilities** (D45-D46) declarado por el módulo.

## Reglas de negocio críticas

- **Cálculo de estado del HPS derivado**: `APuntoDeCaducar`, `Caducado`, `Solicitado`, `PendienteRenovacion` se calculan vía `DatosHPSCalculados` desde fechas (`F_Curso`, `FechaHPSConcesionMinima`, `Requiere_Curso`, `CursoEnVigor`) y flags. La nueva plataforma debe **persistir el valor fuente y el calculado** (no inferir el uno del otro).
- **Transaccionalidad de selección de anexos**: `AnexoSelectionTransactionCoordinator` garantiza atomicidad al seleccionar/mover anexos entre usuarios. Crítico para trazabilidad.
- **Transaccionalidad de lifecycle de usuarios**: `UsuarioLifecycleTransactionCoordinator` coordina alta/baja/modificación de usuarios con histórico y observaciones. Crítico para auditoría.
- **Sincronización histórico ↔ anexos**: `Mod_Sincronizacion_Historico.bas` mantiene coherencia entre `TbUsuariosHistoricos` y `TbAnexosUsuariosHistoricos`.
- **Datos personales sensibles**: `DNI`, `Nombre`, `Apellido_1`, `Apellido_2`, `Telefono`, `Correo_e`, `F_Nacimiento`, `LugarNacimiento` requieren manejo especial (ver D92).
- **FKs conceptuales**: `IDExpediente`, `IDEmpresaUsuario`, `IDEmpresaHPS`, `IDJuridicaContrato`, `IDSolicitud` son referencias por convención (sin FK física). Mantener o formalizar en migración es decisión pendiente (D92).
- **LocalReadAccessAuthorization**: `Test_LocalReadAccessAuthorization.bas` valida que el acceso local a datos sigue las reglas de autorización. Test crítico a portar.

## Roles y funciones diferenciadas

HPS separa explícitamente **Calidad** de **Técnico** (mismo patrón que las otras apps):

- **Calidad**: tareas de validación de HPS, modificación de estados, gestión de histórico.
- **Técnico**: tareas operativas, mantenimiento de datos.

Las pantallas de calidad y técnico tienen worklists distintas en `Form_FormInicial*.cls`. La nueva plataforma debe mantener esta separación como **vista especializada** del módulo (D46), no como aplicaciones distintas.

## Indicadores en tiempo real con kill switch atómico

`Test_RealTimeIndicatorCoherence.bas` valida la coherencia del sistema de indicadores en tiempo real (`Indicador.cls`, `clsIndicadoresBus.cls`, `modIndicadores.bas`). El kill switch atómico se preserva como referencia del **puerto de observabilidad** de la nueva plataforma (métricas Prometheus / OpenTelemetry).

## ⚠️ HALLAZGO CRÍTICO DE SEGURIDAD: caché local en frontend con datos personales

HPS es la **única aplicación del ecosistema con caché en el frontend** (patrón no presente en Lanzadera/Expedientes/Gestion_Riesgos/NoConformidades). El frontend `HPS.accdb` (30 MB) contiene **12 tablas locales** que funcionan como caché sincronizado con el backend, incluyendo:

### Tablas con datos personales (CRÍTICO ⚠️)

- **`TbDatosLocal`** (344 filas): contiene TODOS los datos personales de los usuarios HPS activos (DNI, Nombre, Apellido_1, Apellido_2, Teléfono, Correo_e, F_Nacimiento, LugarNacimiento) + estado HPS desnormalizado por organismo (NAC, OTAN, ESA, UE) + Observaciones.
- **`TbUsuariosHistoricosLocal`** (235 filas): datos personales del histórico de usuarios (DNI, nombres, fechas, correo).

### Tablas auxiliares (sensibilidad media-baja)

- `TbCursosLocal` (caché de cursos).
- `TbSuministradoresLocal` (caché de suministradores).
- `TbDatosLocalParaIndicadores` (caché para indicadores).
- `TbUsuariosSICALocalParaIndicadores` (caché SICA para indicadores).
- `TbConfiguracionHPS` (clave-valor local del frontend: Clave, Valor, Activo, FechaModificacion, UsuarioModificacion).
- `TbVinculosTablas`, `tblInfo`, `tblSettings` (metadatos locales).

### Tabla vacía (descartar)

- `TbUsuariosSICALocal` (0 filas, no se usa).

### Riesgos críticos

1. **`.gitignore` del repo NO excluye `*.accdb`**: solo `*.accde`, `*.mdb`, `*.mde`, `HPST.accdb` (específico). Si `HPS.accdb` se versiona, **344 usuarios con datos personales quedan en el historial de git**. Esto es un riesgo de exposición de PII (Personally Identifiable Information) grave.
2. **El frontend se mueve con los datos**: copia a otra máquina, backup no cifrado, robo de portátil = exposición.
3. **Sin política de retención**: los datos locales pueden persistir después de baja del usuario.
4. **Caché desactualizada**: el frontend puede mostrar datos obsoletos sin que el usuario lo sepa.

### Recomendaciones inmediatas (operativas en el repo `00_HPS`, no desde aquí)

1. **Agregar `*.accdb` al `.gitignore`** del repo `00_HPS` (cambio de configuración operativo).
2. **Verificar que `HPS.accdb` no esté en el historial de git** de `00_HPS` (`git log --all -- HPS.accdb`). Si está, **rotar el repo y aplicar git-filter-repo** para eliminarlo del historial.
3. **Cifrar el frontend** o **extraer datos personales a un esquema separado** que pueda rotarse independientemente.
4. **Documentar la política de privacidad** que aplica tanto al frontend como al backend.

### Recomendaciones de migración (en la nueva plataforma)

- Las 12 tablas locales del frontend **NO se reproducen como caché en cliente**. La nueva plataforma usa caché server-side (Redis u opción detrás del puerto D70-D71).
- `TbDatosLocal` y `TbUsuariosHistoricosLocal` → migrar como **vistas materializadas o queries server-side** con caché detrás del puerto (D70-D71).
- `TbCursosLocal`, `TbSuministradoresLocal`, `TbDatosLocalParaIndicadores`, `TbUsuariosSICALocalParaIndicadores` → idem.
- `TbUsuariosSICALocal` (vacía) → descartar.
- `TbVinculosTablas`, `tblInfo`, `tblSettings`, `TbConfiguracionHPS` → migrar como **config del módulo** (no tablas, sino variables de entorno o config centralizada).

## Riesgos de privacidad/migración

- **Datos personales** (DNI, nombres, fechas, correos, teléfonos) — requieren política explícita (D92). Las plantillas de informes que muestren estos campos deben seguir la misma política.
- **Histórico** (242 filas) — trazabilidad de auditoría. La retención debe ser al menos la misma que la de auditoría (D29).
- **Backends en `00_main/`** (`HPST.accdb`) — NO se usan como autoridad. El backend autoritativo está en `C:\00repos\datos\`. La copia local es legacy.
- **Plantillas de correo** (`Correo.cls`) — pueden contener datos sensibles; revisar antes de portar.
- **APAP y APAP_WEB** — no aparecen ni se mencionan (proyecto personal del desarrollador; regla transversal del blueprint).