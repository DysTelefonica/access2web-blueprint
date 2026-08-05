# Brass — seguridad y reglas

## Autorización

La matriz común se mantiene en [06-autorizacion-legacy-matriz.md](../../06-autorizacion-legacy-matriz.md). Aquí solo se conserva el comportamiento específico de Brass:

- **Tres roles** (Administrador / Calidad / Técnico) — análogo al resto del ecosistema.
- El acceso a funcionalidades específicas (alta de evento, franqueo, calibración) está condicionado por capacidades del rol.
- El sistema de calibración tiene **validaciones de SLA** en el franqueo (D108).

## Permisos por aplicación

Los permisos efectivos se cargan desde `UsuarioAplicacionPermisos` (clase compartida con Lanzadera). El acceso a Brass está condicionado por `IDAplicacion = "6"` (producción). En la nueva plataforma esto se reemplaza por el **catálogo de capabilities** (D45-D46) declarado por el módulo.

## Reglas de negocio críticas

- **Validación SLA en franqueo** (D108): `Evento.Franquear` valida `MotivosNoFranqueableSLA` antes de cerrar el evento. **Crítico para la nueva plataforma**: el SLA debe preservarse.
- **3 tipos de reparación mutuamente excluyentes** (`TipoRepInsitu`, `TipoRepNoSMT`, `TipoRepValvulas` como YesNo): migrar a un enum en PostgreSQL.
- **`Franqueado` es YesNo (type 1) — not `Text 2`**: a diferencia de otras apps, Brass usa `YesNo` real. **Inconsistencia cross-cutting menor** con Lanzadera/Expedientes/Etc. (D102).
- **FKs por texto** (D106): `NODO`, `BUI`, `SUBSISTEMA`, `ALIASTECNICO`, `ORIGINADOR`, `IDParte` son Text. Decisión de diseño legacy que se preserva.
- **`IDEvento` Text(50)** (D105): todas las FKs del sistema se hacen por texto. La nueva plataforma debe decidir si mantiene IDs como string (UUID) o los migra a numéricos.
- **Sistema de calibración regulatorio**: `TbEquiposMedidaCalibraciones` con fechas críticas. Preservar como dominio regulatorio.
- **Sistema de facturación con múltiples involucrados**: 1 principal + 5 detalles + 1 perfiles + 1 conciliación. Migrar a un modelo normalizado en PostgreSQL con FKs explícitas.
- **Sistema BUI / Nodos / Subsistemas** (jerarquía organizacional de 4 niveles). Migrar a modelo recursivo (CTE o `ltree`).
- **Transaccionalidad**: las operaciones de evento (alta, edición, franqueo) usan transacciones DAO implícitas. La nueva plataforma usa transacciones SQLAlchemy `AsyncSession.begin()` (D66, D82).

## Roles y funciones diferenciadas

Brass separa explícitamente **Calidad** de **Técnico** (similar a otras apps). Las pantallas de calidad validan SLA + cierres; las de técnico gestionan altas y partes.

## D104 · Contraseña hardcodeada `"dpddpd"` como REAL en `Variables Globales.bas:560`

**Estado**: PROPUESTO. **CRÍTICO** ⚠️⚠️⚠️. Detalle completo en [Matriz de migración § D104](migration-matrix.md#d104--contraseña-hardcodeada-dpddpd-como-real-en-variables-globalesbas560) y [Integraciones § D104](integrations-automation.md#d104--contraseña-hardcodeada-dpddpd-como-real).

### Hallazgo crítico

`Variables Globales.bas:560`:

```vba
Set db = wks.OpenDatabase(m_URL, False, False, "MS Access;PWD=" & "dpddpd" & "")
```

**`"dpddpd"` es la contraseña REAL del backend** (no un fallback como en Condor D93). Se pasa directamente a `OpenDatabase`.

### Por qué es crítico

1. **Visibilidad en código fuente**: cualquier persona con acceso al repo ve la contraseña.
2. **Si el repo se versiona en git** (que es probable), la contraseña queda en el historial de git. `git log --all -p -- "*.bas" | grep dpddpd` la encuentra.
3. **Rotación de la contraseña NO surte efecto** sin cambio de código. Si se rota la contraseña del backend pero no se actualiza el código, la app deja de funcionar.
4. **HR-3 del arnés dysflow** lo prohíbe: cero secretos en código fuente, cero secretos en repos.

### Acciones inmediatas (operativas en `00_BRASS/00_main/`)

1. **Rotar la contraseña** del backend `Gestion_Brass_Gestion_Datos.accdb` (asumiendo que la real es distinta de `"dpddpd"`).
2. **Modificar `Variables Globales.bas:560`** para usar `Environ$("BRASS_BACKEND_PASSWORD")` (env var) en vez del literal `"dpddpd"`.
3. **Migrar a secret manager** (Vault, AWS Secrets Manager, etc.) en producción.
4. **Auditar el historial de git** del repo `00_BRASS` con `git log --all -p -- "*.bas" | grep "PWD="` para confirmar si la contraseña real fue distinta en algún momento.
5. **Si la contraseña real es distinta**, **aplicar `git-filter-repo`** para eliminar el secreto del historial.

### Recomendación cross-cutting

Aplicar el mismo patrón a **todos los repos** que usen Access. **D104 complementa D93** (Condor fallback). La regla HR-3 del arnés dysflow debe respetarse en **todos los proyectos consumer**.

## D105 · `IDEvento` es `Text(50)`, no Long

`IDEvento` (y todos los IDs principales de Brass) son `Text(50)`. Esto es una **decisión de diseño legacy** que se debe **documentar y preservar** en la nueva plataforma (coherente con D106).

## D106 · FKs por texto (coherente con D105)

`NODO`, `BUI`, `SUBSISTEMA`, `ALIASTECNICO`, `ORIGINADOR`, `IDParte` son `Text` y se usan como FKs. **Inconsistencias de nombre** (`ALIAS` vs `ALIASTECNICO` vs `Alias`, `TIPO` vs `TIPOTECNICO`). Migración: unificar nombres y agregar columnas numéricas como `*IdNum`.

## D107 · Booleanos como YesNo (mejor que Text 2)

A diferencia de otras apps, Brass usa `YesNo` real (type 1) en la mayoría de las columnas booleanas. **Inconsistencia cross-cutting D102**: otras apps usan `Text 2`. Brass está mejor en este aspecto. **Recomendación**: en la nueva plataforma, estandarizar a `BOOLEAN` (Brass ya lo hace correctamente).

## D108 · Workflow de franqueo con SLA (workflow declarativo)

`Evento.Franquear` valida:
- `MotivosNoFranqueable` (validaciones de negocio).
- `MotivosNoFranqueableSLA` (validaciones de SLA).
- `m_FechaFin`, `m_HoraFin` (fecha y hora de cierre, **hardcodeado a 15:00:00** en el código).

Si pasa las validaciones, ejecuta un `UPDATE` que setea `FechaFinal`, `HORAFINALEVENTO`, `CAUSAFIN`, `Franqueado=True`. Es un **workflow de cierre** con validaciones previas.

**Recomendación**: traducir a un **workflow declarativo** en la nueva plataforma (similar a D96 de Condor), con transiciones y roles requeridos. La regla "no franqueable sin causa" debe preservarse.

## Riesgos de privacidad/migración

- **Datos personales** (presumidos en `TbTecnicos`, `TbMaterialSeguimiento`, observaciones de eventos, etc.) — requieren política de manejo de PII.
- **Adjuntos** (`Anexo.cls`, 7592 filas) — pueden contener datos sensibles. Manejo vía object storage (D16) con autorización.
- **Calibraciones** (`TbEquiposMedidaCalibraciones`) — datos de cumplimiento normativo. Retención al menos la misma que la de auditoría (D29).
- **Facturas con múltiples involucrados** (`TbFactura*`) — datos financieros. Retención al menos la misma que la de auditoría.
- **APAP y APAP_WEB** — no aparecen ni se mencionan (proyecto personal del desarrollador; regla transversal del blueprint).