# HPS — matriz de migración (scaffold de preservación)

Esta matriz no decide el esquema PostgreSQL. Su regla es conservadora: todo campo/registro se conserva hasta que negocio apruebe una disposición. Disposiciones específicas de HPS en § D92.

| Fuente | Significado/candidato de dominio | Transformación | Estado | Validación/rechazo | Decisión abierta |
|---|---|---|---|---|---|
| `TbUsuarios` (345 filas) | usuario HPS activo | separar identidad, datos personales, fechas, flags, FKs conceptuales | mapeado preliminar | unicidad por `DNI` (?), reconciliación | **campos personales sensibles (D92)** |
| `TbUsuariosHistoricos` (242 filas) | histórico de usuarios | snapshot conservado para auditoría | mapeado preliminar | trazabilidad; integridad con `TbUsuarios` actual | **duplicación frontend/backend (D92)** |
| `TbUsuariosSICA` | usuario en sistema SICA externo | FK externa + snapshot | external reference | reconciliación con `SICA` (fuera de scope) | ownership externo |
| `TbUsuariosEntidades` | relación usuario ↔ entidad | entidad de relación N:N | mapeado preliminar | cardinalidad | versionado |
| `TbHPS` (1280 filas) | relación usuario ↔ curso HPS | FK a `TbUsuarios` + metadatos del curso | mapeado preliminar | cardinalidad; **duplicación frontend/backend (D92)** | versionado |
| `TbHPSEquivalencia` | equivalencias entre cursos HPS | entidad de catálogo | mapeado preliminar | unicidad | versionado |
| `TbHPSGrado` | grado del HPS | catálogo | mapeado preliminar | unicidad | versionado |
| `TbMotivoHPS` | motivos de alta/baja HPS | catálogo | mapeado preliminar | unicidad | versionado |
| `TbAuxCursos` | cursos obligatorios | catálogo | mapeado preliminar | unicidad | versionado |
| `TbAnexosUsuariosHPS` | anexos de un usuario HPS | FK a `TbUsuarios` + metadatos + path externo | external reference | comprobar existencia de fichero/URL sin copiar datos | repositorio destino (D16) |
| `TbAnexosUsuariosHistoricos` | anexos del histórico | snapshot conservado | mapeado preliminar | trazabilidad | retención |
| `TbAnexosUsuariosSICA` | anexos SICA | FK externa + metadatos | external reference | reconciliación con SICA | ownership externo |
| `TbUsuarioAnexos` | relación usuario ↔ anexo (alternativa) | entidad de relación | mapeado preliminar | cardinalidad; ⚠️ `TbAnexosUsuariosHPS` parece solapar | consolidar o separar |
| `TbObservaciones` (330 filas) | observaciones sobre usuarios | FK a `TbUsuarios` + texto + fecha + autor | mapeado preliminar | trazabilidad; ⚠️ FK PK-to-PK | reescribir FK por `IDUsuario` |
| `TbObservacionesHistoricas` | histórico de observaciones | snapshot conservado | mapeado preliminar | trazabilidad; ⚠️ FK PK-to-PK | reescribir FK por `IDUsuarioHistorico` |
| `TbConsultas` | queries SQL pre-armadas | entidad de query parametrizable | mapeado preliminar | unicidad; SQL injection | deprecate o migrar a query builder |
| `TbJuridicasContratacion` | catálogo de jurídicas | entidad de catálogo | mapeado preliminar | unicidad; coherencia con Lanzadera | versionado |
| **Copia de TbExpedientes** | copia legacy | conservar en zona legacy | **legacy copy (D92)** | comparar recuentos y uso en código | archivo vs dominio |
| **Copia de TbExpedienteLugares** | copia legacy | conservar en zona legacy | **legacy copy (D92)** | comparar recuentos y uso en código | archivo vs dominio |
| **Copia de TbUsuarios** | copia legacy | conservar en zona legacy | **legacy copy (D92)** | comparar recuentos y uso en código | archivo vs dominio |
| **Copia de TbUsuariosEntidades** | copia legacy | conservar en zona legacy | **legacy copy (D92)** | comparar recuentos y uso en código | archivo vs dominio |
| **Errores de pegado** | sentinel de errores en pegado masivo | preservar hasta decisión | **sentinel (D92)** | comparar recuentos | transformar en log estructurado canónico (D27) o mantener |
| catálogos compartidos con Lanzadera (vía `getdbLanzadera`) | identidad / permisos / entorno | identidad vía adaptador unificado (D9-D10) | mapeado preliminar | coherencia | acoplamiento directo a romper (D86/D87) |
| tablas E2E/temp/aux (si las hay) | soporte técnico | migrar solo si se requiere trazabilidad; conservar dump si no | technical-only candidate | documentar antes de excluir | decisión explícita, no obsolescencia |
| integraciones externas (Lanzadera, SICA, HPS_Solicitudes, AGEDYS, correos) | referencias externas | claves externas y snapshot contractual | external reference | reconciliar por identificador, sin asumir propiedad | ownership y sincronización |

## Ledger

- **Mapeado preliminar:** cabecera de usuarios, histórico, HPS, observaciones, anexos, indicadores.
- **Necesita decisión:** flags frente a tipo, formalización de FK conceptuales, retención de histórico.
- **Legacy copy:** 4 tablas "Copia de..." — ver D92.
- **Sentinel:** `Errores de pegado` — ver D92.
- **External reference:** anexos, SICA, identidad, Lanzadera, HPS_Solicitudes.
- **No hay campos/filas declarados obsoletos** (salvo los marcados en D92).

## D92 · Disposiciones específicas de HPS

**Estado**: PROPUESTO. Disposiciones pendientes de aprobación con negocio antes de fase SDD.

### 1. Campos personales sensibles en `TbUsuarios`

Las columnas marcadas con **dato personal (D92)** en [data-model.md](data-model.md#esquema-real-tbusuarios-27-columnas-tipos-reales) requieren política de manejo:
- `DNI`, `Nombre`, `Apellido_1`, `Apellido_2`, `Telefono`, `Correo_e`, `F_Nacimiento`, `LugarNacimiento`.

Recomendación provisional:
- **Preservar en PostgreSQL** sin transformaciones para mantener paridad funcional.
- **Enmascarar en logs y observabilidad** (no loguear valores completos).
- **Documentar en matriz de capabilities** quién puede ver cada campo (D45).
- **Evaluar encriptación en reposo** como follow-up de seguridad (no bloqueante para la migración inicial).

### 2. Tablas "Copia de..." (4 legacy copies)

Las 4 tablas `Copia de Tb*` son legado de operaciones de copia previas a cambios masivos. Recomendación provisional:
- **Inventariar contenido** (count_rows) en una segunda pasada.
- **Migrar a una zona `legacy` del esquema PostgreSQL** con retención indefinida.
- **No crear endpoints ni UI** sobre estas tablas.
- **Marcar como deprecated** en la documentación del módulo.

### 3. Sentinel `Errores de pegado`

Tabla para capturar errores en operaciones de pegado masivo (imports batch). Recomendación provisional:
- **Evaluar** si la lógica de captura es específica de HPS o replicable.
- Si es replicable, **migrar a logs estructurados canónicos** (D27).
- Si es específica de HPS, **preservar como tabla de auditoría** con retención limitada.

### 4. ⚠️ HALLAZGO CRÍTICO: caché local en frontend con datos personales

El frontend `HPS.accdb` (30 MB) tiene **12 tablas locales** que funcionan como caché sincronizado con el backend. Las más sensibles (D92):

- `TbDatosLocal` (344 filas) — **contiene TODOS los datos personales de los usuarios HPS activos** (DNI, Nombre, Apellidos, Teléfono, Correo, Fecha de Nacimiento, Lugar de Nacimiento, Observaciones) + estado HPS desnormalizado por organismo (NAC, OTAN, ESA, UE). Es la representación completa de un usuario HPS con todos los datos sensibles.
- `TbUsuariosHistoricosLocal` (235 filas) — **datos personales del histórico de usuarios** (DNI, nombres, fechas).
- `TbCursosLocal`, `TbSuministradoresLocal`, `TbDatosLocalParaIndicadores`, `TbUsuariosSICALocalParaIndicadores` — cachés auxiliares con datos de negocio (no personales).
- `TbUsuariosSICALocal` (0 filas, vacía, no se usa).
- `TbVinculosTablas`, `tblInfo`, `tblSettings`, `TbConfiguracionHPS` — metadatos locales del frontend.

**Causa**: HPS usa el frontend como caché local porque Access + red + tablas vinculadas es lento. Patrón legacy que NO debe reproducirse en la nueva plataforma.

**Riesgos de seguridad**:
1. **El `.gitignore` del repo NO excluye `*.accdb`** (solo `*.accde`, `*.mdb`, `*.mde`, `HPST.accdb`). Si `HPS.accdb` se versiona, **344 usuarios con datos personales quedan en el historial de git**.
2. **Si el frontend se copia a otra máquina**, los datos personales se mueven (robo de portátil, backup no cifrado, etc.).
3. **La caché local se desactualiza** y puede contener datos obsoletos que el usuario ve como "actuales".
4. **No hay política de retención** sobre los datos locales — pueden persistir después de baja del usuario.

**Recomendaciones inmediatas**:
1. **Agregar `*.accdb` al `.gitignore`** del repo (este hallazgo aplica a `00_HPS` y a cualquier otro repo que siga el patrón).
2. **Verificar que `HPS.accdb` no esté en el historial de git** (comprobar con `git log --all -- HPS.accdb` en el repo afectado — operación a hacer en `00_HPS`, no desde aquí).
3. **Considerar cifrar el frontend o extraer datos personales** a un esquema separado que pueda rotarse/cifrarse independientemente.
4. **Documentar la política de privacidad** que aplica tanto al frontend como al backend.

**Recomendaciones de migración**:
- `TbDatosLocal` y `TbUsuariosHistoricosLocal` → migrar a **vistas materializadas o queries server-side** en PostgreSQL con caché detrás del puerto (D70-D71). NO persistir en cliente.
- `TbCursosLocal`, `TbSuministradoresLocal`, `TbDatosLocalParaIndicadores`, `TbUsuariosSICALocalParaIndicadores` → idem.
- `TbUsuariosSICALocal` (vacía) → descartar.
- `TbVinculosTablas`, `tblInfo`, `tblSettings`, `TbConfiguracionHPS` → migrar como config del módulo (no tablas, sino variables de entorno o config centralizada).

### 5. Booleanos como Text(2)

Las columnas `CursoEnVigor`, `Requiere_Curso`, `RequiereComunicacionConcesion` son `VARCHAR(2)` con valores 'Sí'/'No' (inconsistencia detectada también en NoConformidades). Recomendación:
- **Estandarizar a `BOOLEAN` en PostgreSQL**.
- **Regla de migración**: `'Sí' → TRUE`, `'No' → FALSE`, `NULL → NULL`, otros → `FALSE DEFAULT`.
- Aplicar la misma regla en Lanzadera, Expedientes, Gestion_Riesgos, NoConformidades (cross-cutting).

### 6. FKs conceptuales sin FK física

`IDExpediente`, `IDEmpresaUsuario`, `IDEmpresaHPS`, `IDJuridicaContrato`, `IDSolicitud` son FKs conceptuales (por convención de código, no por constraint). Recomendación:
- **Mantener como referencia conceptual** (D86 / D87) en primera iteración.
- **Evaluar formalización** en iteración posterior si la trazabilidad lo requiere.

### 7. Histórico con anexos

`Mod_Sincronizacion_Historico.bas` + `Test_HistoricoAdjuntosTransactionWrapper.bas` indican sincronización atómica entre histórico de usuarios y anexos. Preservar el patrón transaccional en la nueva plataforma.