# NoConformidades — seguridad y reglas

## Autorización

La matriz común se mantiene en [06-autorizacion-legacy-matriz.md](../../06-autorizacion-legacy-matriz.md). Aquí solo se conserva el comportamiento específico de NoConformidades:

- **Tres roles** (Administrador, Calidad, Técnico), calculados en `EVE` por lookup en `m_ObjEntorno.ColUsuariosAdministradores` y `UsuarioConectadoEsDeCalidad`. El resto cae en Técnico.
- `EsAdministrador` habilita la Ribbon (`DoCmd.ShowToolbar "Ribbon", acToolbarYes`); en cualquier otro caso se oculta.
- `Form_Form0BDOpciones` distingue tres menús: `Parte de Proyecto` (NC de Proyecto), `Auditorías` (NC de Auditoría) y `Técnicos` (vistas técnicas).

## Permisos por aplicación

Los permisos efectivos se cargan desde `UsuarioAplicacionPermisos` (clase compartida con Lanzadera). El acceso a NoConformidades está condicionado por `IDAplicacion = "8"` (producción) o `"81"` (pruebas). En la nueva plataforma esto se reemplaza por el **catálogo de capabilities** (D45-D46) declarado por el módulo.

## Reglas de negocio críticas

- **Alta de NC condicionada**: la lógica de validación está en `NCProyectoOperaciones.Alta` y `NCAuditoriaOperaciones.Alta`; ver detalles en código (no inspeccionado en detalle en esta pasada).
- **AC/AR condicionados**: cada acción tiene responsable, fecha y motivo; las acciones no se pueden borrar sin trazabilidad.
- **Control de eficacia con veredicto**: el control de eficacia genera un veredicto (público/privado) con responsable y fecha; se requiere motivo explícito cuando no se requiere control.
- **Replanificación con motivo**: toda replanificación requiere motivo, autor y fecha; se conserva histórico.
- **Estado del riesgo calculado** (en vinculaciones): cuando una NC se vincula a un riesgo, se preserva el código de riesgo como FK conceptual; no se infiere el uno del otro.
- **Transaccionalidad**: alta/edición combina cabecera, ACs, ARs, replanificaciones, documentos y notas en operaciones DAO; la nueva plataforma usa transacciones SQLAlchemy (`AsyncSession.begin()`) coherentes con D66 y D82.

## Roles y funciones diferenciadas

Como Gestion_Riesgos, NoConformidades separa explícitamente **Calidad** de **Técnico**:

- **Calidad**: acceso a `Form_Form0BDOpcionesAuditorias`, `Form_FormAuditoriasGestion`, `Form_FormNCAuditoria*`. Tareas de auditoría, validación de cierres, control de eficacia.
- **Técnico**: acceso a `Form_Form0BDOpcionesParteProyectos`, `Form_FormNCProyecto*`. Tareas de proyecto, alta/edición de NC operativas.

## Seguridad específica por NC

- `SegNCAuditoria.cls` y `SegNCProyecto.cls` son clases dedicadas a seguridad por NC individual (visibilidad, edición,谁能 ver/editar).
- La nueva plataforma debe traducir esto a **políticas contextuales** (D45) evaluadas en el código del módulo, no en el catálogo.

## D89 · Diagnóstico del fallo de `list_objects` de Dysflow — INVALIDADO

**Estado**: PROPUESTO. **INVALIDADO post-verificación con Dysflow live (2026-08-05).**

### Lo que realmente pasó

El inventario Dysflow sobre NoConformidades **funciona correctamente**. No hay fallo. El "fallo de inventario" era una diagnosis errónea basada en análisis estático de `.dysflow/project.json` y `backends.json`, sin haber contactado el runtime.

### Secuencia correcta para invocar Dysflow sobre NoConformidades

1. **Cargar `dysflow-usage` skill** (canonical tool names, write-flag matrix, error codes).
2. **Cargar `dysflow-arnes` skill** (HR-0 a HR-13, hard rules).
3. **`get_capabilities({})`** — fuente de verdad del runtime (versión, tools, write-gates).
4. **`register_worktree({cwd})`** o `resolve_project({cwd, projectId, projectChoiceReason, recoveryToken})` para registrar/seleccionar el worktree.
5. **`migrate_project_config({cwd, apply:true})`** si el config tiene `allowWrites` top-level (T18 caps-block migration).
6. **Llamadas de inventario con `accessPath` absoluto explícito**:
   ```js
   dysflow.list_objects({cwd, accessPath: "<abs/path/frontend.accdb>", outputMode: "summary"});
   dysflow.list_tables({cwd, accessPath: "<abs/path/frontend.accdb>", outputMode: "full"});
   dysflow.get_schema({cwd, accessPath, table: "<tableName>"});
   dysflow.get_relationships({cwd, accessPath});
   dysflow.count_rows({cwd, accessPath, table: "<tableName>"});
   ```
7. Para herramientas sobre el backend, también pasar `backendPath: "<abs/path/backend.accdb>"`.

### Resultado real del inventario (post-D89 invalidado)

- **42 tablas** en `NoConformidades_Datos.accdb`.
- **438 filas en `TbNoConformidades`** (NC de Proyecto).
- **55 filas en `TbNoConformidadesAuditoria`** (NC de Auditoría).
- **14 FK relationships** entre tablas de usuario.
- **44 columnas en `TbNoConformidades`** (ver [data-model.md § Esquema real](data-model.md#esquema-real-tbnoConformidades)).

### Lecciones aprendidas (capturadas en Engram)

- Cargar `dysflow-usage` ANTES de cualquier llamada o diagnóstico sobre Dysflow.
- Cargar `dysflow-arnes` para HR-0..HR-13.
- `get_capabilities({})` es siempre la primera llamada.
- Si `resolve_project` devuelve `outcome: "ambiguous"`, pedir al humano que elija uno de los `availableProjects` (HR-11) y reintentar con el trio.
- El error `CONFIG_MISSING_ACCESS_PATH` se resuelve pasando `accessPath` absoluto explícito.

### Estado del hallazgo en el blueprint

D89 queda registrado como **invalidado**. La diagnosis original fue retirada de [Matriz de migración § D89](migration-matrix.md#d89-diagnóstico-del-fallo-de-list_objects-de-dysflow) y reemplazada por los datos reales del inventario.

## D90 · Riesgo de seguridad: `backends.json` con contraseña en claro

**Estado**: PROPUESTO (riesgo grave). El archivo `00_main/backends.json` contiene una entrada `ACCESS_VBA_PASSWORD` con valor en claro. Este archivo está versionado en git.

### Riesgos concretos

- **Exposición de credenciales en el historial de git**: cualquier push al remoto expone la contraseña. Si el repo es público o semi-público, el daño es inmediato.
- **Riesgo de rotación nunca ejecutada**: si la contraseña real del VBA cambia, `backends.json` queda desfasado.
- **Uso indebido por terceros**: cualquier persona con acceso al repo puede abrir el binario y leer datos de NC.

### Pasos operativos para remediar

1. **Rotar la contraseña VBA del binario** antes de cualquier remediación.
2. **Eliminar `backends.json` del repo** (`git rm` + `.gitignore` + commit de limpieza).
3. **Reemplazar por `backends.example.json`** sin valor, con comentarios que documenten las claves esperadas (`ACCESS_VBA_PASSWORD`).
4. **Forzar el uso de env vars**: el código VBA ya lee `Environ$("ACCESS_VBA_PASSWORD")` (`Variables Globales.bas:250`); basta con que el runner de Dysflow y los entornos exporten la variable.
5. **Documentar D90 en el blueprint** como hallazgo transversal: cualquier credencial versionada en `backends.json` o equivalente es un riesgo a corregir lote por lote.

### Estado del hallazgo en el blueprint

Registrado en [Matriz de migración § D90](migration-matrix.md#d90-riesgo-de-seguridad--backendsjson-con-contraseña-en-claro) como decisión propuesta. La remediación es **urgente** antes de continuar con cualquier lote que involucre Dysflow sobre NoConformidades.

## D91 · Caché selectivo maduro como referencia

La capa de caché de NoConformidades es la **referencia principal del blueprint para el puerto de caché** de la nueva plataforma (D70-D71). Ver detalle en [Matriz de migración § D91](migration-matrix.md#d91-caché-selectivo-maduro-como-referencia-del-puerto-de-caché).

## Riesgos de privacidad/migración

- No se han incluido nombres, correos, contraseñas, cadenas de conexión, hashes, hosts ni filas personales.
- **El valor de la contraseña en `backends.json` NO se reproduce en este artefacto** (es un riesgo de seguridad, ver D90).
- Las URLs de anexos a SharePoint y las rutas UNC locales se mantienen en backend con metadatos; **no se exponen en este artefacto**.
- Las plantillas de correo (`Correo.cls`) pueden contener datos sensibles; revisar antes de portar.
- `TbLogNCAuditoria`, `TbLogNCProyecto`, `TbSegTareas*`, `TbLogCache` son eventos con potencial valor probatorio; **la retención debe ser al menos la misma que la de auditoría** (D29).