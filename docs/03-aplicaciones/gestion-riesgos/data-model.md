# Gestion_Riesgos — modelo físico y diccionario de datos

## Autoridad y fecha

Fuente física: `C:\00repos\datos\Gestion_Riesgos_Datos.accdb`, pendiente de inspección Dysflow read-only en una segunda pasada. La primera pasada inspeccionó el código fuente exportado en `00_main\src\` vía CodeGraph-VBA; el **inventario completo de tablas y columnas se obtendrá al ejercitar Dysflow** sobre el binario y el backend autoritativo. Este documento cosechó **47 clases del dominio** en `src/classes/` que dan una imagen estructural sólida pero incompleta.

## Diccionario completo

El diccionario fuente completo, con **cada tabla y cada campo**, se obtendrá al ejecutar Dysflow `list_tables`, `get_schema`, `get_relationships` y `count_rows` sobre `Gestion_Riesgos_Datos.accdb`. Esta página añade la **estructura por clases** y los **perfiles agregados** declarados en código; no se han copiado valores de filas.

### Entidades de dominio principales (extraídas del código)

| Clase | Tabla backend probable | Rol |
|---|---|---|
| `Riesgo.cls` | `TbRiesgos` (presumido; pendiente confirmar) | Unidad principal del dominio |
| `RiesgoBiblioteca.cls` | `TbRiesgosBiblioteca` (presumido) | Causa raíz reutilizable |
| `RiesgoExterno.cls` | `TbRiesgosExternos` (presumido) | Riesgo importado de otro sistema |
| `RiesgoMaterializacion.cls` | `TbRiesgosMaterializaciones` (presumido) | Materialización /撤回 |
| `Edicion.cls` | `TbEdiciones` (presumido) | Unidad de versionado |
| `Proyecto.cls` | `TbProyectos` (presumido) | Padre de ediciones |
| `ProyectoSuministrador.cls` | `TbProyectosSuministradores` (presumido) | Suministradores por proyecto |
| `EdicionSuministrador.cls` | `TbProyectosEdicionesSuministradores` (referenciado en código) | Suministradores por edición |
| `PM.cls` / `PC.cls` | `TbPlanesMitigacion` / `TbPlanesContingencia` (presumidos) | Planes de mitigación / contingencia |
| `PMAccion.cls` / `PCAccion.cls` | `TbPlanesMitigacionAcciones` / `TbPlanesContingenciaAcciones` (presumidos) | Acciones del plan |
| `PMAccionReversa.cls` / `PCAccionReversa.cls` | `TbPlanesMitigacionAccionesReversa` / `TbPlanesContingenciaAccionesReversa` (presumidos) | Reversa /撤销 |
| `Anexo.cls` / `AnexoAntiguo.cls` | `TbAnexos` / `TbAnexosAntiguos` (presumidos) | Evidencias |
| `CCCambio.cls` / `CCDocumentoCambio.cls` / `CCVersion.cls` | `TbControlCambios*` (presumidos) | Control de cambios |
| `Cambio.cls` / `CambioExplicacion.cls` / `CarenciasExplicacion.cls` | `TbCambios*` / `TbCarenciasExplicacion` (presumidos) | Cambios y explicaciones |
| `Documento.cls` | `TbDocumentos` (presumido) | Documentos vinculados |
| `Correo.cls` / `EdicionCorreoRevision.cls` | `TbCorreos*` (presumidos) | Correos |
| `Pedido.cls` | `TbPedidos` (presumido) | Pedidos |
| `AreaImpacto.cls` / `Juridica.cls` / `OrganoContratacion.cls` / `RAC.cls` | catálogos correspondientes | Catálogos de clasificación |
| `MitigacionValor.cls` | `TbMitigacionValores` (presumido) | Valoración |
| `NC.cls` | `TbNoConformidades` (presumido) | Vínculo con No Conformidades |
| `TareasCalidad.cls` / `ArbolTareasCalidad.cls` | `TbTareasCalidad` / estructura árbol (presumidos) | Worklists de calidad |
| `TareasTecnico.cls` / `ArbolTareasTecnico.cls` | `TbTareasTecnico` / estructura árbol (presumidos) | Worklists de técnico |
| `PublicacionLog.cls` | `TbPublicacionLog` (presumido) | Log de publicabilidad |
| `UltimoProyecto.cls` | `TbUltimoProyecto` (presumido) | Último proyecto por usuario |
| `Usuario.cls` / `UsuarioAplicacionPermisos.cls` / `Entorno.cls` | `TbUsuariosAplicaciones*` (Lanzadera, vía `getdbLanzadera`) | Identidad / permisos / entorno |
| `Expediente.cls` / `ExpedienteEntidad.cls` / `ExpedienteResponsable.cls` | tablas de Expedientes (compartidas con Lanzadera) | Vínculo con Expedientes |
| `InformeRiesgoPDFServicio.cls` | sin tabla; genera artefacto | Informe PDF/HTML |
| `HTML.cls` | sin tabla; helpers | Helpers HTML |

### Relaciones físicas confirmadas (en código)

- `EdicionSuministrador.IDEdicion → Constructor.getEdicion(...)` (FK conceptual a `TbEdiciones`).
- `EdicionSuministrador.IDSuministrador → Constructor.getSuministrador(...)` (FK a `TbSuministradores`).
- `EdicionSuministrador.IDAnexo → Constructor.getAnexo(...)` (FK a `TbAnexos`).
- `EdicionSuministrador.ProyectoSuministrador → Constructor.getSuministradorEnProyecto(Me.Edicion.IDProyecto, Me.IDSuministrador)` (join a través de proyecto).
- `Riesgo.Edicion → Constructor.getEdicion(Me.IDEdicion)` (FK).
- `Riesgo.RiesgoEdicionAnterior / RiesgoEdicionPrimera / RiesgoEdicionSiguiente` (FKs entre ediciones del mismo código de riesgo).
- `AnexoAntiguo.Edicion`, `Correo.Edicion`, `EdicionCorreoRevision.Edicion` → todas vía `Constructor.getEdicion`.
- **La relación física entre `TbProyectosEdicionesSuministradores` y `TbProyectos` no aparece como FK explícita en el código Access**; se modela como join a través de `IDEdicion → IDProyecto`. Esto es una oportunidad de mejora en el modelo de PostgreSQL.

### Semántica Access que debe conservarse

- Tipos 1/3/4/7/8/10/12 observados: Boolean, Integer/Long, Currency, DateTime, Text y Memo según columna; confirmar mapeo final por campo en la segunda pasada Dysflow.
- `Sí/No` se almacena frecuentemente como texto de longitud 2, no como Boolean; **no convertirlo silenciosamente**.
- `Null` y cadena vacía se distinguen en formularios, DTO y `Registrar`.
- IDs calculados con `DameID("TbXxx", "IDXxx", getdb(), m_Error)` antes de `.AddNew` — la nueva plataforma debe reproducir el patrón (idempotente, atómico).
- `IDAplicacion = "5"` (producción) / `"51"` (pruebas) se asigna por la TempVar `EnPruebas`; en la nueva plataforma viene de la configuración (D9).

### Perfil agregado (privacidad segura)

El perfilado de filas por tabla queda **pendiente para la segunda pasada Dysflow** sobre `Gestion_Riesgos_Datos.accdb`. En esta primera pasada se cosecharon **47 clases** y se confirmó que:

- `getdb()` tiene **308 callers** (vs 135 en Expedientes) — Gestion_Riesgos es mucho más intensivo en DAO.
- `getdbLanzadera()` se invoca desde `Constructor.bas` (al menos 1 caller en `getUsuario`) — la aplicación lee directamente la BD de Lanzadera.
- `DameID()` aparece en múltiples clases para generación de IDs antes de `.AddNew`.

Una vez completada la segunda pasada, este apartado se completará con `count_rows` y comprobaciones de unicidad, nulos, huérfanos y rangos de fechas.

## Rendimiento del árbol de riesgos — causa raíz y opciones de implementación

### Síntoma

Cuando un usuario carga la pantalla de gestión de riesgos con **muchas ediciones** y **muchos riesgos por edición**, el árbol jerárquico tarda "un mundo" en aparecer. La queja es transversal a usuarios de calidad y técnico.

### Causa raíz (evidenciada en código)

1. **Control UI legacy ActiveX**: `Form_FormRiesgosGestion.m_Arbol` es `MSComctlLib.TreeView` (línea 117 de `Form_FormRiesgosGestionRiesgo.cls`). El control se llena en `CargarArbol`, que itera sobre las ediciones y, por cada edición, sobre los riesgos. Cada nodo del TreeView requiere su propio `Add` (comunicación con el OCX), lo que es lento cuando hay miles de nodos.

2. **Doble modelo de cadena jerárquica activo en producción**: la TempVar `CadenaJerarquicaModelo` admite valores `"nuevo"` y `"antiguo"` (líneas 248-249 de `Variables Globales.bas`). Ambos modelos coexisten en el código. Cada uno tiene su propia ruta de carga, con diferentes cuellos de botella; el síntoma aparece en ambos.

3. **Recargas del árbol disparadas desde cambios pequeños**: `EstablecerContadoresCalidad → CargarArbol` (visto en blast radius) y `ComandoEliminar_Click → CargarArbol` (línea 123 de `Form_FormRiesgosGestionRiesgo.cls`). Cada edición del usuario puede requerir recarga del árbol completo.

4. **DAO intensivo**: 308 callers de `getdb()` significa que las operaciones CRUD básicas (alta/borrado/edición de riesgo) hacen múltiples viajes a Access; en Access + tablas vinculadas cada viaje es lento. En escenarios de batch (creación de muchos riesgos) el usuario percibe la suma de latencias.

5. **Carga ansiosa**: el árbol se carga completo aunque el usuario solo abra una edición. No hay paginación ni lazy expansion.

### Opciones para la nueva plataforma

| Opción | Stack | Latencia esperada | Complejidad | Compatibilidad legacy |
|---|---|---|---|---|
| **A) HTMX + CTE recursivo en PostgreSQL + lazy expansion por nivel** | FastAPI + Jinja2 + HTMX (`hx-get` por nivel) | Centenas de ms por nivel (consulta indexada) | Media | Alta — `CargarArbol` se reemplaza por endpoint `GET /ediciones/{id}/arbol?nivel=...` |
| **B) Web Component cliente + endpoint con snapshot JSON** | FastAPI + JSON + componente cliente (Alpine.js o Web Component ligero) | Centenas de ms para snapshot completo; lazy nativo en cliente | Media | Media — el formato JSON debe incluir la jerarquía completa |
| **C) Server-Sent Events para push de cambios del árbol** | FastAPI + SSE + HTMX (`hx-ext="sse"`) | Latencia mínima en updates | Alta | Baja — introduce SSE, rechazado en sección 1 del prompt externo (D69) |
| **D) Replicar MSComctlLib.TreeView en frontend con librería JS** | FastAPI + librería JS tipo `react-arborist` o similar | Depende de la librería; suele ser buena | Alta | Baja — atado a SPA, rompe el principio SSR |

**Recomendación provisional**: **Opción A** (HTMX + CTE recursivo + lazy expansion). Es coherente con el stack aprobado (D66-D69), preserva SSR, elimina el OCX legacy, y la complejidad es manejable para 2 desarrolladores.

### Esquema de la propuesta A

```sql
-- Tabla de ediciones (extracto; el modelo completo en migration-matrix.md)
CREATE TABLE ediciones (
    id_edicion BIGINT PRIMARY KEY,
    id_proyecto BIGINT NOT NULL REFERENCES proyectos(id_proyecto),
    codigo TEXT NOT NULL,
    es_activo BOOLEAN NOT NULL DEFAULT TRUE,
    fecha_publicacion DATE,
    ...
);

-- Tabla de riesgos
CREATE TABLE riesgos (
    id_riesgo BIGINT PRIMARY KEY,
    id_edicion BIGINT NOT NULL REFERENCES ediciones(id_edicion),
    codigo_riesgo TEXT NOT NULL,
    descripcion TEXT,
    estado TEXT,
    fecha_detectado DATE,
    fecha_retirado DATE,
    fecha_materializado DATE,
    ...
);

-- Índice crítico para la carga del árbol
CREATE INDEX idx_riesgos_id_edicion ON riesgos(id_edicion);
CREATE INDEX idx_riesgos_codigo ON riesgos(codigo_riesgo);

-- CTE recursivo: árbol de un proyecto completo
WITH RECURSIVE arbol_ediciones AS (
    SELECT id_edicion, id_proyecto, codigo, es_activo, fecha_publicacion, 1 AS nivel
    FROM ediciones
    WHERE id_proyecto = $1 AND id_edicion_padre IS NULL
    UNION ALL
    SELECT e.id_edicion, e.id_proyecto, e.codigo, e.es_activo, e.fecha_publicacion, ae.nivel + 1
    FROM ediciones e
    INNER JOIN arbol_ediciones ae ON e.id_edicion_padre = ae.id_edicion
)
SELECT a.*, r.id_riesgo, r.codigo_riesgo, r.estado
FROM arbol_ediciones a
LEFT JOIN riesgos r ON r.id_edicion = a.id_edicion
ORDER BY a.nivel, a.id_edicion, r.id_riesgo;
```

Endpoint FastAPI correspondiente:

```python
@router.get("/proyectos/{id_proyecto}/arbol")
async def get_arbol(
    id_proyecto: int,
    nivel: int = 1,
    expandir: bool = False,
    session: AsyncSession = Depends(get_db),
) -> TemplateResponse:
    """Devuelve fragmento HTML con el nivel N del árbol de ediciones/riesgos.
    Si expandir=True, incluye los hijos; si no, solo el nivel pedido.
    """
    # ... implementar con CTE recursivo paginado por nivel
```

El fragmento HTML se intercambia con `hx-get` cuando el usuario expande un nodo. El árbol completo deja de existir en memoria: cada nivel es una consulta indexada.

### Implicaciones

- **D88** (registrada en [migration-matrix.md](migration-matrix.md#d88-modelo-de-árbol-de-riesgos-en-la-nueva-plataforma)): se descarta el modelo de cadena jerárquica `nuevo/antiguo` y el control `MSComctlLib.TreeView`; se adopta el modelo HTMX + CTE recursivo + lazy expansion por nivel. Esta decisión queda **propuesta** y debe revisarse antes de la fase SDD.
- `CadenaJerarquicaModelo`, `JPMesesAvisoEntreEdiciones`, `JPDiasPreviosParaElAviso`, `CalDiaInicialMesAviso` se conservan como configuración del módulo (con sus defaults); algunos se reinterpretan (ej. los avisos entre ediciones podrían implementarse como una tarea del scheduler en lugar de un flag temporal).
- La validación de carga (latencia esperada < 1 segundo por nivel con índices) se hará en la fase de diseño con datos sintéticos representativos antes de promover a implementación.