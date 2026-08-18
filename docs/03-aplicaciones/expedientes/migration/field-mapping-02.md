# Field Mapping Tranche 02 — Tablas 16–30

Mapeo campo-a-campo de las tablas 16-30 del backend legacy `Expedientes_datos.accdb` hacia las capacidades del spec `expedientes-web-migration`.

**Fuente legacy**: `docs/03-aplicaciones/expedientes/ERD/schema.sql` (extraído con Jackcess)
**Spec destino**: `openspec/changes/expedientes-web-migration/specs/`
**Design**: `openspec/changes/expedientes-web-migration/design.md` (D-EXP-1..9)

## Resumen de tablas

| # | Tabla legacy | Tipo | Capacidad destino | Estado |
|---|---|---|---|---|
| 16 | TbEjercitos | Catálogo | EXP-CAP-017 (Ejércitos) | Migrar |
| 17 | TbEstados | Catálogo | EXP-CAP-005 (Estado) | Migrar |
| 18 | TbExpAgedys | AGEDYS | EXP-CAP-051 (AGEDYS) | Cuarentena |
| 19 | TbExpAGEDYS1 | AGEDYS | EXP-CAP-051 (AGEDYS) | Cuarentena |
| 20 | TbExpedientes | **Tabla principal** | EXP-CAP-001..007 (Ciclo) | Migrar |
| 21 | TbExpedientes_antes | Histórico | No migrar (histórico legacy) | Rechazar |
| 22 | TbExpedientesAnexos | Anexos | EXP-CAP-010 (Anexos) | Migrar |
| 23 | TbExpedientesAnualidades | Anualidades | EXP-CAP-011 (Anualidades) | Migrar |
| 24 | TbExpedientesCadenaContratacion | Cadena | EXP-CAP-009 (Modificados) | Migrar |
| 25 | TbExpedientesCodigoCompras | Códigos | EXP-CAP-009 (Modificados) | Migrar |
| 26 | TbExpedientesComerciales | Relación | EXP-CAP-015 (Comerciales) | Migrar |
| 27 | TbExpedientesConEntidades | Relación | EXP-CAP-013 (Entidades) | Migrar |
| 28 | TbExpedientesCPVs | Relación | EXP-CAP-016 (CPV) | Migrar |
| 29 | TbExpedientesE2E | E2E | EXP-CAP-033..042 (E2E) | Migrar |
| 30 | TbExpedientesHitos | Hitos | EXP-CAP-008 (Hitos) | Migrar |

## Detalle por tabla

### 16. TbEjercitos → MIGRAR

**Razón**: Catálogo de ejércitos. Relacionada con EXP-CAP-017 (Ejércitos). Tabla maestra simple.

**Campos clave**:
- `IDEjercito` (LONG) → `ejercitos.id`
- `Ejercito` (TEXT) → `ejercitos.nombre`
- `Descripcion` (MEMO) → `ejercitos.descripcion`

**Acción**: Migrar a `ejercitos` del nuevo schema. Sin transformaciones.

### 17. TbEstados → MIGRAR

**Razón**: Catálogo de estados del expediente. Relacionada con EXP-CAP-005 (Estado y garantía). Tabla maestra que define el ciclo de vida.

**Campos clave**:
- `IDEstado` (LONG) → `estados.id`
- `Estado` (TEXT) → `estados.nombre`
- `DESCRIPCION` (MEMO) → `estados.descripcion`

**Acción**: Migrar a `estados` del nuevo schema. Sin transformaciones.

### 18. TbExpAgedys → CUARENTENA

**Razón**: Tabla de integración AGEDYS. Relacionada con EXP-CAP-051 (AGEDYS). Requiere análisis adicional para determinar el contenido y la relevancia.

**Campos**: TBD (verificar en schema.sql)
**Acción**: Cuarentena. Analizar contenido antes de decidir migración.

### 19. TbExpAGEDYS1 → CUARENTENA

**Razón**: Tabla de integración AGEDYS (variante). Relacionada con EXP-CAP-051 (AGEDYS). Mismo criterio que #18.

**Campos**: TBD (verificar en schema.sql)
**Acción**: Cuarentena. Analizar contenido antes de decidir migración.

### 20. TbExpedientes → MIGRAR (TABLA PRINCIPAL)

**Razón**: Tabla principal del expediente. Relacionada con EXP-CAP-001..007 (Ciclo de vida). Contiene todos los datos del expediente.

**Campos clave**:
- `IDExpediente` (LONG) → `expedientes.id`
- `IDExpedientePadre` (LONG) → `expedientes.parent_id` (jerarquía EXP-CAP-006)
- `Nemotecnico` (TEXT) → `expedientes.nemotecnico` (EXP-CAP-007)
- `Titulo` (MEMO) → `expedientes.titulo`
- `ImporteLicitacion` (DOUBLE) → `expedientes.importe_licitacion`
- `ImporteContratacion` (DOUBLE) → `expedientes.importe_contratacion`
- `CodProyecto` (TEXT) → `expedientes.cod_proyecto`
- `CodExp` (TEXT) → `expedientes.cod_exp`
- `CodExpLargo` (TEXT) → `expedientes.cod_exp_largo`
- `CodS4H` (TEXT) → `expedientes.cod_s4h`
- `FechaInicioContrato` (SHORT_DATE_TIME) → `expedientes.fecha_inicio_contrato`
- `FechaFinContrato` (SHORT_DATE_TIME) → `expedientes.fecha_fin_contrato`
- `FechaFinGarantia` (SHORT_DATE_TIME) → `expedientes.fecha_fin_garantia` (EXP-CAP-005)
- `EsAM` (TEXT) → `expedientes.es_am` (BOOLEAN, EXP-CAP-006)
- `EsLote` (TEXT) → `expedientes.es_lote` (BOOLEAN, EXP-CAP-006)
- `EsBasado` (TEXT) → `expedientes.es_basado` (BOOLEAN, EXP-CAP-006)
- `EsExpediente` (TEXT) → `expedientes.es_expediente` (BOOLEAN)
- `Ordinal` (TEXT) → `expedientes.ordinal` (EXP-CAP-007)
- `IdGradoClasificacion` (LONG) → FK a `grados_clasificacion.id`
- `IDOrganoContratacion` (LONG) → FK a `organos_contratacion.id`
- `IDOficinaPrograma` (LONG) → FK a `oficinas_programa.id`
- `IDEjercito` (LONG) → FK a `ejercitos.id`
- `AccesoSharepoint` (MEMO) → `expedientes.acceso_sharepoint`
- `Observaciones` (MEMO) → `expedientes.observaciones`
- `FechaCreacion` (SHORT_DATE_TIME) → `expedientes.created_at`
- `IDUsuarioCreacion` (TEXT) → `expedientes.created_by` (transformar a FK numérica)
- `FechaUltimoCambio` (SHORT_DATE_TIME) → `expedientes.updated_at`
- `IDUsuarioUltimoCambio` (TEXT) → `expedientes.updated_by` (transformar a FK numérica)
- `IDEstado` (LONG) → FK a `estados.id` (EXP-CAP-005)
- `Ambito` (TEXT) → `expedientes.ambito`
- `NPedido` (TEXT) → `expedientes.n_pedido`
- `IDResponsableCalidad` (LONG) → FK a `responsables.id`
- `Adjudicado` (TEXT) → `expedientes.adjudicado` (BOOLEAN)
- `EnPeriodoDeAdjudicacion` (TEXT) → `expedientes.en_periodo_adjudicacion` (BOOLEAN)
- `Tipo` (TEXT) → `expedientes.tipo`
- `TipoInforme` (TEXT) → `expedientes.tipo_informe`
- `AGEDYSAplica` (TEXT) → `expedientes.agedys_aplica` (BOOLEAN)
- `AGEDYSGenerico` (TEXT) → `expedientes.agedys_generico` (BOOLEAN)
- `HPSAplica` (TEXT) → `expedientes.hps_aplica` (BOOLEAN)
- `CadenaPecal` (TEXT) → `expedientes.cadena_pecal`
- `Pecal` (TEXT) → `expedientes.pecal` (BOOLEAN)
- `POSTAGEDO` (TEXT) → `expedientes.post_agedo` (BOOLEAN)
- `APLICAESTADO` (TEXT) → `expedientes.aplica_estado` (BOOLEAN)
- `FECHAINICIOLICITACION` (SHORT_DATE_TIME) → `expedientes.fecha_inicio_licitacion`
- `FECHAOFERTA` (SHORT_DATE_TIME) → `expedientes.fecha_oferta`
- `FECHAADJUDICACION` (SHORT_DATE_TIME) → `expedientes.fecha_adjudicacion`
- `FECHAFIRMACONTRATO` (SHORT_DATE_TIME) → `expedientes.fecha_firma_contrato`
- `GARANTIAMESES` (TEXT) → `expedientes.garantia_meses` (transformar a INTEGER)
- `FECHACERTIFICACION` (SHORT_DATE_TIME) → `expedientes.fecha_certificacion`
- `FECHAPERDIDA` (SHORT_DATE_TIME) → `expedientes.fecha_perdida`
- `FECHADESESTIMADA` (SHORT_DATE_TIME) → `expedientes.fecha_desestimada`
- `ESTADO` (TEXT) → `expedientes.estado_texto` (redundante con IDEstado, consolidar)
- `CodigoActividad` (TEXT) → `expedientes.codigo_actividad`
- `HPSAplicaTareaS4H` (TEXT) → `expedientes.hps_aplica_tarea_s4h` (BOOLEAN)

**Acción**: Migrar a `expedientes` del nuevo schema. Transformaciones:
- `IDUsuarioCreacion/IDUsuarioUltimoCambio` TEXT → FK numérica
- Booleanos TEXT(2) → BOOLEAN (D102)
- `GARANTIAMESES` TEXT → INTEGER
- `ESTADO` TEXT consolidar con `IDEstado` FK

### 21. TbExpedientes_antes → RECHAZAR

**Razón**: Tabla histórica (`_antes`). Contiene datos de expedientes anteriores. No tiene capacidad asociada en el spec.

**Campos**: TBD (verificar en schema.sql)
**Acción**: No migrar. Documentar como tabla histórica legacy.

### 22. TbExpedientesAnexos → MIGRAR

**Razón**: Anexos de expedientes. Relacionada con EXP-CAP-010 (Anexos y referencias). Contiene los anexos de cada expediente.

**Campos clave**:
- `IDExpediente` (LONG) → FK a `expedientes.id`
- `IDAnexo` (LONG) → `expediente_anexos.id`
- `NombreAnexo` (TEXT) → `expediente_anexos.nombre`
- `RutaAnexo` (MEMO) → `expediente_anexos.ruta` (transformar a referencia opaca D-EXP-6)
- `FechaAnexo` (SHORT_DATE_TIME) → `expediente_anexos.fecha`

**Acción**: Migrar a `expediente_anexos` del nuevo schema. Transformar `RutaAnexo` a referencia opaca (D-EXP-6: DocumentStoragePort).

### 23. TbExpedientesAnualidades → MIGRAR

**Razón**: Anualidades de expedientes. Relacionada con EXP-CAP-011 (Anualidades). Contiene los importes por año.

**Campos clave**:
- `IDExpediente` (LONG) → FK a `expedientes.id`
- `Ano` (LONG) → `expediente_anualidades.ano`
- `Importe` (DOUBLE) → `expediente_anualidades.importe`
- `TipoImpuesto` (TEXT) → `expediente_anualidades.tipo_impuesto`

**Acción**: Migrar a `expediente_anualidades` del nuevo schema. Sin transformaciones.

### 24. TbExpedientesCadenaContratacion → MIGRAR

**Razón**: Cadena de contratación. Relacionada con EXP-CAP-009 (Modificados e historial). Contiene la cadena de subcontratación.

**Campos clave**:
- `IDExpediente` (LONG) → FK a `expedientes.id`
- `IDCadena` (LONG) → `expediente_cadenas.id`
- `Nivel` (LONG) → `expediente_cadenas.nivel`
- `Descripcion` (MEMO) → `expediente_cadenas.descripcion`

**Acción**: Migrar a `expediente_cadenas` del nuevo schema. Sin transformaciones.

### 25. TbExpedientesCodigoCompras → MIGRAR

**Razón**: Códigos de compras. Relacionada con EXP-CAP-009 (Modificados e historial). Contiene los códigos de compras asociados.

**Campos clave**:
- `IDExpediente` (LONG) → FK a `expedientes.id`
- `CodigoCompra` (TEXT) → `expediente_codigos_compras.codigo`

**Acción**: Migrar a `expediente_codigos_compras` del nuevo schema. Sin transformaciones.

### 26. TbExpedientesComerciales → MIGRAR

**Razón**: Relación expediente-comercial. Relacionada con EXP-CAP-015 (Comerciales). Tabla de relación many-to-many.

**Campos clave**:
- `IDExpediente` (LONG) → FK a `expedientes.id`
- `IDComercial` (LONG) → FK a `comerciales.id`

**Acción**: Migrar a `expediente_comerciales` del nuevo schema. Sin transformaciones.

### 27. TbExpedientesConEntidades → MIGRAR

**Razón**: Relación expediente-entidades (vista desnormalizada). Relacionada con EXP-CAP-013 (Entidades y jurídicas). Contiene cadenas concatenadas para JOIN rápido.

**Campos clave**:
- `IDExpediente` (LONG) → FK a `expedientes.id`
- `Clasificacion` (TEXT) → `expediente_entidades.clasificacion`
- `OrganoContratacion` (TEXT) → `expediente_entidades.organo_contratacion`
- `OficinaPrograma` (TEXT) → `expediente_entidades.oficina_programa`
- (23 columnas totales — vista desnormalizada)

**Acción**: Migrar a `expediente_entidades` del nuevo schema. La vista desnormalizada se normaliza en el nuevo schema (tablas separadas por tipo de entidad).

### 28. TbExpedientesCPVs → MIGRAR

**Razón**: Relación expediente-CPV. Relacionada con EXP-CAP-016 (CPV). Tabla de relación many-to-many.

**Campos clave**:
- `IDExpediente` (LONG) → FK a `expedientes.id`
- `IDCPV` (LONG) → FK a `cpv.id`

**Acción**: Migrar a `expediente_cpvs` del nuevo schema. Sin transformaciones.

### 29. TbExpedientesE2E → MIGRAR

**Razón**: Datos E2E de expedientes. Relacionada con EXP-CAP-033..042 (E2E). Contiene los datos de exportación E2E.

**Campos clave**:
- `IDExpediente` (LONG) → FK a `expedientes.id`
- `HashActual` (TEXT) → `expediente_e2e.hash_actual` (EXP-CAP-036)
- `HashUltimaExportacion` (TEXT) → `expediente_e2e.hash_ultima_exportacion` (EXP-CAP-036)
- `OrdinalE2E` (TEXT) → `expediente_e2e.ordinal` (EXP-CAP-042)

**Acción**: Migrar a `expediente_e2e` del nuevo schema. Sin transformaciones.

### 30. TbExpedientesHitos → MIGRAR

**Razón**: Hitos de expedientes. Relacionada con EXP-CAP-008 (Hitos temporales). Contiene los hitos temporales de cada expediente.

**Campos clave**:
- `IDExpediente` (LONG) → FK a `expedientes.id`
- `IDHito` (LONG) → `expediente_hitos.id`
- `TipoHito` (TEXT) → `expediente_hitos.tipo`
- `FechaHito` (SHORT_DATE_TIME) → `expediente_hitos.fecha`
- `Importe` (DOUBLE) → `expediente_hitos.importe`

**Acción**: Migrar a `expediente_hitos` del nuevo schema. Sin transformaciones.

## Transformaciones globales

| Transformación | Aplica a | Razón |
|---|---|---|
| `IDUsuario*` TEXT → FK numérica | TbExpedientes | El legacy usa TEXT para IDs de usuario; el nuevo schema usa FK numéricas |
| Booleanos TEXT(2) → BOOLEAN | TbExpedientes (EsAM, EsLote, EsBasado, etc.) | D102: booleanos Text(2) → BOOLEAN |
| `GARANTIAMESES` TEXT → INTEGER | TbExpedientes | Tipo numérico en el nuevo schema |
| `ESTADO` TEXT → consolidar con `IDEstado` | TbExpedientes | Redundancia: el estado se define por FK, no por texto |
| `RutaAnexo` → referencia opaca | TbExpedientesAnexos | D-EXP-6: DocumentStoragePort guarda por referencia opaca |
| Vista desnormalizada → normalizar | TbExpedientesConEntidades | El nuevo schema usa tablas separadas por tipo de entidad |

## Cuarentena

| Tabla | Razón | Acción requerida |
|---|---|---|
| TbExpAgedys | Integración AGEDYS, contenido desconocido | Analizar contenido antes de decidir migración |
| TbExpAGEDYS1 | Integración AGEDYS (variante), contenido desconocido | Analizar contenido antes de decidir migración |

## Rechazos

| Tabla | Razón |
|---|---|
| TbExpedientes_antes | Tabla histórica legacy, sin capacidad asociada |
