# Gestion_Riesgos — matriz de migración (scaffold de preservación)

Esta matriz no decide el esquema PostgreSQL. Su regla es conservadora: todo campo/registro se conserva hasta que negocio apruebe una disposición. La queja del árbol de ediciones/riesgos se aborda en [D88](#d88-modelo-de-árbol-de-riesgos-en-la-nueva-plataforma) con propuesta concreta.

| Fuente | Significado/candidato de dominio | Transformación | Estado | Validación/rechazo | Decisión abierta |
|---|---|---|---|---|---|
| `TbRiesgos` (presumido) | riesgo por edición | separar identidad, descripción, fechas, estado, valoración,Flags causa-raíz; conservar valores originales | mapeado preliminar | recuento 1:1, IDs y códigos reconciliados | catálogo de estados final; fuente de verdad entre flags y Tipo |
| `IDRiesgoPadre` (si existe) | jerarquía entre riesgos | FK autorreferente preservando ID legacy | mapeado preliminar | cero huérfanos; detectar ciclos | política de ciclos |
| `CodigoRiesgo` | identificador funcional | conservar columnas originales y normalizadas | necesita decisión | unicidad y vacíos; no inventar sustituto | claves canónicas |
| Flags `EsAM/EsLote/EsBasado/EsExpediente` y `Tipo` | clasificación/tipo | tabla de valores + valor legacy | necesita decisión | coherencia tipo/jerarquía | fuente de verdad entre flags y Tipo |
| `TbRiesgosBiblioteca` | causa raíz reutilizable | entidad de catálogo con versionado | mapeado preliminar | unicidad de código; trazabilidad | versionado de biblioteca |
| `TbRiesgosExternos` | riesgos importados | clave externa + snapshot | external reference | reconciliación por código de riesgo | ownership de la importación |
| `TbRiesgosMaterializaciones` | evento de materialización | evento con timestamps, responsable,撤回 | mapeado preliminar | secuencialidad; vínculo con PM/PC | retención/legal |
| `TbEdiciones` | unidad de versionado por proyecto | identidad + fechas + flag activo + FK proyecto | mapeado preliminar | árbol de ediciones coherente | jerarquía entre ediciones |
| `TbProyectos` | proyecto | identidad + responsables (autorizado/técnico/calidad) | mapeado preliminar | unicidad; coherencia con Lanzadera | ownership |
| `TbProyectosEdicionesSuministradores` | suministradores por edición | FK a ediciones, suministradores, anexo (opcional) | mapeado preliminar | cardinalidad; cero huérfanos | transiciones de estado del anexo |
| `TbPlanesMitigacion` (presumido) | PM por riesgo | FK a riesgo + responsable + calendario | mapeado preliminar | integridad referencial | versionado del plan |
| `TbPlanesContingencia` (presumido) | PC por riesgo | FK a riesgo + responsable + calendario | mapeado preliminar | integridad referencial | versionado del plan |
| `TbPlanesMitigacionAcciones` / `TbPlanesContingenciaAcciones` (presumidos) | acciones del plan | FK a plan + responsable + fechas | mapeado preliminar | secuencialidad | dependencias entre acciones |
| `TbPlanesMitigacionAccionesReversa` / `TbPlanesContingenciaAccionesReversa` (presumidos) | reversa de acciones | evento con motivo y autor | mapeado preliminar | trazabilidad | retención de la reversa |
| `TbAnexos` | evidencias | metadatos + referencia de almacenamiento | external reference | comprobar existencia de fichero/URL sin copiar datos | repositorio destino (D16) |
| `TbAnexosAntiguos` | evidencias antiguas | snapshot conservado en zona legacy | historical-only candidate | comparar recuentos | archivo vs dominio |
| `TbControlCambios*` (CCCambio, CCDocumentoCambio, CCVersion) | control de versiones | entidad versionada + relación con documentos | mapeado preliminar | secuencialidad; FKs | política de versionado |
| `TbCambios` / `TbCambiosExplicacion` / `TbCarenciasExplicacion` | cambios y explicaciones | evento + explicación | mapeado preliminar | trazabilidad | retención/legal |
| `TbPublicacionLog` | log de publicabilidad | evento de evaluación + veredicto | mapeado preliminar | secuencialidad; vínculo con riesgo | retención |
| `TbTareasCalidad` / `TbTareasTecnico` | worklists | entidad de tarea + estado | mapeado preliminar | cardinalidad con riesgos | catálogo de tareas |
| `TbRiesgosEstadosHistoricos` (presumido) | histórico de estados | evento de cambio + timestamp | mapeado preliminar | secuencialidad | retención |
| `TbMitigacionValores` | valoración | catálogo | mapeado preliminar | unicidad | versionado |
| `TbAreaImpacto`, `TbJuridicas`, `TbOrganosContratacion`, `TbRAC` | catálogos de clasificación | entidades de catálogo | mapeado preliminar | unicidad; coherencia con Lanzadera | versionado |
| `TbPedidos` | pedidos vinculados | FK + metadatos | external reference | reconciliación con HPS | ownership |
| `TbNoConformidades` (lectura) | vínculo con NC | FK + snapshot | external reference | reconciliación por código | ownership |
| `TbUltimoProyecto` | último proyecto por usuario | configuración por usuario | mapeado preliminar | unicidad por usuario | limpieza al logout |
| catálogos compartidos con Lanzadera (vía `getdbLanzadera`) | identidad / permisos / entorno | identidad vía adaptador unificado (D9) | mapeado preliminar | coherencia | acoplamiento directo a romper (ver D86/D87) |
| tablas E2E/temp/aux (si las hay) | soporte técnico | migrar solo si se requiere trazabilidad; conservar dump si no | technical-only candidate | documentar antes de excluir | decisión explícita, no obsolescencia |
| tablas copia/históricas/auxiliares | histórico o soporte | conservar en zona legacy hasta clasificación | needs business decision | comparar recuentos y uso en código | archivo vs dominio |
| integraciones externas (Lanzadera, Expedientes, HPS, NC, AGEDYS, correos) | referencias externas | claves externas y snapshot contractual | external reference | reconciliar por identificador, sin asumir propiedad | ownership y sincronización |

## Ledger

- **Mapeado preliminar:** cabecera, ediciones, proyectos, riesgos, planes, anexos, publicabilidad, worklists, control de cambios, catálogos.
- **Necesita decisión:** flags frente a tipo, biblioteca, histórico, E2E, anexos antiguos, contratos externos.
- **Candidato técnico:** temporales y helpers de exportación; no retirado.
- **Referencia externa:** anexos, identidad, Lanzadera (vía `getdbLanzadera`), Expedientes, HPS, NC, AGEDYS y correo.
- **No hay campos/filas declarados obsoletos.**

## D88 · Modelo de árbol de riesgos en la nueva plataforma

**Decisión propuesta**: descartar el modelo de cadena jerárquica `nuevo/antiguo` y el control ActiveX `MSComctlLib.TreeView`; adoptar el modelo **HTMX + CTE recursivo en PostgreSQL + lazy expansion por nivel** (FastAPI + Jinja2 SSR). Justificación detallada en [data-model.md § Rendimiento del árbol](data-model.md#rendimiento-del-árbol-de-riesgos--causa-raíz-y-opciones-de-implementación).

| Aspecto | Decisión |
|---|---|
| Componente UI | Fragmento HTML renderizado por Jinja2; expansión con `hx-get` |
| Fuente de datos | CTE recursivo en PostgreSQL, índice en `(id_proyecto, id_edicion_padre)` y `(id_edicion)` |
| Carga | Ansiosa para el primer nivel del proyecto; lazy por nivel al expandir |
| Persistencia del estado | URL params (`?nivel=...&expandido=...`) o fragmento hidden; sin estado cliente |
| Compatibilidad legacy | `CargarArbol` se reemplaza por endpoint `GET /proyectos/{id}/arbol?nivel=N` |
| Validación previa a implementación | Medir latencia con datos sintéticos representativos (< 1 s por nivel esperado) |
| Estado | **PROPUESTO** — pendiente revisión con usuario antes de fase SDD |

Cada fila futura debe expandirse a `source table.field` para los 49 esquemas (cuando estén cosechados en la segunda pasada Dysflow), con transformación, regla de reconciliación, rechazo y disposición aprobada.