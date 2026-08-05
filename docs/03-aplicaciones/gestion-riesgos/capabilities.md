# Gestion_Riesgos — capacidades observadas

## Resultado

La aplicación cubre un agregado de gestión de riesgos por proyecto/edición con control de cambios, publicabilidad, planes de mitigación/contingencia y evidencias de suministrador. La paridad futura debe incluir como mínimo las capacidades siguientes; ninguna se marca como retirada.

| Dominio | Capacidades evidenciadas | Evidencia principal |
|---|---|---|
| Arranque e identidad | `EVE()` con `ReiniciarLasVariables`, `Constructor.getUsuario`, carga de `Entorno.ColItems`, IDAplicacion `5` (producción) / `51` (pruebas) | `src/modules/Variables Globales.bas:204` (`EVE`), `:248-272` (TempVars e IDAplicacion); `src/modules/Constructor.bas:2970` (`getUsuario`) |
| Roles | Tres: Administrador (calculado por `m_ObjEntorno.UsuarioConectadoEsAdministrador`), Calidad (`EsUsuarioCalidad`), Técnico (resto) | `src/modules/Variables Globales.bas:343-352` |
| Configuración temporal | `CadenaJerarquicaModelo` (nuevo/antiguo), `JPMesesAvisoEntreEdiciones=3`, `JPDiasPreviosParaElAviso=15`, `CalDiaInicialMesAviso=2`, `Publicabilidad_Usar_Cache=No`, `DatosEnLocal=No`, `EnDesarrollo=No`, `EnPruebas=No` | `src/modules/Variables Globales.bas:247-272` |
| Proyecto | alta/edición, responsables (autorizado/técnico/calidad), ámbito, oficinas, RAC | `Proyecto.cls`, `ProyectoSuministrador.cls`, `Form_FormProyectosGestion.cls` |
| Edición | ciclo de vida, `EsActivo`, `FechaPublicacion`, suministros vinculados, proyectos asociados, histórico | `Edicion.cls` (referenciado por 27 archivos), `Form_Form0BDOpciones*.cls` |
| Riesgo | alta/edición/borrado condicionado, descripción, priorización, origen, causa raíz (biblioteca o libre), impacto global, fecha detectado/retirado/materializado | `Riesgo.cls`, `RiesgoBiblioteca.cls`, `Form_FormRiesgosGestionRiesgo.cls`, `Form_FormGestionRiesgos*.cls` |
| Estados del riesgo | `EnumRiesgoEstado` (calculado), cambio manual, histórico de estados, `RiesgoAltoOMuyAlto`, valoración calculada | `Riesgo.cls` (línea 71 `m_eESTADOCalculado`), `Form_FormRiesgosGestionRiesgo.cls:265` |
| Causa raíz y biblioteca | `RequiereRiesgoDeBibliotecaCalculado` (Sí/No por proyecto), adaptación de UI (`Descripcion.Height`) | `Riesgo.cls`, `Form_FormRiesgosGestionRiesgo.cls:230-243` |
| Riesgo externo | alta, edición, vínculo con expediente, `ProvieneDeRiesgoExterno` | `RiesgoExterno.cls` |
| Materialización | alta, edición,撤回 (`RiesgoMaterializacion`) | `RiesgoMaterializacion.cls`, `Form_FormCalidadRiesgoMaterializaciones.cls` |
| Planes de mitigación (PM) | alta, edición, acciones, reversas, calendario, responsable | `PM.cls`, `PMAccion.cls`, `PMAccionReversa.cls`, `Form_FormPlanPrincipal.cls` |
| Planes de contingencia (PC) | alta, edición, acciones, reversas, responsable | `PC.cls`, `PCAccion.cls`, `PCAccionReversa.cls`, `Form_FormPlanPrincipal.cls` |
| Acciones / reversas | sub-formularios, responsables, fechas, motivo | `PMAccion.cls`, `PMAccionReversa.cls`, `PCAccion.cls`, `PCAccionReversa.cls` |
| Evidencias y anexos | alta por calidad/técnico, anexo único por edición-suministrador, `MotivoAnexarTecnicoNoOK` (chequeo de URL local, existencia y apertura de fichero) | `Anexo.cls`, `AnexoAntiguo.cls`, `EdicionSuministrador.cls:286-381` (`AnexarPorTecnico`, `MotivoAnexarTecnicoNoOK`) |
| Suministradores por edición | alta por calidad, edición de evidencia, baja por calidad, árbol propio | `EdicionSuministrador.cls`, `Form_FormGestionRiesgosSuministradores.cls` |
| Control de cambios (CC) | versiones, documentos, cambios, `CCCambio`, `CCDocumentoCambio`, `CCVersion` | `CCCambio.cls`, `CCDocumentoCambio.cls`, `CCVersion.cls`, `Form_FormControlCambiosGestion.cls` |
| Publicabilidad | veredicto (`NoPublicable`/`Publicable`), `tPublicabilidadRiesgoDatos`, checks, log | `Form_FormRiesgosGestionRiesgo.cls:173`, `Form_FormPublicacionCalidad.cls`, `PublicacionLog.cls` |
| Calidad — tareas | seis worklists (aceptado/retirado/visado, materializaciones, retipificación, detalle de edición), filtros | `TareasCalidad.cls`, `ArbolTareasCalidad.cls`, `Form_FormCalidadTarea*.cls`, `Form_FormCalidadTareas.cls`, `Form_FormCalidadTareasDetalleEdicion.cls` |
| Calidad — riesgos aceptados/retirados/visados/materializados/retipificación | tareas específicas por estado, formularios dedicados | `Form_FormCalidadRiesgoAceptadoRetiradoVisado.cls`, `Form_FormCalidadRiesgoMaterializaciones.cls` |
| Técnico — tareas | worklist propio, árbol | `TareasTecnico.cls`, `ArbolTareasTecnico.cls`, `Form_Form0BDOpcionesTecnico.cls` |
| Indicadores | pantalla dedicada, salida HTML/Excel | `Form_FormIndicador.form.txt`, `Form_formIndicadorProyectos.form.txt` |
| Informes | salida por tipo (HTML, PDF, Excel) vía `InformeRiesgoPDFServicio`, selector de salida | `InformeRiesgoPDFServicio.cls`, `Form_FormInformeTipoSalida.cls` |
| Histórico de estados del riesgo | `lstEstadosHistoricos` con `getEstadosDiferentesHastaEdicion` (lista estado;fecha) | `Form_FormRiesgosGestionRiesgo.cls:294-329` |
| Catálogos | alta/edición/baja de: `AreaImpacto`, `CarenciasExplicacion`, `Juridica`, `MitigacionValor`, `NC`, `OrganoContratacion`, `Pedido`, `RAC` | sus respectivas `.cls` |
| Exportación e impresión | Excel, HTML, generación de informes PDF | `InformeRiesgoPDFServicio.cls` |
| Integraciones | Lanzadera (identidad y permisos vía `getdbLanzadera`), Expedientes (vínculo por código), HPS, No Conformidades, AGEDYS (vía código compartido), correos | `Constructor.bas` (callers de `getdbLanzadera`), `Expediente.cls`, `ExpedienteResponsable.cls` (compartidos con Lanzadera), `NC.cls` |
| Acoplamiento DAO directo | `getdb()` con **308 callers** (más del doble que Expedientes con 135). La aplicación es intensísima en DAO. | `src/modules/Variables Globales.bas:485` (`getdb`) |

## Reglas de conservación

- La lectura del árbol jerárquico (`CargarArbol`) **se dispara desde cualquier cambio de edición, riesgo, plan o materialización** (no solo al abrir la pantalla). Esto amplifica el coste de la lentitud reportada: cada edición del usuario puede requerir recarga del árbol completo.
- El flag `CadenaJerarquicaModelo` admite dos modelos distintos de carga jerárquica (`"nuevo"` / `"antiguo"`); ambos modelos están activos en producción y ambos pueden ser causantes del síntoma de lentitud según el path de acceso. La nueva plataforma debe **descartar ambos** en favor de un modelo único basado en PostgreSQL + CTE recursivo (ver [Matriz de migración § D88](migration-matrix.md#d88-modelo-de-árbol-de-riesgos-en-la-nueva-plataforma)).
- La identidad y permisos se resuelven **directamente contra la base de datos de Lanzadera** (`getdbLanzadera()`), no contra un servicio. La nueva plataforma debe exponer identidad vía el adaptador de autenticación unificado (D9–D10) y eliminar este acoplamiento directo.
- `IDAplicacion = "5"` (producción) y `"51"` (pruebas) se asignan por el flag `EnPruebas`; la nueva plataforma lee `IDAplicacion` de la configuración, no de la TempVar.
- Las seis worklists de calidad y el árbol de tareas técnicas se rigen por `PintarTareas` análogo al de Expedientes; el patrón `Polling HTMX con `hx-trigger="every 30s"` + botón refresh manual` (D69) se aplica también aquí.
- El inventario incluye capacidades raras/técnicas: `InformeRiesgoPDFServicio`, control de versiones CC, `Publicabilidad_Usar_Cache` (hoy "No"), `DatosEnLocal` (hoy "No"), `EnDesarrollo` (hoy "No").

## Evidencia previa

Se han cosechado PRD, Discovery Map, Architecture Overview, ERD, OpenSpec CAP-001..055, UAT y releases antes de inspeccionar staging. CodeGraph-VBA sobre `00_main` se consultó primero; las afirmaciones divergentes entre esos documentos y el código quedan abiertas, no resueltas por intención. La inspección Dysflow sobre el binario y el backend autoritativo queda pendiente para una segunda pasada.