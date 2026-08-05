# Gestion_Riesgos — formularios, navegación y call paths

## Navegación principal

```text
Form_frmSplash
  -> EVE()
  -> Form_Form0BDOpciones | Form_Form0BDOpcionesTecnico
      -> Form_FormProyectosGestion
          -> Form_FormGestionRiesgos
              -> Form_FormGestionRiesgosAutorizados | Form_FormGestionRiesgosDatosGenerales
                 | Form_FormGestionRiesgosRiesgosOferta | Form_FormGestionRiesgosSuministradores
              -> Form_FormRiesgosGestionRiesgo (detalle de un riesgo)
                  -> Form_FormPlanPrincipal (PM/PC)
                  -> Form_FormPlanAcciones (acciones)
      -> Form_FormControlCambiosGestion
      -> Form_FormPublicacionCalidad
          -> Form_FormPublicacionCalidadNotas
      -> Form_FormIndicador | Form_formIndicadorProyectos
      -> Form_FormInformeTipoSalida
      -> Form_FormAnexos | Form_FormAnexos1
      -> Form_FormCalidadTareas
          -> Form_FormCalidadTareasDetalleEdicion
          -> Form_FormCalidadTareaRiesgosAceptadosRetirados
          -> Form_FormCalidadTareaRiesgosMaterializadosPorDecidir
          -> Form_FormCalidadTareaRiesgosRetipificacion
          -> Form_FormCalidadRiesgoAceptadoRetiradoVisado
          -> Form_FormCalidadRiesgoMaterializaciones
          -> Form_FormCalidadTareaExplicacion
```

## Call paths críticos

| Capacidad | Camino observado | Persistencia / efecto |
|---|---|---|
| Inicio | `frmSplash.Form_Timer → EVE → Constructor.getUsuario → Entorno.ColItems` | TempVars (incluye `CadenaJerarquicaModelo`, `JPMesesAvisoEntreEdiciones`, `JPDiasPreviosParaElAviso`, IDAplicacion 5/51), sesión, cachés, tareas |
| Carga del árbol | `Form_FormGestionRiesgos.Form_Load → CargarArbol(p_Refrescando) → CargarArbolPM/CargarArbolPC/CargarArbolTareasCalidad` | `MSComctlLib.TreeView`, `m_ColRiesgosAplicados`, persistencia visual; rendimiento sensible (ver [data-model.md § Árbol](data-model.md#rendimiento-del-árbol-de-riesgos--causa-raíz-y-opciones-de-implementación)) |
| Alta/edición de riesgo | `Form_FormRiesgosGestionRiesgo.ComandoDetalle_Click → DoCmd.OpenForm "FormRiesgo" → m_ObjRiesgoActivo.EstablecerDatos` | transacción DAO; cabecera, hijos, planes, histórico de estados |
| Borrado de riesgo | `Form_FormRiesgosGestionRiesgo.ComandoEliminar_Click → m_ObjRiesgoActivo.Borrar → Form_FormGestionRiesgos.CargarArbol` | limpia el árbol, refresca selección |
| Alta PM/PC | `Form_FormRiesgosGestionRiesgo.ComandoAltaPM/PC_Click → DoCmd.OpenForm "FormPlanPrincipal" → m_FormPlan.RaiseEvent PlanNuevo → Form_FormGestionRiesgos.CargarArbolPM/CargarArbolPC` | plan nuevo, recálculo del árbol |
| Anexo técnico | `EdicionSuministrador.AnexarPorTecnico(p_URLLocal) → MotivoAnexarTecnicoNoOK → Anexo.Registrar` → `getdb().Execute UPDATE TbProyectosEdicionesSuministradores SET IDAnexo = ...` | fichero en disco (chequeo de `fso.FileExists` y `FicheroAbierto`); anexo escrito en backend |
| Histórico de estados | `Form_FormRiesgosGestionRiesgo.EstablecerDatos → getEstadosDiferentesHastaEdicion → lstEstadosHistoricos.AddItem` | lista visual; no persiste por sí misma |
| Publicabilidad | `Form_FormRiesgosGestionRiesgo.EstablecerDatos → ConstruirDatosPublicabilidadRiesgo → EvaluarPublicabilidadRiesgo` | veredicto visual; `PublicacionLog` |
| Control de cambios | `Form_FormControlCambiosGestion → CCCambio → CCDocumentoCambio → CCVersion` | versiones + documentos + cambios |
| Calidad — tareas | `Form_FormCalidadTareas → Form_FormCalidadTareasDetalleEdicion → TareasCalidad / ArbolTareasCalidad` | worklists de aceptados/retirados/visados/materializados/retipificación |
| Técnico — tareas | `Form_Form0BDOpcionesTecnico → TareasTecnico / ArbolTareasTecnico` | worklist y árbol propios |
| Indicadores | `Form_FormIndicador / Form_formIndicadorProyectos` | salida HTML/Excel |
| Informes | `Form_FormRiesgosGestionRiesgo.ComandoVerInformeRiesgo_Click → GenerarInformeRiesgoHTML` | `InformeRiesgoPDFServicio` produce artefacto |

## Inventario normalizado

- **Formularios**: 30 clases `Form_*.cls` + 3 definiciones `.form.txt` (`Form_FormIndicador.form.txt`, `Form_formIndicadorProyectos.form.txt`, `Form_FormInformeTipoSalida.form.txt`); incluye splash, opciones generales, gestión de riesgos, gestión de proyectos, detalle de riesgo, planes PM/PC, control de cambios, publicación, indicadores, informes, anexos, tareas de calidad y técnico, gestión de suministradores y detalle de edición.
- **Clases**: 47 clases en `src/classes/` (dominio, infraestructura, helpers, entidades externas compartidas con Lanzadera). El inventario CodeGraph incluye `Riesgo*`, `Edicion`, `Proyecto*`, `PM`, `PC`, `Anexo*`, `Cambio*`, `CC*`, `Tareas*`, `ArbolTareas*`, `Entorno`, `Usuario*`, `Correo`, `InformeRiesgoPDFServicio`, `PublicacionLog`.
- **Módulos**: módulos de bootstrap/factory/DAO (`Variables Globales.bas`, `Constructor.bas`, `Funciones Generales.bas`), helpers UI, exportadores JSON/Excel/HTML/PDF, configuración de backend, utilidades y módulos de test.
- **Reports/macros/queries**: hay `src/reports/` con reports clásicos de Access. Queries exportadas forman parte del delta y deben mantenerse como evidencia separada. Macros embebidas requieren revisión del binario si no aparecen como fuente.

## Nota de evidencia

CodeGraph-VBA se consultó primero sobre `00_main` y devolvió call paths dinámicos. Dysflow `list_objects` y `get_schema` sobre el binario `Gestion_Riesgos.accdb` y el backend autoritativo `Gestion_Riesgos_Datos.accdb` quedan **pendientes para una segunda pasada** por límite de tiempo operativo en esta sesión; la inspección actual se basa exclusivamente en el código fuente exportado. La inspección de UI se mantiene read-only y no se han alterado formularios.