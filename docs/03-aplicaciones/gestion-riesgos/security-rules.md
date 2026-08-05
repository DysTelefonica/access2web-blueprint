# Gestion_Riesgos — seguridad y reglas

## Autorización

La matriz común se mantiene en [06-autorizacion-legacy-matriz.md](../../06-autorizacion-legacy-matriz.md). Aquí solo se conserva el comportamiento específico de Gestion_Riesgos:

- **Tres roles** (no dos como en Expedientes): Administrador, Calidad, Técnico. La asignación se calcula en `EVE()`:
  - `EsAdministrador = m_ObjEntorno.UsuarioConectadoEsAdministrador` (lookup en `m_ObjEntorno.ColUsuariosAdministradores`).
  - `EsCalidad = m_ObjEntorno.UsuarioConectadoEsDeCalidad` (solo si no es Administrador).
  - `EsTecnico` es el resto (implícito).
- Las pantallas técnicas (`Form_Form0BDOpcionesTecnico`, `Form_FormGestionRiesgosAutorizados`) son de solo lectura en el flujo observado; los datos vienen de `getdbLanzadera` con `IDAplicacion = "5"` / `"51"`.
- El responsable autorizado del proyecto se evalúa en `Proyecto.UsuarioAutorizado`; condiciona `ComandoAltaPM/PC.Enabled` y `ComandoEliminar.Visible/Enabled` en `Form_FormRiesgosGestionRiesgo.EstablecerDatos` (líneas 187-222).

## Permisos por aplicación

Los permisos efectivos se cargan desde `UsuarioAplicacionPermisos` (clase compartida con Lanzadera). El acceso a Gestion_Riesgos está condicionado por `IDAplicacion = "5"` en `TbUsuariosAplicacionesPermisos` (producción) o `"51"` (pruebas). En la nueva plataforma esto se reemplaza por el **catálogo de capabilities** (D45-D46) declarado por el módulo.

## Reglas de negocio críticas

- **Alta de PM/PC requiere riesgo completo**: `ComandoAltaPM/PC_Click` verifica `m_ObjRiesgoActivo.FichaRiesgoCompleta = EnumSiNo.No` y muestra mensaje si falta información; ofrece abrir el detalle para completarlo.
- **Borrado de riesgo condicionado**: `ComandoEliminar_Click` verifica `m_ObjRiesgoActivo.Borrable = EnumSiNo.No` y `Edicion.UsuarioConectadoAutorizado = EnumSiNo.No` antes de permitir el borrado.
- **Anexo técnico condicionado**: `EdicionSuministrador.MotivoAnexarTecnicoNoOK` valida que:
  - `Edicion.EsActivo = EnumSiNo.Sí` (no se anexa en edición no activa).
  - `fso.FileExists(p_URLLocal)` (existe el fichero local).
  - `Not FicheroAbierto(p_URLLocal)` (fichero no abierto en otra app).
  - `fso.FileExists(m_URLFinal)` con `Not FicheroAbierto(m_URLFinal)` (no colisión en destino).
- **Suministrador único por edición**: `EdicionSuministrador.MotivoNoOK` rechaza alta duplicada (`Constructor.getSuministradorEnEdicion(...) Is Nothing` debe ser `True` para proceder).
- **Estado del riesgo calculado**: `Riesgo.ESTADOCalculadoTexto`, `Riesgo.FechaFinGarantiaCalculada`, `Riesgo.RiesgoAltoOMuyAltoTexto` se derivan de fechas y flags. La nueva plataforma debe **persistir el valor fuente y el calculado** (no inferir el uno del otro).
- **Histórico de estados preservado**: el `lstEstadosHistoricos` se carga con `getEstadosDiferentesHastaEdicion`; la nueva plataforma debe garantizar que los cambios manuales y automáticos quedan ambos en el histórico.
- **Publicabilidad auditable**: `ConstruirDatosPublicabilidadRiesgo + EvaluarPublicabilidadRiesgo` produce un veredicto; el log (`PublicacionLog`) registra la decisión con contexto.
- **Transaccionalidad**: alta/edición de riesgo combina cabecera, planes, anexos y registros asociados en operaciones DAO; la nueva plataforma usa transacciones SQLAlchemy (`AsyncSession.begin()`) coherentes con D66 y D82.

## Roles y funciones diferenciadas

A diferencia de Expedientes (que tiene dos roles: Administrador y Técnico), Gestion_Riesgos separa explícitamente **Calidad** de **Técnico**:

- **Calidad**: tareas de aceptación, retirada, visado, materialización, retipificación. Acceso a `Form_FormCalidadTareas` y derivados.
- **Técnico**: tareas propias, árbol propio. Acceso a `Form_Form0BDOpcionesTecnico`.

Las pantallas de calidad y técnico tienen **worklists distintas** (`TareasCalidad` / `ArbolTareasCalidad` vs `TareasTecnico` / `ArbolTareasTecnico`); la nueva plataforma debe mantener esta separación como **vista especializada** del módulo (D46), no como aplicaciones distintas.

## Riesgos de privacidad/migración

- No se han incluido nombres, correos, contraseñas, cadenas de conexión, hashes, hosts ni filas personales. **Los comentarios `m_Command = "..."` con correos de usuarios reales en `Variables Globales.bas` se omiten en este artefacto** (son artefactos de desarrollo que NO deben migrar).
- El acoplamiento directo con `getdbLanzadera()` debe eliminarse en la nueva plataforma; la identidad se resuelve por el adaptador unificado (D9–D10).
- Las URLs de anexos a SharePoint y las rutas UNC locales se mantienen en backend con metadatos; **no se exponen en este artefacto**.
- Las plantillas de correo (`Correo.cls`) pueden contener datos sensibles; revisar antes de portar.
- `TbRiesgosMaterializaciones`, `TbCambios`, `TbCarenciasExplicacion`, `TbPublicacionLog` son eventos con potencial valor probatorio; **la retención debe ser al menos la misma que la de auditoría** (D29).