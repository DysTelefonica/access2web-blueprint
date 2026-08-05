# Brass — formularios, navegación y call paths

## Navegación principal

```text
Form_frmSplash (presumido)
  -> EVE (Variables Globales.bas:264)
     -> getdb() (Variables Globales.bas:535) — ⚠️ PWD="dpddpd" hardcodeado (D104)
     -> TbConfiguracionBackends
  -> Form_Form0BDOpciones (opciones generales)
      ├─> Form_FormEventoAlta (alta de evento)
      │    ├─> Form_FormEventoEqMedida (asignar equipos de medida al evento)
      │    └─> Form_FormActividadAlta (alta de actividad del evento)
      ├─> Form_FormEventoEdicion (edición)
      │    └─> Form_FormEventoEqMedidaEdicion (edición equipos de medida)
      ├─> Form_FormFacturaAlta (alta de factura)
      │    ├─> Form_FormFacturaActividades (actividades involucradas)
      │    ├─> Form_FormFacturaEventos (eventos involucrados)
      │    ├─> Form_FormFacturaGastos (gastos involucrados)
      │    ├─> Form_FormFacturaMateriales (materiales involucrados)
      │    └─> Form_FormFacturaSubcontrataciones (subcontrataciones involucradas)
      ├─> Form_FormPlanificacion (planificación)
      ├─> Form_FormParte (parte de trabajo)
      ├─> Form_FormMaterial (gestión de materiales)
      ├─> Form_FormTecnico (gestión de técnicos)
      ├─> Form_FormEquipo (gestión de equipos)
      ├─> Form_FormEquipoMedidaCalibracion (gestión de calibración)
      └─> Form_FormInforme (informes)
```

## Call paths críticos

| Capacidad | Camino observado | Persistencia / efecto |
|---|---|---|
| Inicio | `frmSplash.Form_Timer → EVE → getdb() (con PWD="dpddpd") → TempVars + Entorno.ColItems → m_ObjUsuarioConectado` | TempVars, sesión, validación fail-fast |
| Alta de evento | `Form_FormEventoAlta → EventoDAO.Insert → TbEventos → TbEventosEquipoMedida (asignar)` | transacción DAO multi-tabla |
| Edición de evento | `Form_FormEventoEdicion → EventoDAO.Update → TbEventos + TbEventosEquipoMedida` | idem |
| **Franqueo con SLA** | `Form_FormFranquear → Evento.Franquear(p_CausaFin) → ValidarSLA → MotivosNoFranqueable → UPDATE TbEventos SET FechaFinal, HORAFINALEVENTO, CAUSAFIN, Franqueado=True` | cierre de evento con auditoría |
| Equipos de medida | `Form_FormEquipoMedidaCalibracion → EquipoMedidaDAO → TbEquiposMedida + TbEquiposMedidaCalibraciones` | gestión de calibración (regulatorio) |
| Calibración de equipo | `Form_FormEquipoMedidaCalibracion → EquipoMedidaCalibracionDAO.Insert → TbEquiposMedidaCalibraciones` | registro de calibración con fecha |
| Materiales en evento | `Form_FormActividadAlta → MaterialDAO.Insert → TbMaterial + TbActividadesLotes` | transacción |
| Factura | `Form_FormFacturaAlta → FacturaDAO.Insert → TbFacturaPrincipal + 5 tablas de detalle (Actividades, Eventos, Gastos, Materiales, Subcontrataciones)` | factura con múltiples involucrados |
| Subcontratación | `Form_FormSubcontratacion → SubcontratacionDAO → TbSubcontrataciones` | gestión de subcontratación |
| Técnico asignado | `Form_FormTecnicoAsignacion → TecnicoDAO → TbTecnicos` | asignación de técnico a evento |
| Parte de trabajo | `Form_FormParte → ParteDAO → TbPartesPpal + TbPartesDetalle` | parte de trabajo |
| **Validación SLA** (servicio compartido) | `Evento.MotivosNoFranqueableSLA() → ValidarSLA(p_EnFranqueo:=True) → reglas de negocio` | mantiene reglas de SLA en un solo lugar |
| **Consultas de inventario** | `Constructor.getActividadesNoDeEventosParaFactura / getMaterialesParaFactura / getActividadesDeEventosParaFactura / getTipoTecnicoActivo` (4+ helpers) | queries SQL para selección |

## Inventario normalizado

- **Formularios**: ~30+ archivos `Form_*.cls` cada uno con `.form.txt` (estimación; los forms principales son `Form_FormEvento*`, `Form_FormFactura*`, `Form_FormActividad*`, `Form_FormPlanificacion*`, `Form_FormParte*`, `Form_FormMaterial*`, `Form_FormTecnico*`, `Form_FormEquipo*`, `Form_FormEquipoMedida*`, `Form_FormInforme*`).
- **Patrón `.cls + .form.txt`**: presente.
- **Clases** (28 en `src/classes/`):
  - **Domain principal**: `Evento`, `Actividad`, `Factura`, `FacturaPerfiles`, `Material`, `MaterialSeguimiento`, `Tecnico`, `TecnicoHoras`, `TipoTecnico`, `TipoTecnicoPrecio`, `Usuario`, `UsuarioAplicacionPermisos`, `Equipo`, `EquipoMedida`, `EquipoMedidaCalibracion`, `EventoEquipoMedida`, `EventoEquipoMedidaAntes`, `Subcontratacion`, `Parte`, `Anexo`, `Informe`, `GestorInforme`, `Planificacion`, `LIbranza` (typo: "Libranza"), `TipoDia`, `Fiesta`, `Entorno`, `Actividad`.
  - **Compartidas con Lanzadera**: `UsuarioAplicacionPermisos`, `Entorno`.
- **Módulos** (estimación): `Variables Globales.bas` (EVE + getdb), `Constructor.bas` (factories: `getEvento`, `getActividad`, `getFacturaDeActividad`, `getFacturaDeEvento`, `getEquipoMedida`, `getEquipoMedidaCalibracion`, etc.), `FUNCIONES UTILES.bas`, `ValidarSLA.bas` (presumido), `Instalador.bas`, `Automatizacion*.bas` (presumido).
- **Tests VBA**: no inspeccionados en esta pasada.

## Nota de evidencia

CodeGraph-VBA se consultó primero sobre el repo y devolvió 101 símbolos en 6 archivos para la query inicial (incluyendo `EVE`, `getUsuario`, `Constructor.getEvento`, `Constructor.getFacturaDeActividad`, `EventoEquipoMedida`, `EventoEquipoMedidaAntes`). Dysflow read-only se ejecutó después de `setup_project` + `register_worktree` + `accessPath` + `backendPath` absolutos: 60 tablas, 27 FKs, 5777 eventos, 25639 actividades, 7592 anexos. La inspección de UI se mantiene read-only y no se han alterado formularios.