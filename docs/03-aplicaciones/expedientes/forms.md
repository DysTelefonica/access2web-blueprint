# Expedientes — formularios, navegación y call paths

## Navegación principal

```text
Form_frmSplash
  -> EVE()
  -> Form_Form0BDOpciones | Form_Form0BDOpcionesTecnicos
      -> Form_FormExpedientesGestion | Form_FormExpedientesGestionTecnica
          -> Form_FormExpedienteAltaTipo -> Form_FormExpediente
              -> subforms General / Fechas / Entidades / Suministradores /
                 Hitos / Modificados / Documentación
      -> Form_FormTareas -> detalle
      -> Form_Form0BDGestorEntidades -> master/detail de catálogos
```

## Call paths críticos

| Capacidad | Camino observado | Persistencia/efecto |
|---|---|---|
| Inicio | `frmSplash.Form_Timer → EVE → LeeConfiguracionLocal → getUsuario → Entorno` | TempVars, sesión, cachés, tareas |
| Alta/edición | `FormExpediente.ComandoRegistrar_Click → RellenarDTO → ExpedienteOperaciones.Registrar` | transacción DAO, cabecera e hijos |
| Árbol | `FormExpedienteSuministradores → ExpedienteSuministradorServicio/Repositorio` | `TbExpedientesSuministradores`, jerarquía `IdPadre` |
| Búsqueda | `FormExpedientesGestion.Filtrar → Helper/constructor → ExpedienteCompleto` | lista y navegación a detalle |
| Tareas | `EVE/Registrar → PintarTareas → Entorno → FormTareas` | seis worklists y contadores |
| Exportación | `FormEleccionTipoConsulta → helper Excel` o `ExpedienteJsonExporter` | fichero externo, batch e historial |
| Auditoría | `Registrar → UltimoCambioOperaciones.Registrar` y `CambioOperaciones` | `TbUltimoCambio`, `TbCambios` |

## Inventario normalizado

- **Formularios:** 46 pares/clases `.cls` y definiciones `.form.txt` observables en staging, incluidos startup, menús, expediente, catálogos, tareas, E2E y utilidades.
- **Clases:** 59 clases de dominio/servicio/infraestructura y helpers/mocks; el inventario CodeGraph incluye `Expediente*`, `Entorno`, `Usuario*`, operaciones y entidades externas.
- **Módulos:** módulos de bootstrap/factory/DAO, helpers UI, exportadores JSON/Excel/E2E, configuración de backend, utilidades y módulos de test.
- **Reports/macros/queries:** no hay `.report.txt` en el árbol exportado inspeccionado; queries exportadas son parte del delta staging/main y deben mantenerse como evidencia separada. Las macros embebidas requieren revisión del binario si no aparecen como fuente.

## Nota de evidencia

CodeGraph-VBA se consultó primero sobre staging y devolvió call paths dinámicos. Dysflow `list_objects` devolvió 303 elementos binarios; su respuesta resumida no se usa para inventar nombres que no estén en el source tree. La inspección de UI se mantiene read-only y no se han alterado formularios.
