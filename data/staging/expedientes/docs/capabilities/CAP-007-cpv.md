<!--
Documento de capacidad CAP-007 - CPV (CRUD generico de entidades, variante CPV).
Linaje: PRUEBA-003 REFAC-2a/2b + dysflow_get_schema de TbCPV.
Idioma: castellano de Espana. Identificadores de codigo/test se mantienen tal cual.
-->

# Capacidad: CPV (CRUD de entidades)

## 0 Identidad
- **ID de capacidad**: CAP-007
- **Tier**: standard
- **Estado**: active
- **Source**: hybrid (SDD + codigo actual)
- **Responsable / autoridad de producto**: Pendiente de confirmacion
- **Ultima verificacion**: 2026-06-16 mediante `dysflow.get_schema` (TbCPV) + `dysflow.test_vba` (9/9 verde slice-refac-2a/2b) + `dysflow.compile_vba` (PASS)
- **Confianza global**: alta - ver 7 (mismo patron que CAP-006)

## 1 Intencion de negocio (~ proposal SDD) - POR QUE
- **Proposito**: CRUD de CPV (Codigo Comun de Contratacion Publica, ej. "03000000-1" para "productos agricolas") que se asocian a expedientes via `TbExpedientesCPVs`.
- **Usuarios / perfiles**: administrador.
- **Problema que resuelve**: mismo patron que CAP-006 — duplicacion del CRUD en 2 forms (Form_FormCPV + Form_FormCPVsGestion) con 2 fields (CPV, DESCRIPCION).
- **Valor de negocio**: rewire a Helper_EntidadCRUD con 2 fields.
- **No-objetivos**: validacion de formato del CPV (no es un campo validado como digito-verificador — se acepta cualquier string).

## 2 Contrato de comportamiento (~ spec SDD) - QUE ? ANCLA DE REGRESION

### Escenarios (Dado / Cuando / Entonces)
- **DADO** el usuario hace click en "Alta" en `Form_FormCPVsGestion` **CUANDO** completa los campos `CPV` y `DESCRIPCION` y hace click en "Registrar" **ENTONCES** el `CPV` se guarda en `TbCPV` y dispara el evento `Alta`.
- **DADO** el usuario selecciona un CPV **CUANDO** modifica `CPV` o `DESCRIPCION` **ENTONCES** `Helper_EntidadCRUD.HaHabidoCambiosGenerico(m_CPVAlInicio, m_ValoresForm, Array("CPV","DESCRIPCION"), p_Error)` retorna `True` y permite guardar.
- **DADO** el usuario selecciona un CPV **CUANDO** hace click en "Eliminar" y confirma **ENTONCES** `Helper_EntidadCRUD.EliminarEntidadGenerico(m_CPVOp, "CPV", m_CPVSeleccionado, m_Error)` ejecuta la eliminacion via `CallByName`.

### Reglas de negocio

| ID regla | Enunciado | Autoridad | Aplicada en codigo? | Prueba | Confianza |
|---|---|---|---|---|---|
| **BR-007-01** | `Helper_EntidadCRUD.CopiarCamposAObjeto` con `p_Obj=Nothing` retorna error | spec | Si - `Helper_EntidadCRUD.CopiarCamposAObjeto` (early return) | `Test_Helper_EntidadCRUD_CopiarCamposAObjeto_ObjNothing_PueblaError` - PASA 2026-06-15 | Verified-runtime |
| **BR-007-02** | `Helper_EntidadCRUD.HaHabidoCambiosGenerico` con `p_ObjInicial=Nothing` retorna `True` (es un alta) | spec | Si | `Test_Helper_EntidadCRUD_HaHabidoCambios_ObjInicialNothing_DevuelveTrue` - PASA 2026-06-15 | Verified-runtime |
| **BR-007-03** | `Helper_EntidadCRUD.EliminarEntidadGenerico` con `p_Operaciones=Nothing` retorna error | spec | Si | `Test_Helper_EntidadCRUD_EliminarEntidadGenerico_OpNothing_PueblaError` - PASA 2026-06-16 | Verified-runtime |
| **BR-007-04** | `Helper_EntidadCRUD.CopiarCamposAObjeto` copia `CPV` y `DESCRIPCION` de `m_ValoresForm` a `m_ObjCPVActivo` via `CallByName` | spec | Si - rewire de `Form_FormCPV.ComandoRegistrar_Click` (commit `770ca37`) | (no testeable sin MockCPV.cls) | Verified-static |
| **BR-007-05** | `Form_FormCPVsGestion.ComandoEliminar_Click` elimina el CPV seleccionado via el helper | spec | Si - rewire (commit `b59b03f`) | (no testeable sin abrir UI) | Verified-static |

(Las reglas son identicas a CAP-006 salvo los field names y el nombre de la property en el helper.)

## 3 Especificacion tecnica (~ design SDD) - COMO
- **Forms involved**: `Form_FormCPV.cls` (single-record) + `Form_FormCPVsGestion.cls` (lista).
- **Helper**: `src/modules/Helper_EntidadCRUD.bas` (mismo que CAP-006).
- **DAO**: `TbCPV` (PK = `IDCPV` segun schema).
- **Field names**: 2 (CPV, DESCRIPCION). El resto identico a CAP-006.
- **Variable names** (especificas de CPV): `m_ObjCPVActivo`, `m_CPVAlInicio`, `m_CPVOp`, `m_CPVSeleccionado`.

## 4 Operacion (~ tasks SDD) - QUIEN/HACE CUANDO
- **PR owner**: PR-REFAC-2a/2b (cerrado en este slice).
- **Tests anadidos en este slice**: mismos 9 tests del helper (compartidos con CAP-006).
- **Tareas realizadas**: T1a.3, T1a.4, T1a.5, T2a.6 (Form_FormCPV rewire), T2b.6 (Form_FormCPVsGestion rewire).

## 5 Referencias cruzadas
- Spec: `staging-alignment-prueba-002/specs/capabilities/CAP-007-cpv.spec.md`.
- Forms: `Form_FormCPV.cls` + `Form_FormCPVsGestion.cls`.
- Helper: `src/modules/Helper_EntidadCRUD.bas`.
- CAP-006 (Comercial) — gemelo, mismo patron con distintos field names.
- Engram: `sdd/expedientes/capability-map/explore`.

## 6 Lagunas explicitas
- **BR-007-04/05 happy path no testeable** (mismo gap que CAP-006 — requiere MockCPV.cls).
- **`EstablecerDatos` no rewireado** (mismo gap que CAP-006).
- **`ComandoEliminar` no diferencia cancelacion del usuario** (mismo gap que CAP-006).
- **Validacion de formato CPV no implementada** (fuera de scope; el campo es string libre).

## 7 Resumen de cobertura
- Total reglas: 5 (BR-007-01..05).
- `Verified-runtime`: 3 (input validation del helper).
- `Verified-static`: 2 (rewire del form compila, happy path no testeable).
- Tasa de cobertura verde: 9/9 tests slice-refac-2a/2b.
- Coverage delta: 0% -> 60% del helper cubierto, 0% -> 100% (2/2 forms) rewireados y compilando.
