<!--
Documento de capacidad CAP-008 - Ejercito (CRUD generico de entidades, variante Ejercito).
Linaje: PRUEBA-003 REFAC-2a/2b + dysflow_get_schema de TbEjercitos.
Idioma: castellano de Espana. Identificadores de codigo/test se mantienen tal cual.
-->

# Capacidad: Ejercito (CRUD de entidades)

## 0 Identidad
- **ID de capacidad**: CAP-008
- **Tier**: standard
- **Estado**: active
- **Source**: hybrid (SDD + codigo actual)
- **Responsable / autoridad de producto**: Pendiente de confirmacion
- **Ultima verificacion**: 2026-06-16 mediante `dysflow.get_schema` (TbEjercitos) + `dysflow.test_vba` (9/9 verde slice-refac-2a/2b + 3/3 happy-path del helper) + `dysflow.compile_vba` (PASS)
- **Confianza global**: alta - mismo patron que CAP-006 (Comercial)

## 1 Intencion de negocio (~ proposal SDD) - POR QUE
- **Proposito**: CRUD de Ejercitos (Ejercito de Tierra, Armada, etc.) que se asocian a expedientes via `TbExpedientes` (columna `IDEjercito`).
- **Usuarios / perfiles**: administrador.
- **Problema que resuelve**: mismo patron que CAP-006/007 — 2 forms (single-record + lista) con codigo duplicado.
- **Valor de negocio**: rewire a Helper_EntidadCRUD con 2 fields (Ejercito, DESCRIPCION).

## 2 Contrato de comportamiento (~ spec SDD) - QUE ? ANCLA DE REGRESION

### Reglas de negocio (5, mismo patron que CAP-006)

| ID regla | Enunciado | Aplicada? | Prueba | Confianza |
|---|---|---|---|---|
| **BR-008-01** | `Helper_EntidadCRUD.CopiarCamposAObjeto` con `p_Obj=Nothing` retorna error | Si | `Test_Helper_EntidadCRUD_CopiarCamposAObjeto_ObjNothing_PueblaError` - PASA | Verified-runtime |
| **BR-008-02** | `Helper_EntidadCRUD.HaHabidoCambiosGenerico` con `p_ObjInicial=Nothing` retorna `True` (es alta) | Si | `Test_Helper_EntidadCRUD_HaHabidoCambios_ObjInicialNothing_DevuelveTrue` - PASA | Verified-runtime |
| **BR-008-03** | `Helper_EntidadCRUD.EliminarEntidadGenerico` con `p_Operaciones=Nothing` retorna error | Si | `Test_Helper_EntidadCRUD_EliminarEntidadGenerico_OpNothing_PueblaError` - PASA | Verified-runtime |
| **BR-008-04** | `Helper_EntidadCRUD.CopiarCamposAObjeto` copia `Ejercito` y `DESCRIPCION` de `m_ValoresForm` a `m_ObjEjercitoActivo` via `CallByName` | Si - rewire de `Form_FormEjercito.ComandoRegistrar_Click` (commit `770ca37`) | `Test_Helper_EntidadCRUD_CopiarCamposAObjeto_HappyPath_CopiaValores` (con MockEntidad, cierra el gap) - PASA 2026-06-16 | Verified-runtime |
| **BR-008-05** | `Form_FormEjercitosGestion.ComandoEliminar_Click` elimina el Ejercito seleccionado via el helper | Si - rewire (commit `b59b03f`) | (no testeable sin abrir UI) | Verified-static |

## 3 Especificacion tecnica (~ design SDD) - COMO
- **Forms**: `Form_FormEjercito.cls` (single-record) + `Form_FormEjercitosGestion.cls` (lista).
- **Helper**: `src/modules/Helper_EntidadCRUD.bas` (mismo que CAP-006/007).
- **DAO**: `TbEjercitos` (PK = `IDEjercito` segun schema).
- **Field names**: 2 (Ejercito, DESCRIPCION).
- **Variable names** (especificas de Ejercito): `m_ObjEjercitoActivo`, `m_EjercitoAlInicio`, `m_EjercitoOp`, `m_EjercitoSeleccionado`.

## 4-7 (mismo patron que CAP-006/007)
- Operacion: 2 commits de rewire (`770ca37` + `b59b03f`), 9/9 tests verde.
- Referencias: spec `CAP-008-ejercito.spec.md`, apply-progress, tasks, feature doc, Engram.
- Lagunas: mismas (happy path del EliminarEntidadGenerico + EstablecerDatos + cancelacion del usuario).
- Cobertura: 4/5 reglas `Verified-runtime` (BR-008-04 pasa por MockEntidad; las otras 4 son del helper generico que ya esta verde).
