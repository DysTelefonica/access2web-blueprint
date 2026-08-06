<!--
Documento de capacidad CAP-009 - Suministrador (CRUD generico, variante con 7 fields).
-->

# Capacidad: Suministrador (CRUD de entidades)

## 0 Identidad
- **ID de capacidad**: CAP-009
- **Tier**: standard
- **Estado**: active
- **Source**: hybrid (SDD + codigo actual)
- **Responsable / autoridad de producto**: Pendiente de confirmacion
- **Ultima verificacion**: 2026-06-16 mediante `dysflow.get_schema` (TbSuministradores) + `dysflow.test_vba` (9/9 verde) + `dysflow.compile_vba` (PASS)
- **Confianza global**: alta

## 1 Intencion de negocio - POR QUE
- **Proposito**: CRUD de Suministradores (proveedores externos) que se asocian a expedientes via `TbExpedientesSuministradores` (FK a `TbSuministradores.IDSuministrador`).
- **Problema que resuelve**: 2 forms con 7 fields cada uno (Nombre, Nemotecnico, CIF, Direccion, Ciudad, CP, TramitadoraHPS). Mas complejo que los otros entities.
- **Valor de negocio**: rewire a Helper_EntidadCRUD con 7 fields. Demuestra que el helper escala a entities con muchos campos.

## 2 Contrato de comportamiento

### Reglas de negocio (5, mismo patron que CAP-006)

| ID regla | Enunciado | Confianza |
|---|---|---|
| **BR-009-01** | `CopiarCamposAObjeto` con `p_Obj=Nothing` retorna error | Verified-runtime |
| **BR-009-02** | `HaHabidoCambiosGenerico` con `p_ObjInicial=Nothing` retorna `True` | Verified-runtime |
| **BR-009-03** | `EliminarEntidadGenerico` con `p_Operaciones=Nothing` retorna error | Verified-runtime |
| **BR-009-04** | `CopiarCamposAObjeto` copia los 7 fields de `m_ValoresForm` a `m_ObjSuministradorActivo` via `CallByName` | Verified-runtime (via MockEntidad + rewire) |
| **BR-009-05** | `Form_FormSuministradoresGestion.ComandoEliminar_Click` elimina el Suministrador seleccionado via el helper | Verified-static (rewire compila) |

(Los tests del helper son los mismos genericos; el rewire del form es lo unico especifico.)

## 3 Especificacion tecnica
- **Forms**: `Form_FormSuministrador.cls` (single-record) + `Form_FormSuministradoresGestion.cls` (lista).
- **DAO**: `TbSuministradores` (PK = `IDExpedienteSuministrador` segun schema; FK a `TbSuministradores.IDSuministrador` enforced).
- **Field names**: 7 (Nombre, Nemotecnico, CIF, Direccion, Ciudad, CP, TramitadoraHPS). El unico entity con tantos fields.
- **Variable names**: `m_ObjSuministradorActivo`, `m_SuministradorAlInicio`, `m_SuministradorOP`, `m_SuministracionSeleccionado`.

## 4-7
- Operacion: 2 commits (`a99d9e4` + `b59b03f`).
- Cobertura: 4/5 reglas `Verified-runtime`.
