<!--
Documento de capacidad CAP-010 - LugarEjecucion.
-->

# Capacidad: Lugar de ejecucion (CRUD de entidades)

## 0 Identidad
- **ID de capacidad**: CAP-010
- **Tier**: standard
- **Estado**: active
- **Source**: hybrid (SDD + codigo actual)
- **Ultima verificacion**: 2026-06-16
- **Confianza global**: alta

## 1 Intencion de negocio
- **Proposito**: CRUD de Lugares de ejecucion (geografia: Madrid, Barcelona, etc.) via `TbExpedientesLugaresEjecucion`.
- **Diferencia notable**: el variable se llama `m_ObjLugarEjecucionActiva` (femenino, no masculino). Detalle de nomenclatura que el rewire respeto.

## 2 Contrato

| ID regla | Enunciado | Confianza |
|---|---|---|
| **BR-010-01** | `CopiarCamposAObjeto` con `p_Obj=Nothing` retorna error | Verified-runtime |
| **BR-010-02** | `HaHabidoCambiosGenerico` con `p_ObjInicial=Nothing` retorna `True` | Verified-runtime |
| **BR-010-03** | `EliminarEntidadGenerico` con `p_Operaciones=Nothing` retorna error | Verified-runtime |
| **BR-010-04** | `CopiarCamposAObjeto` copia `LugarEjecucion` y `DESCRIPCION` (2 fields) | Verified-runtime |
| **BR-010-05** | `Form_FormLugarEjecucionGestion.ComandoEliminar_Click` elimina via helper | Verified-static |

## 3 Especificacion tecnica
- **Forms**: `Form_FormLugarEjecucion.cls` + `Form_FormLugarEjecucionGestion.cls`.
- **DAO**: `TbLugaresEjecucion` (PK = `IDLugarEjecucion` segun schema).
- **Field names**: 2 (LugarEjecucion, DESCRIPCION).

## 4-7
- Operacion: 2 commits (`7d814d2` + `b59b03f`).
- Cobertura: 4/5 reglas `Verified-runtime`.
