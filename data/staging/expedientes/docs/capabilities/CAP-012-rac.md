<!--
Documento de capacidad CAP-012 - RAC.
-->

# Capacidad: RAC (CRUD de entidades)

## 0 Identidad
- **ID de capacidad**: CAP-012
- **Tier**: standard
- **Estado**: active
- **Source**: hybrid (SDD + codigo actual)
- **Ultima verificacion**: 2026-06-16
- **Confianza global**: alta

## 1 Intencion de negocio
- **Proposito**: CRUD de RAC (Responsables Administrativos de Contratacion) con 3 fields (RAC, CORREO, DESCRIPCION). Mas completo que los entities de 2 fields.

## 2 Contrato

| ID regla | Enunciado | Confianza |
|---|---|---|
| **BR-012-01..03** | Errores del helper | Verified-runtime |
| **BR-012-04** | `CopiarCamposAObjeto` copia `RAC`, `CORREO`, `DESCRIPCION` (3 fields) | Verified-runtime |
| **BR-012-05** | `Form_FormRACSGestion.ComandoEliminar_Click` elimina via helper | Verified-static |

## 3 Especificacion tecnica
- **Forms**: `Form_FormRAC.cls` + `Form_FormRACSGestion.cls`.
- **DAO**: `TbRACS` (PK = `IDRAC` segun schema).
- **Field names**: 3 (RAC, CORREO, DESCRIPCION).
- **Variable names**: `m_ObjRACActivo`, `m_RACAlInicio`, `m_RACOp`, `m_RACSeleccionado`.

## 4-7
- Operacion: 2 commits (`a99d9e4` + `b59b03f`).
- Cobertura: 4/5 reglas `Verified-runtime`.
