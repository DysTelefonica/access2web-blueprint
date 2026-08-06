<!--
Documento de capacidad CAP-011 - PECAL.
-->

# Capacidad: PECAL (CRUD de entidades)

## 0 Identidad
- **ID de capacidad**: CAP-011
- **Tier**: standard
- **Estado**: active
- **Source**: hybrid (SDD + codigo actual)
- **Ultima verificacion**: 2026-06-16
- **Confianza global**: alta

## 1 Intencion de negocio
- **Proposito**: CRUD de PECAL (categorias contractuales) via `TbExpedientesPECAL`.

## 2 Contrato

| ID regla | Enunciado | Confianza |
|---|---|---|
| **BR-011-01..03** | Errores del helper | Verified-runtime |
| **BR-011-04** | `CopiarCamposAObjeto` copia `PECAL` y `DESCRIPCION` (2 fields) | Verified-runtime |
| **BR-011-05** | `Form_FormPECALESGestion.ComandoEliminar_Click` elimina via helper | Verified-static |

## 3 Especificacion tecnica
- **Forms**: `Form_FormPECAL.cls` + `Form_FormPECALESGestion.cls`.
- **DAO**: `TbPECAL` (PK = `IDPECAL` segun schema).
- **Field names**: 2 (PECAL, DESCRIPCION).
- **Variable names** (singular femenino): `m_ObjPECALActiva`, `m_PECALAlInicio`, `m_PECALOp`, `m_PECALSeleccionada`.

## 4-7
- Operacion: 2 commits (`a99d9e4` + `b59b03f`).
- Cobertura: 4/5 reglas `Verified-runtime`.
