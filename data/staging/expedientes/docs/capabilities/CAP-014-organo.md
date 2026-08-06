<!--
Documento de capacidad CAP-014 - OrganoContratacion.
-->

# Capacidad: Organo de contratacion (CRUD de entidades)

## 0 Identidad
- **ID de capacidad**: CAP-014
- **Tier**: standard
- **Estado**: active
- **Source**: hybrid (SDD + codigo actual)
- **Ultima verificacion**: 2026-06-16
- **Confianza global**: alta

## 1 Intencion de negocio
- **Proposito**: CRUD de Organos de contratacion (Ministerio de Defensa, Junta de Compras, etc.) via `TbExpedientes` (columna `IDOrganoContratacion`).

## 2 Contrato

| ID regla | Enunciado | Confianza |
|---|---|---|
| **BR-014-01..03** | Errores del helper | Verified-runtime |
| **BR-014-04** | `CopiarCamposAObjeto` copia `OrganoContratacion` y `DESCRIPCION` (2 fields) | Verified-runtime |
| **BR-014-05** | `Form_FormOrganoContratacionGestion.ComandoEliminar_Click` elimina via helper | Verified-static |

## 3 Especificacion tecnica
- **Forms**: `Form_FormOrganoContratacion.cls` + `Form_FormOrganoContratacionGestion.cls`.
- **DAO**: `TbOrganosContratacion` (PK = `IDOrganoContratacion` segun schema).
- **Field names**: 2 (OrganoContratacion, DESCRIPCION).

## 4-7
- Operacion: 2 commits (`a99d9e4` + `b59b03f`).
- Cobertura: 4/5 reglas `Verified-runtime`.
