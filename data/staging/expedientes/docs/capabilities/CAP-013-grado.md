<!--
Documento de capacidad CAP-013 - GradoClasificacion.
-->

# Capacidad: Grado de clasificacion (CRUD de entidades)

## 0 Identidad
- **ID de capacidad**: CAP-013
- **Tier**: standard
- **Estado**: active
- **Source**: hybrid (SDD + codigo actual)
- **Ultima verificacion**: 2026-06-16
- **Confianza global**: alta

## 1 Intencion de negocio
- **Proposito**: CRUD de Grados de clasificacion (Público, Reservado, Confidencial, Secreto) via `TbExpedientes` (columna `IdGradoClasificacion`).

## 2 Contrato

| ID regla | Enunciado | Confianza |
|---|---|---|
| **BR-013-01..03** | Errores del helper | Verified-runtime |
| **BR-013-04** | `CopiarCamposAObjeto` copia `GradoClasificacion` y `DESCRIPCION` (2 fields) | Verified-runtime |
| **BR-013-05** | `Form_FormGradosClasificacionGestion.ComandoEliminar_Click` elimina via helper | Verified-static |

## 3 Especificacion tecnica
- **Forms**: `Form_FormGradoClasificacion.cls` + `Form_FormGradosClasificacionGestion.cls`.
- **DAO**: `TbGradosClasificacion` (PK = `IdGradoClasificacion` segun schema, nota lowercase 'd').
- **Field names**: 2 (GradoClasificacion, DESCRIPCION).
- **Variable names**: `m_ObjGradoClasificacionActivo`, `m_GradoClasificacionAlInicio`, `m_GradoClasificacionOp`, `m_GradoClasificacionSeleccionado`.

## 4-7
- Operacion: 2 commits (`a99d9e4` + `b59b03f`).
- Cobertura: 4/5 reglas `Verified-runtime`.
