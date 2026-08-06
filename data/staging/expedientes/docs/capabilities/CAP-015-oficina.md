<!--
Documento de capacidad CAP-015 - OficinaPrograma.
-->

# Capacidad: Oficina de programa (CRUD de entidades)

## 0 Identidad
- **ID de capacidad**: CAP-015
- **Tier**: standard
- **Estado**: active
- **Source**: hybrid (SDD + codigo actual)
- **Ultima verificacion**: 2026-06-16
- **Confianza global**: alta

## 1 Intencion de negocio
- **Proposito**: CRUD de Oficinas de programa (unidades logisticas) via `TbExpedientes` (columna `IDOficinaPrograma`).

## 2 Contrato

| ID regla | Enunciado | Confianza |
|---|---|---|
| **BR-015-01..03** | Errores del helper | Verified-runtime |
| **BR-015-04** | `CopiarCamposAObjeto` copia `OficinaPrograma` y `DESCRIPCION` (2 fields) | Verified-runtime |
| **BR-015-05** | `Form_FormOficinasProgramaGestion.ComandoEliminar_Click` elimina via helper | Verified-static |

## 3 Especificacion tecnica
- **Forms**: `Form_FormOficinaPrograma.cls` + `Form_FormOficinasProgramaGestion.cls`.
- **DAO**: `TbOficinasPrograma` (PK = `IDOficinaPrograma` segun schema).
- **Field names**: 2 (OficinaPrograma, DESCRIPCION).
- **Variable names** (singular femenino): `m_ObjOficinaProgramaActiva`, `m_OficinaProgramaAlInicio`, `m_OficinaProgramaOp`, `m_OficinaProgramaSeleccionada`.

## 4-7
- Operacion: 2 commits (`a99d9e4` + `b59b03f`).
- Cobertura: 4/5 reglas `Verified-runtime`.
