# Matriz campo a campo — catalogos

> Parte de la [matriz campo a campo](../matriz-campos.md). Reglas de mapeo, leyenda y límites de la fuente en ese documento. Commit fijado `ec6a8c3`.

## TbComerciales

3 columnas · 26 filas en la extracción del snapshot · citada en 4 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Evidencia en `src` | Destino | Tipo objetivo | Estado |
|---|---|---|---|---|---|
| `TbComerciales.IDComercial` | LONG(4) | F `ComercialOperaciones.cls:63` | `expedientes.TbComerciales.IDComercial` | BIGINT | existente (migración 0002) |
| `TbComerciales.Comercial` | TEXT(255) | F `ComercialOperaciones.cls:64` | `expedientes.TbComerciales.Comercial` | VARCHAR(255) | existente (migración 0002) |
| `TbComerciales.Descripcion` | MEMO | F `ComercialOperaciones.cls:66` | `expedientes.TbComerciales.Descripcion` | TEXT | existente (migración 0002) |

## TbCPV

3 columnas · 73 filas en la extracción del snapshot · citada en 4 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Evidencia en `src` | Destino | Tipo objetivo | Estado |
|---|---|---|---|---|---|
| `TbCPV.IDCPV` | LONG(4) | F `CPVOperaciones.cls:59` | `expedientes.TbCPV.IDCPV` | BIGINT | existente (migración 0002) |
| `TbCPV.CPV` | TEXT(255) | F `CPVOperaciones.cls:60` | `expedientes.TbCPV.CPV` | VARCHAR(255) | existente (migración 0002) |
| `TbCPV.DESCRIPCION` | MEMO | F `CPVOperaciones.cls:62` | `expedientes.TbCPV.DESCRIPCION` | TEXT | propuesta |

## TbEjercitos

3 columnas · 5 filas en la extracción del snapshot · citada en 4 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Evidencia en `src` | Destino | Tipo objetivo | Estado |
|---|---|---|---|---|---|
| `TbEjercitos.IDEjercito` | LONG(4) | F `EjercitoOperaciones.cls:63` | `expedientes.TbEjercitos.IDEjercito` | BIGINT | existente (migración 0002) |
| `TbEjercitos.Ejercito` | TEXT(255) | F `EjercitoOperaciones.cls:64` | `expedientes.TbEjercitos.Ejercito` | VARCHAR(255) | existente (migración 0002) |
| `TbEjercitos.Descripcion` | MEMO | F `EjercitoOperaciones.cls:66` | `expedientes.TbEjercitos.Descripcion` | TEXT | existente (migración 0002) |

## TbEstados

3 columnas · 9 filas en la extracción del snapshot · citada en 1 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Evidencia en `src` | Destino | Tipo objetivo | Estado |
|---|---|---|---|---|---|
| `TbEstados.IDEstado` | LONG(4) | S `ExpedienteJsonExporter.bas:968` | `expedientes.TbEstados.IDEstado` | BIGINT | existente (migración 0002) |
| `TbEstados.Estado` | TEXT(255) | F `ExpedienteJsonExporter.bas:841` | `expedientes.TbEstados.Estado` | VARCHAR(255) | existente (migración 0002) |
| `TbEstados.DESCRIPCION` | MEMO | F `ExpedienteJsonExporter.bas:785` | `expedientes.TbEstados.DESCRIPCION` | TEXT | existente (migración 0002) |

## TbGradosClasificacion

3 columnas · 5 filas en la extracción del snapshot · citada en 5 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Evidencia en `src` | Destino | Tipo objetivo | Estado |
|---|---|---|---|---|---|
| `TbGradosClasificacion.IdGradoClasificacion` | LONG(4) | F `GradoClasificacionOperaciones.cls:63` | `expedientes.TbGradosClasificacion.IdGradoClasificacion` | BIGINT | propuesta |
| `TbGradosClasificacion.GradoClasificacion` | TEXT(255) | F `GradoClasificacionOperaciones.cls:64` | `expedientes.TbGradosClasificacion.GradoClasificacion` | VARCHAR(255) | propuesta |
| `TbGradosClasificacion.Descripcion` | MEMO | F `GradoClasificacionOperaciones.cls:66` | `expedientes.TbGradosClasificacion.Descripcion` | TEXT | existente (migración 0002) |

## TbJefaturas

3 columnas · 3 filas en la extracción del snapshot · citada en 0 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Evidencia en `src` | Destino | Tipo objetivo | Estado |
|---|---|---|---|---|---|
| `TbJefaturas.IDJefatura` | LONG(4) | — | `expedientes.TbJefaturas.IDJefatura` | BIGINT | existente (migración 0002) |
| `TbJefaturas.Jefatura` | TEXT(255) | — | `expedientes.TbJefaturas.Jefatura` | VARCHAR(255) | existente (migración 0002) |
| `TbJefaturas.DESCRIPCION` | MEMO | — | `expedientes.TbJefaturas.DESCRIPCION` | TEXT | propuesta |

## TbJuridicas

4 columnas · 17 filas en la extracción del snapshot · citada en 2 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Evidencia en `src` | Destino | Tipo objetivo | Estado |
|---|---|---|---|---|---|
| `TbJuridicas.IDJuridica` | LONG(4) | S `Instalador.bas:1237` | `expedientes.TbJuridicas.IDJuridica` | BIGINT | existente (migración 0002) |
| `TbJuridicas.Juridica` | TEXT(255) | — | `expedientes.TbJuridicas.Juridica` | VARCHAR(255) | existente (migración 0002) |
| `TbJuridicas.DESCRIPCION` | MEMO | F `ExpedienteJsonExporter.bas:785` | `expedientes.TbJuridicas.DESCRIPCION` | TEXT | existente (migración 0002) |
| `TbJuridicas.IDSuministrador` | LONG(4) | F `ExpedienteJsonExporter.bas:765` | `expedientes.TbJuridicas.IDSuministrador` | BIGINT | propuesta |

## TbLugaresEjecucion

3 columnas · 34 filas en la extracción del snapshot · citada en 4 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Evidencia en `src` | Destino | Tipo objetivo | Estado |
|---|---|---|---|---|---|
| `TbLugaresEjecucion.IDLugarEjecucion` | LONG(4) | F `LugarEjecucionOperaciones.cls:63` | `expedientes.TbLugaresEjecucion.IDLugarEjecucion` | BIGINT | existente (migración 0002) |
| `TbLugaresEjecucion.LugarEjecucion` | MEMO | F `LugarEjecucionOperaciones.cls:64` | `expedientes.TbLugaresEjecucion.LugarEjecucion` | TEXT | existente (migración 0002) |
| `TbLugaresEjecucion.Descripcion` | MEMO | F `LugarEjecucionOperaciones.cls:66` | `expedientes.TbLugaresEjecucion.Descripcion` | TEXT | existente (migración 0002) |

## TbOficinasPrograma

3 columnas · 37 filas en la extracción del snapshot · citada en 5 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Evidencia en `src` | Destino | Tipo objetivo | Estado |
|---|---|---|---|---|---|
| `TbOficinasPrograma.IDOficinaPrograma` | LONG(4) | F `OficinaProgramaOperaciones.cls:62` | `expedientes.TbOficinasPrograma.IDOficinaPrograma` | BIGINT | existente (migración 0002) |
| `TbOficinasPrograma.OficinaPrograma` | TEXT(255) | F `OficinaProgramaOperaciones.cls:63` | `expedientes.TbOficinasPrograma.OficinaPrograma` | VARCHAR(255) | existente (migración 0002) |
| `TbOficinasPrograma.Descripcion` | MEMO | F `OficinaProgramaOperaciones.cls:65` | `expedientes.TbOficinasPrograma.Descripcion` | TEXT | existente (migración 0002) |

## TbOrganosContratacion

3 columnas · 46 filas en la extracción del snapshot · citada en 4 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Evidencia en `src` | Destino | Tipo objetivo | Estado |
|---|---|---|---|---|---|
| `TbOrganosContratacion.IDOrganoContratacion` | LONG(4) | F `OrganoContratacionOperaciones.cls:63` | `expedientes.TbOrganosContratacion.IDOrganoContratacion` | BIGINT | existente (migración 0002) |
| `TbOrganosContratacion.OrganoContratacion` | TEXT(255) | F `OrganoContratacionOperaciones.cls:64` | `expedientes.TbOrganosContratacion.OrganoContratacion` | VARCHAR(255) | existente (migración 0002) |
| `TbOrganosContratacion.Descripcion` | MEMO | F `OrganoContratacionOperaciones.cls:66` | `expedientes.TbOrganosContratacion.Descripcion` | TEXT | existente (migración 0002) |

## TbPECAL

3 columnas · 6 filas en la extracción del snapshot · citada en 5 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Evidencia en `src` | Destino | Tipo objetivo | Estado |
|---|---|---|---|---|---|
| `TbPECAL.IDPECAL` | LONG(4) | F `ExpedienteOperaciones.cls:1445` | `expedientes.TbPECAL.IDPECAL` | BIGINT | existente (migración 0002) |
| `TbPECAL.PECAL` | TEXT(255) | F `ExpedienteOperaciones.cls:1616` | `expedientes.TbPECAL.PECAL` | VARCHAR(255) | existente (migración 0002) |
| `TbPECAL.DESCRIPCION` | MEMO | F `ExpedienteOperaciones.cls:1806` | `expedientes.TbPECAL.DESCRIPCION` | TEXT | propuesta |

## TbRACS

4 columnas · 33 filas en la extracción del snapshot · citada en 4 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Evidencia en `src` | Destino | Tipo objetivo | Estado |
|---|---|---|---|---|---|
| `TbRACS.IDRAC` | LONG(4) | F `RACOperaciones.cls:63` | `expedientes.TbRACS.IDRAC` | BIGINT | existente (migración 0002) |
| `TbRACS.RAC` | TEXT(255) | F `RACOperaciones.cls:64` | `expedientes.TbRACS.RAC` | VARCHAR(255) | existente (migración 0002) |
| `TbRACS.CORREO` | TEXT(255) | F `RACOperaciones.cls:66` | `expedientes.TbRACS.CORREO` | VARCHAR(255) | existente (migración 0002) |
| `TbRACS.DESCRIPCION` | MEMO | F `RACOperaciones.cls:69` | `expedientes.TbRACS.DESCRIPCION` | TEXT | existente (migración 0002) |

## TbResponsablesPorRol

3 columnas · 11 filas en la extracción del snapshot · citada en 2 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Evidencia en `src` | Destino | Tipo objetivo | Estado |
|---|---|---|---|---|---|
| `TbResponsablesPorRol.IDResponsablePorRol` | LONG(4) | — | `expedientes.TbResponsablesPorRol.IDResponsablePorRol` | BIGINT | propuesta |
| `TbResponsablesPorRol.IDUsuario` | LONG(4) | F `modResponsablePorRolHelper.bas:260` | `expedientes.TbResponsablesPorRol.IDUsuario` | BIGINT | propuesta |
| `TbResponsablesPorRol.Rol` | TEXT(50) | S `modResponsablePorRolHelper.bas:179` | `expedientes.TbResponsablesPorRol.Rol` | VARCHAR(50) | existente (migración 0002) |

## TbSuministradores

10 columnas · 72 filas en la extracción del snapshot · citada en 7 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Evidencia en `src` | Destino | Tipo objetivo | Estado |
|---|---|---|---|---|---|
| `TbSuministradores.IDSuministrador` | LONG(4) | F `SuministradorOperaciones.cls:96` | `expedientes.TbSuministradores.IDSuministrador` | BIGINT | existente (migración 0002) |
| `TbSuministradores.Nombre` | TEXT(255) | F `Expediente.cls:865` | `expedientes.TbSuministradores.Nombre` | VARCHAR(255) | propuesta |
| `TbSuministradores.CIF` | TEXT(255) | F `SuministradorOperaciones.cls:98` | `expedientes.TbSuministradores.CIF` | VARCHAR(255) | propuesta |
| `TbSuministradores.DESCRIPCION` | MEMO | F `SuministradorOperaciones.cls:100` | `expedientes.TbSuministradores.DESCRIPCION` | TEXT | existente (migración 0002) |
| `TbSuministradores.TramitadoraHPS` | TEXT(2) | F `SuministradorOperaciones.cls:103` | `expedientes.TbSuministradores.TramitadoraHPS` | VARCHAR(2) · Sí/No como texto | propuesta |
| `TbSuministradores.Nemotecnico` | TEXT(255) | F `Expediente.cls:864` | `expedientes.TbSuministradores.Nemotecnico` | VARCHAR(255) | propuesta |
| `TbSuministradores.Direccion` | MEMO | F `SuministradorOperaciones.cls:109` | `expedientes.TbSuministradores.Direccion` | TEXT | propuesta |
| `TbSuministradores.CP` | TEXT(255) | F `SuministradorOperaciones.cls:112` | `expedientes.TbSuministradores.CP` | VARCHAR(255) | propuesta |
| `TbSuministradores.Ciudad` | TEXT(255) | F `SuministradorOperaciones.cls:115` | `expedientes.TbSuministradores.Ciudad` | VARCHAR(255) | propuesta |
| `TbSuministradores.ConsorcioPropio` | TEXT(2) | S `ExpedienteSuministradorServicio.cls:297` | `expedientes.TbSuministradores.ConsorcioPropio` | VARCHAR(2) · Sí/No como texto | propuesta |
