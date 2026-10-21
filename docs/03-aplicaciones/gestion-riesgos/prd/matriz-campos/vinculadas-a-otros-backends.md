# Vinculadas a otros backends — matrices campo a campo

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Issue:** #1079 · Fase 3 · **Matriz:** [matriz-campos.md](../matriz-campos.md)

12 tablas. Cada fila está en estado **propuesta**: la topología PostgreSQL global y el ERD de origen siguen abiertos (D-GR-4 del [README](../README.md)).

### TbDPD

- Columnas: **0**. Filas: sin recuento en el esquema disponible.
- Evidencia en `src`: **0** fichero(s) de producción — ninguna: la tabla no aparece en el código, solo en el ERD.
- **Tabla vinculada**: aporta `AGEDYS_DATOS.accdb`. No es propiedad de esta aplicación; ver `coexistencia.md` (tanda posterior).

| Columna | Tipo de origen | Clase de dato |
|---|---|---|

### TbDocumentos

- Columnas: **0**. Filas: sin recuento en el esquema disponible.
- Evidencia en `src`: **2** fichero(s) de producción — `Documento.cls`, `Constructor.bas`
- **Tabla vinculada**: aporta `AGEDO20_Datos.accdb`. No es propiedad de esta aplicación; ver `coexistencia.md` (tanda posterior).

| Columna | Tipo de origen | Clase de dato |
|---|---|---|

### TbDocumentosID

- Columnas: **0**. Filas: sin recuento en el esquema disponible.
- Evidencia en `src`: **0** fichero(s) de producción — ninguna: la tabla no aparece en el código, solo en el ERD.
- **Tabla vinculada**: aporta `AGEDO20_Datos.accdb`. No es propiedad de esta aplicación; ver `coexistencia.md` (tanda posterior).

| Columna | Tipo de origen | Clase de dato |
|---|---|---|

### TbExpedientes

- Columnas: **0**. Filas: sin recuento en el esquema disponible.
- Evidencia en `src`: **1** fichero(s) de producción — `Constructor.bas`
- **Tabla vinculada**: aporta `AGEDYS_DATOS.accdb`. No es propiedad de esta aplicación; ver `coexistencia.md` (tanda posterior).

| Columna | Tipo de origen | Clase de dato |
|---|---|---|

### TbExpedientes1

- Columnas: **0**. Filas: sin recuento en el esquema disponible.
- Evidencia en `src`: **1** fichero(s) de producción — `Constructor.bas`
- **Tabla vinculada**: aporta `Expedientes_datos.accdb`. No es propiedad de esta aplicación; ver `coexistencia.md` (tanda posterior).

| Columna | Tipo de origen | Clase de dato |
|---|---|---|

### TbExpedientesResponsables

- Columnas: **0**. Filas: sin recuento en el esquema disponible.
- Evidencia en `src`: **2** fichero(s) de producción — `ExpedienteResponsable.cls`, `Constructor.bas`
- **Tabla vinculada**: aporta `Expedientes_datos.accdb`. No es propiedad de esta aplicación; ver `coexistencia.md` (tanda posterior).

| Columna | Tipo de origen | Clase de dato |
|---|---|---|

### TbExpedientesSuministradores

- Columnas: **0**. Filas: sin recuento en el esquema disponible.
- Evidencia en `src`: **1** fichero(s) de producción — `Constructor.bas`
- **Tabla vinculada**: aporta `Expedientes_datos.accdb`. No es propiedad de esta aplicación; ver `coexistencia.md` (tanda posterior).

| Columna | Tipo de origen | Clase de dato |
|---|---|---|

### TbNoConformidades

- Columnas: **0**. Filas: sin recuento en el esquema disponible.
- Evidencia en `src`: **4** fichero(s) de producción — `NC.cls`, `FormRiesgoNC.form.txt`, `Form_FormRiesgoNC.cls`, `Constructor.bas`
- **Tabla vinculada**: aporta `NoConformidades_Datos.accdb`. No es propiedad de esta aplicación; ver `coexistencia.md` (tanda posterior).

| Columna | Tipo de origen | Clase de dato |
|---|---|---|

### TbProyectos1

- Columnas: **0**. Filas: sin recuento en el esquema disponible.
- Evidencia en `src`: **0** fichero(s) de producción — ninguna: la tabla no aparece en el código, solo en el ERD.
- **Tabla vinculada**: aporta `Gestion_Riesgos_Datos.accdb`. No es propiedad de esta aplicación; ver `coexistencia.md` (tanda posterior).

| Columna | Tipo de origen | Clase de dato |
|---|---|---|

### TbResponsablesExpedientes

- Columnas: **0**. Filas: sin recuento en el esquema disponible.
- Evidencia en `src`: **0** fichero(s) de producción — ninguna: la tabla no aparece en el código, solo en el ERD.
- **Tabla vinculada**: aporta `AGEDYS_DATOS.accdb`. No es propiedad de esta aplicación; ver `coexistencia.md` (tanda posterior).

| Columna | Tipo de origen | Clase de dato |
|---|---|---|

### TbSuministradores

- Columnas: **0**. Filas: sin recuento en el esquema disponible.
- Evidencia en `src`: **1** fichero(s) de producción — `Constructor.bas`
- **Tabla vinculada**: aporta `Expedientes_datos.accdb`. No es propiedad de esta aplicación; ver `coexistencia.md` (tanda posterior).

| Columna | Tipo de origen | Clase de dato |
|---|---|---|

### TbSuministradoresSAP

- Columnas: **0**. Filas: sin recuento en el esquema disponible.
- Evidencia en `src`: **0** fichero(s) de producción — ninguna: la tabla no aparece en el código, solo en el ERD.
- **Tabla vinculada**: aporta `AGEDYS_DATOS.accdb`. No es propiedad de esta aplicación; ver `coexistencia.md` (tanda posterior).

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
