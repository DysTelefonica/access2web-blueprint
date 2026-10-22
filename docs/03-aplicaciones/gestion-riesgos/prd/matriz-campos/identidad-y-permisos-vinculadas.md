# Identidad y permisos (vinculadas a Lanzadera) — matrices campo a campo

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Issue:** #1079 · Fase 3 · **Matriz:** `matriz-campos.md` (índice, tanda posterior)

3 tablas. Cada fila está en estado **propuesta**: la topología PostgreSQL global y el ERD de origen siguen abiertos (D-GR-4 del [README](../README.md)).

### TbAplicaciones

- Columnas: **0**. Filas: sin recuento en el esquema disponible.
- Evidencia en `src`: **0** fichero(s) de producción — ninguna: la tabla no aparece en el código, solo en el ERD.
- **Tabla vinculada**: aporta `Lanzadera_Datos.accdb`. No es propiedad de esta aplicación; ver `coexistencia.md` (tanda posterior).

| Columna | Tipo de origen | Clase de dato |
|---|---|---|

### TbUsuariosAplicaciones

- Columnas: **0**. Filas: sin recuento en el esquema disponible.
- Evidencia en `src`: **2** fichero(s) de producción — `Usuario.cls`, `Constructor.bas`
- **Tabla vinculada**: aporta `Lanzadera_Datos.accdb`. No es propiedad de esta aplicación; ver `coexistencia.md` (tanda posterior).

| Columna | Tipo de origen | Clase de dato |
|---|---|---|

### TbUsuariosAplicacionesPermisos

- Columnas: **0**. Filas: sin recuento en el esquema disponible.
- Evidencia en `src`: **2** fichero(s) de producción — `UsuarioAplicacionPermisos.cls`, `Constructor.bas`
- **Tabla vinculada**: aporta `Lanzadera_Datos.accdb`. No es propiedad de esta aplicación; ver `coexistencia.md` (tanda posterior).

| Columna | Tipo de origen | Clase de dato |
|---|---|---|
