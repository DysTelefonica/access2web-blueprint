# Matriz campo a campo — equipos, calibraciones y materiales

> Parte de la [matriz campo a campo](../matriz-campos.md). Reglas de mapeo, leyenda y límites de la fuente en ese documento. Commit fijado `3f77d58`.

## TbAuxMateriales

14 columnas · 0 filas en la extracción del snapshot · citada en 1 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbAuxMateriales.NODO` | TEXT(255) → VARCHAR(255) | — | F `Funciones Generales.bas:3438` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxMateriales.BUI` | TEXT(255) → VARCHAR(255) | — | F `Funciones Generales.bas:3439` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxMateriales.SUBSISTEMA` | TEXT(255) → VARCHAR(255) | — | F `Funciones Generales.bas:3440` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxMateriales.EQUIPO` | TEXT(255) → VARCHAR(255) | — | F `Funciones Generales.bas:3441` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxMateriales.IDEVENTO` | TEXT(255) → VARCHAR(255) | — | F `Funciones Generales.bas:3355` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxMateriales.TIPOACCION` | TEXT(255) → VARCHAR(255) | — | F `Funciones Generales.bas:3582` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxMateriales.PN` | TEXT(255) → VARCHAR(255) | — | F `Funciones Generales.bas:3583` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxMateriales.NS` | TEXT(255) → VARCHAR(255) | — | F `Funciones Generales.bas:3584` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxMateriales.MATERIAL` | TEXT(255) → VARCHAR(255) | — | F `Funciones Generales.bas:3585` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxMateriales.FECHAENTREGA` | TEXT(50) → VARCHAR(50) | — | F `Funciones Generales.bas:3586` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxMateriales.COSTE` | TEXT(50) → VARCHAR(50) | económico | F `Funciones Generales.bas:3587` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxMateriales.PrecioSinIVA` | TEXT(50) → VARCHAR(50) | económico | — | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxMateriales.REPARADOPOR` | TEXT(255) → VARCHAR(255) | — | F `Funciones Generales.bas:3589` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbAuxMateriales.CODREGISTRODIFERENCIA` | MEMO → TEXT | — | — | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbEquipos

7 columnas · 806 filas en la extracción del snapshot · citada en 22 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbEquipos.IDEquipo` | LONG(4) → INTEGER | — | F `Form_FormEquipoAlta.cls:299` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEquipos.NODO` | TEXT(50) → VARCHAR(50) | — | F `Form_FormEquipoAlta.cls:300` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEquipos.BUI` | TEXT(50) → VARCHAR(50) | — | F `Form_FormEquipoAlta.cls:301` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEquipos.SUBSISTEMA` | TEXT(50) → VARCHAR(50) | — | F `Form_FormEquipoAlta.cls:302` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEquipos.Equipo` | TEXT(255) → VARCHAR(255) | — | F `Form_FormEquipoAlta.cls:303` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEquipos.Descripcion` | MEMO → TEXT | — | F `Form_FormEquipoAlta.cls:305` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEquipos.FechaObsoleto` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Form_FormEquipoEdicion.cls:565` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbEquiposCalibrables

7 columnas · 27 filas en la extracción del snapshot · citada en 6 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbEquiposCalibrables.IDEquipoCalibrable` | LONG(4) → INTEGER | — | F `Form_FormCalibracionAlta.cls:201` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEquiposCalibrables.Equipo` | TEXT(255) → VARCHAR(255) | — | F `Form_FormCalibracionAlta.cls:85` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEquiposCalibrables.Modelo` | TEXT(255) → VARCHAR(255) | — | F `Form_FormCalibracionAlta.cls:86` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEquiposCalibrables.NS` | TEXT(50) → VARCHAR(50) | — | F `Form_FormCalibracionAlta.cls:87` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEquiposCalibrables.BUI` | TEXT(50) → VARCHAR(50) | — | F `Form_FormCalibracionAlta.cls:88` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEquiposCalibrables.Observaciones` | MEMO → TEXT | — | F `Form_FormCalibracionEquipoAlta.cls:166` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEquiposCalibrables.FechaBajaParaCalibracion` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Form_FormCalibracionEquipoEdicion.cls:29` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbEquiposCalibrablesFechas

6 columnas · 94 filas en la extracción del snapshot · citada en 5 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbEquiposCalibrablesFechas.IDCalibracion` | LONG(4) → INTEGER | — | F `Form_FormCalibracionAlta.cls:200` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEquiposCalibrablesFechas.IDEquipoCalibrable` | LONG(4) → INTEGER | — | F `Form_FormCalibracionAlta.cls:201` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEquiposCalibrablesFechas.FechaCalibrado` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Form_FormCalibracionAlta.cls:202` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEquiposCalibrablesFechas.FechaFinCalibrado` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Form_FormCalibracionAlta.cls:203` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEquiposCalibrablesFechas.URLDocumentosAnexados` | TEXT(255) → VARCHAR(255) | — | S `Form_FormCalibracionGestion.cls:419` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEquiposCalibrablesFechas.FechaComunicacionCliente` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Form_FormCalibracionEdicion.cls:28` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbEquiposMedida

9 columnas · 14 filas en la extracción del snapshot · citada en 2 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbEquiposMedida.IDEquipoMedida` | LONG(4) → INTEGER | — | F `EquipoMedida.cls:377` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEquiposMedida.Nombre` | TEXT(255) → VARCHAR(255) | personal | F `EquipoMedida.cls:394` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEquiposMedida.NS` | TEXT(255) → VARCHAR(255) | — | F `EquipoMedida.cls:397` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEquiposMedida.PN` | TEXT(255) → VARCHAR(255) | — | F `EquipoMedida.cls:400` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEquiposMedida.FechaInicioServicio` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `EquipoMedida.cls:385` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEquiposMedida.FechaFinServicio` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `EquipoMedida.cls:382` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEquiposMedida.Marca` | TEXT(255) → VARCHAR(255) | — | F `EquipoMedida.cls:388` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEquiposMedida.Modelo` | TEXT(255) → VARCHAR(255) | — | F `EquipoMedida.cls:391` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEquiposMedida.Descripcion` | MEMO → TEXT | — | F `EquipoMedida.cls:379` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbEquiposMedidaCalibraciones

6 columnas · 28 filas en la extracción del snapshot · citada en 2 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbEquiposMedidaCalibraciones.IDCalibracion` | LONG(4) → INTEGER | — | F `EquipoMedidaCalibracion.cls:137` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEquiposMedidaCalibraciones.IDEquipoMedida` | LONG(4) → INTEGER | — | F `EquipoMedidaCalibracion.cls:138` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEquiposMedidaCalibraciones.FechaCalibracion` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `EquipoMedidaCalibracion.cls:139` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEquiposMedidaCalibraciones.FechaFinCalibracion` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `EquipoMedidaCalibracion.cls:140` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEquiposMedidaCalibraciones.EmpresaCalibradora` | TEXT(255) → VARCHAR(255) | — | F `EquipoMedidaCalibracion.cls:141` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbEquiposMedidaCalibraciones.Observaciones` | MEMO → TEXT | — | F `EquipoMedidaCalibracion.cls:143` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbMaterial

15 columnas · 2351 filas en la extracción del snapshot · citada en 14 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbMaterial.IDMaterial` | TEXT(255) → VARCHAR(255) | — | F `Form_FormEventoCopia.cls:241` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbMaterial.IDEquipo` | LONG(4) → INTEGER | — | F `Form_FormEventosAltaMasivaPorNS.cls:580` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbMaterial.FechaIncio` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Form_FormEventoCopia.cls:255` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbMaterial.Garantia` | TEXT(50) → VARCHAR(50) | — | F `Form_FormEventoCopia.cls:256` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbMaterial.IDActividad` | TEXT(50) → VARCHAR(50) | — | F `Form_FormEventoCopia.cls:196` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbMaterial.TipoAccion` | TEXT(255) → VARCHAR(255) | — | F `Form_FormEventoCopia.cls:250` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbMaterial.Material` | TEXT(255) → VARCHAR(255) | — | F `Form_FormEventoCopia.cls:248` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbMaterial.PN` | TEXT(50) → VARCHAR(50) | — | F `Form_FormEventoCopia.cls:251` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbMaterial.NS` | TEXT(50) → VARCHAR(50) | — | F `Form_FormEventoCopia.cls:252` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbMaterial.COSTE` | DOUBLE(8) → DOUBLE PRECISION | económico | F `Form_FormEventoCopia.cls:249` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbMaterial.ReparadoPor` | TEXT(255) → VARCHAR(255) | — | F `Form_FormEventoCopia.cls:253` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbMaterial.FechaEntrega` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Form_FormEventoCopia.cls:258` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbMaterial.Descripcion` | MEMO → TEXT | — | F `Form_FormEventoCopia.cls:254` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbMaterial.CodImportacion` | TEXT(50) → VARCHAR(50) | — | S `Form_FormEventoCopia.cls:167` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbMaterial.EsReparacion` | TEXT(2) → VARCHAR(2) | — | F `Form_FormEventosAltaMasivaPorNS.cls:439` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbMaterialSeguimiento

11 columnas · 2049 filas en la extracción del snapshot · citada en 6 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbMaterialSeguimiento.IDSeguimiento` | TEXT(50) → VARCHAR(50) | — | F `Form_FormEventoCopia.cls:436` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbMaterialSeguimiento.IDMaterial` | TEXT(50) → VARCHAR(50) | — | F `Form_FormEventoCopia.cls:241` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbMaterialSeguimiento.FechaInicioIntervencion` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Form_FormEventoCopia.cls:438` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbMaterialSeguimiento.Destino` | TEXT(255) → VARCHAR(255) | — | F `Form_FormEventoCopia.cls:439` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbMaterialSeguimiento.Motivo` | MEMO → TEXT | — | F `Form_FormEventoCopia.cls:440` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbMaterialSeguimiento.FechaFinIntervencion` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Form_FormEventoCopia.cls:441` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbMaterialSeguimiento.Resultado` | MEMO → TEXT | — | F `Form_FormEventoCopia.cls:442` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbMaterialSeguimiento.UltimaIntervencion` | TEXT(2) → VARCHAR(2) | — | F `Form_FormEventoCopia.cls:443` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbMaterialSeguimiento.TituloAnexo` | TEXT(255) → VARCHAR(255) | — | — | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbMaterialSeguimiento.DescripcionAnexo` | MEMO → TEXT | — | — | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbMaterialSeguimiento.NombreArchivoAnexo` | TEXT(255) → VARCHAR(255) | personal | — | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbRepuestosTipo

1 columnas · 3 filas en la extracción del snapshot · citada en 0 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbRepuestosTipo.TIPOREPUESTOS` | TEXT(255) → VARCHAR(255) | — | — | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
