# Matriz campo a campo — facturación, gastos y subcontratación

> Parte de la [matriz campo a campo](../matriz-campos.md). Reglas de mapeo, leyenda y límites de la fuente en ese documento. Commit fijado `3f77d58`.

## TbFacturaActividadesInvolucradas

18 columnas · 25107 filas en la extracción del snapshot · citada en 7 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbFacturaActividadesInvolucradas.IDFactura` | LONG(4) → INTEGER | — | F `Form_FormEventoGestion.cls:1241` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaActividadesInvolucradas.IDActividad` | TEXT(50) → VARCHAR(50) | — | F `Form_FormEventoGestion.cls:2777` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaActividadesInvolucradas.NODO` | TEXT(50) → VARCHAR(50) | — | F `Funciones Generales.bas:3438` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaActividadesInvolucradas.BUI` | TEXT(50) → VARCHAR(50) | — | F `Constructor.bas:2834` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaActividadesInvolucradas.SUBSISTEMA` | TEXT(255) → VARCHAR(255) | — | F `Funciones Generales.bas:3440` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaActividadesInvolucradas.EQUIPO` | TEXT(255) → VARCHAR(255) | — | F `Funciones Generales.bas:3441` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaActividadesInvolucradas.IDEVENTO` | TEXT(50) → VARCHAR(50) | — | F `Constructor.bas:288` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaActividadesInvolucradas.TIPOTECNICO` | TEXT(50) → VARCHAR(50) | personal | F `Funciones Generales.bas:1991` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaActividadesInvolucradas.ALIASTECNICO` | TEXT(50) → VARCHAR(50) | personal | F `Form_FormActividadGestion.cls:39` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaActividadesInvolucradas.PERFIL` | TEXT(50) → VARCHAR(50) | — | F `Factura.cls:2700` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaActividadesInvolucradas.PrecioHoraLaborablePerfil` | DOUBLE(8) → DOUBLE PRECISION | económico | F `Factura.cls:2701` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaActividadesInvolucradas.PrecioHoraExtraPerfil` | DOUBLE(8) → DOUBLE PRECISION | económico | F `Factura.cls:2702` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaActividadesInvolucradas.TIPOACTIVIDAD` | TEXT(255) → VARCHAR(255) | — | F `Factura.cls:2703` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaActividadesInvolucradas.DESCRIPCION` | MEMO → TEXT | — | F `Funciones Generales.bas:3450` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaActividadesInvolucradas.FECHAALTA` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Funciones Generales.bas:826` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaActividadesInvolucradas.HORASLABORABLES` | DOUBLE(8) → DOUBLE PRECISION | — | F `Form_FormActividadGestion.cls:154` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaActividadesInvolucradas.HORASEXTRAS` | DOUBLE(8) → DOUBLE PRECISION | — | F `Funciones Generales.bas:3448` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaActividadesInvolucradas.UBICACION` | TEXT(50) → VARCHAR(50) | — | F `Funciones Generales.bas:3449` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbFacturaEventosInvolucrados

21 columnas · 5514 filas en la extracción del snapshot · citada en 7 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbFacturaEventosInvolucrados.IDFactura` | LONG(4) → INTEGER | — | F `Form_FormEventoGestion.cls:1241` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaEventosInvolucrados.IDEvento` | TEXT(50) → VARCHAR(50) | — | F `Form_FormInformeSLA.cls:139` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaEventosInvolucrados.NODO` | TEXT(50) → VARCHAR(50) | — | F `Funciones Generales.bas:3438` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaEventosInvolucrados.BUI` | TEXT(50) → VARCHAR(50) | — | F `Form_FormInformeSLA.cls:145` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaEventosInvolucrados.SUBSISTEMA` | TEXT(255) → VARCHAR(255) | — | F `Form_FormInformeSLA.cls:148` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaEventosInvolucrados.EQUIPO` | TEXT(255) → VARCHAR(255) | — | F `Funciones Generales.bas:3441` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaEventosInvolucrados.PMPR` | TEXT(50) → VARCHAR(50) | — | F `Funciones Generales.bas:3727` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaEventosInvolucrados.TIPOEVENTO` | TEXT(50) → VARCHAR(50) | — | F `Funciones Generales.bas:3728` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaEventosInvolucrados.CRITICIDAD` | FLOAT(4) → DOUBLE PRECISION | — | F `Form_FormInformeSLA.cls:151` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaEventosInvolucrados.TIPOTECNICO` | TEXT(50) → VARCHAR(50) | personal | F `Funciones Generales.bas:1991` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaEventosInvolucrados.ORIGINADOR` | TEXT(50) → VARCHAR(50) | — | F `Funciones Generales.bas:3731` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaEventosInvolucrados.FECHAALTA` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Funciones Generales.bas:826` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaEventosInvolucrados.HORAALTA` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Funciones Generales.bas:3733` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaEventosInvolucrados.TIEMPORESPUESTA` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Funciones Generales.bas:3734` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaEventosInvolucrados.FECHAFIN` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Funciones Generales.bas:3735` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaEventosInvolucrados.HORAFIN` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Funciones Generales.bas:3736` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaEventosInvolucrados.CAUSAFIN` | MEMO → TEXT | — | F `Funciones Generales.bas:3737` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaEventosInvolucrados.CONTACTO` | TEXT(255) → VARCHAR(255) | personal | F `Funciones Generales.bas:3738` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaEventosInvolucrados.IDEVENTOSGENERADOS` | TEXT(255) → VARCHAR(255) | — | — | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaEventosInvolucrados.DESCRIPCION` | MEMO → TEXT | — | F `Funciones Generales.bas:3450` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaEventosInvolucrados.ParaRegularizar` | TEXT(255) → VARCHAR(255) | — | — | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbFacturaGastosInvolucrados

8 columnas · 519 filas en la extracción del snapshot · citada en 4 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbFacturaGastosInvolucrados.IDFactura` | LONG(4) → INTEGER | — | F `Funciones Generales.bas:6322` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaGastosInvolucrados.IDGasto` | LONG(4) → INTEGER | — | F `Factura.cls:2998` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaGastosInvolucrados.TipoGasto` | TEXT(255) → VARCHAR(255) | — | F `Factura.cls:3000` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaGastosInvolucrados.ImporteUnitario` | DOUBLE(8) → DOUBLE PRECISION | económico | F `Form_FormGastoGestion.cls:507` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaGastosInvolucrados.NumeroUnidades` | DOUBLE(8) → DOUBLE PRECISION | — | F `Form_FormGastoGestion.cls:507` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaGastosInvolucrados.FechaImputacionGasto` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Factura.cls:3010` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaGastosInvolucrados.ALIAS` | TEXT(50) → VARCHAR(50) | personal | F `Funciones Generales.bas:2828` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaGastosInvolucrados.Observaciones` | MEMO → TEXT | — | F `Factura.cls:3014` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbFacturaMaterialesInvolucrados

15 columnas · 1986 filas en la extracción del snapshot · citada en 6 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbFacturaMaterialesInvolucrados.IDFactura` | LONG(4) → INTEGER | — | F `Form_FormEventoGestion.cls:1241` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaMaterialesInvolucrados.IDMaterial` | TEXT(50) → VARCHAR(50) | — | F `Form_FormEventoGestion.cls:1389` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaMaterialesInvolucrados.NODO` | TEXT(50) → VARCHAR(50) | — | F `Funciones Generales.bas:3438` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaMaterialesInvolucrados.BUI` | TEXT(50) → VARCHAR(50) | — | F `Constructor.bas:2834` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaMaterialesInvolucrados.SUBSISTEMA` | TEXT(255) → VARCHAR(255) | — | F `Funciones Generales.bas:3440` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaMaterialesInvolucrados.EQUIPO` | TEXT(255) → VARCHAR(255) | — | F `Form_FormMaterialIntervencionEdicion.cls:373` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaMaterialesInvolucrados.IDEvento` | TEXT(50) → VARCHAR(50) | — | F `Form_FormMaterialIntervencionEdicion.cls:374` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaMaterialesInvolucrados.TIPOACCION` | TEXT(255) → VARCHAR(255) | — | F `Funciones Generales.bas:3582` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaMaterialesInvolucrados.Material` | TEXT(255) → VARCHAR(255) | — | F `Form_FormMaterialIntervencionEdicion.cls:375` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaMaterialesInvolucrados.PN` | TEXT(50) → VARCHAR(50) | — | F `Funciones Generales.bas:3583` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaMaterialesInvolucrados.NS` | TEXT(50) → VARCHAR(50) | — | F `Form_FormMaterialesGestion.cls:1256` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaMaterialesInvolucrados.DESCRIPCIONEQUIPO` | TEXT(255) → VARCHAR(255) | — | F `Factura.cls:2915` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaMaterialesInvolucrados.FECHAENTREGA` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Form_FormMaterialesGestion.cls:1255` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaMaterialesInvolucrados.PRECIOSINIVA` | DOUBLE(8) → DOUBLE PRECISION | económico | F `Factura.cls:2924` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaMaterialesInvolucrados.ReparadoPor` | TEXT(255) → VARCHAR(255) | — | F `Funciones Generales.bas:3589` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbFacturaPrincipal

19 columnas · 53 filas en la extracción del snapshot · citada en 4 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbFacturaPrincipal.IDFactura` | LONG(4) → INTEGER | — | F `Factura.cls:2468` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaPrincipal.FechaFinalEventos` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Factura.cls:2469` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaPrincipal.FechaFactura` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Factura.cls:2470` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaPrincipal.Titulo1` | TEXT(255) → VARCHAR(255) | — | F `Factura.cls:2476` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaPrincipal.Titulo2` | TEXT(255) → VARCHAR(255) | — | F `Factura.cls:2479` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaPrincipal.Titulo3` | TEXT(255) → VARCHAR(255) | — | F `Factura.cls:2482` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaPrincipal.TipoImpositivo` | DOUBLE(8) → DOUBLE PRECISION | — | F `Factura.cls:2488` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaPrincipal.Recargo` | DOUBLE(8) → DOUBLE PRECISION | — | F `Factura.cls:2495` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaPrincipal.PrecioHoras` | DOUBLE(8) → DOUBLE PRECISION | económico | F `Factura.cls:2498` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaPrincipal.PrecioSuministro` | DOUBLE(8) → DOUBLE PRECISION | económico | F `Factura.cls:2501` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaPrincipal.PrecioSuministroConRecargo` | DOUBLE(8) → DOUBLE PRECISION | económico | F `Factura.cls:2504` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaPrincipal.PrecioDietasYGastos` | DOUBLE(8) → DOUBLE PRECISION | económico | F `Factura.cls:2507` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaPrincipal.PrecioIncurrido` | DOUBLE(8) → DOUBLE PRECISION | económico | F `Factura.cls:2510` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaPrincipal.PrecioSubcontrataciones` | DOUBLE(8) → DOUBLE PRECISION | económico | F `Factura.cls:2513` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaPrincipal.PrecioSubcontratacionesConRecargo` | DOUBLE(8) → DOUBLE PRECISION | económico | F `Factura.cls:2516` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaPrincipal.PrecioBaseImponible` | DOUBLE(8) → DOUBLE PRECISION | económico | F `Factura.cls:2519` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaPrincipal.PrecioImpuestos` | DOUBLE(8) → DOUBLE PRECISION | económico | F `Factura.cls:2522` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaPrincipal.ImporteAFacturar` | DOUBLE(8) → DOUBLE PRECISION | económico | F `Factura.cls:2525` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaPrincipal.NombreFacturaAnexa` | TEXT(255) → VARCHAR(255) | personal | F `Factura.cls:2531` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbFacturaPrincipalPerfiles

7 columnas · 304 filas en la extracción del snapshot · citada en 2 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbFacturaPrincipalPerfiles.IDFacturaPerfil` | LONG(4) → INTEGER | — | F `Factura.cls:3290` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaPrincipalPerfiles.IDFactura` | LONG(4) → INTEGER | — | F `Factura.cls:2468` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaPrincipalPerfiles.Perfil` | TEXT(255) → VARCHAR(255) | — | F `Factura.cls:2700` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaPrincipalPerfiles.PrecioHoraLab` | DOUBLE(8) → DOUBLE PRECISION | económico | F `Factura.cls:3293` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaPrincipalPerfiles.PrecioHoraExt` | DOUBLE(8) → DOUBLE PRECISION | económico | F `Factura.cls:3294` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaPrincipalPerfiles.NHorasLab` | DOUBLE(8) → DOUBLE PRECISION | — | F `Factura.cls:3295` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaPrincipalPerfiles.NHorasExt` | DOUBLE(8) → DOUBLE PRECISION | — | F `Factura.cls:3296` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbFacturaSubcontratacionesInvolucradas

7 columnas · 74 filas en la extracción del snapshot · citada en 4 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbFacturaSubcontratacionesInvolucradas.IDFactura` | LONG(4) → INTEGER | — | F `Funciones Generales.bas:6322` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaSubcontratacionesInvolucradas.IDSubcontratacion` | LONG(4) → INTEGER | — | F `Factura.cls:3068` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaSubcontratacionesInvolucradas.FechaAlta` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Funciones Generales.bas:826` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaSubcontratacionesInvolucradas.FechaFin` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Funciones Generales.bas:3735` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaSubcontratacionesInvolucradas.Empresa` | TEXT(255) → VARCHAR(255) | — | F `Factura.cls:3074` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaSubcontratacionesInvolucradas.Descripcion` | MEMO → TEXT | — | F `Funciones Generales.bas:3450` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbFacturaSubcontratacionesInvolucradas.Importe` | DOUBLE(8) → DOUBLE PRECISION | económico | F `Factura.cls:3081` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbGastos

7 columnas · 519 filas en la extracción del snapshot · citada en 6 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbGastos.IDGasto` | LONG(4) → INTEGER | — | F `Form_FormGastoAlta.cls:225` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbGastos.TipoGasto` | TEXT(255) → VARCHAR(255) | — | F `Form_FormGastoAlta.cls:226` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbGastos.ImporteUnitario` | DOUBLE(8) → DOUBLE PRECISION | económico | F `Form_FormGastoAlta.cls:92` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbGastos.NumeroUnidades` | DOUBLE(8) → DOUBLE PRECISION | — | F `Form_FormGastoAlta.cls:228` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbGastos.FechaImputacionGasto` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Form_FormGastoAlta.cls:229` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbGastos.ALIAS` | TEXT(50) → VARCHAR(50) | personal | F `Form_FormGastoAlta.cls:231` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbGastos.Observaciones` | MEMO → TEXT | — | F `Form_FormGastoAlta.cls:234` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbGastosImportePorTipo

6 columnas · 3 filas en la extracción del snapshot · citada en 2 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbGastosImportePorTipo.TipoGasto` | TEXT(255) → VARCHAR(255) | — | F `Form_FormGastoAlta.cls:226` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbGastosImportePorTipo.FechaValorInicial` | SHORT_DATE_TIME(8) → TIMESTAMP | — | — | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbGastosImportePorTipo.FechaValorFinal` | SHORT_DATE_TIME(8) → TIMESTAMP | — | — | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbGastosImportePorTipo.ImporteUnitario` | DOUBLE(8) → DOUBLE PRECISION | económico | F `Form_FormGastoAlta.cls:92` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbGastosImportePorTipo.FechaFinalImporte` | TEXT(50) → VARCHAR(50) | económico | S `Form_FormGastoAlta.cls:88` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbGastosImportePorTipo.Observaciones` | MEMO → TEXT | — | F `Form_FormGastoAlta.cls:234` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbGuiaConciliaciones

7 columnas · 0 filas en la extracción del snapshot · citada en 0 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbGuiaConciliaciones.IDGuia` | LONG(4) → INTEGER | — | — | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbGuiaConciliaciones.Parte` | TEXT(255) → VARCHAR(255) | — | — | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbGuiaConciliaciones.Nodo` | TEXT(255) → VARCHAR(255) | — | — | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbGuiaConciliaciones.BUI` | TEXT(255) → VARCHAR(255) | — | — | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbGuiaConciliaciones.SubSistema` | TEXT(255) → VARCHAR(255) | — | — | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbGuiaConciliaciones.EquipoID` | TEXT(255) → VARCHAR(255) | — | — | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbGuiaConciliaciones.IDEvento` | TEXT(255) → VARCHAR(255) | — | — | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbSubcontrataciones

7 columnas · 75 filas en la extracción del snapshot · citada en 4 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbSubcontrataciones.IDSubcontratacion` | LONG(4) → INTEGER | — | F `Form_FormSubContratacionAlta.cls:131` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbSubcontrataciones.FechaAlta` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Form_FormSubContratacionAlta.cls:132` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbSubcontrataciones.FechaFin` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Form_FormSubContratacionAlta.cls:138` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbSubcontrataciones.Empresa` | TEXT(255) → VARCHAR(255) | — | F `Form_FormSubContratacionAlta.cls:133` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbSubcontrataciones.Descripcion` | MEMO → TEXT | — | F `Form_FormSubContratacionAlta.cls:135` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbSubcontrataciones.Importe` | DOUBLE(8) → DOUBLE PRECISION | económico | F `Form_FormSubContratacionAlta.cls:141` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbSubcontrataciones.Observaciones` | MEMO → TEXT | — | — | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
