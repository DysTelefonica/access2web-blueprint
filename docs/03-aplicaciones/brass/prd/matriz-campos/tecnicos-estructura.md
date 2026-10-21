# Matriz campo a campo — técnicos, estructura y configuración

> Parte de la [matriz campo a campo](../matriz-campos.md). Reglas de mapeo, leyenda y límites de la fuente en ese documento. Commit fijado `3f77d58`.

## Tb0FiltroGestion

5 columnas · 6551 filas en la extracción del snapshot · citada en 1 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `Tb0FiltroGestion.ID` | TEXT(255) → VARCHAR(255) | — | F `GestorInforme.cls:111` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `Tb0FiltroGestion.Fecha` | TEXT(255) → VARCHAR(255) | — | F `GestorInforme.cls:112` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `Tb0FiltroGestion.Datos` | MEMO → TEXT | — | F `GestorInforme.cls:114` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `Tb0FiltroGestion.Usuario` | TEXT(255) → VARCHAR(255) | personal | F `GestorInforme.cls:117` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `Tb0FiltroGestion.Visto` | TEXT(2) → VARCHAR(2) | — | F `GestorInforme.cls:118` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbBUI

1 columnas · 18 filas en la extracción del snapshot · citada en 4 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbBUI.BUI` | TEXT(255) → VARCHAR(255) | — | F `Form_FormBUIGestion.cls:212` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbHerramientaDocAyuda

2 columnas · 4 filas en la extracción del snapshot · citada en 1 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbHerramientaDocAyuda.NombreFormulario` | TEXT(255) → VARCHAR(255) | personal | S `Funciones Generales.bas:30` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbHerramientaDocAyuda.NombreArchivoAyuda` | TEXT(255) → VARCHAR(255) | personal | S `Funciones Generales.bas:15` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbNodoBUI

2 columnas · 19 filas en la extracción del snapshot · citada en 6 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbNodoBUI.NODO` | TEXT(50) → VARCHAR(50) | — | F `Form_FormEquipoAlta.cls:300` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbNodoBUI.BUI` | TEXT(50) → VARCHAR(50) | — | F `Form_FormBUIGestion.cls:212` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbNodos

1 columnas · 6 filas en la extracción del snapshot · citada en 1 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbNodos.NODO` | TEXT(50) → VARCHAR(50) | — | F `Form_FormNodoGestion.cls:216` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbSistema

1 columnas · 2 filas en la extracción del snapshot · citada en 0 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbSistema.SISTEMA` | TEXT(50) → VARCHAR(50) | — | — | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbSubsistemaBui

2 columnas · 162 filas en la extracción del snapshot · citada en 11 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbSubsistemaBui.SUBSISTEMA` | TEXT(255) → VARCHAR(255) | — | F `Form_FormEquipoAlta.cls:302` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbSubsistemaBui.BUI` | TEXT(50) → VARCHAR(50) | — | F `Form_FormEquipoAlta.cls:301` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbTecnicos

8 columnas · 115 filas en la extracción del snapshot · citada en 6 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbTecnicos.ALIAS` | TEXT(50) → VARCHAR(50) | personal | F `Form_FormTecnicoGestion.cls:549` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbTecnicos.TIPO` | TEXT(50) → VARCHAR(50) | — | F `Form_FormTecnicoGestion.cls:550` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbTecnicos.NOMBRE` | TEXT(255) → VARCHAR(255) | personal | F `Form_FormTecnicoGestion.cls:549` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbTecnicos.TELEFONO` | TEXT(50) → VARCHAR(50) | personal | F `Tecnico.cls:274` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbTecnicos.EMAIL` | TEXT(255) → VARCHAR(255) | personal | F `Tecnico.cls:278` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbTecnicos.FECHAALTA` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Form_FormTecnicoGestion.cls:550` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbTecnicos.FECHABAJA` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Form_FormTecnicoGestion.cls:551` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbTecnicos.DNI` | TEXT(255) → VARCHAR(255) | personal | F `Tecnico.cls:287` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbTecnicosAusencias

6 columnas · 1236 filas en la extracción del snapshot · citada en 9 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbTecnicosAusencias.IDAusencia` | LONG(4) → INTEGER | — | F `LIbranza.cls:229` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbTecnicosAusencias.Alias` | TEXT(50) → VARCHAR(50) | personal | F `Form_FormTecnicoAusenciasGestion.cls:45` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbTecnicosAusencias.FechaLibranza` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Form_FormTecnicoAusenciasGestion.cls:19` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbTecnicosAusencias.Observaciones` | MEMO → TEXT | — | F `LIbranza.cls:233` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbTecnicosAusencias.Motivo` | TEXT(255) → VARCHAR(255) | — | F `Form_FormTecnicoAusencias.cls:205` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbTecnicosAusencias.horas` | DOUBLE(8) → DOUBLE PRECISION | — | F `LIbranza.cls:236` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbTecnicosFiestas

2 columnas · 104 filas en la extracción del snapshot · citada en 8 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbTecnicosFiestas.FechaFiesta` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Form_FormTecnicoFiestasGestion.cls:18` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbTecnicosFiestas.Observaciones` | MEMO → TEXT | — | F `Fiesta.cls:95` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbTipoTecnico

1 columnas · 10 filas en la extracción del snapshot · citada en 4 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbTipoTecnico.TIPO` | TEXT(255) → VARCHAR(255) | — | F `Tecnico.cls:271` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbTipoTecnicoPrecios

6 columnas · 19 filas en la extracción del snapshot · citada en 5 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbTipoTecnicoPrecios.IDTipoTecnicoFecha` | LONG(4) → INTEGER | personal | F `Form_FormTipoTecnicoPreciosGestion.cls:352` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbTipoTecnicoPrecios.TIPOTECNICO` | TEXT(50) → VARCHAR(50) | personal | F `Form_FormTecnico.cls:121` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbTipoTecnicoPrecios.FechaInicial` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Form_FormTipoTecnicoPreciosGestion.cls:353` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbTipoTecnicoPrecios.FechaFinal` | SHORT_DATE_TIME(8) → TIMESTAMP | — | F `Form_FormTipoTecnicoPreciosGestion.cls:353` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbTipoTecnicoPrecios.PrecioHoraLab` | DOUBLE(8) → DOUBLE PRECISION | económico | F `Form_FormTipoTecnicoPreciosGestion.cls:354` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
| `TbTipoTecnicoPrecios.PrecioHoraExtra` | DOUBLE(8) → DOUBLE PRECISION | económico | F `Form_FormTipoTecnicoPreciosGestion.cls:355` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |

## TbUbicacion

1 columnas · 3 filas en la extracción del snapshot · citada en 3 ficheros de `src` (sin pruebas).

| Origen | Tipo origen | Clase de dato | Evidencia en `src` | Destino (esquema.tabla.columna) | Estado |
|---|---|---|---|---|---|
| `TbUbicacion.UBICACION` | TEXT(50) → VARCHAR(50) | — | F `Form_FormActividadAlta.cls:136` | pendiente: esquema web de Brass sin definir | propuesta (tipo); destino pendiente |
