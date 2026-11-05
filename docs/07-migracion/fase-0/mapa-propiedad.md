# Mapa de propiedad de tablas y vínculos entre backends

Fuente: la [copia de referencia](copia-referencia.md), leída con DAO sobre copias temporales. Una tabla **propia** vive físicamente en su backend; un **vínculo** es una referencia a la tabla propia de otro backend. Toda propuesta que cambie una tabla (clave, relación, cuarentena, borrado) se aplica en el backend que la tiene como propia; las que actúan sobre un vínculo o sobre el código se aplican en el consumidor. El listado completo tabla a tabla se regenera con `scripts/dysflow-erd/inventory-data.ps1 -Mode Ownership`.

## Resumen por backend

| Backend | Tablas propias | Vínculos que tiene a otros backends | Vínculos que recibe de otros backends |
|---|---:|---:|---:|
| `AGEDYS_DATOS` | 102 | 34 | 25 |
| `AGEDO20_Datos` | 66 | 2 | 5 |
| `Gestion_Riesgos_Datos` | 64 | 14 | 7 |
| `Gestion_Brass_Gestion_Datos` | 57 | 0 | 0 |
| `Expedientes_datos` | 47 | 23 | 45 |
| `NoConformidades_Datos` | 36 | 11 | 3 |
| `Seguridad_datos` | 36 | 0 | 0 |
| `Lanzadera_Datos` | 35 | 1 | 20 |
| `GestionContratos_datos` | 20 | 0 | 0 |
| `HPST` | 18 | 9 | 9 |
| `condor_datos` | 15 | 10 | 0 |
| `Registro_Ent_Salida_Datos` | 11 | 1 | 2 |
| `Solicitudes_HPS_datos` | 11 | 9 | 2 |
| `Registro_Datos` | 5 | 0 | 2 |
| `SICA_datos` | 4 | 0 | 0 |
| `Tareas_datos1` | 4 | 13 | 0 |
| `Control_Cambios_datos` | 3 | 0 | 0 |
| `Correos_datos` | 1 | 0 | 7 |

Total: 535 tablas propias y 127 vínculos. Ningún vínculo apunta a otro vínculo.

## Qué backend vincula a cuál

| Backend que vincula | Backend dueño | Vínculos | Tablas |
|---|---|---:|---|
| `AGEDO20_Datos` | `Lanzadera_Datos` | 2 | `TbUsuariosAplicaciones`, `tbUsuarios` |
| `AGEDYS_DATOS` | `AGEDO20_Datos` | 3 | `TbDocExpProceso`, `TbDocumentos`, `TbDocumentosID` |
| `AGEDYS_DATOS` | `Correos_datos` | 1 | `TbCorreosEnviados` |
| `AGEDYS_DATOS` | `Expedientes_datos` | 23 | `TbCPV`, `TbComerciales`, `TbEjercitos`, `TbEstados`, `TbExpedientes`, `TbExpedientesAnexos` y 17 más |
| `AGEDYS_DATOS` | `Lanzadera_Datos` | 4 | `TbAplicaciones`, `TbConexionesRegistro`, `TbUsuariosAplicaciones`, `TbUsuariosAplicacionesPermisos` |
| `AGEDYS_DATOS` | `Registro_Datos` | 1 | `Salidas` |
| `AGEDYS_DATOS` | `Registro_Ent_Salida_Datos` | 2 | `TbEntradas_RES`, `TbSalidas_RES` |
| `Expedientes_datos` | `AGEDYS_DATOS` | 14 | `TbComunicados`, `TbDpDInformeCondicionamiento`, `TbExpedientesAnualidades`, `TbFacturasDetalle`, `TbFacturasPrincipal`, `TbGTVContratos` y 8 más |
| `Expedientes_datos` | `Correos_datos` | 1 | `TbCorreosEnviados` |
| `Expedientes_datos` | `Gestion_Riesgos_Datos` | 3 | `TbProyectos`, `TbProyectosEdiciones` |
| `Expedientes_datos` | `HPST` | 2 | `TbUsuarios`, `TbUsuariosHistoricos` |
| `Expedientes_datos` | `Lanzadera_Datos` | 2 | `TbUsuariosAplicaciones`, `TbUsuariosAplicacionesPermisos` |
| `Expedientes_datos` | `NoConformidades_Datos` | 1 | `TbNoConformidades` |
| `Gestion_Riesgos_Datos` | `AGEDO20_Datos` | 2 | `TbDocumentos`, `TbDocumentosID` |
| `Gestion_Riesgos_Datos` | `AGEDYS_DATOS` | 3 | `TbProyectos`, `TbResponsablesExpedientes`, `TbSuministradoresSAP` |
| `Gestion_Riesgos_Datos` | `Correos_datos` | 1 | `TbCorreosEnviados` |
| `Gestion_Riesgos_Datos` | `Expedientes_datos` | 4 | `TbExpedientes`, `TbExpedientesResponsables`, `TbExpedientesSuministradores`, `TbSuministradores` |
| `Gestion_Riesgos_Datos` | `Lanzadera_Datos` | 3 | `TbAplicaciones`, `TbUsuariosAplicaciones`, `TbUsuariosAplicacionesPermisos` |
| `Gestion_Riesgos_Datos` | `NoConformidades_Datos` | 1 | `TbNoConformidades` |
| `HPST` | `Correos_datos` | 1 | `TbCorreosEnviados` |
| `HPST` | `Expedientes_datos` | 4 | `TbExpedientes`, `TbExpedientesConEntidades`, `TbExpedientesLugaresEjecucion`, `TbSuministradores` |
| `HPST` | `Lanzadera_Datos` | 2 | `TbUsuariosAplicaciones`, `TbUsuariosAplicacionesPermisos` |
| `HPST` | `Solicitudes_HPS_datos` | 2 | `TbSolicitudes`, `TbSolicitudesFechas` |
| `Lanzadera_Datos` | `Correos_datos` | 1 | `TbCorreosEnviados` |
| `NoConformidades_Datos` | `Correos_datos` | 1 | `TbCorreosEnviados` |
| `NoConformidades_Datos` | `Expedientes_datos` | 4 | `TbExpedientes`, `TbExpedientesJuridicas`, `TbExpedientesResponsables`, `TbJuridicas` |
| `NoConformidades_Datos` | `Gestion_Riesgos_Datos` | 4 | `TbProyectos`, `TbProyectosEdiciones`, `TbRiesgos`, `TbRiesgosNC` |
| `NoConformidades_Datos` | `Lanzadera_Datos` | 2 | `TbUsuariosAplicaciones`, `TbUsuariosAplicacionesPermisos` |
| `Registro_Ent_Salida_Datos` | `Registro_Datos` | 1 | `Salidas` |
| `Solicitudes_HPS_datos` | `Expedientes_datos` | 2 | `TbExpedientes`, `TbExpedientesLugaresEjecucion` |
| `Solicitudes_HPS_datos` | `HPST` | 7 | `TbHPS`, `TbHPSEquivalencia`, `TbHPSGrado`, `TbMotivoHPS`, `TbUsuarios`, `TbUsuariosEntidades` y 1 más |
| `Tareas_datos1` | `AGEDYS_DATOS` | 8 | `TbDPDOfertaSuministrador`, `TbFacturasDetalle`, `TbNPedido`, `TbProyectos`, `TbSolicitudesOfertasPrevias`, `TbSuministradoresSAP` y 2 más |
| `Tareas_datos1` | `Expedientes_datos` | 2 | `TbExpedientes`, `TbExpedientesResponsables` |
| `Tareas_datos1` | `Lanzadera_Datos` | 3 | `TbUsuariosAplicaciones`, `TbUsuariosAplicacionesPermisos`, `TbUsuariosAplicacionesTareas` |
| `condor_datos` | `Correos_datos` | 1 | `TbCorreosEnviados` |
| `condor_datos` | `Expedientes_datos` | 6 | `TbExpedientes`, `TbExpedientesRACS`, `TbExpedientesResponsables`, `TbExpedientesSuministradores`, `TbRACS`, `TbSuministradores` |
| `condor_datos` | `Lanzadera_Datos` | 2 | `TbUsuariosAplicaciones`, `TbUsuariosAplicacionesPermisos` |
| `condor_datos` | `NoConformidades_Datos` | 1 | `TbNoConformidades` |

## Nombres de tabla propios en más de un backend

Mismo nombre, tablas distintas. Algunos son nombres genéricos sin relación entre sí; otros son copias de los mismos datos, que hay que resolver antes del traslado único porque en PostgreSQL cada tabla tiene un solo dueño.

| Tabla | Backends que la tienen como propia |
|---|---|
| `Copia de TbExpedientes` | `Expedientes_datos`, `Solicitudes_HPS_datos` |
| `Entradas` | `AGEDO20_Datos`, `Registro_Datos` |
| `Salidas` | `AGEDO20_Datos`, `Registro_Datos` |
| `Tb0HerramientaDocAyuda` | `Lanzadera_Datos`, `Seguridad_datos` |
| `TbAnexos` | `Gestion_Brass_Gestion_Datos`, `Gestion_Riesgos_Datos`, `NoConformidades_Datos` |
| `TbAuditorias` | `NoConformidades_Datos`, `Seguridad_datos` |
| `TbConexiones` | `AGEDYS_DATOS`, `Lanzadera_Datos`, `NoConformidades_Datos`, `Registro_Ent_Salida_Datos` |
| `TbConfiguracion` | `Gestion_Riesgos_Datos`, `Lanzadera_Datos`, `NoConformidades_Datos`, `Registro_Datos`, `Solicitudes_HPS_datos` |
| `TbCorreosEnviados` | `Correos_datos`, `GestionContratos_datos`, `Solicitudes_HPS_datos`, `Tareas_datos1` |
| `TbDatosEconomicosExpedientes` | `AGEDYS_DATOS`, `Expedientes_datos` |
| `TbExpedientesAnualidades` | `AGEDYS_DATOS`, `Expedientes_datos` |
| `TbExpedientesCodigoCompras` | `AGEDYS_DATOS`, `Expedientes_datos` |
| `TbHerramientaDocAyuda` | `AGEDYS_DATOS`, `Gestion_Brass_Gestion_Datos`, `Gestion_Riesgos_Datos`, `NoConformidades_Datos` |
| `TbHPSGrado` | `HPST`, `Solicitudes_HPS_datos` |
| `TbLog` | `AGEDO20_Datos`, `AGEDYS_DATOS`, `Gestion_Riesgos_Datos`, `NoConformidades_Datos`, `Registro_Ent_Salida_Datos`, `Seguridad_datos` |
| `TbNCAccionesRealizadas` | `AGEDO20_Datos`, `NoConformidades_Datos` |
| `TbProyectos` | `AGEDYS_DATOS`, `Gestion_Riesgos_Datos` |
| `TbTareas` | `Gestion_Riesgos_Datos`, `Tareas_datos1` |
| `TbTareasExplicaciones` | `Gestion_Riesgos_Datos`, `NoConformidades_Datos` |
| `TbUltimoCambio` | `Expedientes_datos`, `Solicitudes_HPS_datos` |

## Propuestas contrastadas con el mapa

Las fichas se escribieron por aplicación, y cada una propuso cambios sobre las tablas que ve su aplicación, incluidas las vinculadas. Se contrastaron las 159 propuestas con este mapa: 27 nombran tablas que no son propias del backend de su aplicación. Se clasifican por lo que hace la acción.

### Mal colocadas: cambian una tabla desde un backend que no es su dueño

| Propuesta | Problema | Corrección aplicada en el catálogo |
|---|---|---|
| `CON-C5` | Clave primaria de `TbCorreosEnviados`, propia de `Correos_datos`, desde Condor | Fusionada en `COR-C1`, del dueño (decisión 22) |
| `NOC-N3` | La misma clave primaria, desde No Conformidades | Fusionada en `COR-C1` |
| `EXP-E2` | Incluye `TbCorreosEnviados` (Correos) y `TbFacturasPrincipal` (AGEDYS) | Se quitan: están en `COR-C1` y en `AGT-A1` |
| `GRI-R2` | Incluye `TbCorreosEnviados` | Se quita: está en `COR-C1` |
| `NOC-N4` | Su cambio a identidad incluye `TbCorreosEnviados` | Esa tabla la cubre la decisión 22 |
| `EXP-E9` | Integridad referencial hacia `TbUsuariosAplicaciones`, propia de `Lanzadera_Datos` | Son dos ficheros: Access no puede declararla; solo se declara en PostgreSQL |

### «Tablas inexistentes» que existen en otro backend

| Propuesta | Tabla | Backend que la tiene como propia |
|---|---|---|
| `BR-B5` | `TbDocumentos` | `AGEDO20_Datos` |
| `GRI-R5` | `TbAplicacionesVersiones`, `TbCambioDocumentos` | `Control_Cambios_datos` |
| `GRI-R5` | `TbExpedientesConEntidades`, `TbExpedientesRACS` | `Expedientes_datos` |
| `REG-G5` | `tbUsuarios` | `Lanzadera_Datos` (y otra distinta en `HPST`) |

El código las busca en un backend que no las vincula: no faltan, falta el vínculo o el código apunta al backend equivocado.

### Bien colocadas

Las otras 18 actúan sobre el vínculo del consumidor (quitarlo, reapuntarlo o añadirlo) o sobre su código, y se aplican donde están: `EXP-E3`, `NOC-N6`, `NOC-N8`, `SHO-S1`, `SHO-S3`, `SHO-S4`, `SHO-S7`, `SHO-S8`, `SHO-S10`, `GRI-R4`, `AGE-E1`, `BFA-S3`, `BFA-S6`, `BR-B2`, `LAN-A1`, `REG-G12`, `HPS-H7` y `HPS-H8`.
