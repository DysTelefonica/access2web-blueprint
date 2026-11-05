# Comprobaciones de datos del grupo (a)

Los recuentos proceden de la [copia de referencia](copia-referencia.md). El inventario de tablas se reproduce con `scripts/dysflow-erd/inventory-data.ps1`, que comprueba antes el SHA-256 de cada backend contra el manifiesto y trabaja sobre copias temporales; su salida no se versiona porque se regenera contra la misma foto. El inventario recorre solo tablas propias: un vínculo se mide en la tabla a la que apunta.
Las fichas identifican las tablas y sus propietarios; los datos del front quedan fuera de la foto. Estas comprobaciones no autorizan aplicar las propuestas.

## Resumen

| Propuesta | Tipo | Veredicto |
|---|---|---|
| CON-C5 | Clave primaria | Bloquea |
| COR-C1 | Clave primaria | Bloquea |
| EXP-E2 | Clave primaria | Bloquea; una tabla no tiene candidata |
| GRI-R2 | Clave primaria | Bloquea; hay tablas sin candidata |
| NOC-N3 | Clave primaria | Bloquea |
| AGE-E5 | Clave primaria | Front: 136 filas; ver [tablas de los fronts](#tablas-de-los-fronts) |
| AGT-A1 | Clave primaria | Bloquea; hay tablas sin candidata |
| BFA-S2 | Clave primaria | Sin candidata |
| BFA-S5 | Clave primaria | Bloquea; hay tablas sin candidata |
| BR-B3 | Clave primaria | Sin candidata en una tabla |
| REG-G3 | Clave primaria | Se puede aplicar: todas tienen clave |
| SEG-G1 | Clave primaria | Se puede aplicar |
| SHO-S1 | Vacías | Tiene datos: revisar |
| AGT-A4 | Vacías | Front: 139 filas; revisar |
| AGT-A5 | Vacías | Front: 35 filas; revisar |
| BR-B1 | Vacías | Tiene datos: revisar |
| CON-C1 | Vacías | Vacías |
| EXP-E5 | Vacías | Tiene datos: revisar; `TbExpedientesAnualidades` está duplicada |
| GRI-R3 | Vacías | Tiene datos: revisar |
| HPS-H1 | Vacías | Tiene datos: revisar |
| HPS-H2 | Vacías | Tiene datos: revisar; la copia del front tiene 99 filas |
| HPS-H3 | Vacías | Front: 344 y 50 filas; revisar |
| LAN-A2 | Vacías | Vacías |
| LAN-A3 | Vacías | Tiene datos: revisar |
| LAN-A4 | Vacías | Tiene datos: revisar; excluir `tbUsuarios` |
| NOC-N1 | Vacías | Tiene datos: revisar |
| NOC-N2 | Vacías | Tiene datos: revisar |
| NOC-N7 | Vacías | Front: 23 filas; revisar |
| REG-G1 | Vacías | Tiene datos: revisar |
| SEG-G2 | Vacías | Tiene datos: revisar |
| BFA-S4 | Duplicidad | Comparación pendiente |
| AGT-A3 | Duplicidad | Dos copias idénticas sobran; `TbHerramientaDocAyuda` difiere |
| SHO-S3 | Duplicidad | Idénticas: la copia se puede sustituir por un vínculo |
| LAN-A5 | Integridad referencial | Bloquea: 9 huérfanos |
| SEG-G5 | Integridad referencial | Se puede aplicar |
| AGT-A8 | Integridad referencial | Se puede aplicar en los tres pares medidos |

## Claves primarias

«Duplicados» es el exceso de filas no nulas sobre los valores distintos: filas − nulos − distintos. Una candidata en una tabla vacía cumple la comprobación de la foto, no garantiza unicidad futura. «Sin candidata» significa que el inventario no recoge una columna autonumérica o `ID*`; necesita una decisión de clave, posiblemente compuesta. Una clave existente no tiene candidatas medidas. Que una columna sea única en los datos no la convierte en la clave correcta: cada candidata se confirma con el significado de la tabla antes de declararla.

### Correos: CON-C5, COR-C1 y NOC-N3

| Backend | Tabla | Filas | Clave actual | Mejor candidata | Resultado |
|---|---|---:|---|---|---|
| `Correos_datos.accdb` | `TbCorreosEnviados` | 35395 | Ninguna | `IDCorreo` | Bloquea: 0 nulos, 35376 distintos, 19 duplicados |

Los 19 duplicados son parejas de correos distintos que recibieron el mismo `IDCorreo` porque cada aplicación lo calcula al escribir. La solución decidida es renumerar y pasar a autonumérico (decisión 22).

### EXP-E2

| Backend | Tabla | Filas | Clave actual | Mejor candidata | Resultado |
|---|---|---:|---|---|---|
| `Correos_datos.accdb` | `TbCorreosEnviados` | 35395 | Ninguna | `IDCorreo` | Bloquea: 0 nulos, 35376 distintos, 19 duplicados |
| `AGEDYS_DATOS.accdb` | `TbFacturasPrincipal` | 2706 | Ninguna | — | Sin candidata |
| `Expedientes_datos.accdb` | `TbExpAgedys` | 116 | Ninguna | `IdExpediente` | Se puede aplicar: 0 nulos, 116 distintos |
| `Expedientes_datos.accdb` | `TbExpAGEDYS1` | 116 | Ninguna | `IdExpediente` | Se puede aplicar: 0 nulos, 116 distintos |
| `Expedientes_datos.accdb` | `TbExpedientes1` | — | — | — | No es una tabla: era un vínculo residual de Expedientes que apuntaba a otro vínculo de `AGEDYS_DATOS` y volvía a `TbExpedientes`. El propietario lo mandó retirar de la copia el 2026-11-05; en producción sigue, sin uso salvo una rutina sin llamadores de `Instalador.bas` |
| `Expedientes_datos.accdb` | `E2E_export_20260723_142133_lista` | 89 | Ninguna | `IDExpediente` | Bloquea: 1 nulo, 88 distintos, 0 duplicados |

### GRI-R2

| Backend | Tabla | Filas | Clave actual | Mejor candidata | Resultado |
|---|---|---:|---|---|---|
| `Gestion_Riesgos_Datos.accdb` | `ProyectoRespCalidad` | 16 | Ninguna | — | Sin candidata |
| `Gestion_Riesgos_Datos.accdb` | `TbAuxProyectosRiesgos` | 242 | Ninguna | `IDProyecto` | Bloquea: 0 nulos, 18 distintos, 224 duplicados |
| `Gestion_Riesgos_Datos.accdb` | `TbCachePublicabilidadEdicion` | 0 | Ninguna | `IDEdicion` o `IDRiesgo` | Se puede aplicar: 0 nulos, 0 distintos; solo en la foto |
| `Correos_datos.accdb` | `TbCorreosEnviados` | 35395 | Ninguna | `IDCorreo` | Bloquea: 0 nulos, 35376 distintos, 19 duplicados |
| `Gestion_Riesgos_Datos.accdb` | `TbRiesgosValoracion` | 25 | Ninguna | — | Sin candidata |

### AGE-E5

`TbAuxGTVContratos` es propia del front: fuera de la foto. La foto no permite medir sus filas, clave ni candidata.

### AGT-A1

Todas las tablas de esta sección pertenecen a `AGEDYS_DATOS.accdb`.

| Tabla | Filas | Clave actual | Mejor candidata | Resultado |
|---|---:|---|---|---|
| `TbAuxExp` | 40 | Ninguna | `IdExpediente` | Se puede aplicar: 0 nulos, 40 distintos |
| `TbAuxPrado` | 468 | Ninguna | — | Sin candidata |
| `TbAvalesDatosGenerales` | 1 | Ninguna | — | Sin candidata |
| `TbAvalesDatosGeneralesNuevos` | 1 | Ninguna | — | Sin candidata |
| `TbDpDCartaSolicitudOfertaDatosFijos` | 1 | Ninguna | — | Sin candidata |
| `TbDpDCartaSolicitudOfertaPlantillasSegunReqBasicos` | 25 | Ninguna | — | Sin candidata |
| `TbFacturasPrincipal` | 2706 | Ninguna | — | Sin candidata |
| `TbLog` | 19731 | Ninguna | `ID` | Bloquea: 0 nulos, 19725 distintos, 6 duplicados |
| `TbProductos` | 408 | Ninguna | — | Sin candidata |
| `TbProductosObsoletos` | 25 | Ninguna | — | Sin candidata |
| `TbProductos_antes` | 424 | Ninguna | — | Sin candidata |

### BFA-S2 y BFA-S5

| Propuesta | Backend | Tabla | Filas | Clave actual | Mejor candidata | Resultado |
|---|---|---|---:|---|---|---|
| BFA-S2 | `SICA_datos.accdb` | `TbAperturas` | 7800 | Ninguna | — | Sin candidata |
| BFA-S2 | `SICA_datos.accdb` | `TbArchivos` | 0 | Ninguna | — | Sin candidata |
| BFA-S2 | `SICA_datos.accdb` | `TbCierres` | 54 | Ninguna | — | Sin candidata |
| BFA-S2 | `SICA_datos.accdb` | `TbUltimaImportacion` | 1 | Ninguna | — | Sin candidata |
| BFA-S5 | `AGEDO20_Datos.accdb` | `Entradas` | 14125 | Ninguna | — | Sin candidata |
| BFA-S5 | `AGEDO20_Datos.accdb` | `Salidas` | 11609 | Ninguna | — | Sin candidata |
| BFA-S5 | `AGEDO20_Datos.accdb` | `TbDocExpProcesoRegistroTArchivos` | 810 | Ninguna | `IDProceso` | Bloquea: 0 nulos, 255 distintos, 555 duplicados |
| BFA-S5 | `AGEDO20_Datos.accdb` | `TbCorreoAcciones` | 910 | Ninguna | `IDCorreo` | Se puede aplicar: 0 nulos, 910 distintos |
| BFA-S5 | `AGEDO20_Datos.accdb` | `TbAuxCodRepetidos` | 170 | Ninguna | `IDRepetido` | Se puede aplicar: 0 nulos, 170 distintos |
| BFA-S5 | `AGEDO20_Datos.accdb` | `TbMatClasificadoEntregaDetalleAux` | 22 | Ninguna | `IDEntregaDetalle` | Se puede aplicar: 0 nulos, 22 distintos |
| BFA-S5 | `AGEDO20_Datos.accdb` | `TbCartas` | 20 | Ninguna | `IDCarta` o `IDProceso` | Se puede aplicar: 0 nulos, 20 distintos |
| BFA-S5 | `AGEDO20_Datos.accdb` | `TbVersionCliente` | 1 | Ninguna | — | Sin candidata |

### BR-B3 y SEG-G1

| Propuesta | Backend | Tabla | Filas | Clave actual | Mejor candidata | Resultado |
|---|---|---|---:|---|---|---|
| BR-B3 | `Gestion_Brass_Gestion_Datos.accdb` | `TbAuxActividad` | 0 | Ninguna | `IDEQUIPO` o `IDEVENTO` | Se puede aplicar: 0 nulos, 0 distintos; solo en la foto |
| BR-B3 | `Gestion_Brass_Gestion_Datos.accdb` | `TbAuxManteniminetosPreventivosCalendario` | 93 | Ninguna | `IDEQUIPO` | Se puede aplicar: 0 nulos, 93 distintos |
| BR-B3 | `Gestion_Brass_Gestion_Datos.accdb` | `TbAuxMateriales` | 0 | Ninguna | `IDEVENTO` | Se puede aplicar: 0 nulos, 0 distintos; solo en la foto |
| BR-B3 | `Gestion_Brass_Gestion_Datos.accdb` | `TbAuxPlanificacion` | 81 | Ninguna | `IDPlanificacion` | Se puede aplicar: 0 nulos, 81 distintos |
| BR-B3 | `Gestion_Brass_Gestion_Datos.accdb` | `TbBuiIDEvento` | 11 | Ninguna | — | Sin candidata |
| SEG-G1 | `Seguridad_datos.accdb` | `TbNombramientoJS` | 17 | Ninguna | `IDNombramientoJS` | Se puede aplicar: 0 nulos, 17 distintos |
| SEG-G1 | `Seguridad_datos.accdb` | `Tb0Vinculo` | 0 | Ninguna | `ID` | Se puede aplicar: 0 nulos, 0 distintos; solo en la foto |

### REG-G3

Las 11 tablas propias de `Registro_Ent_Salida_Datos.accdb` ya tienen clave primaria; no hay candidatas medidas. El vínculo `Salidas` pertenece a `Registro_Datos.accdb` y no forma parte de esta propuesta.

| Tabla | Filas | Clave actual | Resultado |
|---|---:|---|---|
| `TbConexiones` | 12 | `Usuario` | Se puede aplicar: ya tiene clave |
| `TbEntradaDestinatarios_RES` | 37 | `IDEntrada+Destinatario` | Se puede aplicar: ya tiene clave |
| `TbEntradaPersonasContacto_RES` | 0 | `IDEntrada+PersonaContacto` | Se puede aplicar: ya tiene clave |
| `TbEntradas_indexacion` | 0 | `IDEntrada` | Se puede aplicar: ya tiene clave |
| `TbEntradas_RES` | 635 | `IDEntrada` | Se puede aplicar: ya tiene clave |
| `TbEntradas_RES_Indexacion` | 2 | `IDEntrada` | Se puede aplicar: ya tiene clave |
| `TbLog` | 3341 | `IDLog` | Se puede aplicar: ya tiene clave |
| `TbSalidaDestinatarios_RES` | 5 | `IDSalida+Destinatario` | Se puede aplicar: ya tiene clave |
| `TbSalidas_Indexacion` | 0 | `IDSalida` | Se puede aplicar: ya tiene clave |
| `TbSalidas_RES` | 4053 | `IDSalida` | Se puede aplicar: ya tiene clave |
| `TbSalidas_RES_Indexacion` | 1063 | `IDSalida` | Se puede aplicar: ya tiene clave |

## Tablas candidatas a cuarentena o retirada

Los recuentos son de la foto, no una autorización de cuarentena. Las tablas vinculadas se buscan en su backend de origen; si el nombre exacto no aparece, no se sustituye por otro.

| Propuesta y backend | Tablas y filas | Veredicto |
|---|---|---|
| SHO-S1 · `Solicitudes_HPS_datos.accdb` | `TbLogs` 0; `Copia de TbExpedientes` 70 | Tiene datos: revisar `Copia de TbExpedientes` 70 |
| AGT-A4 · front | `Copia de TbRequisitosSTG`: fuera de la foto (front) | Fuera de la foto (front) |
| AGT-A5 · front | `TbAuxNPedido`: fuera de la foto (front) | Fuera de la foto (front) |
| BR-B1 · `Gestion_Brass_Gestion_Datos.accdb` | `TbAuxManteniminetosPreventivosCalendario` 93; `TbCausaFin` 4; `TbGuiaConciliaciones` 0; `TbRepuestosTipo` 3; `TbResultadoVerificacion` 3; `TbSistema` 2; `TbTipoAsistencia` 4 | Tiene datos: revisar todas salvo `TbGuiaConciliaciones` |
| CON-C1 · `condor_datos.accdb` | `tbHistorialRechazos` 0 | Vacías |
| HPS-H1 · `HPST.accdb` | `TbAuxCursos` 90; `TbHPSEquivalencia` 3; `TbJuridicasContratacion` 3 | Tiene datos: revisar las tres |
| HPS-H2 · `HPST.accdb` y front | `TbUsuariosEntidades` 352; `Copia de TbUsuariosEntidades`: fuera de la foto (front) | Tiene datos: revisar `TbUsuariosEntidades` 352 |
| HPS-H3 · front | `TbDatosLocalParaIndicadores` y `TbUsuariosSICALocalParaIndicadores`: fuera de la foto (front) | Fuera de la foto (front) |
| LAN-A2 · `Lanzadera_Datos.accdb` | `TbAplicacionesVideos` 0; `TbCuestionarioPreguntas` 0; `TbCuestionarios` 0; `TbCuestionaroRespuestas` 0; `TbUsuarioAplicacionesSolicitud` 0; `TbVideosCuestionario` 0 | Vacías |
| LAN-A3 · `Lanzadera_Datos.accdb` | `TbTablasAVincular` 411 | Tiene datos: revisar `TbTablasAVincular` 411 |
| NOC-N2 · `NoConformidades_Datos.accdb` | `TbAnexosAuditoria` 8; `TbAnexosNCAuditorias` 0 | Tiene datos: revisar `TbAnexosAuditoria` 8 |
| NOC-N7 · front | `TbTipologiaAux`: fuera de la foto (front) | Fuera de la foto (front) |
| REG-G1 · `Registro_Ent_Salida_Datos.accdb` | `TbEntradaPersonasContacto_RES` 0; `TbEntradas_indexacion` 0; `TbEntradas_RES_Indexacion` 2; `TbSalidas_Indexacion` 0; `TbConexiones` 12 | Tiene datos: revisar `TbEntradas_RES_Indexacion` 2 y `TbConexiones` 12 |
| SEG-G2 · `Seguridad_datos.accdb` | `Tb0AdjuntosIndexaciones` 1; `Tb0Conexiones` 3; `Tb0RegistrosPendientes` 2; `Tb0TipoRegistrosDescripcion` 21; `Tb0TipoRegistroUso` 0; `Tb0Vinculo` 0; `TbAuditorias` 0; `TbBuzonTEOficio` 0; `TbCarta` 0; `TbRecibosMP` 0 | Tiene datos: revisar las cuatro primeras |

### EXP-E5

La ficha enumera 11 propias y 14 vinculadas. En las vinculadas se indica el backend de origen cuando la ficha permite identificarlo; los alias que no aparecen con ese nombre no se equiparan a otras tablas.

| Backend | Tablas y filas |
|---|---|
| `Expedientes_datos.accdb` | `Copia de TbExpedientes` 0; `Copia de TbExpedientesConEntidades` 365; `E2E_export_20260723_142133_lista` 89; `ListaPrevia` 99; `TbAusExpPostAGEDO` 203; `TbAuxEstadosMartina` 143; `TbAuxNemotecnico` 0; `TbExpAGEDYS1` 116; `TbExpedientes_antes` 457; `TbExpedientesCadenaContratacion` 0; `TbJefaturas` 3 |
| `AGEDYS_DATOS.accdb` | `TbComunicados` 0; `TbDpDInformeCondicionamiento` 1152; `TbExpedientesAnualidades1`: vínculo a la `TbExpedientesAnualidades` propia de `AGEDYS_DATOS`, 223 (ver la duplicidad abajo); `TbFacturasDetalle` 5896; `TbFacturasPrincipal` 2706; `TbGTVContratos` 594; `TbGTVContratosExpedientes` 595; `TbNPedido` 2850; `TbResponsablesExpedientes` 996; `TbSolicitudesOfertasPrevias` 2327; `TbSuministradoresSAP` 50040; `TbVisadoFacturas_Nueva` 5809; `TbVisadosGenerales` 2323 |
| `Gestion_Riesgos_Datos.accdb` | `TbGestionRiesgos`: alias del vínculo a `TbProyectos`, 64 |

Veredicto: Tiene datos: revisar las tablas con recuento distinto de cero.

**Duplicidad encontrada.** Hay dos tablas `TbExpedientesAnualidades` propias y distintas: 223 filas en `AGEDYS_DATOS.accdb` y 173 en `Expedientes_datos.accdb`. Cada backend vincula la del otro como `TbExpedientesAnualidades1`. El propietario decide que la tabla viva de forma única en Expedientes (decisión 21). La comparación muestra que divergen: ningún `IDAnualidad` común, 157 pares expediente y año en las dos sin ninguna fila idéntica, 66 solo en AGEDYS y 16 solo en Expedientes. Los pasos de conciliación están en el [catálogo de conflictos](catalogo-accionables-conflictos.md#qué-necesita-decidir-el-propietario).

### GRI-R3

`Gestion_Riesgos_Datos.accdb`: las 30 tablas de la ficha son las siguientes. `ProyectoRespCalidad` y `TbTareas` se revisan aparte y no se incluyen entre las 30.

| Grupo | Tablas y filas |
|---|---|
| Identificadores históricos | `TbIDAccionPlanContingencia` 251; `TbIDAccionPlanMitigacion` 1074; `TbIDAnexos` 107; `TbIDBibliotecaRiesgos` 51; `TbIDCorreosEnviados` 264; `TbIDDocumentos` 27; `TbIDEdiciones` 209; `TbIDLog` 1077; `TbIDPlanContingencia` 154; `TbIDPlanMitigacion` 640; `TbIDProyectos` 71; `TbIDRiesgos` 889; `TbIDRiesgosAIntegrar` 58; `TbIDTareas` 53 |
| Valores posibles | `TbValoresPosiblesContingencia` 4; `TbValoresPosiblesEstadoRiesgo` 7; `TbValoresPosiblesMitigacion` 4; `TbValoresPosiblesPlazoCalidadCosteVulnerabilidad` 5; `TbValoresPosiblesValoracion` 4 |
| Resto | `TbAnexosAntigua` 2; `TbAux` 13; `TbAuxSumExp` 6; `TbFelix` 47; `TbLog` 5817; `TbOrigenesRiesgosDetalles` 5; `TbProyectosResponsablesCalidad` 39; `TbProyectosSuministradores` 2; `TbProyectosTipo` 0; `TbRiesgosPorTipoProyecto` 0; `TbTablas` 41 |

Veredicto: Tiene datos: revisar todas salvo `TbProyectosTipo` y `TbRiesgosPorTipoProyecto`.

### LAN-A4

`Lanzadera_Datos.accdb`: `TbAplicacionesEdiciones` 1262; `TbAplicacionesEdicionesCambios` 1364; `TbAplicacionesEstados` 6; `TbAplicacionesParametros` 60; `TbAplicacionesPerfiles` 34; `TbConexiones` 24; `TbDetalleVersiones` 28; `TbParametros` 53; `TbPermisos` 205; `TbUbicaciones` 8; `TbUsuarioConfiguracion` 14; `tbUsuarios` 156; `TbUsuariosTareasDiarias` 8; `TbVideosVisionados` 138; `TbUsuariosCorreosEnvio` 1042.

Veredicto: Tiene datos: revisar las 15. **No poner `tbUsuarios` en cuarentena**: AGEDO la consume, conforme a la decisión 14. Las demás siguen sujetas al contraste de consumidores.

### NOC-N1

`NoConformidades_Datos.accdb`: `TbAnexos` 228; `TbNCARAvisos` 432; `TbNoConformidadesIngresoPorLotesTemporal` 166; `TbNoConformidadesIngresoPorLotesDetalle` 166; `TbNoConformidadesIngresoPorLotesPrincipal` 1; `TbTareasExplicaciones` 10; `TbConsultasPorFechas` 8; `TbConfiguracion` 1; `TbNCInformacionRAC` 2; `TbNCInformacionRACAnexos` 1; `TbAuditoriaLog` 0; `TbCacheIndicadoresProyectoHeader` 0; `TbCacheIndicadoresProyectoDetalle` 0; `TbCacheListadoNC` 0; `TbConexiones` 0.

Veredicto: Tiene datos: revisar las diez primeras.

## Duplicidades

| Propuesta | Copias y filas en la foto | Comparación pendiente |
|---|---|---|
| BFA-S4 | `TbCorreosEnviados`: `Tareas_datos1.accdb` 1620; `Correos_datos.accdb` 35395 | Comparar destinatarios y fechas y determinar si registran el mismo hecho. |
| AGT-A3 | `AGEDYS_DATOS.accdb`: `TbAuxExp` 40; `TbHerramientaDocAyuda` 10; `TbProductosObsoletos` 25; `TbSTGParametrosParaFormulario` 15; `TbSuministradoresListaFavoritos` 0. Las cinco copias del front están fuera de la foto. | Comparar contenido, esquema y recuentos con el front; comprobar sus lectores. |
| SHO-S3 | `TbHPSGrado`: `Solicitudes_HPS_datos.accdb` 13; `HPST.accdb` 13 | Comparada: misma estructura y las 13 filas idénticas. La copia de Solicitudes HPS se puede sustituir por un vínculo a la de `HPST.accdb` sin perder datos; el vínculo `TbHPSGrado1` no es una copia local. |

## Integridad referencial

Huérfano es una fila hija cuyo valor no nulo no existe en la tabla madre. Las claves de cada tabla se leen en su propio backend, sobre copias temporales, y se comparan sin distinguir mayúsculas ni espacios en los extremos. Que una madre no tenga hijas no es un error.

| Propuesta | Relación (hija → madre) | Filas hijas | Huérfanos | Veredicto |
|---|---|---:|---:|---|
| LAN-A5 | `TbConexionUltimaAppAbierta.IDUltimaAplicacionAbierta` → `TbAplicaciones.IDAplicacion` | 1469 | 9 (un solo valor) | Bloquea: limpiar antes, como prevé la ficha |
| SEG-G5 | `Tb0RegistrosRelacionados.IDTipoRegistroARelacionar` → `Tb0TipoRegistros.IDTipoRegistro` | 133 | 0 | Se puede aplicar |
| SEG-G5 | `Tb0RegistrosRelacionados.IDTipoRegistroRelacionado` → `Tb0TipoRegistros.IDTipoRegistro` | 133 | 0 | Se puede aplicar |
| AGT-A8 | `TbNPedido.CODPPD` → `TbProyectos.CODPROYECTOS` | 2850 | 0 | Se puede aplicar |
| AGT-A8 | `TbVisadoFacturas_Nueva.IDFactura` → `TbFacturasDetalle.IDFactura` | 5809 | 0 | Se puede aplicar |
| AGT-A8 | `TbVisadoFacturas_Nueva.NFactura` → `TbFacturasDetalle.NFactura` | 5809 | 1 | La unión correcta es por `IDFactura` |
| AGT-A8 | `TbProyectos.IDExpediente` → `TbExpedientes.IDExpediente` (`Expedientes_datos.accdb`) | 3078 | 0 | Datos limpios; al estar en dos ficheros, Access no permite la relación: solo se declara en PostgreSQL |

Correcciones a las fichas que salen de esta medición:

- **SEG-G5.** `Tb0RegistrosRelacionados` no tiene `IDTipoRegistro`: tiene dos columnas de tipo, `IDTipoRegistroARelacionar` e `IDTipoRegistroRelacionado`. La unión con las 21 tablas de tipo de registro también está limpia: cada uno de los 266 extremos de `Tb0RegistrosRelacionados` (dos por fila) existe en la tabla que `Tb0TipoRegistros.NombreTabla` asigna a su tipo, por la columna `NombreID`; las 21 tablas se leen y no hay ningún huérfano.
- **AGT-A8.** La unión no es por `CODPPD` en las dos tablas: el SQL del front une `TbNPedido.CODPPD` con `TbProyectos.CODPROYECTOS` en 70 sitios.
- En sentido contrario hay 228 proyectos sin pedido y 87 facturas sin visado: son madres sin hijas, no huérfanos.

## Tablas de los fronts

Las tablas propias de un front no están en la copia de referencia, que solo contiene backends. Se midieron en los fronts de los repositorios, sobre copias temporales: AGEDYS Técnico (`00_AGEDYS`, `350896d`), AGEDYS ECO (`00_AGEDYS_ECO`, `a263ead`), HPS (`00_HPS`, `dc194f4`) y No Conformidades (`00_NO_CONFORMIDADES`, `d834997`). Es el contenido del front publicado, con el que arranca cada usuario: en producción cada usuario trabaja sobre su propia copia del front, y sus filas pueden diferir.

| Propuesta | Front | Tabla | Filas |
|---|---|---|---:|
| AGE-E5 | AGEDYS ECO | `TbAuxGTVContratos` | 136 |
| AGT-A4 | AGEDYS Técnico | `Copia de TbRequisitosSTG` | 139 |
| AGT-A5 | AGEDYS Técnico | `TbAuxNPedido` | 35 |
| HPS-H2 | HPS | `Copia de TbUsuariosEntidades` | 99 |
| HPS-H3 | HPS | `TbDatosLocalParaIndicadores` | 344 |
| HPS-H3 | HPS | `TbUsuariosSICALocalParaIndicadores` | 50 |
| NOC-N7 | No Conformidades | `TbTipologiaAux` | 23 |

Ninguna está vacía en el front publicado: todas necesitan revisar su uso antes de retirarse.

AGT-A3 compara las cinco tablas que existen a la vez en el front de AGEDYS Técnico y en `AGEDYS_DATOS.accdb`:

| Tabla | Front | Backend | Filas idénticas | Lectura |
|---|---:|---:|---:|---|
| `TbProductosObsoletos` | 25 | 25 | 25 | Idénticas: la copia del front sobra |
| `TbSTGParametrosParaFormulario` | 15 | 15 | 15 | Idénticas: la copia del front sobra |
| `TbHerramientaDocAyuda` | 33 | 10 | 0 | Difieren: el front tiene más filas; hay que decidir cuál vale |
| `TbAuxExp` | 0 | 40 | 0 | Tabla de trabajo: vacía en el front publicado |
| `TbSuministradoresListaFavoritos` | 0 | 0 | 0 | Vacías en los dos |
