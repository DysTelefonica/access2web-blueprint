# Fase 0: grafo de dependencias entre backends

Issue #1581, tarea T6 y T14. Épica `docs/07-migracion/epica-coexistencia-postgresql.md` (fase 0).
Versión transversal: sustituye al primer recuento, que solo cubría ocho backends. Cubre los **17 backends** del ecosistema, con un único método y una única cifra.

## Método

Recuento determinista sobre el ERD de cada backend, obtenido con `scripts/dysflow-erd/extract-erd.ps1` (copia temporal, DAO en solo lectura). Para cada tabla vinculada se toma su `Connect` y se resuelve el backend de destino; de ahí sale el censo de consumidores. No se consultan tablas de sistema y no se escribe ninguna base. El recuento de uso por código, cuando hace falta, es un recuento de coincidencias literales sobre `src/` de cada front, y se declara como tal.

Este documento es el **recuento único** del que habla la fase 0: cuando dos fichas den cifras distintas para el mismo hecho, manda este documento, y la diferencia se corrige en la ficha.

## Inventario por backend

| Backend | Tablas propias | Vinculadas | Relaciones | Consultas | Sin clave primaria |
|---|---|---|---|---|---|
| `agedys` | 102 | 35 | 4 | 5 | 12 |
| `agedo` | 66 | 2 | 12 | 1 | 8 |
| `gestion-riesgos` | 64 | 14 | 19 | 35 | 5 |
| `brass` | 57 | 0 | 28 | 0 | 5 |
| `expedientes` | 47 | 24 | 15 | 11 | 6 |
| `no-conformidades` | 36 | 11 | 14 | 4 | 3 |
| `seguridad` | 36 | 0 | 4 | 1 | 2 |
| `lanzadera` | 35 | 1 | 6 | 4 | 2 |
| `contratos` | 20 | 0 | 3 | 0 | 12 |
| `hps` | 18 | 9 | 8 | 8 | 1 |
| `condor` | 15 | 10 | 5 | 2 | 1 |
| `registro-salida` | 11 | 1 | 1 | 1 | 0 |
| `registro` | 5 | 0 | 1 | 1 | 0 |
| `sica` | 4 | 0 | 1 | 0 | 4 |
| `tareas` | 4 | 13 | 1 | 0 | 1 |
| `control-cambios` | 3 | 0 | 3 | 0 | 0 |
| `correos` | 1 | 0 | 1 | 1 | 1 |

## Tablas propias de un backend consumidas por dos o más backends

Estas son las tablas que no se pueden tocar sin coordinar a todos sus consumidores. El orden es por número de consumidores.

| Tabla | Dueño | Consumidores | Backends consumidores |
|---|---|---|---|
| `TbUsuariosAplicaciones` | `lanzadera` | 8 | `agedo`, `agedys`, `condor`, `expedientes`, `gestion-riesgos`, `hps`, `no-conformidades`, `tareas` |
| `TbCorreosEnviados` | `correos` | 7 | `agedys`, `condor`, `expedientes`, `gestion-riesgos`, `hps`, `lanzadera`, `no-conformidades` |
| `TbUsuariosAplicacionesPermisos` | `lanzadera` | 7 | `agedys`, `condor`, `expedientes`, `gestion-riesgos`, `hps`, `no-conformidades`, `tareas` |
| `TbExpedientes` | `expedientes` | 6 | `agedys`, `condor`, `gestion-riesgos`, `hps`, `no-conformidades`, `tareas` |
| `TbExpedientesResponsables` | `expedientes` | 5 | `agedys`, `condor`, `gestion-riesgos`, `no-conformidades`, `tareas` |
| `TbSuministradoresSAP` | `agedys` | 3 | `expedientes`, `gestion-riesgos`, `tareas` |
| `TbSuministradores` | `expedientes` | 3 | `condor`, `gestion-riesgos`, `hps` |
| `TbNoConformidades` | `no-conformidades` | 3 | `condor`, `expedientes`, `gestion-riesgos` |
| `TbDocumentos` | `agedo` | 2 | `agedys`, `gestion-riesgos` |
| `TbDocumentosID` | `agedo` | 2 | `agedys`, `gestion-riesgos` |
| `TbFacturasDetalle` | `agedys` | 2 | `expedientes`, `tareas` |
| `TbNPedido` | `agedys` | 2 | `expedientes`, `tareas` |
| `TbResponsablesExpedientes` | `agedys` | 2 | `expedientes`, `gestion-riesgos` |
| `TbSolicitudesOfertasPrevias` | `agedys` | 2 | `expedientes`, `tareas` |
| `TbVisadoFacturas_Nueva` | `agedys` | 2 | `expedientes`, `tareas` |
| `TbVisadosGenerales` | `agedys` | 2 | `expedientes`, `tareas` |
| `TbExpedientesJuridicas` | `expedientes` | 2 | `agedys`, `no-conformidades` |
| `TbExpedientesLugaresEjecucion` | `expedientes` | 2 | `agedys`, `hps` |
| `TbExpedientesRACS` | `expedientes` | 2 | `agedys`, `condor` |
| `TbExpedientesSuministradores` | `expedientes` | 2 | `condor`, `gestion-riesgos` |
| `TbJuridicas` | `expedientes` | 2 | `agedys`, `no-conformidades` |
| `TbRACS` | `expedientes` | 2 | `agedys`, `condor` |
| `TbAplicaciones` | `lanzadera` | 2 | `agedys`, `gestion-riesgos` |
| `Salidas` | `registro` | 2 | `agedys`, `registro-salida` |

## Tablas de un backend consumidas por un solo backend

Candidatas a estar mal ubicadas: su único consumidor es otro backend, y la regla de propiedad sitúa la tabla en el backend del front que la necesita. Ninguna se mueve sin comprobar el uso por código: el recuento de vínculos no ve el SQL guardado en tablas, que ya ha tapiado un falso positivo en AGEDYS (`TbFacturasDetalleNoDYS`).

| Tabla | Dueño | Único consumidor |
|---|---|---|
| `TbDocExpProceso` | `agedo` | `agedys` |
| `TbComunicados` | `agedys` | `expedientes` |
| `TbDPD` | `agedys` | `gestion-riesgos` |
| `TbDPDOfertaSuministrador` | `agedys` | `tareas` |
| `TbDpDInformeCondicionamiento` | `agedys` | `expedientes` |
| `TbExpedientes1` | `agedys` | `expedientes` |
| `TbExpedientesAnualidades1` | `agedys` | `expedientes` |
| `TbFacturasPrincipal` | `agedys` | `expedientes` |
| `TbGTVContratos` | `agedys` | `expedientes` |
| `TbGTVContratosExpedientes` | `agedys` | `expedientes` |
| `TbProyectos` | `agedys` | `tareas` |
| `TbProyectos_AGEDYS` | `agedys` | `expedientes` |
| `TbSolicitudesOfertasPrevias1` | `agedys` | `agedys` |
| `TbCPV` | `expedientes` | `agedys` |
| `TbComerciales` | `expedientes` | `agedys` |
| `TbEjercitos` | `expedientes` | `agedys` |
| `TbEstados` | `expedientes` | `agedys` |
| `TbExpedientesAnexos` | `expedientes` | `agedys` |
| `TbExpedientesAnualidades1` | `expedientes` | `agedys` |
| `TbExpedientesCPVs` | `expedientes` | `agedys` |
| `TbExpedientesComerciales` | `expedientes` | `agedys` |
| `TbExpedientesConEntidades` | `expedientes` | `hps` |
| `TbExpedientesJefaturas` | `expedientes` | `agedys` |
| `TbExpedientesPECAL` | `expedientes` | `agedys` |
| `TbGradosClasificacion` | `expedientes` | `agedys` |
| `TbJefaturas` | `expedientes` | `agedys` |
| `TbLugaresEjecucion` | `expedientes` | `agedys` |
| `TbOficinasPrograma` | `expedientes` | `agedys` |
| `TbOrganosContratacion` | `expedientes` | `agedys` |
| `TbPECAL` | `expedientes` | `agedys` |
| `TbGestionRiesgos` | `gestion-riesgos` | `expedientes` |
| `TbProyectos` | `gestion-riesgos` | `no-conformidades` |
| `TbProyectosEdiciones` | `gestion-riesgos` | `no-conformidades` |
| `TbProyectosEdiciones_Riesgos` | `gestion-riesgos` | `expedientes` |
| `TbProyectos_Riesgos` | `gestion-riesgos` | `expedientes` |
| `TbRiesgos` | `gestion-riesgos` | `no-conformidades` |
| `TbRiesgosNC` | `gestion-riesgos` | `no-conformidades` |
| `TbUsuariosHistoricos_HPS` | `hps` | `expedientes` |
| `TbUsuarios_HPS` | `hps` | `expedientes` |
| `TbConexionesRegistro` | `lanzadera` | `agedys` |
| `TbUsuariosAplicacionesTareas` | `lanzadera` | `tareas` |
| `tbUsuarios` | `lanzadera` | `agedo` |
| `TbEntradas_RES` | `registro-salida` | `agedys` |
| `TbSalidas_RES` | `registro-salida` | `agedys` |

## Veredictos

**R-1 (cada front consulta solo su backend).** Se **refuta en los fronts** y se confirma solo en parte de los backends. En el lado de los fronts hay tres excepciones medidas: el front de AGEDYS Técnico vincula ocho backends distintos, el front de HPS y el de Registros enlazan rutas de producción, y 11 fronts abren `Lanzadera_Datos` directamente desde código (inventario en `aperturas-directas-lanzadera.md`). En el lado de los backends, el patrón real es el inverso del enunciado: los backends contienen vínculos hacia otros backends, de modo que el grafo se lee desde ellos, pero no describe a quién consulta el front.

**R-2 (una tabla vive en el backend de la aplicación que la usa).** No se confirma. Hay 44 tablas cuyo único consumidor es otro backend (tabla anterior) y 24 consumidas por dos o más backends. El veredicto tabla por tabla está en la ficha de cada aplicación y el catálogo único consolida las propuestas.

**Consecuencia para la migración.** Con la D92 la web va a leer en PostgreSQL las mismas tablas que Access vincula por ODBC. Por eso el orden de migración no puede ser el del catálogo de la lanzadera: los consumidores de cada tabla compartida tienen que migrarse antes o a la vez que su dueño, y ese orden es una decisión del propietario.

## Tensión de propiedad por tabla compartida

| Tabla | Dueño | Consumidores | Por qué importa |
|---|---|---|---|
| `TbUsuariosAplicaciones` | `lanzadera` | 8 | Identidad y permisos: ningún backend puede moverse sin esta tabla. |
| `TbUsuariosAplicacionesPermisos` | `lanzadera` | 7 | Idéntico razonamiento que la anterior. |
| `TbCorreosEnviados` | `correos` | 7 | Sin clave primaria, 35 394 filas y escrita desde al menos cinco aplicaciones. Cuatro fichas piden clave sobre ella. |
| `TbExpedientes` | `expedientes` | 6 | El expediente compartido es el objeto central de varias aplicaciones. |
| `TbExpedientesResponsables` | `expedientes` | 5 | Igual que el expediente. |
| `TbSuministradoresSAP` | `agedys` | 3 | AGEDYS es dueño, pero Expedientes, Gestión de Riesgos y Tareas la consumen. |
| `TbDocumentos`, `TbDocumentosID` | `agedo` | 2 | AGEDO está fuera del alcance de la aplicación, pero Gestión de Riesgos, que sí migra, las consume. |

## Lo que no se pudo verificar

1. El consumo por SQL guardado en tablas no aparece en este censo: el recuento es de vínculos. Por eso el catálogo de accionables exige contraste con las fichas antes de proponer una retirada.
2. Los vínculos en producción (`\datoste\…`) no son alcanzables desde el entorno de desarrollo: los campos se leyeron del ERD, no del origen.
3. Los recuentos de filas son de copias de desarrollo. Antes de aprobar cualquier acción hay que verificarlos en una copia reciente de producción.
