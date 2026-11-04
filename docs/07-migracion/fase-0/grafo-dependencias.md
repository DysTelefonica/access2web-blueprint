# Fase 0: grafo de dependencias entre backends

Issue #1581, tarea T6. Épica `docs/07-migracion/epica-coexistencia-postgresql.md` (fase 0).
Método: ERD del front y del backend de cada aplicación con `scripts/dysflow-erd/extract-erd.ps1`, que abre una **copia temporal** en solo lectura con DAO y vuelca tablas propias, vínculos con su `Connect` y su `SourceTableName`, relaciones, consultas y campos. No se consultan tablas de sistema. Ninguna base se escribe.

## Vínculos por backend

| Backend | Tablas propias | Tablas vinculadas | Destinos de los vínculos |
|---|---|---|---|
| `Lanzadera_Datos` | 35 | 1 | `Correos_datos` (1: `TbCorreosEnviados`) |
| `Gestion_Brass_Gestion_Datos` | 57 | 0 | ninguna |
| `Expedientes_datos` | 47 | 24 | `AGEDYS_DATOS` (15), `Gestion_Riesgos_Datos` (3), `HPST` (2), `Lanzadera_Datos` (2), `Correos_datos` (1), `NoConformidades_Datos` (1) |
| `condor_datos` | 15 | 10 | `Expedientes_datos` (6), `Lanzadera_Datos` (2), `NoConformidades_Datos` (1), `Correos_datos` (1) |
| `NoConformidades_Datos` | 36 | 11 | `Expedientes_datos` (4), `Gestion_Riesgos_Datos` (4), `Lanzadera_Datos` (2), `Correos_datos` (1) |
| `Gestion_Riesgos_Datos` | 64 | 14 | `Expedientes_datos` (4), `Lanzadera_Datos` (3), `AGEDYS_DATOS` (3), `AGEDO20_Datos` (2), `NoConformidades_Datos` (1), `Correos_datos` (1) |
| `HPST` | 18 | 9 | `Expedientes_datos` (4), `Solicitudes_HPS_datos` (2), `Lanzadera_Datos` (2), `Correos_datos` (1) |
| `Registro_Datos` | 5 | 0 | ninguna |

## Tablas que cada backend toma de otro

| Backend | Tablas ajenas que consume |
|---|---|
| `Expedientes_datos` | De AGEDYS: `TbComunicados`, `TbDpDInformeCondicionamiento`, `TbExpedientes1`, `TbExpedientesAnualidades1`, `TbFacturasDetalle`, `TbFacturasPrincipal`, `TbGTVContratos`, `TbGTVContratosExpedientes`, `TbNPedido`, `TbProyectos_AGEDYS`, `TbResponsablesExpedientes`, `TbSolicitudesOfertasPrevias`, `TbSuministradoresSAP`, `TbVisadoFacturas_Nueva`, `TbVisadosGenerales`. De Gestión de Riesgos: `TbGestionRiesgos`, `TbProyectosEdiciones_Riesgos`, `TbProyectos_Riesgos`. De HPS: `TbUsuariosHistoricos_HPS`, `TbUsuarios_HPS`. De No Conformidades: `TbNoConformidades`. De Lanzadera: `TbUsuariosAplicaciones`, `TbUsuariosAplicacionesPermisos`. De Correos: `TbCorreosEnviados` |
| `condor_datos` | De Expedientes: `TbExpedientes`, `TbExpedientesRACS`, `TbExpedientesResponsables`, `TbExpedientesSuministradores`, `TbRACS`, `TbSuministradores`. De No Conformidades: `TbNoConformidades`. De Lanzadera: `TbUsuariosAplicaciones`, `TbUsuariosAplicacionesPermisos`. De Correos: `TbCorreosEnviados` |
| `NoConformidades_Datos` | De Expedientes: `TbExpedientes`, `TbExpedientesJuridicas`, `TbExpedientesResponsables`, `TbJuridicas`. De Gestión de Riesgos: `TbProyectos`, `TbProyectosEdiciones`, `TbRiesgos`, `TbRiesgosNC`. De Lanzadera: `TbUsuariosAplicaciones`, `TbUsuariosAplicacionesPermisos`. De Correos: `TbCorreosEnviados` |
| `Gestion_Riesgos_Datos` | De Expedientes: `TbExpedientes`, `TbExpedientesResponsables`, `TbExpedientesSuministradores`, `TbSuministradores`. De AGEDYS: `TbDPD`, `TbResponsablesExpedientes`, `TbSuministradoresSAP`. De AGEDO: `TbDocumentos`, `TbDocumentosID`. De No Conformidades: `TbNoConformidades`. De Lanzadera: `TbAplicaciones`, `TbUsuariosAplicaciones`, `TbUsuariosAplicacionesPermisos`. De Correos: `TbCorreosEnviados` |
| `HPST` | De Expedientes: `TbExpedientes`, `TbExpedientesConEntidades`, `TbExpedientesLugaresEjecucion`, `TbSuministradores`. De Lanzadera: `TbUsuariosAplicaciones`, `TbUsuariosAplicacionesPermisos`. De Correos: `TbCorreosEnviados` |
| `Lanzadera_Datos` | De Correos: `TbCorreosEnviados` |

## Tablas compartidas y su punto de tensión

| Tabla | Vive en | La consumen |
|---|---|---|
| `TbUsuariosAplicaciones`, `TbUsuariosAplicacionesPermisos` | `Lanzadera_Datos` | Expedientes, Condor, No Conformidades, Gestión de Riesgos, HPS |
| `TbCorreosEnviados` | `Correos_datos` | Los seis backends del ecosistema |
| `TbExpedientes`, `TbExpedientesResponsables`, `TbSuministradores` | `Expedientes_datos` | Condor, No Conformidades, Gestión de Riesgos, HPS |
| `TbNoConformidades` | `NoConformidades_Datos` | Expedientes, Condor, Gestión de Riesgos |
| `TbProyectos`, `TbProyectosEdiciones`, `TbRiesgos` | `Gestion_Riesgos_Datos` | No Conformidades |

Cada tabla compartida tiene una consecuencia operativa: su backend no puede moverse sincoordenar a todos sus consumidores, y un cambio de esquema en ella afecta a varias aplicaciones a la vez.

## Veredictos

**R-1 (cada front consulta solo su backend).** Se confirma **en los backends**: cada uno contiene sus tablas propias y los vínculos hacia otros backends. Se **refuta en los fronts**, con excepciones ya medidas: 11 fronts abren `Lanzadera_Datos` directamente desde código (inventario en `aperturas-directas-lanzadera.md`); el front de Brass vincula su propio backend; el front de HPS y el de Registros vinculan rutas de producción; los fronts de Condor, No Conformidades y Gestión de Riesgos no tienen vínculos y acceden por código.

**R-2 (una tabla vive en el backend de la aplicación que la usa).** No se confirma de forma general: el grafo muestra tablas cuyo uso se reparte entre aplicaciones. El veredicto por tabla está en la ficha de cada aplicación, y la lista consolidada de candidatas a reasignación es la tarea T7.

**Dependencias fuera del alcance.** `AGEDYS_DATOS` (15 tablas desde Expedientes, 3 desde Gestión de Riesgos) y `AGEDO20_Datos` (2 desde Gestión de Riesgos) están fuera de la migración. Sus vínculos deben seguir funcionando durante toda la coexistencia: la web no puede tocar esas tablas ni exigir su reescritura.

## Riesgos

| Riesgo | Ámbito | Evidencia |
|---|---|---|
| Vínculos que apuntan a la ruta de producción `\\datoste\…` desde copias de desarrollo | Expedientes (un vínculo AGEDYS), HPS (5 vínculos), Registros (13 vínculos), Brass (57 vínculos) | ERD de los fronts |
| Dos rutas distintas al mismo backend de Registros, con distinta capitalización | Registros | ERD del front: 6 vínculos a `…\Registros\Registro_Ent_Salida_Datos.accdb` y 1 a `…\REGISTROS\…` |
| `TbCorreosEnviados` en seis backends a la vez | Correos | Tabla compartida |
| `TbExpedientes1` enlazado al AGEDYS de producción desde desarrollo | Expedientes | ERD del backend |

Ninguna de estas acciones se aplica: la puerta de descubrimiento global sigue cerrada (#1576).

## Lo que no se pudo verificar

1. El consumo real de cada tabla por cada front: requiere el veredicto por aplicación, en curso (T5).
2. Quién escribe en las tablas compartidas y con qué frecuencia: no se ha observado el runtime.
3. Los vínculos de `AGEDYS`, `AGEDO` y `Correos`: sus backends no están en el alcance de esta fase.
