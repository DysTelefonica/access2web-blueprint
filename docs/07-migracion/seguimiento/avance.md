# Avance de la migración por fases

> Fichero generado. No edite este fichero a mano: se regenera con `python scripts/migration_progress.py` a partir de `fases.json` y de las issues de GitHub. Fecha de corte: 2026-11-25.

| Fase | Nombre | Estado | Hechas | Total |
|---|---|---|---|---|
| 0 | Inventario | revisada | 4 | 4 |
| 1 | Saneamiento en Access | en curso | 6 | 189 |
| 2 | Unificar el acceso a datos | pendiente | 0 | 5 |
| 3 | Ensayo en Citrix con todos los backends | revisada (diseño revisado, pendiente de ejecutar) | 0 | 7 |
| 4 | Traslado único de todos los datos a PostgreSQL | pendiente | 0 | 5 |
| 5 | Expedientes a web | pendiente | 0 | 3 |

## Fase 0. Inventario

**Estado:** revisada · **Hechas:** 4/4

- [x] Inventario de las once aplicaciones del alcance (fase-0/catalogo-aplicaciones.md)
- [x] Catálogo de accionables con su coste medido (fase-0/catalogo-accionables.md)
- [x] Comprobaciones de datos del grupo (a) contra la copia de referencia (fase-0/comprobaciones-grupo-a.md)
- [x] Validación del propietario de la evidencia de la fase 0, 2026-11-15 (fase-0/cierre.md, decisión 28)

## Fase 1. Saneamiento en Access

**Estado:** en curso · **Hechas:** 6/189

### Preparación

- [ ] Ensayo del cambio de IDCorreo en una copia, con el perfil production-test de Automatización
- [ ] Confirmar el perfil de Automatización que usa el servidor real
- [ ] Validación del plan de aplicación por el propietario (fase-1/plan-aplicacion.md)

### Bloque A. IDCorreo (0/14)

- [ ] [access2web-blueprint#1740](https://github.com/DysTelefonica/access2web-blueprint/issues/1740) refactor(platform): IDCorreo autonumérico en la cola común de correo
- [ ] [access2web-blueprint#1741](https://github.com/DysTelefonica/access2web-blueprint/issues/1741) refactor(platform): IDCorreo autonumérico y clave primaria en TbCorreosEnviados de Correos_datos
- [ ] [AGEDYS#5](https://github.com/DysTelefonica/AGEDYS/issues/5) refactor: dejar de calcular IDCorreo al encolar correos
- [ ] [AGEDYS_ECO#12](https://github.com/DysTelefonica/AGEDYS_ECO/issues/12) refactor: dejar de calcular IDCorreo al encolar correos
- [ ] [Automatizacion#53](https://github.com/DysTelefonica/Automatizacion/issues/53) fix(correo): dejar de calcular IDCorreo y soportar IDCorreo autonumérico en las colas de correo
- [ ] [condor#29](https://github.com/DysTelefonica/condor/issues/29) refactor: dejar de calcular IDCorreo al encolar correos
- [ ] [Expedientes#279](https://github.com/DysTelefonica/Expedientes/issues/279) refactor: dejar de calcular IDCorreo al encolar correos
- [ ] [GESTION_RIESGOS#338](https://github.com/DysTelefonica/GESTION_RIESGOS/issues/338) refactor: dejar de calcular IDCorreo al encolar correos
- [ ] [HPS#52](https://github.com/DysTelefonica/HPS/issues/52) refactor: dejar de calcular IDCorreo al encolar correos
- [ ] [HPS_SOLICITUDES#2](https://github.com/DysTelefonica/HPS_SOLICITUDES/issues/2) refactor: dejar de calcular IDCorreo al encolar correos
- [ ] [Lanzadera#27](https://github.com/DysTelefonica/Lanzadera/issues/27) refactor: dejar de calcular IDCorreo al encolar correos
- [ ] [No_conformidades#186](https://github.com/DysTelefonica/No_conformidades/issues/186) refactor: dejar de calcular IDCorreo al encolar correos
- [ ] [REGISTROS#2](https://github.com/DysTelefonica/REGISTROS/issues/2) refactor: dejar de calcular IDCorreo al encolar correos
- [ ] [seguridad#10](https://github.com/DysTelefonica/seguridad/issues/10) refactor: dejar de calcular IDCorreo al encolar correos

### Lanzadera (5/12)

- [x] [Lanzadera#24](https://github.com/DysTelefonica/Lanzadera/issues/24) refactor: retirar los vídeos-guías, los cuestionarios y la ayuda de la Lanzadera
- [ ] [Lanzadera#25](https://github.com/DysTelefonica/Lanzadera/issues/25) refactor: retirar el registro de conexiones y de aperturas de la Lanzadera
- [x] [Lanzadera#26](https://github.com/DysTelefonica/Lanzadera/issues/26) refactor: retirar la configuración del directorio de ayuda de las aplicaciones
- [ ] [Lanzadera#28](https://github.com/DysTelefonica/Lanzadera/issues/28) refactor: sustituir la asignación MAX+1 por una fuente de ID atómica
- [ ] [Lanzadera#29](https://github.com/DysTelefonica/Lanzadera/issues/29) refactor: extraer credenciales del código y definir cómo se proporcionan en Citrix
- [ ] [Lanzadera#30](https://github.com/DysTelefonica/Lanzadera/issues/30) refactor: revisar la vida de las conexiones antes de ODBC
- [ ] [Lanzadera#31](https://github.com/DysTelefonica/Lanzadera/issues/31) refactor: decidir qué hacer con getEsUsuarioAutorizado de Contratos antes de conservarla o
- [ ] [Lanzadera#32](https://github.com/DysTelefonica/Lanzadera/issues/32) refactor: declarar con integridad referencial las relaciones que ya cumplen los datos (0
- [x] [Lanzadera#33](https://github.com/DysTelefonica/Lanzadera/issues/33) refactor: borrar las consultas de prueba Consulta1 a Consulta4 del backend.
- [ ] [Lanzadera#34](https://github.com/DysTelefonica/Lanzadera/issues/34) refactor: poner en cuarentena las 13 tablas huérfanas de Lanzadera_Datos
- [x] [Lanzadera#35](https://github.com/DysTelefonica/Lanzadera/issues/35) test: dejar en verde la suite base sobre el binario de producción
- [x] [Lanzadera#39](https://github.com/DysTelefonica/Lanzadera/issues/39) feat: módulo de migraciones del backend ejecutable por entorno (local, pruebas, producción)

### Expedientes (0/14)

- [ ] [Expedientes#278](https://github.com/DysTelefonica/Expedientes/issues/278) refactor: retirar los códigos de compras de expediente
- [ ] [Expedientes#282](https://github.com/DysTelefonica/Expedientes/issues/282) refactor: dejar `TbExpedientesAnualidades` única en `Expedientes_datos`
- [ ] [Expedientes#283](https://github.com/DysTelefonica/Expedientes/issues/283) refactor: declarar clave primaria en las seis tablas que no la tienen
- [ ] [Expedientes#284](https://github.com/DysTelefonica/Expedientes/issues/284) refactor: resolver las tres escrituras del front en tablas de otro backend
- [ ] [Expedientes#285](https://github.com/DysTelefonica/Expedientes/issues/285) refactor: retirar del front los cuatro vínculos que su código no usa
- [ ] [Expedientes#286](https://github.com/DysTelefonica/Expedientes/issues/286) refactor: aislar el vínculo TbExpedientes1
- [ ] [Expedientes#287](https://github.com/DysTelefonica/Expedientes/issues/287) refactor: obtener el SQL de las once consultas guardadas del backend y mapear sus tablas
- [ ] [Expedientes#288](https://github.com/DysTelefonica/Expedientes/issues/288) refactor: reconciliar TbExpedientesDatosAGEDYS
- [ ] [Expedientes#289](https://github.com/DysTelefonica/Expedientes/issues/289) refactor: sustituir el generador DameID (máximo más uno) por una identidad
- [ ] [Expedientes#290](https://github.com/DysTelefonica/Expedientes/issues/290) refactor: declarar la integridad referencial que falta hacia TbUsuariosAplicaciones
- [ ] [Expedientes#291](https://github.com/DysTelefonica/Expedientes/issues/291) refactor: revisar el RowSource residual de Form_FormE2ECola y las rutinas dormidas del Instalador
- [ ] [Expedientes#292](https://github.com/DysTelefonica/Expedientes/issues/292) refactor: decidir el destino de TbAuxEstadosMartina
- [ ] [Expedientes#293](https://github.com/DysTelefonica/Expedientes/issues/293) refactor: poner en cuarentena las 9 tablas huérfanas de Expedientes_datos
- [ ] [Expedientes#294](https://github.com/DysTelefonica/Expedientes/issues/294) refactor: poner en cuarentena las 25 tablas que el VBA del front no usa

### Condor (0/8)

- [ ] [condor#30](https://github.com/DysTelefonica/condor/issues/30) refactor: declarar TbConfiguracionBackends como configuración propia del front
- [ ] [condor#31](https://github.com/DysTelefonica/condor/issues/31) refactor: retirar los cuatro abridores externos sin llamadores
- [ ] [condor#32](https://github.com/DysTelefonica/condor/issues/32) refactor: sustituir los 9 generadores Max + 1 por autonumérico o identidad.
- [ ] [condor#33](https://github.com/DysTelefonica/condor/issues/33) refactor: declarar las cuatro relaciones que faltan en el backend
- [ ] [condor#34](https://github.com/DysTelefonica/condor/issues/34) refactor: reescribir el filtro de usuarios de los dos combos de log sin MSysObjects
- [ ] [condor#35](https://github.com/DysTelefonica/condor/issues/35) refactor: decidir el destino de la contraseña del backend
- [ ] [condor#36](https://github.com/DysTelefonica/condor/issues/36) refactor: poner en cuarentena las 1 tablas huérfanas de condor_datos
- [ ] [condor#37](https://github.com/DysTelefonica/condor/issues/37) refactor: poner en cuarentena tbSandboxLog en el front

### Seguridad (0/12)

- [ ] [seguridad#11](https://github.com/DysTelefonica/seguridad/issues/11) refactor: declarar clave primaria en TbNombramientoJS (17 registros
- [ ] [seguridad#12](https://github.com/DysTelefonica/seguridad/issues/12) refactor: resolver las dos referencias muertas
- [ ] [seguridad#13](https://github.com/DysTelefonica/seguridad/issues/13) refactor: sustituir los dos generadores Max + 1 (src/modules/Constructor.bas:762-774 y
- [ ] [seguridad#14](https://github.com/DysTelefonica/seguridad/issues/14) refactor: declarar la integridad referencial que el ERD no recoge
- [ ] [seguridad#15](https://github.com/DysTelefonica/seguridad/issues/15) refactor: mudar al backend las 4 tablas de trabajo del front que alimentan los cinco formularios
- [ ] [seguridad#16](https://github.com/DysTelefonica/seguridad/issues/16) refactor: retirar la credencial literal de las tres cadenas de conexión (src/modules/Variables de
- [ ] [seguridad#17](https://github.com/DysTelefonica/seguridad/issues/17) refactor: reescribir el SQL que no compila en PostgreSQL
- [ ] [seguridad#18](https://github.com/DysTelefonica/seguridad/issues/18) refactor: parametrizar el SQL que concatena valores
- [ ] [seguridad#19](https://github.com/DysTelefonica/seguridad/issues/19) refactor: retirar o reubicar Tb0AuxOrdenarColeccion (3 registros) y TbAuxURLNoAlcanzables (100
- [ ] [seguridad#20](https://github.com/DysTelefonica/seguridad/issues/20) refactor: decidir el destino del patrón «una tabla por tipo de registro»
- [ ] [seguridad#21](https://github.com/DysTelefonica/seguridad/issues/21) refactor: fijar la semántica de comparación de cadenas
- [ ] [seguridad#22](https://github.com/DysTelefonica/seguridad/issues/22) refactor: poner en cuarentena las 10 tablas huérfanas de Seguridad_datos

### HPS (0/12)

- [ ] [HPS#31](https://github.com/DysTelefonica/HPS/issues/31) [Bug]: remove frontend cache tables: filter results as ListBoxes fed by backend queries
- [ ] [HPS#33](https://github.com/DysTelefonica/HPS/issues/33) [Bug]: retire write-only TbUsuariosEntidades backend table and its dead readers
- [ ] [HPS#53](https://github.com/DysTelefonica/HPS/issues/53) refactor: sustituir los generadores Max + 1 por autonumérico
- [ ] [HPS#54](https://github.com/DysTelefonica/HPS/issues/54) refactor: retirar los 5 vínculos del front que apuntan a la ruta de producción (\\datoste\…)
- [ ] [HPS#55](https://github.com/DysTelefonica/HPS/issues/55) refactor: retirar la credencial literal de las cinco cadenas de conexión
- [ ] [HPS#56](https://github.com/DysTelefonica/HPS/issues/56) refactor: añadir a HPST el vínculo a TbExpedientesSuministradores de Expedientes_datos.accdb
- [ ] [HPS#57](https://github.com/DysTelefonica/HPS/issues/57) refactor: retirar de HPST los 2 vínculos que ninguna aplicación usa
- [ ] [HPS#58](https://github.com/DysTelefonica/HPS/issues/58) refactor: eliminar las ramas muertas de DameIDAnexo que apuntan a TbAnexosHistoricos y
- [ ] [HPS#59](https://github.com/DysTelefonica/HPS/issues/59) refactor: revisar las referencias a tablas inexistentes TbObservacionesLocal y
- [ ] [HPS#60](https://github.com/DysTelefonica/HPS/issues/60) refactor: retirar del front las 4 tablas locales sin ninguna referencia en el repositorio
- [ ] [HPS#61](https://github.com/DysTelefonica/HPS/issues/61) refactor: fijar por escrito el contrato de las 4 cachés vivas del front (TbDatosLocal
- [ ] [HPS#62](https://github.com/DysTelefonica/HPS/issues/62) refactor: poner en cuarentena las 3 tablas huérfanas de HPST

### Solicitudes HPS (0/15)

- [ ] [HPS_SOLICITUDES#1](https://github.com/DysTelefonica/HPS_SOLICITUDES/issues/1) [Bug]: read real HPS data instead of the flattened TbUsuariosEntidades table
- [ ] [HPS_SOLICITUDES#3](https://github.com/DysTelefonica/HPS_SOLICITUDES/issues/3) refactor: retirar el vínculo roto TbUsuarioSolicitud y la consulta guardada Consulta2 que lo usa
- [ ] [HPS_SOLICITUDES#4](https://github.com/DysTelefonica/HPS_SOLICITUDES/issues/4) refactor: resolver la duplicidad de TbHPSGrado
- [ ] [HPS_SOLICITUDES#5](https://github.com/DysTelefonica/HPS_SOLICITUDES/issues/5) refactor: declarar la propiedad de las cuatro tablas de HPST que esta aplicación escribe
- [ ] [HPS_SOLICITUDES#6](https://github.com/DysTelefonica/HPS_SOLICITUDES/issues/6) refactor: retirar la credencial literal de las cinco cadenas OpenDatabase (src/modules/Variables
- [ ] [HPS_SOLICITUDES#7](https://github.com/DysTelefonica/HPS_SOLICITUDES/issues/7) refactor: revisar los 10 vínculos del backend
- [ ] [HPS_SOLICITUDES#8](https://github.com/DysTelefonica/HPS_SOLICITUDES/issues/8) refactor: añadir al backend los vínculos que el front usa por código
- [ ] [HPS_SOLICITUDES#9](https://github.com/DysTelefonica/HPS_SOLICITUDES/issues/9) refactor: retirar del front la tabla TbCorreosEnviadosParaPruebas
- [ ] [HPS_SOLICITUDES#10](https://github.com/DysTelefonica/HPS_SOLICITUDES/issues/10) refactor: separar los dos TbCorreosEnviados
- [ ] [HPS_SOLICITUDES#11](https://github.com/DysTelefonica/HPS_SOLICITUDES/issues/11) refactor: retirar del backend los dos vínculos que ninguna aplicación usa
- [ ] [HPS_SOLICITUDES#12](https://github.com/DysTelefonica/HPS_SOLICITUDES/issues/12) refactor: corregir RegistrarTbMotivoHPS (src/modules/FUNCIONES UTILES.bas:1422-1440)
- [ ] [HPS_SOLICITUDES#13](https://github.com/DysTelefonica/HPS_SOLICITUDES/issues/13) refactor: corregir la generación del identificador de TbConfiguracion
- [ ] [HPS_SOLICITUDES#14](https://github.com/DysTelefonica/HPS_SOLICITUDES/issues/14) refactor: fijar una sola versión declarada
- [ ] [HPS_SOLICITUDES#15](https://github.com/DysTelefonica/HPS_SOLICITUDES/issues/15) refactor: revisar los datos personales incrustados en el front
- [ ] [HPS_SOLICITUDES#16](https://github.com/DysTelefonica/HPS_SOLICITUDES/issues/16) refactor: poner en cuarentena las 2 tablas huérfanas de Solicitudes_HPS_datos

### Gestión de Riesgos (0/10)

- [ ] [GESTION_RIESGOS#339](https://github.com/DysTelefonica/GESTION_RIESGOS/issues/339) refactor: declarar clave primaria en las 5 tablas sin ella
- [ ] [GESTION_RIESGOS#340](https://github.com/DysTelefonica/GESTION_RIESGOS/issues/340) refactor: retirar del backend los 5 vínculos que el front no usa
- [ ] [GESTION_RIESGOS#341](https://github.com/DysTelefonica/GESTION_RIESGOS/issues/341) refactor: resolver las 13 referencias a tablas inexistentes o mal nombradas
- [ ] [GESTION_RIESGOS#342](https://github.com/DysTelefonica/GESTION_RIESGOS/issues/342) refactor: sacar la contraseña del backend de la tabla del front
- [ ] [GESTION_RIESGOS#343](https://github.com/DysTelefonica/GESTION_RIESGOS/issues/343) refactor: sustituir el generador Max + 1 (DameID) por identidad en las 22 tablas listadas.
- [ ] [GESTION_RIESGOS#344](https://github.com/DysTelefonica/GESTION_RIESGOS/issues/344) refactor: fijar la semántica de comparación de texto antes de migrar
- [ ] [GESTION_RIESGOS#345](https://github.com/DysTelefonica/GESTION_RIESGOS/issues/345) refactor: decidir el nivel de aislamiento de las 571 aperturas de recordset
- [ ] [GESTION_RIESGOS#346](https://github.com/DysTelefonica/GESTION_RIESGOS/issues/346) refactor: sustituir los 45 literales de fecha con # de los 8 ficheros citados por parámetros o
- [ ] [GESTION_RIESGOS#347](https://github.com/DysTelefonica/GESTION_RIESGOS/issues/347) refactor: poner en cuarentena las 30 tablas huérfanas de Gestion_Riesgos_Datos
- [ ] [GESTION_RIESGOS#348](https://github.com/DysTelefonica/GESTION_RIESGOS/issues/348) refactor: decidir el tratamiento de las 35 consultas guardadas

### No Conformidades (0/10)

- [ ] [No_conformidades#189](https://github.com/DysTelefonica/No_conformidades/issues/189) refactor: sustituir el generador MAX + 1 (constructor.getID
- [ ] [No_conformidades#190](https://github.com/DysTelefonica/No_conformidades/issues/190) refactor: registrar el dueño real de TbRiesgosNC (backend de Gestión de Riesgos) y de
- [ ] [No_conformidades#191](https://github.com/DysTelefonica/No_conformidades/issues/191) refactor: revisar TbAnexosAuditoria y TbAnexosNCAuditorias
- [ ] [No_conformidades#192](https://github.com/DysTelefonica/No_conformidades/issues/192) refactor: revisar las referencias a tablas que no existen
- [ ] [No_conformidades#193](https://github.com/DysTelefonica/No_conformidades/issues/193) refactor: retirar NCProyectoOperaciones.ActualizarRiesgosNC
- [ ] [No_conformidades#194](https://github.com/DysTelefonica/No_conformidades/issues/194) refactor: revisar las 4 consultas guardadas del backend (ColumnaFalta
- [ ] [No_conformidades#195](https://github.com/DysTelefonica/No_conformidades/issues/195) refactor: decidir cómo se reproduce la comparación sin distinguir mayúsculas
- [ ] [No_conformidades#196](https://github.com/DysTelefonica/No_conformidades/issues/196) refactor: decidir dónde vive TbConfiguracionBackends (tabla de arranque del front
- [ ] [No_conformidades#197](https://github.com/DysTelefonica/No_conformidades/issues/197) refactor: poner en cuarentena las 15 tablas huérfanas de NoConformidades_Datos
- [ ] [No_conformidades#198](https://github.com/DysTelefonica/No_conformidades/issues/198) refactor: poner en cuarentena TbTipologiaAux del front (23 filas

### Registros (0/20)

- [ ] [REGISTROS#1](https://github.com/DysTelefonica/REGISTROS/issues/1) refactor: dejar de registrar aperturas y conexiones en la Lanzadera
- [ ] [REGISTROS#3](https://github.com/DysTelefonica/REGISTROS/issues/3) refactor: unificar el acceso a TbSalidas_RES en el front de Registros
- [ ] [REGISTROS#4](https://github.com/DysTelefonica/REGISTROS/issues/4) refactor: declarar explícitamente que TbSalidas_RES y TbEntradas_RES son de Registro y que AGEDYS
- [ ] [REGISTROS#5](https://github.com/DysTelefonica/REGISTROS/issues/5) refactor: sustituir los generadores Max + 1 por autonumérico o identidad
- [ ] [REGISTROS#6](https://github.com/DysTelefonica/REGISTROS/issues/6) refactor: declarar clave primaria en las tablas del backend propio que no la tengan
- [ ] [REGISTROS#7](https://github.com/DysTelefonica/REGISTROS/issues/7) refactor: reponer los dos vínculos que la matriz declara y el front no tiene
- [ ] [REGISTROS#8](https://github.com/DysTelefonica/REGISTROS/issues/8) refactor: resolver las referencias que no tienen tabla
- [ ] [REGISTROS#9](https://github.com/DysTelefonica/REGISTROS/issues/9) refactor: decidir el destino de Registro_Datos.accdb y de sus 5 tablas
- [ ] [REGISTROS#10](https://github.com/DysTelefonica/REGISTROS/issues/10) refactor: corregir el RowSource de TbExpedientes (columnas Expediente y TÍTULOEXP
- [ ] [REGISTROS#11](https://github.com/DysTelefonica/REGISTROS/issues/11) refactor: declarar en Registro_Ent_Salida_Datos las relaciones que los datos ya cumplen
- [ ] [REGISTROS#12](https://github.com/DysTelefonica/REGISTROS/issues/12) refactor: revisar las 613 filas huérfanas de TbLog.IDEntrada y las 5 tablas de indexación vacías.
- [ ] [REGISTROS#13](https://github.com/DysTelefonica/REGISTROS/issues/13) refactor: retirar la credencial literal de las tres cadenas de apertura (Variables de
- [ ] [REGISTROS#14](https://github.com/DysTelefonica/REGISTROS/issues/14) refactor: retirar la dirección de correo nominal fija escrita en el código (Correo.bas:44)
- [ ] [REGISTROS#15](https://github.com/DysTelefonica/REGISTROS/issues/15) refactor: quitar el vínculo del front a TbSalidas_RES_Indexacion
- [ ] [REGISTROS#16](https://github.com/DysTelefonica/REGISTROS/issues/16) refactor: aislar el desarrollo de la producción
- [ ] [REGISTROS#17](https://github.com/DysTelefonica/REGISTROS/issues/17) refactor: normalizar las rutas
- [ ] [REGISTROS#18](https://github.com/DysTelefonica/REGISTROS/issues/18) refactor: mover las 3 consultas guardadas del front al backend y parametrizar Consulta2 (año
- [ ] [REGISTROS#19](https://github.com/DysTelefonica/REGISTROS/issues/19) refactor: renombrar Tb*_RES a una nomenclatura estable
- [ ] [REGISTROS#20](https://github.com/DysTelefonica/REGISTROS/issues/20) refactor: poner en cuarentena las 5 tablas huérfanas de Registro_Ent_Salida_Datos
- [ ] [REGISTROS#21](https://github.com/DysTelefonica/REGISTROS/issues/21) refactor: poner en cuarentena las 3 tablas huérfanas de Registro_Datos

### Brass (0/7)

- [ ] [BRASS#6](https://github.com/DysTelefonica/BRASS/issues/6) refactor: retirar el código dormido de aperturas de la Lanzadera
- [ ] [BRASS#7](https://github.com/DysTelefonica/BRASS/issues/7) refactor: revisar las referencias a tablas inexistentes (TbDocumentos
- [ ] [BRASS#8](https://github.com/DysTelefonica/BRASS/issues/8) refactor: las 4 tablas que viven en el front (TbActividadesLotes
- [ ] [BRASS#9](https://github.com/DysTelefonica/BRASS/issues/9) refactor: declarar clave primaria en las 5 tablas auxiliares sin ella
- [ ] [BRASS#10](https://github.com/DysTelefonica/BRASS/issues/10) refactor: retirar la credencial literal del código del front (Variables Globales.bas:573) y
- [ ] [BRASS#11](https://github.com/DysTelefonica/BRASS/issues/11) refactor: revisar que el front no escriba en el share de producción desde desarrollo
- [ ] [BRASS#12](https://github.com/DysTelefonica/BRASS/issues/12) refactor: poner en cuarentena las 7 tablas huérfanas de Gestion_Brass_Gestion_Datos

### APGEDYS (0/41)

- [ ] [AGEDYS#3](https://github.com/DysTelefonica/AGEDYS/issues/3) refactor: retirar el código dormido de aperturas y el vínculo a TbConexionesRegistro
- [ ] [AGEDYS#4](https://github.com/DysTelefonica/AGEDYS/issues/4) refactor: retirar el código muerto de códigos de compras de expediente
- [ ] [AGEDYS#6](https://github.com/DysTelefonica/AGEDYS/issues/6) refactor: dejar `TbExpedientesAnualidades` única en `Expedientes_datos`
- [ ] [AGEDYS#7](https://github.com/DysTelefonica/AGEDYS/issues/7) refactor: declarar clave primaria en las 11 tablas del backend que no la tienen
- [ ] [AGEDYS#8](https://github.com/DysTelefonica/AGEDYS/issues/8) refactor: retirar la credencial literal de las ocho cadenas de apertura del front
- [ ] [AGEDYS#9](https://github.com/DysTelefonica/AGEDYS/issues/9) refactor: dejar una sola copia de las cinco tablas que el front y el backend tienen con el mismo
- [ ] [AGEDYS#10](https://github.com/DysTelefonica/AGEDYS/issues/10) refactor: corregir la referencia a TbGTVContratos1 en src/classes/GTVContrato.cls:81
- [ ] [AGEDYS#11](https://github.com/DysTelefonica/AGEDYS/issues/11) refactor: sustituir el generador Max + 1 (DameID1
- [ ] [AGEDYS#12](https://github.com/DysTelefonica/AGEDYS/issues/12) refactor: declarar en el backend las relaciones que el front ya exige por SQL
- [ ] [AGEDYS#13](https://github.com/DysTelefonica/AGEDYS/issues/13) refactor: retirar o reparar las 13 consultas guardadas de depuración
- [ ] [AGEDYS#14](https://github.com/DysTelefonica/AGEDYS/issues/14) refactor: registrar el dueño y reubicar las tres tablas que solo consume otro backend
- [ ] [AGEDYS#15](https://github.com/DysTelefonica/AGEDYS/issues/15) refactor: retirar del front los 67 vínculos que su código no usa (de los 127).
- [ ] [AGEDYS#16](https://github.com/DysTelefonica/AGEDYS/issues/16) refactor: declarar Tabla1
- [ ] [AGEDYS#17](https://github.com/DysTelefonica/AGEDYS/issues/17) refactor: retirar los tres abridores sin llamadores
- [ ] [AGEDYS#18](https://github.com/DysTelefonica/AGEDYS/issues/18) refactor: corregir la condición local/remoto invertida de getdbSalida (:558-563)
- [ ] [AGEDYS#19](https://github.com/DysTelefonica/AGEDYS/issues/19) refactor: parametrizar y portar las 23 consultas guardadas de TbConsultasDetalle
- [ ] [AGEDYS#20](https://github.com/DysTelefonica/AGEDYS/issues/20) refactor: normalizar la nomenclatura de las siete tablas de histórico y auxiliares del backend
- [ ] [AGEDYS#21](https://github.com/DysTelefonica/AGEDYS/issues/21) refactor: poner en cuarentena las 57 tablas huérfanas de AGEDYS_DATOS
- [ ] [AGEDYS#22](https://github.com/DysTelefonica/AGEDYS/issues/22) refactor: poner en cuarentena Copia de TbRequisitosSTG del front
- [ ] [AGEDYS#23](https://github.com/DysTelefonica/AGEDYS/issues/23) refactor: poner en cuarentena TbAuxNPedido del front
- [ ] [AGEDYS_ECO#10](https://github.com/DysTelefonica/AGEDYS_ECO/issues/10) refactor: retirar el código dormido de aperturas de la Lanzadera
- [ ] [AGEDYS_ECO#11](https://github.com/DysTelefonica/AGEDYS_ECO/issues/11) refactor: retirar la gestión de códigos de compras de expediente
- [ ] [AGEDYS_ECO#13](https://github.com/DysTelefonica/AGEDYS_ECO/issues/13) refactor: dejar `TbExpedientesAnualidades` única en `Expedientes_datos`
- [ ] [AGEDYS_ECO#14](https://github.com/DysTelefonica/AGEDYS_ECO/issues/14) refactor: unificar el acceso a TbSalidas_RES
- [ ] [AGEDYS_ECO#15](https://github.com/DysTelefonica/AGEDYS_ECO/issues/15) refactor: migrar los 24 recordsets de tipo tabla a sentencias SELECT explícitas.
- [ ] [AGEDYS_ECO#16](https://github.com/DysTelefonica/AGEDYS_ECO/issues/16) refactor: mudar al backend las 7 tablas de trabajo de la función «Consultas» (TbAuxProyectos
- [ ] [AGEDYS_ECO#17](https://github.com/DysTelefonica/AGEDYS_ECO/issues/17) refactor: retirar la credencial literal de las siete cadenas de apertura (src/modules/Variables
- [ ] [AGEDYS_ECO#18](https://github.com/DysTelefonica/AGEDYS_ECO/issues/18) refactor: corregir la referencia muerta TbDPDDocAnexada (src/classes/AnexoOperaciones.cls:32)
- [ ] [AGEDYS_ECO#19](https://github.com/DysTelefonica/AGEDYS_ECO/issues/19) refactor: corregir los dos generadores de identificador que apuntan al front
- [ ] [AGEDYS_ECO#20](https://github.com/DysTelefonica/AGEDYS_ECO/issues/20) refactor: declarar la clave primaria de TbAuxGTVContratos
- [ ] [AGEDYS_ECO#21](https://github.com/DysTelefonica/AGEDYS_ECO/issues/21) refactor: declarar en PostgreSQL las 43 tablas que el front necesita
- [ ] [AGEDYS_ECO#22](https://github.com/DysTelefonica/AGEDYS_ECO/issues/22) refactor: sustituir los dos generadores Max + 1 (src/modules/FUNCIONES UTILES.bas:534 y
- [ ] [AGEDYS_ECO#23](https://github.com/DysTelefonica/AGEDYS_ECO/issues/23) refactor: reescribir el SQL que no compila en PostgreSQL
- [ ] [AGEDYS_ECO#24](https://github.com/DysTelefonica/AGEDYS_ECO/issues/24) refactor: parametrizar el SQL que concatena valores
- [ ] [AGEDYS_ECO#25](https://github.com/DysTelefonica/AGEDYS_ECO/issues/25) refactor: decidir el destino de la consulta guardada Consulta1 del front
- [ ] [AGEDYS_ECO#26](https://github.com/DysTelefonica/AGEDYS_ECO/issues/26) refactor: retirar el código muerto
- [ ] [AGEDYS_ECO#27](https://github.com/DysTelefonica/AGEDYS_ECO/issues/27) refactor: verificar la resolución de las rutas de getdbSalida()
- [ ] [AGEDYS_ECO#28](https://github.com/DysTelefonica/AGEDYS_ECO/issues/28) refactor: normalizar la matriz de vínculos de la Lanzadera para AGEDYS ECO
- [ ] [AGEDYS_ECO#29](https://github.com/DysTelefonica/AGEDYS_ECO/issues/29) refactor: fijar la semántica de comparación de cadenas
- [ ] [AGEDYS_ECO#30](https://github.com/DysTelefonica/AGEDYS_ECO/issues/30) refactor: decidir el destino de Option Compare Database y de los 215 recordsets sin opciones de
- [ ] [AGEDYS_ECO#31](https://github.com/DysTelefonica/AGEDYS_ECO/issues/31) refactor: poner en cuarentena las 3 tablas del front sin uso (TbAuxExpPrado

### Transversal (1/11)

- [ ] [access2web-blueprint#1752](https://github.com/DysTelefonica/access2web-blueprint/issues/1752) refactor: localizar el front de Tareas o declarar el backend como material archivado
- [ ] [access2web-blueprint#1753](https://github.com/DysTelefonica/access2web-blueprint/issues/1753) refactor: fijar clave primaria en TbAperturas (7 800 filas) y en las otras tres tablas de SICA
- [ ] [access2web-blueprint#1754](https://github.com/DysTelefonica/access2web-blueprint/issues/1754) refactor: declarar TbUsuariosAplicacionesTareas como tabla de Tareas
- [ ] [access2web-blueprint#1755](https://github.com/DysTelefonica/access2web-blueprint/issues/1755) refactor: fijar clave primaria en las ocho tablas sin clave de AGEDO
- [ ] [access2web-blueprint#1758](https://github.com/DysTelefonica/access2web-blueprint/issues/1758) refactor: declarar las relaciones de Control_Cambios_datos con integridad referencial
- [ ] [access2web-blueprint#1759](https://github.com/DysTelefonica/access2web-blueprint/issues/1759) refactor: revisar si SICA y Control de cambios siguen teniendo consumidor
- [ ] [access2web-blueprint#1760](https://github.com/DysTelefonica/access2web-blueprint/issues/1760) refactor: revisar la matriz de vínculos de la lanzadera para este par de backends
- [x] [access2web-blueprint#1790](https://github.com/DysTelefonica/access2web-blueprint/issues/1790) docs(migracion): plan de aplicación del saneamiento de la fase 1 y decisión 29
- [ ] [CONTRATOS#1](https://github.com/DysTelefonica/CONTRATOS/issues/1) refactor: dejar de registrar aperturas en la Lanzadera
- [ ] [CONTRATOS#2](https://github.com/DysTelefonica/CONTRATOS/issues/2) refactor: retirar la cola local de correo y la clase Correo sin uso
- [ ] [CONTRATOS#3](https://github.com/DysTelefonica/CONTRATOS/issues/3) refactor: poner en cuarentena las 1 tablas huérfanas de GestionContratos_datos

## Fase 2. Unificar el acceso a datos

**Estado:** pendiente · **Hechas:** 0/5

- [ ] Adaptar a ODBC sin DSN getdb() con OpenDatabase y clave (Lanzadera)
- [ ] Adaptar a ODBC sin DSN la revinculación desde TbTablasAVincular (aplicaciones antiguas)
- [ ] Adaptar a ODBC sin DSN el resolvedor propio de cada aplicación moderna
- [ ] Adaptar a ODBC sin DSN la apertura directa de Lanzadera_Datos con ruta escrita en el código (11 fronts)
- [ ] Cambio de .accdb a PostgreSQL, y vuelta, solo por configuración, demostrado fuera de producción en cada front

## Fase 3. Ensayo en Citrix con todos los backends

**Estado:** revisada (diseño revisado, pendiente de ejecutar) · **Hechas:** 0/7

- [ ] Validación local de Citrix con RDS Session Host (H-3, primer paso)
- [ ] Ensayo completo en preproducción con todos los backends sobre PostgreSQL
- [ ] Conversión de tipos verificada
- [ ] Tablas vinculadas actualizables (con clave primaria)
- [ ] Todas las aplicaciones Access funcionando sobre PostgreSQL
- [ ] Rendimiento aceptable en Citrix
- [ ] Hipótesis H-1, H-3 y H-4 confirmadas o refutadas con evidencia

## Fase 4. Traslado único de todos los datos a PostgreSQL

**Estado:** pendiente · **Hechas:** 0/5

- [ ] Plan del traslado
- [ ] Repetición del ensayo de la fase 3 con una copia reciente de producción
- [ ] Reapuntado de los consumidores que no migran: AGEDO, Tareas y el envío de correos del servidor (decisiones 14 y 16)
- [ ] Decisión sobre cada despliegue de prueba que comparte la unidad de red con producción
- [ ] Traslado único en producción

## Fase 5. Expedientes a web

**Estado:** pendiente · **Hechas:** 0/3

- [ ] Entregables del corte que fija la coexistencia de Expedientes (03-aplicaciones/expedientes/prd/coexistencia.md)
- [ ] Botón de Expedientes deshabilitado en la Lanzadera Access y habilitado en la web
- [ ] Reversión probada: rehabilitar el botón de Access
