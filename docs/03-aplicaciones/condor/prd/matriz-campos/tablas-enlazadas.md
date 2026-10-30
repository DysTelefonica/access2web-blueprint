# Tablas enlazadas de Condor

> Parte de la [matriz campo a campo](../matriz-campos.md). Commit fijado `c88347e`.

Nueve de las tablas que Condor usa no están en su backend: son tablas de **otras aplicaciones**, enlazadas desde el fichero principal. El volcado del ERD las declara como `LINKED/SKIPPED` porque los ficheros externos no estaban en la máquina de extracción, así que aquí no hay columnas que listar: la evidencia es la cita en `src`, que es lo que permite afirmar que la aplicación depende de ellas.

Esto convierte la coexistencia en un problema de datos compartidos y no solo de código: mientras Condor siga enlazando las tablas de identidad de Lanzadera y las de expediente, cada cambio de esquema en esas aplicaciones rompe a Condor.

## Tablas enlazadas

| Tabla | Aplicación propietaria | Evidencia en `src` | Riesgo para la migración |
|---|---|---|---|
| `TbCorreosEnviados` | Correos | `CorreoRepositorio.bas:18` | el registro de correo se crea en Condor, pero la tabla es de la aplicación de Correos |
| `TbExpedientes` | Lanzadera | `ExpedienteRepositorio.bas:18` · `ExpedienteServicio.cls:35` · `UsuarioServicio.cls:117` | el expediente es el agregado padre de una solicitud y Condor lo lee para componer el código |
| `TbExpedientesRACS` | Lanzadera | `ExpedienteRepositorio.bas:24` · `ExpedienteServicio.cls:41` | de aquí sale el RAC asignado por expediente, que se muestra en el visor de estado |
| `TbExpedientesResponsables` | Lanzadera | `ExpedienteRepositorio.bas:21` · `ExpedienteServicio.cls:38` · `UsuarioServicio.cls:117` | de aquí salen los responsables técnico y de calidad de las notificaciones |
| `TbExpedientesSuministradores` | Lanzadera | `ExpedienteRepositorio.bas:59` · `SolicitudServicio.cls:150` | de aquí salen los suministros del expediente |
| `TbNoConformidades` | No Conformidades | `NoConformidadRepositorio.bas:18` | Condor escribe en esta tabla el código de la concesión asociada al formalizar |
| `TbRACS` | Lanzadera | `ExpedienteRepositorio.bas:24` · `ExpedienteServicio.cls:41` | catálogo de RAC, leído para el nombre del asignado |
| `TbSuministradores` | Lanzadera | `SuministradorRepositorio.bas:16` · `SuministradorServicio.cls:25` | catálogo de suministros, leído para los valores por defecto del bloque |
| `TbUsuariosAplicaciones` | Lanzadera | `UsuarioRepositorio.bas:41` · `ExpedienteRepositorio.bas:19` · `ExpedienteServicio.cls:36` | tabla de usuarios e identidad de la que sale el usuario de la sesión |
| `TbUsuariosAplicacionesPermisos` | Lanzadera | `AplicacionRepositorio.bas:24` · `UsuarioServicio.cls:135` | de aquí salen los permisos que se traducen a rol |

## Lectura y escritura por tabla

| Tabla | Lee | Escribe | Dónde |
|---|---|---|---|
| `TbCorreosEnviados` | sí | no | `NotificacionServicio` |
| `TbExpedientes` | sí | no | `ExpedienteServicio` y `ExpedienteRepositorio` |
| `TbExpedientesRACS` | sí | no | `ExpedienteServicio.GetRacAsignadoPorExpediente` |
| `TbExpedientesResponsables` | sí | no | `ExpedienteServicio` |
| `TbExpedientesSuministradores` | sí | no | `ExpedienteServicio.getSuministradoresPorExpediente` |
| `TbNoConformidades` | sí | sí, al formalizar | `NoConformidadServicio` |
| `TbRACS` | sí | no | `ExpedienteRepositorio.GetRacAsignadoPorExpediente` |
| `TbSuministradores` | sí | no | `SuministradorServicio` |
| `TbUsuariosAplicaciones` | sí | no | `UsuarioServicio.getUsuarioConPermisos` |
| `TbUsuariosAplicacionesPermisos` | sí | no | `AplicacionRepositorio.getAplicacionesPermisos` |

## Qué decide el destino

1. **Identidad.** `TbUsuariosAplicaciones` y `TbUsuariosAplicacionesPermisos` son de Lanzadera y son la fuente de la sesión y del rol. La decisión de qué modelo de identidad adopte el roadmap afecta a Condor antes que a ninguna otra de sus decisiones de datos.
2. **Expediente.** `TbExpedientes` y sus tres tablas de relación son el agregado padre. Condor solo las lee, pero compone el código de la solicitud a partir de un campo de texto del expediente (`ExpedienteServicio.getClaveParaCodigoSolicitud`), de modo que un cambio de formato en Lanzadera altera los códigos ya emitidos.
3. **Escritura entre bases.** `TbNoConformidades` es la única tabla externa que Condor escribe, y lo hace guardando una cadena con el código de la concesión. Ese vínculo por texto debería pasar a una clave estable.
4. **Correo.** Condor registra en la tabla de Correos y no envía: la frontera entre las dos aplicaciones está en el registro, no en el envío.
