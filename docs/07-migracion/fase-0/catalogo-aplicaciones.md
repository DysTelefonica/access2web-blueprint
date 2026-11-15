# Fase 0: catálogo de aplicaciones de la lanzadera

Issue #1581, tarea T17. Épica `docs/07-migracion/epica-coexistencia-postgresql.md` (fase 0).
Fuente: tabla `TbAplicaciones` de `Lanzadera_Datos`, leída con DAO en solo lectura sobre una copia temporal. La tabla es la autoridad de lo que la lanzadera puede abrir: 22 filas, de las que 17 son aplicaciones de producción, 4 son despliegues de prueba y 1 es una aplicación antigua.

## Aplicaciones de producción

| Id | Nombre | Front | Backend | Visible en la lanzadera | En el alcance |
|---|---|---|---|---|---|
| 1 | Sirdee | `GestionPS.accdb` | `GestionPS_Datos.accdb` | No | No (decisión 6: otro departamento) |
| 2 | AGEDYS ECO | `Gestion Proyectos Eco 30.accdb` | `AGEDYS_DATOS.accdb` | Sí | Sí, como parte de APGEDYS |
| 3 | AGEDYS TÉCNICO | `GestionTecnicaProyectos.accdb` | `AGEDYS_DATOS.accdb` | Sí | Sí, como parte de APGEDYS |
| 4 | AGEDO | `Agedo20.accdb` | `AGEDO20_Datos.accdb` | Sí | No (decisión 6: no se usa) |
| 5 | Gestión de Riesgos | `Gestion_Riesgos.accdb` | `Gestion_Riesgos_Datos.accdb` | Sí | Sí |
| 6 | Brass | `Gestion_Brass_Gestion.accdb` | `Gestion_Brass_Gestion_Datos.accdb` | Sí | Sí |
| 8 | No Conformidades | `NoConformidades.accdb` | `NoConformidades_Datos.accdb` | Sí | Sí |
| 10 | Registro | `Registro_Ent_Salida.accdb` | `Registro_Ent_Salida_Datos.accdb` | Sí | Sí |
| 12 | Lanzadera | `Lanzadera.accdb` | `Lanzadera_Datos.accdb` | No (es la lanzadera) | Sí, y es la primera en coexistir |
| 16 | Seguridad | `Seguridad.accdb` | `Seguridad_datos.accdb` | Sí | Sí |
| 17 | HPS | `HPS.accdb` | `HPST.accdb` | Sí | Sí |
| 18 | Gestión Contratos | `GestionContratos.accdb` | `ConClaTab.accdb` | Sí | No |
| 19 | Expedientes | `Expedientes.accdb` | `Expedientes_datos.accdb` | Sí | Sí |
| 20 | SICA | `SICA.accdb` | `SICA_Datos.accdb` | Sí | No |
| 21 | **AGEDYS** | `AGEDYS.accdb` | `AGEDYS_DATOS.accdb` | **No** | **Residuo, fuera** |
| 22 | Solicitudes HPS | `Solicitudes_HPS.accdb` | `Solicitudes_HPS_datos.accdb` | Sí | Sí |
| 23 | Condor | `Condor.accdb` | `Condor_datos.accdb` | Sí | Sí |

**En el alcance: doce**, contando la Lanzadera. Once de las capturas del propietario más la lanzadera misma.

## Aplicaciones que el propietario marcó como fuera del alcance

- **Sirdee** (id 1): otro departamento, por decisión 6 de la épica.
- **AGEDO** (id 4): por decisión 6 de la épica, «no se usa». Sus tablas (`TbDocumentos`, `TbDocumentosID`) las consume Gestión de Riesgos, que sí migra.
- **SICA** (id 20) y **Gestión Contratos** (id 18): por decisión del propietario sobre el catálogo.

## Lo que el volcado destapa

### 1. Existe una tercera aplicación AGEDYS, invisible y sin checkout

La fila 21 declara `AGEDYS.accdb` con backend `AGEDYS_DATOS.accdb` y `ConIconoEnLanzadera = No`: por eso no aparece en la rejilla de la lanzadera. **No se ha encontrado ningún checkout de `AGEDYS.accdb`** en `C:\00repos\codigo`, y su backend es el mismo que comparten las dos mitades de APGEDYS.

Esto tiene dos lecturas y hay que resolver cuál: o es la aplicación matriz de la que cuelgan AGEDYS TÉCNICO y AGEDYS ECO, y entonces hereda las 102 tablas de `AGEDYS_DATOS`; o es un residuo retirado.

**Resuelto (2026-11-05, decisión 15 de la épica):** el propietario confirma que es un residuo que no se usa. No entra en la migración y el reparto de `AGEDYS_DATOS` se hace solo entre las dos mitades de APGEDYS.

### 2. «RegistroAntiguo», un segundo registro en el mismo backend

La fila 101 declara `Registro.accdb` con backend `Registro_Datos.accdb` y sin icono en la lanzadera. La ficha de [Registros](registros.md) ya detectó que ese bloque está roto, no muerto: el catálogo de `TbAplicaciones` distingue la aplicación 10 (Registro) de la 101 (RegistroAntiguo). Como `Registro_Datos.accdb` es el backend que consume `Registro_Ent_Salida_Datos`, la decisión es si se retira o se conserva.

### 3. Cuatro despliegues de prueba, y un backend que no estaba inventariado

| Id | Nombre | Front | Backend |
|---|---|---|---|
| 231 | CONDOR PRUEBA | `Condor_prueba.accdb` | **`condor_datos_prueba.accdb`** |
| 81 | No Conformidades Prueba | `NoConformidades.accdb` | `NoConformidades_Datos.accdb` |
| 191 | EXPEDIENTES PRUEBA | `Expedientes.accdb` | `Expedientes_datos.accdb` |
| 171 | HPS PRUEBA | `HPS.accdb` | `HPST.accdb` |

Tres comparten backend con su aplicación de producción, así que no añaden inventario. **Condor PRUEBA tiene backend propio**, `condor_datos_prueba.accdb`, que no está en `C:\00repos\datos` y que no se ha inventariado.

### 4. Dos discrepancias con lo ya inventariado

- El catálogo asigna a Gestión Contratos el backend `ConClaTab.accdb`, **no** `GestionContratos_datos.accdb`, que es el que se inventarió. Y `Control_Cambios_datos.accdb` no aparece en el catálogo: existe como fichero, pero ninguna aplicación lo declara.
- El backend `Tareas_datos1.accdb` existe y enlaza 13 tablas de cuatro backends, pero **no tiene fila en el catálogo**: no hay aplicación Tareas en la lanzadera. Es un backend sin aplicación conocida.

### 5. La credencial de cada aplicación está en el catálogo

Las 22 filas guardan la contraseña del backend en el campo `Pass`, en claro. Es la misma credencial de los vínculos y del código, y ya está registrada como hallazgo de seguridad en la ficha de Brass (`B4`). No se reproduce aquí.

## Consecuencia para el alcance

El catálogo confirma las once aplicaciones del propietario más la lanzadera, y añade **tres puntos**: la tercera aplicación AGEDYS, el RegistroAntiguo y el backend de Condor PRUEBA. Ninguno cambia el alcance. El primero queda resuelto como residuo (decisión 15); los otros dos siguen abiertos.

## Lo que no se pudo verificar

1. El contenido de `condor_datos_prueba.accdb`: no está en `C:\00repos\datos` ni hay checkout.
2. Si `ConClaTab.accdb` es el backend real de Contratos o el nombre cambió: no hay checkout de ese nombre.
3. Las ediciones desplegadas por aplicación (`TbAplicacionesEdiciones`) se leyeron pero no se inventariaron: son miles de filas de historial de publicación, útiles para el orden de corte pero no para el inventario de tablas.
