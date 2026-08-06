# Capacidad: usuarios, permisos y navegaciÃ³n

## Â§0 Identidad
- **ID de capacidad**: `CAP-USERS-PERMS-NAV`
- **Tier**: critical
- **Estado**: active / inventario documental inicial
- **Source**: reverse-engineered
- **Responsable / autoridad de producto**: Pendiente de confirmaciÃ³n â€” administraciÃ³n de aplicaciÃ³n
- **Ãšltima verificaciÃ³n**: 2026-06-15 mediante inspecciÃ³n estÃ¡tica; no se ejecutÃ³ Dysflow/Access
- **Confianza global**: mixta â€” reglas bÃ¡sicas visibles en cÃ³digo; matriz completa pendiente

## Â§1 IntenciÃ³n de negocio
- **PropÃ³sito**: Controlar quÃ© usuarios acceden a menÃºs, altas, configuraciÃ³n, vistas de proyecto/auditorÃ­a e indicadores, y ofrecer navegaciÃ³n entre dominios.
- **Usuarios / perfiles**: Administradores, calidad, tÃ©cnicos, secretarÃ­a/economÃ­a si procede, usuarios sin acceso y revisores UAT.
- **Problema que resuelve**: Evita que acciones sensibles se ejecuten por perfiles no autorizados y hace reproducible el recorrido de la aplicaciÃ³n.
- **Valor de negocio / por quÃ© existe**: La aplicaciÃ³n contiene flujos crÃ­ticos; navegaciÃ³n y permisos deben ser explÃ­citos para evitar regresiones de seguridad.
- **No-objetivos**: No sustituye a una matriz de IAM corporativa ni define autenticaciÃ³n fuera de Access.
- **Origen de la intenciÃ³n**: CÃ³digo exportado; reglas de producto pendientes.
- **Referencia de tracker de origen**: Issue #67; tests de rutas/backend y cache readiness como evidencia adyacente.

## Â§2 Contrato de comportamiento

### Escenarios (Dado / Cuando / Entonces)
- **DADO** que el usuario abre el menÃº principal **CUANDO** elige Parte de Proyectos o AuditorÃ­as **ENTONCES** se abre el menÃº de dominio correspondiente.
- **DADO** que un usuario tÃ©cnico intenta crear NC Proyecto, AuditorÃ­a o NC AuditorÃ­a **CUANDO** ejecuta la acciÃ³n de alta **ENTONCES** la acciÃ³n se bloquea con â€œNo tiene autorizaciÃ³n para esa acciÃ³nâ€.
- **DADO** que un usuario abre el menÃº en modo pruebas **CUANDO** es administrador **ENTONCES** la cinta puede mostrarse; si no, se oculta.
- **DADO** que un tÃ©cnico abre gestiÃ³n de NC Proyecto **CUANDO** se precarga el listado **ENTONCES** se filtra por `ResponsableTelefonica = m_ObjUsuarioConectado.Nombre`.
- **DADO** que un usuario no tÃ©cnico abre gestiÃ³n de NC AuditorÃ­a **CUANDO** se precarga el listado **ENTONCES** se filtra por `RESPONSABLEIMPLANTACION = m_ObjUsuarioConectado.Nombre`.

> **PrecondiciÃ³n transversal**: cualquier prueba runtime de BR-UPN-* requiere backend y cachÃ© en estado seguro. Ver `configuration-backends-runtime` (BR-CFG-5 `AssertSafeBackendForCatalogBootstrap` y BR-CFG-6 auditorÃ­a de routing/kill-switch/indicadores) antes de ejecutar suites contra `TbUsuariosAplicaciones` o `m_ObjUsuarioConectado`. Sin esa precondiciÃ³n, las pruebas de permisos pueden ejecutarse contra un backend inseguro.

### Reglas de negocio
| ID regla | Enunciado (pretendido) | Autoridad | Â¿Aplicada en cÃ³digo? | Prueba | Confianza |
|---|---|---|---|---|---|
| BR-UPN-1 | El menÃº principal enruta a Proyecto y AuditorÃ­as mediante formularios dedicados. | CÃ³digo exportado | SÃ­ â€” `Form_Form0BDOpciones.cls:15` (`DoCmd.OpenForm "Form0BDOpcionesParteProyectos"`) y `Form_Form0BDOpciones.cls:71` (`DoCmd.OpenForm "Form0BDOpcionesAuditorias"`) | FALTA â†’ crear mediante `access-vba-tdd` como contrato de navegaciÃ³n/cableado (sin UI) | Verified-static |
| BR-UPN-2 | Usuario tÃ©cnico no puede ejecutar altas sensibles de Proyecto/AuditorÃ­a. | CÃ³digo exportado | SÃ­ â€” checks `EsTecnico = EnumSino.SÃ­` en `Form_Form0BDOpcionesParteProyectos.cls:46,104,142` y `Form_Form0BDOpcionesAuditorias.cls:50,168,265`; tambiÃ©n `Form_Form0BDTecnicos.cls:124` | FALTA â†’ crear mediante `access-vba-tdd` con fixtures de `TbUsuariosAplicaciones` + inyecciÃ³n de `m_ObjUsuarioConectado` y asserts sobre mensaje de autorizaciÃ³n | Verified-static |
| BR-UPN-3 | Solo administrador ve Ribbon en modo pruebas; en uso normal se oculta. | CÃ³digo exportado | SÃ­ â€” `EsAdministrador = EnumSino.SÃ­` combinado con `PermisoPruebas` en `Form_Form0BDOpciones.cls:115,132`; `PermisoPruebas` declarado en `src/classes/Usuario.cls:36` | FALTA â†’ crear mediante `access-vba-tdd` con coste vÃ­a stub de `m_ObjUsuarioConectado` y asserts sobre visibilidad de Ribbon | Verified-static |
| BR-UPN-4 | La gestiÃ³n de Proyecto precarga NC abiertas y filtra al tÃ©cnico por su nombre. | CÃ³digo exportado | SÃ­ â€” `Form_Form0BDOpcionesParteProyectos.cls:142-143` filtra `Forms("FormNCProyectoGestion").ResponsableTelefonica = m_ObjUsuarioConectado.Nombre` | FALTA â†’ crear mediante `access-vba-tdd` con fixtures de NC y asserts sobre `ResponsableTelefonica` precargado | Verified-static |
| BR-UPN-5 | La gestiÃ³n de AuditorÃ­a precarga NC abiertas y filtra responsable de implantaciÃ³n para no tÃ©cnicos. | CÃ³digo exportado | SÃ­ â€” `Form_Form0BDOpcionesAuditorias.cls:140-141` filtra `Forms("FormNCAuditoriaGestion").RESPONSABLEIMPLANTACION = m_ObjUsuarioConectado.Nombre` | FALTA â†’ crear mediante `access-vba-tdd` con fixtures de NC de auditorÃ­a y asserts sobre el filtro | Verified-static |
| BR-UPN-6 | Roles calculados de usuario: 7 flags `EsUsuario*` (`Administrador`, `Calidad`, `Economia`, `Secretaria`, `Tecnico`, `SinAcceso`, `CalidadAvisos`) en `UsuarioAplicacionPermisos` + `PermisoPruebas` en `Usuario` (8 flags/permisos calculados totales). | CÃ³digo exportado | SÃ­ â€” `src/classes/UsuarioAplicacionPermisos.cls:15-21` (7 flags `EsUsuario*`) y `src/classes/Usuario.cls:36` (`PermisoPruebas`) | FALTA â†’ crear mediante `access-vba-tdd` con fixtures de permisos por rol, asserts sobre cada `*Calculado` y `PermisoPruebas` | Verified-static |
| BR-UPN-7 `#72` | La matriz completa de permisos por acciÃ³n sensible (cerrar/eliminar/rehabilitar/documento/acciÃ³n/informe/configuraciÃ³n) estÃ¡ aprobada por producto. | Producto pendiente | Desconocido | FALTA â†’ crear mediante `access-vba-tdd` tras confirmar matriz; misma matriz referenciada por `cross-cutting-support` BR-XCUT-6 | Intended |

> **Estado de cobertura runtime**: a la fecha de esta revisiÃ³n (2026-06-15) **no existe ningÃºn manifest de pruebas** (`tests/tests.vba*.json`) que cubra permisos, roles calculados, navegaciÃ³n de menÃºs ni bloqueo por rol. Cualquier afirmaciÃ³n `Verified-runtime` para BR-UPN-1..6 estÃ¡ **fuera de alcance** hasta que se creen las pruebas con `access-vba-tdd` (schema-first, fixtures deterministas de `TbUsuariosAplicaciones`, inyecciÃ³n controlada de `m_ObjUsuarioConectado`, asserts sobre mensajes de formulario y estados de control). Ver tambiÃ©n `cross-cutting-support` Â§2 BR-XCUT-6 y Â§5 sobre la ausencia de manifest dedicado.

### Validaciones
- Usuario tÃ©cnico bloqueado en altas sensibles.
- Usuario sin permiso de administraciÃ³n no debe acceder a configuraciÃ³n sensible.
- La navegaciÃ³n de menÃºs debe cerrar/abrir formularios de dominio de forma predecible.
- Cualquier regla de permisos no leÃ­da en cÃ³digo queda como obligaciÃ³n abierta.

### Transiciones de estado
- `MenÃº principal` --(`Parte Proyecto`)--> `MenÃº Proyecto`.
- `MenÃº principal` --(`AuditorÃ­as`)--> `MenÃº AuditorÃ­as`.
- `TÃ©cnico` --(`Alta NC/Alta auditorÃ­a`)--> `AcciÃ³n bloqueada`.
- `Administrador en pruebas` --(`Abrir menÃº`)--> `Ribbon visible`.

### Casos lÃ­mite y de error
- `m_ObjUsuarioConectado` ausente romperÃ­a filtros/captions; necesita pruebas de arranque o guardas.
- El comportamiento de â€œusuario sin accesoâ€ existe como rol, pero no se ha inventariado dÃ³nde bloquea navegaciÃ³n.

### SeÃ±ales de aceptaciÃ³n / presencia
- MenÃºs abren los formularios correctos y aplican filtros iniciales.
- Las acciones sensibles tienen pruebas de rol positivo y negativo.
- La matriz de permisos queda documentada con `Verified-runtime` solo tras pruebas.

## Â§3 Mapa de implementaciÃ³n
- **Puntos de entrada de UI**: `Form_Form0BDOpciones`, `Form_Form0BDOpcionesParteProyectos`, `Form_Form0BDOpcionesAuditorias`, `Form_Form0BDTecnicos`.
- **Puntos de entrada de cÃ³digo**: `Usuario`, `UsuarioAplicacionPermisos`, helpers/globales `EsTecnico`, `EsAdministrador`, `m_ObjUsuarioConectado`, `Entorno.TituloUsuarioConectado`.
- **Datos afectados**: `TbUsuariosAplicaciones`, permisos de aplicaciÃ³n, tablas exactas de permisos pendientes de esquema.
- **Salidas**: formularios abiertos/cerrados, filtros iniciales, mensajes de autorizaciÃ³n, Ribbon visible/oculta.
- **Dependencias e integraciones**: todas las capacidades de dominio.
- **SincronizaciÃ³n fuenteâ†”binario**: no comprobada; tarea solo documental.
- **ValoraciÃ³n de diseÃ±o**: navegaciÃ³n y permisos estÃ¡n acoplados a formularios/globales. Para web deben convertirse en rutas, guards y permisos de dominio explÃ­citos.

## Â§4 Receta de reconstrucciÃ³n
1. Confirmar matriz de permisos por acciÃ³n y rol.
2. Inspeccionar esquema de usuarios/permisos antes de fixtures.
3. Crear pruebas de rol calculado y de bloqueo/autorizaciÃ³n por comando sensible.
4. Crear pruebas de navegaciÃ³n de menÃºs como contrato de formulario/costura, sin automatizaciÃ³n UI innecesaria.
5. Registrar cada nueva regla en esta pÃ¡gina y en la matriz de huecos.

## Â§5 Evidencia y trazabilidad
- **Tests**: **no se localizÃ³ manifest dedicado a navegaciÃ³n/permisos/roles**. BÃºsqueda en `tests/tests.vba*.json` (todos los manifests) no devuelve ninguna coincidencia para `permisos`, `EsTecnico`, `EsAdministrador`, `PermisoPruebas`, `UsuarioAplicacionPermisos`, `TbUsuariosAplicaciones`, `ResponsableTelefonica` ni `RESPONSABLEIMPLANTACION`. La Ãºnica evidencia adyacente estÃ¡ en pruebas de backend/configuraciÃ³n (`tests.vba.e2e.json`, `tests.vba.cache-readiness.json`) y en manifests de formularios helper. Cualquier promociÃ³n a `Verified-runtime` para BR-UPN-1..6 estÃ¡ bloqueada hasta que se creen las pruebas mediante `access-vba-tdd` con schema-first, fixtures deterministas de `TbUsuariosAplicaciones`, inyecciÃ³n controlada de `m_ObjUsuarioConectado` y asserts sobre mensajes/estados de formulario.
- **PrecondiciÃ³n para ejecutar pruebas de permisos**: ver `configuration-backends-runtime` BR-CFG-5 (`AssertSafeBackendForCatalogBootstrap`) y BR-CFG-6 (auditorÃ­a de routing/kill-switch/indicadores) â€” sin esa base, las pruebas de BR-UPN-* pueden ejecutarse contra un backend inseguro.

| Elemento | Ref. tracker | VersiÃ³n de staging (UAT) | Estado UAT | Release de producciÃ³n | Fecha en producciÃ³n | Nota |
|---|---|---|---|---|---|---|
| NavegaciÃ³n Proyecto/AuditorÃ­a | Issue #67 | Pendiente | pending | Pendiente | Pendiente | Falta prueba de contrato. Cableado visible en `src/forms/Form_Form0BDOpciones.cls:15,71`. |
| Bloqueo a tÃ©cnico en altas | Pendiente | Pendiente | pending | Pendiente | Pendiente | Visible en cÃ³digo; falta prueba. Requiere `m_ObjUsuarioConectado` inyectable. |
| Roles calculados (7 + `PermisoPruebas`) | Pendiente | Pendiente | pending | Pendiente | Pendiente | Visible en `UsuarioAplicacionPermisos.cls:15-21` y `Usuario.cls:36`; falta prueba. |
| Matriz completa de permisos | Pendiente | Pendiente | pending | Pendiente | Pendiente | Falta confirmaciÃ³n de producto. Cross-link: `cross-cutting-support` BR-XCUT-6. |

| SÃ­ntoma | Causa probable | ComprobaciÃ³n (Dysflow) | Ancla del documento |
|---|---|---|---|
| TÃ©cnico puede crear NC | RegresiÃ³n de guard de permisos | Crear prueba de bloqueo de alta con fixtures de `TbUsuariosAplicaciones` | BR-UPN-2 |
| MenÃº abre dominio incorrecto | Cableado de navegaciÃ³n roto | Crear prueba de navegaciÃ³n/costura (sin UI) | BR-UPN-1 |
| Usuario ve datos de otro responsable | Filtro inicial por rol roto | Crear prueba de precarga por rol con NC de fixture | BR-UPN-4..5 |
| Ribbon visible para no-admin | RegresiÃ³n de `PermisoPruebas` | Crear prueba de visibilidad con stub de usuario | BR-UPN-3 |
| Rol calculado devuelve valor incorrecto | RegresiÃ³n de `*Calculado` o `PermisoPruebas` | Crear prueba de `UsuarioAplicacionPermisos` con permisos forzados | BR-UPN-6 |

## Â§6 Notas de migraciÃ³n web

### Â§6.1 Conservar (comportamiento de negocio que debe sobrevivir)
- El menÃº principal enruta a Proyecto y AuditorÃ­as mediante formularios dedicados (BR-UPN-1): la web debe traducir cada opciÃ³n de menÃº a una ruta explÃ­cita (`/proyectos`, `/auditorias`), con guard de rol y guard de dominio.
- El usuario tÃ©cnico no puede ejecutar altas sensibles de Proyecto/AuditorÃ­a (BR-UPN-2): la API REST de alta debe devolver `403` con mensaje explÃ­cito ("No tiene autorizaciÃ³n para esa acciÃ³n") cuando el usuario tiene `EsTecnico = SÃ­`. El mensaje debe ser el mismo que ya muestra `Form_Form0BDOpcionesParteProyectos.cls:46,104,142`.
- Solo administrador ve Ribbon en modo pruebas; en uso normal se oculta (BR-UPN-3): la web debe mantener la regla "modo pruebas â‡’ admin visible" y "modo normal â‡’ oculto para no-admin", con `PermisoPruebas` como flag de autorizaciÃ³n.
- La gestiÃ³n de Proyecto precarga NC abiertas y filtra al tÃ©cnico por su nombre (BR-UPN-4): el endpoint de gestiÃ³n de Proyecto debe aplicar el filtro `ResponsableTelefonica = usuario.Nombre` por defecto, sin permitir que el tÃ©cnico vea NC de otro responsable.
- La gestiÃ³n de AuditorÃ­a precarga NC abiertas y filtra responsable de implantaciÃ³n para no tÃ©cnicos (BR-UPN-5): el endpoint de gestiÃ³n de AuditorÃ­a debe aplicar el filtro `RESPONSABLEIMPLANTACION = usuario.Nombre` por defecto para no tÃ©cnicos.
- Los 7 flags `EsUsuario*` (`Administrador`, `Calidad`, `Economia`, `Secretaria`, `Tecnico`, `SinAcceso`, `CalidadAvisos`) en `UsuarioAplicacionPermisos` + `PermisoPruebas` en `Usuario` (BR-UPN-6): la web debe seguir exponiendo los mismos 8 flags/permisos calculados, como atributos del claim/token del usuario.
- La matriz completa de permisos por acciÃ³n sensible (cerrar/eliminar/rehabilitar/documento/acciÃ³n/informe/configuraciÃ³n) estÃ¡ aprobada por producto (BR-UPN-7): la web debe poder consumir esa matriz desde un Ãºnico servicio de autorizaciÃ³n, no como checks dispersos.

### Â§6.2 Transformar (mecanismo legacy que se reformula)
- Sustituir `Form_Form0BDOpciones`, `Form_Form0BDOpcionesParteProyectos`, `Form_Form0BDOpcionesAuditorias`, `Form_Form0BDTecnicos` por un menÃº web declarativo con rutas, guards y permisos; no replicar la cinta (Ribbon) Access.
- Convertir `Usuario` y `UsuarioAplicacionPermisos` en un servicio de identidad + autorizaciÃ³n: el primero resuelve la identidad desde un token, el segundo aplica la matriz de permisos.
- Reemplazar el patrÃ³n de inyecciÃ³n de `m_ObjUsuarioConectado` por middleware de autenticaciÃ³n/autorizaciÃ³n en la capa de aplicaciÃ³n, no por una variable global mutada al inicio.
- Mover los 7 flags `EsUsuario*` y `PermisoPruebas` a claims del JWT/token del usuario, no como propiedades de un objeto VBA.
- Sustituir la cinta (Ribbon) como control de seguridad por un menÃº declarativo con guard de rol en el servidor; la cinta no debe decidir permisos, solo reflejar la decisiÃ³n del servidor.
- Reemplazar la convenciÃ³n de `Forms("FormNCProyectoGestion").ResponsableTelefonica = m_ObjUsuarioConectado.Nombre` por un parÃ¡metro de filtro explÃ­cito en la URL o body de request, con guard en el servidor.

### Â§6.3 NO copiar (deuda legacy de Access que no debe portarse)
- No portar `TempVars` ni globals como mecanismo de inyecciÃ³n de usuario: la web debe usar autenticaciÃ³n por token/sesiÃ³n, no variables globales.
- No usar la cinta (Ribbon) ni la visibilidad de menÃºs como control de seguridad real: la web debe aplicar permisos en el servidor y devolver `403` cuando corresponda.
- No duplicar la lÃ³gica de "quÃ© es un tÃ©cnico" en cada `.cls` de formulario: la web debe tener un Ãºnico servicio de autorizaciÃ³n.
- No migrar la combinaciÃ³n `EsTecnico` + `EsAdministrador` como dos checks booleanos independientes: la web debe tratarlos como roles dentro de una matriz declarativa.
- No portar la dependencia de `Forms(...)` como mecanismo de comunicaciÃ³n entre formularios: la API REST debe recibir parÃ¡metros explÃ­citos en la URL o body.

### Â§6.4 Preguntas abiertas al product owner
- Â¿La matriz de permisos (BR-UPN-7) es la misma para Proyecto y AuditorÃ­a o se diferencia por dominio? Confirmar alcance.
- Â¿Los 7 flags `EsUsuario*` se mantienen como estÃ¡n en la web o se renombran a roles mÃ¡s explÃ­citos? (BR-UPN-6) Confirmar convenciÃ³n.
- Â¿El flag `SinAcceso` bloquea toda la app o solo rutas sensibles? (BR-UPN-6) Hoy se infiere del nombre; la web debe tener un contrato explÃ­cito.
- Â¿La cinta (Ribbon) sobrevive a la migraciÃ³n como artefacto de UI o se elimina? (BR-UPN-3) Si sobrevive, Â¿quiÃ©n la diseÃ±a?
- Â¿Los filtros de precarga por `ResponsableTelefonica` y `RESPONSABLEIMPLANTACION` (BR-UPN-4, BR-UPN-5) son obligatorios o el usuario puede quitarlos? Â¿La respuesta del backend debe filtrar siempre por defecto?
- Â¿La auditorÃ­a de decisiones de autorizaciÃ³n (denegado/permitido + motivo) tiene un SLA de retenciÃ³n? Confirmar antes de definir el servicio.

## Â§7 Registro de confianza
| Hecho | Confianza | Evidencia | Fecha |
|---|---|---|---|
| BR-UPN-1 â€” El menÃº principal enruta a Proyecto y AuditorÃ­as mediante formularios dedicados. | Verified-static | `Form_Form0BDOpciones.cls:15` (`DoCmd.OpenForm "Form0BDOpcionesParteProyectos"`) y `Form_Form0BDOpciones.cls:71` (`DoCmd.OpenForm "Form0BDOpcionesAuditorias"`); FALTA â†’ crear mediante `access-vba-tdd` como contrato de navegaciÃ³n/cableado (sin UI) | 2026-06-15 |
| BR-UPN-2 â€” Usuario tÃ©cnico no puede ejecutar altas sensibles de Proyecto/AuditorÃ­a. | Verified-static | `EsTecnico = EnumSino.SÃ­` en `Form_Form0BDOpcionesParteProyectos.cls:46,104,142` y `Form_Form0BDOpcionesAuditorias.cls:50,168,265`; tambiÃ©n `Form_Form0BDTecnicos.cls:124`; FALTA â†’ crear mediante `access-vba-tdd` con fixtures de `TbUsuariosAplicaciones` + inyecciÃ³n de `m_ObjUsuarioConectado` y asserts sobre mensaje de autorizaciÃ³n | 2026-06-15 |
| BR-UPN-3 â€” Solo administrador ve Ribbon en modo pruebas; en uso normal se oculta. | Verified-static | `EsAdministrador = EnumSino.SÃ­` combinado con `PermisoPruebas` en `Form_Form0BDOpciones.cls:115,132`; `PermisoPruebas` declarado en `src/classes/Usuario.cls:36`; FALTA â†’ crear mediante `access-vba-tdd` con coste vÃ­a stub de `m_ObjUsuarioConectado` y asserts sobre visibilidad de Ribbon | 2026-06-15 |
| BR-UPN-4 â€” La gestiÃ³n de Proyecto precarga NC abiertas y filtra al tÃ©cnico por su nombre. | Verified-static | `Form_Form0BDOpcionesParteProyectos.cls:142-143` filtra `Forms("FormNCProyectoGestion").ResponsableTelefonica = m_ObjUsuarioConectado.Nombre`; FALTA â†’ crear mediante `access-vba-tdd` con fixtures de NC y asserts sobre `ResponsableTelefonica` precargado | 2026-06-15 |
| BR-UPN-5 â€” La gestiÃ³n de AuditorÃ­a precarga NC abiertas y filtra responsable de implantaciÃ³n para no tÃ©cnicos. | Verified-static | `Form_Form0BDOpcionesAuditorias.cls:140-141` filtra `Forms("FormNCAuditoriaGestion").RESPONSABLEIMPLANTACION = m_ObjUsuarioConectado.Nombre`; FALTA â†’ crear mediante `access-vba-tdd` con fixtures de NC de auditorÃ­a y asserts sobre el filtro | 2026-06-15 |
| BR-UPN-6 â€” Roles calculados de usuario: 7 flags `EsUsuario*` (`Administrador`, `Calidad`, `Economia`, `Secretaria`, `Tecnico`, `SinAcceso`, `CalidadAvisos`) en `UsuarioAplicacionPermisos` + `PermisoPruebas` en `Usuario` (8 flags/permisos calculados totales). | Verified-static | `src/classes/UsuarioAplicacionPermisos.cls:15-21` (7 flags `EsUsuario*`) y `src/classes/Usuario.cls:36` (`PermisoPruebas`); FALTA â†’ crear mediante `access-vba-tdd` con fixtures de permisos por rol, asserts sobre cada `*Calculado` y `PermisoPruebas` | 2026-06-15 |
| BR-UPN-7 â€” La matriz completa de permisos por acciÃ³n sensible (cerrar/eliminar/rehabilitar/documento/acciÃ³n/informe/configuraciÃ³n) estÃ¡ aprobada por producto. | Intended | FALTA â†’ crear mediante `access-vba-tdd` tras confirmar matriz; misma matriz referenciada por `cross-cutting-support` BR-XCUT-6 | 2026-06-15 |
| Los menÃºs de Proyecto y AuditorÃ­as existen y enrutan formularios de dominio. | Verified-static | `src/forms/Form_Form0BDOpciones.cls:15,71` (`DoCmd.OpenForm "Form0BDOpcionesParteProyectos"` / `DoCmd.OpenForm "Form0BDOpcionesAuditorias"`) | 2026-06-15 |
| Los tÃ©cnicos estÃ¡n bloqueados en varias altas sensibles. | Verified-static | `EsTecnico = EnumSino.SÃ­` en `Form_Form0BDOpcionesParteProyectos.cls:46,104,142`, `Form_Form0BDOpcionesAuditorias.cls:50,168,265` y `Form_Form0BDTecnicos.cls:124` | 2026-06-15 |
| El Ribbon en modo pruebas se reserva al administrador. | Verified-static | `EsAdministrador = EnumSino.SÃ­` + `PermisoPruebas` en `Form_Form0BDOpciones.cls:115,132`; `PermisoPruebas` declarado en `src/classes/Usuario.cls:36` | 2026-06-15 |
| La gestiÃ³n de Proyecto filtra al tÃ©cnico por su nombre. | Verified-static | `Form_Form0BDOpcionesParteProyectos.cls:143` (`Forms("FormNCProyectoGestion").ResponsableTelefonica = m_ObjUsuarioConectado.Nombre`) | 2026-06-15 |
| La gestiÃ³n de AuditorÃ­a filtra al responsable de implantaciÃ³n. | Verified-static | `Form_Form0BDOpcionesAuditorias.cls:141` (`Forms("FormNCAuditoriaGestion").RESPONSABLEIMPLANTACION = m_ObjUsuarioConectado.Nombre`) | 2026-06-15 |
| Existen 7 flags `EsUsuario*` en `UsuarioAplicacionPermisos` + `PermisoPruebas` en `Usuario` (8 permisos/flags calculados totales). | Verified-static | `src/classes/UsuarioAplicacionPermisos.cls:15-21` y `src/classes/Usuario.cls:36` | 2026-06-15 |
| La matriz completa de permisos estÃ¡ aprobada y probada. | Intended | No hay manifest dedicado; ningÃºn test cubre permisos/roles/navegaciÃ³n. Cross-link: `cross-cutting-support` BR-XCUT-6 | 2026-06-15 |
| Existe cobertura runtime de navegaciÃ³n/permisos/roles. | Intended | No existe manifest; la promociÃ³n a `Verified-runtime` para BR-UPN-1..6 estÃ¡ bloqueada hasta que se creen pruebas con `access-vba-tdd` | 2026-06-15 |

**âš ï¸ Divergencias (intenciÃ³n SDD â‰  realidad del cÃ³digo)**
- Sin divergencia confirmada. Hueco confirmado: hay reglas de autorizaciÃ³n embebidas en formularios, pero no existe una matriz de producto trazada (BR-UPN-7) ni manifest de pruebas que cubra BR-UPN-1..6. La misma matriz es la que `cross-cutting-support` declara como intenciÃ³n en BR-XCUT-6.
