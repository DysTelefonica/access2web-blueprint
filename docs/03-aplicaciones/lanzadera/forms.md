# Lanzadera · Formularios, controles y navegación

La navegación es una composición de formularios contenedores y subformularios mediante `SourceObject`; la lógica de negocio está fuertemente acoplada a eventos `Click`, `Load`, `Open` y `AfterUpdate`.

| Formulario | Capacidad / controles y eventos relevantes | Navegación |
|---|---|---|
| `FormLogin` | `ComandoIniciarSesion_Click` obtiene `Usuario`, delega en `Login` (SHA-256, bloqueo, caducidad y contraseña de un solo uso) y guarda `m_ObjUsuarioConectadoLogin` | Abre `FormUsuarioCambioPass` o `LoginCorrecto`; el menú posterior depende de `m_ObjUsuarioConectadoLogin` |
| `FormMenuPrincipalAdmin` | Menú dinámico, usuario, ubicación, versión; `MENU1_Click`, `SUBMENU1/2_Click`, logout | Aplicaciones, usuarios y configuración de aplicación |
| `FormMenuPrincipalCalidad` | Menú y acceso al contenido formativo según perfil | Aplicaciones/vídeos |
| `FormMenuPrincipalUsuario` | Menú reducido; `MENU1_Click` carga aplicaciones; `MENU3_Click` carga perfil | `FormAplicaciones`, `FormPerfil` |
| `FormAplicaciones` | Tres pestañas `MENU1/2/3`; `FormDetalle` | Oficina, fuera de oficina y pruebas |
| `FormAplicacionesOficina` | Botones dinámicos por `IDAplicacion`; `ComandoActualizar`, lanzamiento de aplicaciones y SICA | `Lanzar` con catálogo/permisos |
| `FormAplicacionesFueraOficina` | Visibilidad/habilitación por permisos y `EjecucionEnOficina`; lanzamiento de Expedientes, Condor, HPS, NC, Brass y otras | `Lanzar` y actualización de versiones |
| `FormAplicacionesPrueba` | Aplicaciones en pruebas; actualización y botones de Condor/NC de prueba | Flujos de prueba específicos |
| `FormAplicacionGeneral` | Pestañas de selección/edición; evento `WithEvents` hacia gestión | `FormAplicacionesGestion` / `FormAplicacionDatosGenerales` |
| `FormAplicacionesGestion` | Filtro, selección y quitar aplicación activa | Emite `AplicacionSeleccionada` |
| `FormAplicacionDatosGenerales` | Alta/edición de nombre, ejecutable, backend, carpeta, perfiles y atributos; `ComandoRegistrar` | Persiste mediante clase `Aplicacion` |
| `FormUsuariosGestion` | Filtro por nombre, lista `ID;Correo;Nombre`, selección y reset de contraseña | Emite `UsuarioSeleccionado`; abre detalle |
| `FormUsuarioGeneral` | Contenedor de datos generales, perfil y aplicaciones del usuario | `FormUsuarioDatosGenerales`, `FormUsuarioPerfilAplicaciones` |
| `FormUsuarioPerfilAplicaciones` | Combo de aplicaciones, seis checks de rol y alta/eliminación de permisos | Persiste `TbUsuariosAplicacionesPermisos` |
| `FormUsuarioCambioPass` / `FormObtenerContraseña` | Cambio obligatorio, contraseña de un solo uso y recuperación por correo | Retorno al login |
| `FormVideosGestion` y `FormVideos*` | Árbol por aplicación, detalle, categorías y visionados; controles de vídeo ActiveX | Contenido formativo; no se probó ejecución |

## Contratos de eventos y navegación

- `Form_FormLogin` y los tres menús dependen de variables globales (`m_ObjUsuarioConectadoLogin`, `m_ObjEntorno`, `m_EnOficina`).
- `Form_FormAplicacionesOficina/FueraOficina` usan `getBoton(Me, ID)` para enlazar controles generados al catálogo; si no existe permiso, cambian `Enabled` e imagen.
- `Form_FormUsuarioPerfilAplicaciones` permite editar solo si `EsAdministradorCalculado = Sí`.
- Los errores normalmente se traducen a `MsgBox` y `CorreoAlAdministrador`; se conserva el patrón `Optional ByRef p_Error`.
- El lanzamiento estático queda trazado como `getBoton` → `Lanzar` → `Constructor.getAplicacion` → `Aplicacion.Lanzar`; este último copia recursos local/remoto y construye `/cmd <correo>` antes de `EjecutarShelllanzar`.
- `UsuarioAplicacionPermisos` tiene callers en `Usuario` y en los formularios de aplicaciones; `getdb` aparece con 76 callers internos. No se infieren consumidores fuera de Lanzadera desde este índice.

El grafo CodeGraph-VBA resolvió eventos, `OpenForm`, `SourceObject` y llamadas estáticas; deja dinámicos DAO, `fso`, `Shell`, `TempVars` y nombres construidos en runtime. Por ello el mapa es `Verified-static`, no un grafo runtime completo. La evidencia Dysflow de comportamiento sigue limitada por el `accessPath` offline.
