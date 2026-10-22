# Form_FormInicial03ConsultasDatos

> **Commit fijado:** `5e64cf9ae1fcb5b01d261eee4f732b5cf611d8ae` · **Fuente:** `src/forms/Form_FormInicial03ConsultasDatos.cls` (34 líneas) y `.form.txt` (3 989 líneas) · **PRD:** [README](../README.md)

Hospedador de datos de la consulta del bloque 03. Es un **subformulario hoja** montado por `Form_FormInicial02ConsultasPrincipal` vía `Me.SubFormResultados.SourceObject = "FormInicial03ConsultasDatos"` (`Form_FormInicial02ConsultasPrincipal.cls:252`, `.form.txt:1245`); el padre se conecta con `WithEvents m_FormInicial03ConsultasDatos As Form_FormInicial03ConsultasDatos` (`Form_FormInicial02ConsultasPrincipal.cls:12`, `.form.txt:1004`). Su `.form.txt` es un datasheet (`DefaultView =2`, `.form.txt:15`; `AllowFormView =0`, `.form.txt:335`) con 103 cajas de texto enlazadas y sus etiquetas — un registro de `TbDatosLocal` por fila. Las citas son `Form_FormInicial03ConsultasDatos.cls:línea` salvo indicación; `FG.bas` abrevía `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormInicial02ConsultasPrincipal` | `Me.SubFormResultados.SourceObject = "FormInicial03ConsultasDatos"` y `Set m_FormInicial03ConsultasDatos = Me.SubFormResultados.Form` | `Form_FormInicial02ConsultasPrincipal.cls:252`, `.form.txt:1245-1246` |

Abre nada por sí mismo. Hereda `RecordSource` del `.form.txt` hasta que `Form_Load` lo sustituye (`:30-33`).

## Controles y eventos

`RecordSource` en diseño: `SELECT TbDatosLocal.* FROM TbDatosLocal;` (`.form.txt:289`). `OnCurrent` y `OnLoad` enganchados a procedimientos del evento (`.form.txt:290, 334`).

| Grupo | Controles | Inicio en `.form.txt` |
|---|---|---|
| Cajas de texto de identidad personal | `ID`, `DNI`, `Nombre`, `Apellido_1`, `Apellido_2`, `Telefono`, `EmpresaUsuario`, `EmpresaTramitadora`, `Correo_e`, `F_Nacimiento`, `Motivo_HPS`, `IDExpediente`, `CodExp` | `.form.txt:549, 581, 613, 645, 677, 709, 741, 773, 805, 837, 869, 901, 933` |
| Cajas de texto de fechas, cursos y avisos | `F_Curso`, `CursoEnVigor`, `F_Baja`, `FAvisoConcesion`, `RequiereComunicacionConcesion`, `FechaHPSConcesionMinima` | `.form.txt:965, 997, 1030, 1063, 1096, 1129` |
| Cajas de texto HPS ESA | `HPS_ESA_SIN_DATOS`, `HPS_ESA_F_Concesion`, `HPS_ESA_F_Caducidad`, `HPS_ESA_Grado`, `HPS_ESA_F_Solicitud`, `HPS_ESA_Especialidad`, `HPS_ESA_Renovacion`, `HPS_ESA_F_Baja`, `HPS_ESA_Activo`, `HPS_ESA_ApuntoDeCaducar`, `HPS_ESA_Baja`, `HPS_ESA_Caducado`, `HPS_ESA_PendienteRenovacion`, `HPS_ESA_Solicitado`, `HPS_ESA_ESTADO`, `HPS_ESA_Irregular`, `HPS_ESA_MOTIVO_IRREGULAR`, `HPS_ESA_MESES_PARA_RENOVAR` | `.form.txt:2641, 2674, 2706, 2738, 2770, 2802, 2834, 2867, 2899, 2932, 2965, 2998, 3032, 3065, 3102, 3136, 3171, 3205` |
| Cajas de texto HPS UE | `HPS_UE_SIN_DATOS`, `HPS_UE_F_Concesion`, `HPS_UE_F_Caducidad`, `HPS_UE_Grado`, `HPS_UE_F_Solicitud`, `HPS_UE_Especialidad`, `HPS_UE_Renovacion`, `HPS_UE_F_Baja`, `HPS_UE_Activo`, `HPS_UE_ApuntoDeCaducar`, `HPS_UE_Baja`, `HPS_UE_Caducado`, `HPS_UE_PendienteRenovacion`, `HPS_UE_Solicitado`, `HPS_UE_ESTADO`, `HPS_UE_Irregular`, `HPS_UE_MOTIVO_IRREGULAR`, `HPS_UE_MESES_PARA_RENOVAR` | `.form.txt:3239, 3274, 3308, 3343, 3377, 3411, 3445, 3480, 3514, 3549, 3584, 3619, 3654, 3689, 3724, 3758, 3795, 3831` |
| Cajas finales | `JuridicaContrato`, `LugarNacimiento`, `CadenaContratistas` | `.form.txt:3867, 3903, 3936` |
| Etiquetas | una por cada caja de texto anterior, sin handler propio | emparejadas a partir de `.form.txt:564` |

Cada caja tiene `ControlSource` con el nombre del campo homónimo (p. ej. `DNI`, `.form.txt:582`). El `.cls` no expone `Sub` alguno para interactuar con estas cajas: depende del padre para `Me.Requery` (`Form_FormInicial02ConsultasPrincipal.cls:315`, `.form.txt:1308`).

## Flujos

### Apertura (`Form_Load`, `:29-33`)

1. Asigna `Me.RecordSource = "SELECT TbDatosLocal.* FROM TbDatosLocal;"` (`:31`) — coincide con el del `.form.txt` y no añade `WHERE`.
2. `Me.Requery` (`:32`).

### Cambio de registro (`Form_Current`, `:7-27`)

1. `m_ID = Nz(Me.ID, "")` (`:11`).
2. `EstablecerSemaforo m_ID` (`:13`): actualiza el semáforo del padre `FormInicial00Principal`>`FormInicial01Lateral` en `ImagenSemaforo` (`FG.bas:2865-2904`; control en `Form_FormInicial01Lateral.form.txt:1805`).
3. `Set m_ObjUsuarioActivo = Constructor.getUsuarioHPS(m_ID, , m_Error)` (`:14`): carga la entidad activa de `TbUsuarios` por `ID` con SQL concatenado `WHERE ID=<ID>` (`Constructor.bas:760-762, 781`); el resultado se asigna a la global `m_ObjUsuarioActivo` (`VariablesEntorno.bas:225`).
4. Si `m_Error <> ""`, eleva 1000 (`:15-17`) para que `errores:` envíe `CorreoAlAdministrador m_Error` y muestre `MsgBox m_Error, vbCritical` (`:21-25`). El evento `UsuarioSeleccionado` queda declarado (`:5`) pero no se lanza en este formulario: lo esperaría un consumidor de `WithEvents` y no se observa consumo en `src/`.

### Errores

La rutina `errores:` (`:19-25`) reenvía cualquier error no 1000 al administrador con `CorreoAlAdministrador m_Error` (`FG.bas:67`); los 1000 se ignoran en el `.cls` y se propagan al `MsgBox` final.

## Reglas, validaciones y cálculos

- La columna `ID` identifica la fila (`Me.ID`, `:11`) y se propaga al semáforo y a `getUsuarioHPS` (`Constructor.bas:760-784`). Sin `WHERE` en `RecordSource`, el subformulario lista toda `TbDatosLocal` (`:31`).
- `EstablecerSemaforo` solo escribe si está abierto `FormInicial00Principal` y si su subformulario lateral es `FormInicial01Lateral` (`FG.bas:2875-2879`); con `p_ID = ""` oculta `ImagenSemaforo` (`FG.bas:2882-2888`).

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Origen del subformulario | `TbDatosLocal` | `SELECT` sin `WHERE` (en `.form.txt` y `Form_Load`) | `.form.txt:289`, `:31` |
| Semáforo del usuario | `TbDatosLocal` (consulta de `URLImagenSemaforo`) | `SELECT` por `ID` | `FG.bas:2891-2897` |
| Entidad activa | `TbUsuarios` (`getUsuarioHPS`) | `SELECT * FROM TbUsuarios WHERE ID=<ID>;` | `Constructor.bas:760-784` |
| `Requery` invocado por el padre tras acciones | `TbDatosLocal` | `Me.Requery` desde el padre | `Form_FormInicial02ConsultasPrincipal.cls:315` |

No hay `INSERT`/`UPDATE`/`DELETE` en este `.cls` (`:1-34`).

## Coexistencia

Pendiente de `../coexistencia.md`. Ya se ve: SQL concatenado en `getUsuarioHPS` (`Constructor.bas:760-762`); literales de Access `acForm`/`acSaveNo` en el padre (`Form_FormInicial02ConsultasPrincipal.cls:27, 55, 147, 220`); comparaciones contra globales de Access (`Application.TempVars`, `Form_FormInicial02ConsultasPrincipal.cls:180`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- El evento `UsuarioSeleccionado` (`:5`) no se lanza en este `.cls`; los consumidores de `WithEvents` declarados en el padre (`Form_FormInicial02ConsultasPrincipal.cls:12`, `.form.txt:1004`) no tienen handler para él en `src/`.
- `Form_Load` reasigna `RecordSource` con un `SELECT` idéntico al del `.form.txt` (`:31`); la doble asignación no aporta valor y se omite si el padre ya lo fija.
- La hoja carga toda `TbDatosLocal` sin `WHERE` (`.form.txt:289`, `:31`); en producción se filtra por los parámetros del padre.
- `Form_Current` depende de que `FormInicial00Principal` esté abierto para fijar el semáforo (`FG.bas:2875`); fuera de ese contexto, `EstablecerSemaforo` es no-op silencioso (`:13`).
- `m_ObjUsuarioActivo` queda en `Nothing` si `Me.ID` es vacío (`:11`) — el padre desactiva entonces `ComandoEditarUsuario` y `ComandoAnexos` (`Form_FormInicial02ConsultasPrincipal.cls:238-239`) sin avisar.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: el evento `UsuarioSeleccionado` se reemplaza por paso de `ID` por argumento al navegar entre registros; el datasheet se sustituye por tabla paginada con `WHERE` reactivo a los parámetros del padre.

## Criterios de aceptación

Propuestos, no aprobados:

1. El subformulario muestra una fila de `TbDatosLocal` por cada `ID` resuelto por la consulta del padre.
2. Al cambiar el registro actual, el semáforo del panel lateral refleja `URLImagenSemaforo` del usuario seleccionado.
3. El objeto global `m_ObjUsuarioActivo` queda cargado con un `UsuarioHPS` cuyo `ID` coincide con `Me.ID`.
4. Los errores que no sean 1000 se notifican al administrador y aparecen en `MsgBox` crítico.
5. La lista de campos del subformulario coincide con las columnas visibles y cada `ControlSource` apunta al campo homónimo.

## Pendientes

- Confirmar si en el binario existe un `ControlSource` añadido por encima del catálogo aquí listado.
- Decidir si el evento `UsuarioSeleccionado` se mantiene en la migración o se sustituye por el patrón `ID` por argumento del padre.
