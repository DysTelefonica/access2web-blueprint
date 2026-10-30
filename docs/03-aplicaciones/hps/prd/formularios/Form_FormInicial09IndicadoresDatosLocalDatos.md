# Form_FormInicial09IndicadoresDatosLocalDatos

> **Commit fijado:** `5e64cf9ae1fcb5b01d261eee4f732b5cf611d8ae` · **Fuente:** `src/forms/Form_FormInicial09IndicadoresDatosLocalDatos.cls` (42 líneas) y `.form.txt` (4067 líneas) · **PRD:** [README](../README.md)

Subformulario hoja del panel de indicadores: la vista de **datos locales** (`TbDatosLocalParaIndicadores`). Se muestra como datasheet (`DefaultView =2`, `.form.txt:14`; `AllowFormView =0`, `.form.txt:330`) y su diseño arranca deliberadamente vacío, con un `RecordSource` en el `.form.txt` que fija `WHERE ID=-1` (`.form.txt:284-285`), hasta que el padre le asigna el `SELECT` real. Es el `SourceObject` por defecto de `SubFormIndicadores` (`Form_FormInicial09Indicadores.form.txt:794`) y el destino de las siete ramas de `EstablecerFuenteDatosIndicadores`, que terminan todas en `frm.RecordSource = m_SQL` más `Requery` (`FG.bas:3086-3134, 3136-3139`). Sus dos procedimientos son `Form_Current` y `Form_Load`; declara `Public Event UsuarioSeleccionado()` (`:5`) sin ningún `Raise` ni `WithEvents` en `src`. Las citas son `Form_FormInicial09IndicadoresDatosLocalDatos.cls:línea` salvo indicación; `FG.bas` abrevía `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormInicial09Indicadores` | `SubFormIndicadores.SourceObject` apunta a este formulario por defecto, y las siete ramas de `EstablecerFuenteDatosIndicadores` lo reasignan al cambiar de indicador | `Form_FormInicial09Indicadores.form.txt:794`, `FG.bas:3086-3134` |
| `Form_FormInicial00Principal` | Contenedor de `SubFormCentral`; `EstablecerFuenteDatosIndicadores` sale sin tocar nada si `SubFormCentral.SourceObject <> "FormInicial09Indicadores"` | `FG.bas:3067-3075` |
| `Form_FormInicial01Lateral` | Destino del semáforo: `EstablecerSemaforo` escribe en el `ImagenSemaforo` de ese subformulario y sale si no es el que está montado | `FG.bas:2874-2880, 2896-2897` |

No abre ningún formulario. El botón «editar» abre `FormUsuario00Principal` o `FormUsuario07SICAEdicion` desde el padre, no desde aquí (`Form_FormInicial09Indicadores.cls:151-162`).

## Controles y eventos

`RecordSource` en diseño: `SELECT TbDatosLocalParaIndicadores.* FROM TbDatosLocalParaIndicadores WHERE (((TbDatosLocalParaIndicadores.ID)=-1));` (`.form.txt:284-285`). `OnCurrent ="[Event Procedure]"` (`.form.txt:286`) y `OnLoad ="[Event Procedure]"` (`.form.txt:329`).

| Grupo | Controles | Inicio en `.form.txt` |
|---|---|---|
| Identidad y expediente | `ID`, `DNI`, `Nombre`, `Apellido_1`, `Apellido_2`, `Telefono`, `EmpresaUsuario`, `EmpresaTramitadora`, `Correo_e`, `F_Nacimiento`, `Motivo_HPS`, `IDExpediente`, `CodExp` | `.form.txt:538` … `.form.txt:941` |
| Curso, HPS y observaciones | `F_Curso`, `CursoEnVigor`, `F_Baja`, `FAvisoConcesion`, `RequiereComunicacionConcesion`, `FechaHPSConcesionMinima`, `Curso_Realizado`, `Requiere_Curso`, `FechaPrimeraConvocatoria`, `FechaSegundaConvocatoria`, `FechaCorreoNoCurso`, `Requiere_PrimeraConvocatoriaCurso`, `Requiere_SegundaConvocatoriaCurso`, `Requiere_CorreoJefeSeguridadCurso`, `Observaciones` | `.form.txt:975` … `.form.txt:1466` |
| Bloque NAC | 18 cajas `HPS_NAC_*`: `SIN_DATOS`, `F_Concesion`, `F_Caducidad`, `Grado`, `F_Solicitud`, `Especialidad`, `Renovacion`, `F_Baja`, `Activo`, `ApuntoDeCaducar`, `Baja`, `Caducado`, `PendienteRenovacion`, `Solicitado`, `ESTADO`, `Irregular`, `MOTIVO_IRREGULAR`, `MESES_PARA_RENOVAR` | `.form.txt:1503` |
| Bloque OTAN | las mismas 18 con prefijo `HPS_OTAN_` | `.form.txt:2126` |
| Bloque ESA | las mismas 18 con prefijo `HPS_ESA_` | `.form.txt:2750` |
| Bloque UE | las mismas 18 con prefijo `HPS_UE_` | `.form.txt:3371` |
| Jurídica de contratación | `JuridicaContrato`, única caja sin etiqueta emparejada | `.form.txt:3999` |
| Etiquetas de columna | `Etiqueta809` … `Etiqueta909`; la numeración salta de `Etiqueta816` a `Etiqueta818` (`.form.txt:789, 828`), de ahí que haya 100 etiquetas para 101 cajas | `.form.txt:553` … `.form.txt:3987` |
| Sección | `Detalle`, con 201 controles con nombre: 101 cajas y 100 etiquetas | `.form.txt:527` |

Cada caja lleva `ControlSource` con el nombre homónimo (`ID`, `.form.txt:544`; `EmpresaTramitadora`, `.form.txt:779`; `JuridicaContrato`, `.form.txt:4007`). Ninguna declara `Locked` ni `Enabled` propios, así que el modo de edición depende de los valores heredados del datasheet.

## Flujos

### Apertura (`Form_Load`, `:21-42`)

1. `CopiarDatosAIndicadores m_Error` (`:27`): sincroniza las tablas físicas `TbDatosLocalParaIndicadores` y `TbUsuariosSICALocalParaIndicadores` desde `TbDatosLocal` y `TbUsuariosSICALocal` (`FG.bas:3032-3046`). Un `Static s_YaCopiadoEnEstaCarga` evita repetir el `DELETE + INSERT` durante la cascada de carga (`FG.bas:3011, 3027-3030`).
2. Si `m_Error` viene informado, relanza el error interno 1000 (`:28-30`) y cae al manejador.
3. `Me.Parent.ComandoEditarUsuario.Enabled = False` (`:32`).
4. **No asigna `RecordSource` ni ejecuta `Requery`**: la hoja conserva el `WHERE ID=-1` del diseño hasta que el padre intervene.

`errores:` (`:35-42`) envía los errores distintos de 1000 al administrador con `CorreoAlAdministrador m_Error` (`FG.bas:67`), baja el cursor de espera y muestra un `MsgBox` crítico. El valor de retorno se guarda en el global `pregunta` (`VariablesEntorno.bas:220`).

### Cambio de registro (`Form_Current`, `:7-19`)

1. `On Error Resume Next`; `m_ID = Nz(Me.ID, "")` (`:9-10`).
2. Si `Err.Number <> 0`, desactiva `Me.Parent.ComandoEditarUsuario` (`:11-12`); si no, lo habilita (`:13-15`).
3. `Err.Clear` (`:16`).
4. `EstablecerSemaforo m_ID` (`:17`): pide la ficha al constructor y pinta el semáforo del lateral; con `m_ID` vacío se limita a ocultarlo (`FG.bas:2882-2886`).

### Selección del indicador (fuera de este formulario)

1. `EstablecerFuenteDatosIndicadores` compone un `SELECT * FROM TbDatosLocalParaIndicadores` (`FG.bas:3076-3077`) más un `WHERE` literal distinto por indicador: curso pendiente (`FG.bas:3085`), primera (`FG.bas:3091`), segunda (`FG.bas:3097`), correo al jefe de seguridad (`FG.bas:3103`), a punto de caducar (`FG.bas:3110-3113`), caducadas (`FG.bas:3119-3122`) y solicitándose (`FG.bas:3128-3131`).
2. En las siete ramas fija `SubFormIndicadores.SourceObject` a este formulario si no lo era (`FG.bas:3086-3134`).
3. Asigna `frm.RecordSource = m_SQL` y ejecuta `frm.Requery` sobre el `.Form` del subformulario (`FG.bas:3136-3139`).

## Reglas, validaciones y cálculos

- **Arranque vacío garantizado por diseño.** El `WHERE ID=-1` está escrito en el `.form.txt` (`.form.txt:284-285`), no en código: al no haber ningún registro con ese `ID`, la hoja no muestra filas hasta que el padre la recarga.
- **Sincronización única por cascada.** `CopiarDatosAIndicadores` borra e inserta las dos tablas una sola vez mientras el `Static` siga activo (`FG.bas:3032-3046, 3048`); el argumento `p_CacheAcabDeRegenerar` reinicia ese `Static` y devuelve sin copiar (`FG.bas:3022-3025`).
- **Botón «editar» gobernado por el error de `Me.ID`.** La decisión depende únicamente de `Err.Number` tras `Nz(Me.ID, "")` (`:9-15`), no del valor: con `RecordSource` que aún no expone `ID`, el botón se desactiva.
- **Semárafo sujeto a dos guardas.** `EstablecerSemaforo` sale sin hacer nada si `FormInicial00Principal` no está abierto o si `SubFormLateral.SourceObject <> "FormInicial01Lateral"` (`FG.bas:2874-2880`); con identificador vacío oculta la imagen sin mensaje (`FG.bas:2882-2886`).
- **El padre fija el botón en su propia carga.** `ComandoEditarUsuario` se deshabilita y se aplica estilo en el `Form_Load` del padre (`Form_FormInicial09Indicadores.cls:581-582`), de modo que ambos formularios compiten por el mismo control.
- **Sin concatenación de entrada del usuario.** Los siete `WHERE` son literales cerrados (`FG.bas:3085-3131`); ningún valor de la hoja entra en el SQL.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Sincronización de las tablas-fuente | `TbDatosLocal`, `TbDatosLocalParaIndicadores`, `TbUsuariosSICALocal`, `TbUsuariosSICALocalParaIndicadores` | `DELETE *` + `INSERT INTO … SELECT` (DAO `Execute`) | `FG.bas:3032-3046` |
| Vista en diseño | `TbDatosLocalParaIndicadores` | `SELECT … WHERE (((TbDatosLocalParaIndicadores.ID)=-1));` | `.form.txt:284-285` |
| Vista real | `TbDatosLocalParaIndicadores` | `RecordSource` asignado por el padre | `FG.bas:3076-3077, 3138` |
| Semáforo de la ficha | `TbDatosLocal` (`getDatosLocal` → `URLImagenSemaforo`) | `SELECT` por `ID` | `FG.bas:2887-2897` |

Sin `INSERT`/`UPDATE`/`DELETE` propios en este `.cls` (`:1-42`).

## Coexistencia

Pendiente de la coexistencia (documento aún no redactado). El patrón de replicación es el de la familia: tablas físicas regeneradas con `DELETE *` + `INSERT INTO … SELECT` (`FG.bas:3032-3046`), siete `WHERE` literales distintos (`FG.bas:3085-3131`), y un semáforo que viaja por el formulario lateral en vez de por la hoja. El lado favorable: ninguno de esos `WHERE` interpola valores, a diferencia de la concatenación que hace el filtro de consultas.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- La rama `HPSASolicitandose` repite `HPS_UE_F_Solicitud` en las dos últimas alternativas y no menciona `HPS_NAC_F_Solicitud` (`FG.bas:3128-3131`). Con el predicado tal como está, un usuario con solicitud solo en NAC no aparece en ese indicador.
- Esa misma rama escribe `NOt` con mayúsculas y mezcla dos convenciones de columna: `Not … Is Null` sobre fechas frente a `='Sí'` sobre las columnas de las seis ramas anteriores (`FG.bas:3085-3122, 3128-3131`).
- `Form_Current` habilita `ComandoEditarUsuario` solo porque `Err.Number = 0` (`:13-15`), sin comprobar que `Me.ID` tenga valor: con el `RecordSource` de diseño y `Nz` resuelto sin error, el botón queda activo sobre una hoja sin ficha seleccionada, y la acción del padre abre `FormUsuario00Principal` sin usuario (`Form_FormInicial09Indicadores.cls:147-155`).
- `On Error Resume Next` (`:9`) silencia el fallo de `Nz(Me.ID, "")`; si el error salta, la rama de deshabilitado actúa sin diagnóstico y el semáforo se oculta.
- Hay una ventana entre `Form_Load` (`:32`) y el `frm.RecordSource` del padre (`FG.bas:3138`) en la que un `Form_Current` puede dejar el botón activo sobre la hoja vacía.
- `Me.Parent` no se comprueba en `:12`, `:14` ni `:32`: abierto como formulario suelto, el error cae al manejador y solo produce un `MsgBox` (`:40-41`).
- `Public Event UsuarioSeleccionado()` (`:5`) no se lanza ni se escucha en `src`: contrato declarado y muerto.
- `CopiarDatosAIndicadores` puede borrar filas recién insertadas por otra cascada si alguien llama con `p_CacheAcabDeRegenerar = True`: el reinicio del `Static` está en `FG.bas:3022-3025` y esa rama no se usa desde este formulario (`:27`).
- La numeración de etiquetas no es contigua: falta `Etiqueta817` entre `Etiqueta816` y `Etiqueta818` (`.form.txt:789, 828`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: derivar la vista de indicadores de una consulta parametrizada en lugar de una tabla física replicada; un catálogo de indicadores con predicado declarado en datos en lugar de siete literales; y una regla explícita para el botón «editar» (deshabilitado hasta que exista ficha cargada) en lugar del detector de error.

## Criterios de aceptación

Propuestos, no aprobados:

1. Al abrirse, la hoja muestra cero filas: el `RecordSource` de diseño con `ID=-1` no aporta ninguna.
2. Al seleccionar un indicador, el padre monta este formulario como `SourceObject`, le asigna el `SELECT` correspondiente y ejecuta `Requery` (`FG.bas:3136-3139`).
3. `CopiarDatosAIndicadores` se ejecuta una sola vez por cascada; las siguientes llamadas no repiten el `DELETE + INSERT` (`FG.bas:3011, 3027-3030`).
4. `ComandoEditarUsuario` del padre queda deshabilitado al cargar y se habilita solo cuando `Me.ID` resuelve sin error.
5. `EstablecerSemaforo` oculta `ImagenSemaforo` cuando el identificador recibido está vacío (`FG.bas:2882-2886`).
6. Si el principal no está abierto o el lateral no es `FormInicial01Lateral`, el semáforo no se toca (`FG.bas:2874-2880`).
7. Los errores distintos de 1000 se envían por correo al administrador y se muestran en un `MsgBox` crítico (`:36-41`).

## Pendientes

- Si algún punto del flujo fuera de `EstablecerFuenteDatosIndicadores` asigna `SourceObject` a este formulario: en `src/` solo constan las siete ramas, y el padre vacía el control en su propia carga (`Form_FormInicial09Indicadores.form.txt:1932`).
- Si `Etiqueta817` existió y se eliminó o si nunca se creó (`.form.txt:789, 828`).
- Qué valor toma `s_YaCopiadoEnEstaCarga` entre recargas de la aplicación: es `Static`, no hay forma de inspeccionarlo desde el formulario.