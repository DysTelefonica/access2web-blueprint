# Form_FormGestionRiesgosAutorizados

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Fuente:** `src/forms/Form_FormGestionRiesgosAutorizados.cls` (92 líneas) y `src/forms/FormGestionRiesgosAutorizados.form.txt` · **PRD:** [README](../README.md)

Subformulario de **solo lectura** del contenedor `FormGestionRiesgos`: muestra quién está autorizado a intervenir en la gestión de riesgos. No escribe nada, no tiene botones y no tiene eventos propios más allá de la carga. Es el más pequeño de los cuatro subformularios. Las citas son `Form_FormGestionRiesgosAutorizados.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `FormGestionRiesgos` | Pestaña «Autorizados» del control de navegación; el destino se fija en diseño | `FormGestionRiesgos.form.txt:503` |
| Nadie más | No hay ningún `DoCmd.OpenForm` en `src` que lo abra de forma independiente | — |

Se carga como `SourceObject` del control `FormDetalle` del contenedor, de modo que **no es un formulario navegable por sí mismo**: no tiene barras de navegación ni botones de cerrar propios. La única forma de entrar es a través de la pestaña.

No acepta parámetros ni escribe variables globales. La salida es nula: ningún `DoCmd.Close` dirigido a terceros.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `Form_Load` | Comprueba que hay proyecto activo, puebla la lista y, ante error, avisa y **cierra el propio subformulario** | `:18-55` |
| `ListaAutorizados` | `ListBox` de 2 columnas | — | Encabezados «Autorizado» / «Con Correo»; una fila por responsable, con el nombre y el indicador de correo | `.form.txt:853`, `:859` (`ColumnCount =2`), `:900` |
| `Etiqueta30` | etiqueta | — | Rótulo «Autorizados» | `.form.txt:889-900` |

Solo hay una sección, `Detalle` (`.form.txt:842`), sin encabezado ni pie. La columna de correo **no es una dirección**: el valor procede de la propiedad `CorreoSiempre` del responsable, que es un indicador de que se le avisa siempre por correo, no el correo en sí.

## Flujos

### Apertura (`Form_Load`, `:18-55`)

1. Pone el reloj de arena y llama a `DoEvents` (`:21-22`).
2. **Si `m_ObjProyectoActivo` es `Nothing`**, compone el aviso «No se tienen datos del proyecto activo» y lanza `Err.Raise 1000` (`:25-29`).
3. Llama a `EstablecerDatos` con `m_Error` por referencia (`:34`); si haDevuelto mensaje, también `Err.Raise 1000` (`:35-37`).
4. Quita el reloj de arena (`:40-42`).

**Flujo de error** (`:45-54`): si el error no es el propio `1000`, arma un mensaje con el número y la descripción, invoca `CorreoAlAdministrador` y muestra un `MsgBox` crítico (`:47-50`); si es el `1000`, muestra un `MsgBox` de advertencia con el mensaje ya compuesto (`:51-52`). En ambos casos **cierra el subformulario** con `DoCmd.Close acForm, Me.Name, acSaveNo` (`:54`). Un fallo de carga deja al usuario en la pestaña de Autorizados sin contenido y sin más explicación que el cuadro de diálogo.

### Poblado (`EstablecerDatos`, `:57-91`)

1. Fija los encabezados de la lista con `lst.RowSource = "Autorizado;Con Correo"` (`:67`).
2. Si el proyecto activo no tiene expediente, **sale sin avisar** (`:68-70`).
3. Si el expediente no tiene colección de responsables, **sale sin avisar** (`:71-73`).
4. Recorre la colección y añade una fila por responsable con el nombre del usuario y el indicador de correo (`:76-81`).

No hay selección ni acción posterior: la lista es puramente informativa.

## Reglas, validaciones y cálculos

- **No hay ninguna validación de escritura.** No se comprueba rol, ni permiso, ni estado de la edición. Es coherente con su función de consulta, pero significa que la autorización real no se verifica en ninguna parte de este formulario.
- **Dos salidas silenciosas.** Sin expediente o sin responsables, `EstablecerDatos` devuelve vacío sin error (`:68-73`), y `Form_Load` **no cierra nada**: el usuario ve una pestaña con la lista en blanco y sin explicación de por qué. Es el caso más silencioso del lote.
- **El encabezado «Con Correo» no identifica al destinatario.** La columna muestra `CorreoSiempre` (`:79`), un booleano presentado como columna de información; conviene no interpretar que sea una dirección de correo.

## Lecturas y escrituras

| Operación | Origen | Evidencia |
|---|---|---|
| Proyecto activo | Variable global de sesión `m_ObjProyectoActivo` | `:25` |
| Expediente y responsables | `m_ObjProyectoActivo.Expediente.Responsables` | `:68-80` |
| Nombre e indicador de correo | `ExpedienteResponsable.Usuario.Nombre`, `.CorreoSiempre` | `:79` |

**Sin escrituras en base de datos.** El único efecto lateral es el aviso al administrador por correo ante errores inesperados (`:49`), que se dispara desde código y no desde una acción del usuario.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** subformulario de solo lectura que, al cargarse, vuelca en una lista de dos columnas los responsables del expediente activo con su nombre y su indicador de correo. Ante cualquier error avisa al usuario, notifica por correo al administrador y se cierra.

**Posibles defectos (sin corregir, a verificar):**

- **H-GR-261.** El mensaje de error de la línea `:46` va sin formato (`: "Al Form_Load se ha producido el error n: "`), y el número va **pegado a la palabra «n»** sin separador: el texto que ve el administrador dice literalmente «el error n: 2345». Es un defecto de redacción del literal, no de la lógica.
- **H-GR-262.** `Err.Raise 1000` (`:28, 36`) no lleva `vbObjectError`, al contrario que la práctica del resto del proyecto; conviene verificar en ejecución si Access ejecuta el manejador o lo trata como error de usuario. Si no lo ejecuta, la pestaña se queda en blanco sin aviso.
- **H-GR-263.** `EstablecerDatos` devuelve `p_Error` solo si el error **no** es `1000` (`:86-89`). Si el propio `EstablecerDatos` lanzara un `Err.Raise 1000`, el mensaje se perdería y `Form_Load` mostraría un `MsgBox` con `m_Error` vacío.
- **H-GR-264.** Los dos `Exit Function` de falta de expediente o de falta de responsables (`:68-73`) no distinguen entre «no aplica» y «no hay datos»: el usuario ve una lista vacía idéntica en ambos casos.
- **H-GR-265.** El cierre del subformulario en el manejador de error (`:54`) se produce **también** en el camino correcto, porque `Form_Load` termina sin dejar nada que mantener. En la práctica el cierre solo importa en el error, pero conviene documentarlo: no hay ningún estado en el que el subformulario sobreviva a su carga.
- **H-GR-266.** `CorreoAlAdministrador` (`:49`) se invoca también en errores del entorno local (por ejemplo, un error de presentación), lo que puede generar ruido en el buzón del administrador si el fallo es del cliente y no del servidor.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: sustituir el buzón por un listado de solo lectura con estado vacío explícito («esta gestión de riesgos todavía no tiene responsables asignados»); mostrar la columna de correo como un indicador booleano etiquetado («aviso siempre por correo») en lugar de un encabezado que invite a leer una dirección; derivar la lista de un permiso del servidor en vez de leerla del expediente; eliminar el subformulario y fundirlo en una sección estática de la página de detalle.

## Criterios de aceptación

Propuestos, no aprobados:

1. Con proyecto activo y expediente con responsables, la lista muestra una fila por responsable con nombre e indicador de correo.
2. Sin expediente o sin responsables, la vista muestra un estado vacío explicativo, no una lista en blanco.
3. Si no hay proyecto activo, se avisa con un mensaje comprensible y no se deja la vista a medio cargar.
4. El contenido es de solo lectura en todo momento.

## Pendientes

- Si `Responsables` es una colección cargada de forma perezosa o eagerly por el modelo `Expediente`; no se ve en este formulario.
- Qué hace exactamente `CorreoAlAdministrador` y si su activación por cada error de cliente es deseable.
- Si el indicador de correo debe mostrarse como columna o como distintivo junto al nombre.
- El esquema de la web: los cuatro subformularios del contenedor son, en la web, muy probablemente pestañas de una misma página (ver [forms](../../forms.md)).