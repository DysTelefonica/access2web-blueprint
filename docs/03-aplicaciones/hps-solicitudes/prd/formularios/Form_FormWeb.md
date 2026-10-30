# Form_FormWeb

> **Commit fijado:** `cfa1937ab05a9e40150ea9236243898de96f84da` · **Fuente:** `src/forms/Form_FormWeb.cls` (56 líneas) y `.form.txt` (15 bloques `Begin`) · **PRD:** [README](../README.md)
>
> **Abreviaturas:** `FU.bas` = `modules/FUNCIONES UTILES.bas`, `VG.bas` = `modules/Variables Globales.bas`. Las citas son `Form_FormWeb.cls:línea` salvo indicación; para el diseño, `Form_FormWeb.form.txt:línea`.

Visor genérico de HTML para la aplicación. No es una ventana de negocio: es el **proyector** que `FU.bas` abre después de generar un fichero HTML temporal o de plantilla, y que se cierra sola en cuanto se abre la siguiente. Es **no enlazado** (sin `RecordSource` efectivo), no declara parámetros ni eventos públicos, y no escribe nada: su único efecto es cargar en un control de explorador web el contenido de un fichero local.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `FU.bas:1336-1392` (`ObtenerHTMLEnLocal`) | Genera el HTML, deja la ruta en `m_URLHTMLActivo`, cierra `FormWeb` si ya estaba abierto y hace `DoCmd.OpenForm "FormWeb"` | `FU.bas:1381-1385` |
| `FU.bas:1721-1748` (`InformeDelRegistro`) | Igual, con la ruta que devuelve `GenerarArchivoConHTML` | `FU.bas:1734, 1738-1741` |
| `Test.bas:192-196` | Réplica de prueba del primer camino | `Test.bas:192-196` |

| Hacia | Cómo | Evidencia |
|---|---|---|
| (ninguno) | El formulario no abre ningún otro formulario, no llama a `DoCmd` de salida y no cierra al usuario: se cierra a sí mismo desde `Form_Load`. | `:22-26, 45` |

No hay forma de cerrar la ventana desde la interfaz: el `.form.txt` no declara ningún control con evento `OnClick` y el `.cls` solo implementa `Form_Load` y `Form_Resize` (`:9, 51`). Los controles que el diseño conserva (botón, grupo de opciones, lista, combo, marco de objeto, imagen, rectángulo, etiqueta) no tienen nombre propio ni handler: son restos de una plantilla reutilizada.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| Formulario | — | `OnLoad` → `Form_Load` | Proyecta el HTML; cuerpo completo en §Flujos | `:9-46` |
| Formulario | — | `OnResize` → `Form_Resize` | Reajusta el explorador al `InsideWidth`/`InsideHeight` del formulario | `:51-56` |
| `Navegador` | `CustomControl` (control de explorador web de Microsoft) | — | Receptor de `Navigate`; es el único control con nombre propio del diseño | `:30, 54-55`, `.form.txt:356` |
| Sin nombre | botón, grupo de opciones, `TextBox`, `ListBox`, `ComboBox`, `UnboundObjectFrame`, `Image`, `Rectangle`, `Label` | — | Restos de plantilla: el diseño los declara pero el `.cls` no los menciona en ninguna línea | `.form.txt:139-332` |
| `Detalle` | sección | — | Sección de detalle vacía; el contenido real lo pinta `Navegador` | `.form.txt:344-347` |

## Flujos

### Normal

1. `Form_Load` activa el reloj de arena y despacha eventos (`:13-15`), vacía `m_Error` y deja el formulario en solo lectura (`:16-17`).
2. `Ajustar Me` dimensiona la ventana al contenido (`:18`; `FU.bas:735-759`); `On Error Resume Next` en `Ajustar` hace que un fallo de ajuste pase desapercibido.
3. Comprueba que el fichero exista: `If Not fso.FileExists(m_URLHTMLActivo)` (`:22`; `fso` es el `FileSystemObject` público de `VG.bas:179`). Si no existe, cierra el formulario sin más (`:23-26`).
4. Navega: `Me.Navegador.Navigate m_URLHTMLActivo` (`:30`), con la ruta ya absoluta y construida por el llamador (`FU.bas:1368` la compone como `URLDirectorioLocal` + un nombre temporal, y `FU.bas:1581-1599` la escribe con `ADODB.Stream` en UTF-8 sobrescribiendo).
5. Quita el reloj de arena, reactiva `AllowEdits` y termina dejando la ventana abierta (`:32-35`).

### Alternativos

- **Segundo uso en la misma sesión:** `FU.bas:1382-1384` y `FU.bas:1738-1740` cierran la instancia previa con `acSaveNo` antes de abrir la nueva, de modo que la ventana se recarga con el HTML nuevo en lugar de acumularse. `FormularioAbierto` (`FU.bas:720`) protege la secuencia.
- **Sin fichero que proyectar:** es la rama de `:22-26`, tratada en §Errores.

### Errores

- `errores:` (`:37-45`) distingue el centinela interno `Err.Number <> 1000`: solo los errores reales llenan `m_Error` y disparan `CorreoAlAdministrador` (`FU.bas:585`); el 1000 se considera ruido del propio código. En ambos casos reactiva `AllowEdits` y quita el reloj de arena (`:42-43`), muestra el `MsgBox` y cierra la ventana con `acSaveNo` (`:44-45`).
- `Form_Resize` (`:51-56`) **no tiene manejador de error**: cualquier excepción al redimensionar el explorador sube al diálogo de error de Access, sin correo al administrador ni cierre controlado.

## Reglas, validaciones y cálculos

- **Precondición única:** el fichero de `m_URLHTMLActivo` tiene que existir en el momento del `Load` (`:22`). La comprobación es sólo de existencia: no se valida el formato ni el tamaño, y no se comprueba nada más tarde.
- **Contrato de salida con el llamador:** `FormWeb` no devuelve nada. Quien confirma el resultado es `FU.bas`, que escribe `"OK"` en el valor de retorno de `InformeDelRegistro` (`:1742`) o propaga `p_Error`; el formulario no participa.
- **Estado global compartido:** la ruta viaja por `Public m_URLHTMLActivo As String` (`VG.bas:206`), que `VG.bas:221` reinicia al inicializar. Con dos exploraciones solapadas en sesiones distintas del mismo proceso, la segunda sobrescribe la ruta de la primera.
- **`pregunta` no es un control:** `pregunta = MsgBox(...)` (`:44`) usa el `Public pregunta As Long` de `VG.bas:182` como variable de descarte del resultado del diálogo, no un control del formulario.

## Lecturas y escrituras

| Operación | Tabla o consulta | Tipo | Evidencia |
|---|---|---|---|
| Existencia del HTML | (ninguna) | `FileSystemObject.FileExists` | `:22`, `VG.bas:179` |
| Carga del HTML | (ninguna) | E/S de fichero vía el control de explorador | `:30` |
| Tamaño de la ventana | (ninguna) | Derivado, sin DAO | `:54-55` |

**Este formulario no lee ni escribe ninguna tabla.** Todo el acceso a datos ocurre aguas arriba, en `FU.bas`: `GenerarHTML` (`FU.bas:1179`) para las plantillas y `GenerarArchivoConHTML` (`FU.bas:1570`) + `HTMLLog` (`FU.bas:1613`) para el informe del registro.

## Coexistencia

Pendiente de la coexistencia (documento aún no redactado). Lo relevante para la conversión: el contrato es **«generar HTML a fichero local y proyectarlo»**, no una consulta a la base de datos. La conversión a PostgreSQL no toca este formulario, pero sí a su dependencia: `GenerarHTML`/`HTMLLog` resuelven plantillas y datos con las reglas de `Entorno` (`m_ObjEntorno.URLPlantillasHTML*`, `URLDirectorioLocal`). En web, el equivalente natural es renderizar la plantilla en servidor y devolver una respuesta, sin escribir en el disco del cliente ni depender del registro de Windows.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-HPS-S-31.** El rótulo de diseño es `Caption ="EXPEDIENTES Versión 2024-001 Administrador"` (`.form.txt:35`), que describe otra aplicación y una versión de expediente, no este visor. Ni `Form_Load` ni `Form_Resize` lo sobrescriben, así que el título que ve el usuario es siempre ese texto.
- **H-HPS-S-32.** Cuando el error capturado es el centinela interno `1000`, `m_Error` sigue vacío (`:38-41`) y se muestra igualmente `MsgBox(m_Error, vbCritical, "Error")` (`:44`): diálogo crítico sin texto, cuyo único efecto es el cierre. El mismo `MsgBox` se dispararía con `m_Error` vacío si `Err.Description` fuese una cadena vacía.
- **H-HPS-S-33.** Si el fichero de `m_URLHTMLActivo` no existe, el formulario se cierra sin ninguna comunicación (`:22-26`): desde el punto de vista del usuario la ventana desaparece al abrirse. El llamador, que sí ha pagado el coste de generar el HTML, no recibe ninguna señal.
- `Form_Resize` carece de manejador de error (`:51-56`) a diferencia de los dos únicos procedimientos que sí lo tienen: un fallo al redimensionar el explorador sale por la vía de Access, sin `CorreoAlAdministrador` ni cierre limpio.
- `Me.AllowEdits = False` (`:17`) se revierte sin condición al final de la ruta normal (`:35`) y también en `errores:` (`:42`); sobre un formulario no enlazado la propiedad es inocua, pero las tres líneas sugieren una restricción de diseño que no existe.
- Restos de la plantilla de origen en el diseño y en el código: botón, grupo de opciones, `TextBox`, `ListBox`, `ComboBox`, `UnboundObjectFrame`, `Image`, `Rectangle` y `Label` sin nombre ni handler (`.form.txt:139-332`), y una ruta absoluta ajena a este entorno comentada en el propio `Form_Load` (`:29`, `https://msaccess/C:\Users\adm.DEFENSA\AppData\...`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: renderizado de la plantilla en servidor y respuesta HTTP en lugar de fichero local + explorador incrustado; apertura en pestaña nueva cuando el contenido deba conservar el estado del expediente; y supresión del formulario en favor de un visor de solo lectura con botón de impresión explícito.

## Criterios de aceptación

Propuestos, no aprobados:

1. Con un fichero válido en `m_URLHTMLActivo`, el visor muestra su contenido completo y dimensiona su área al redimensionar la ventana.
2. Con una ruta que no existe, se informa del motivo y no se abre ninguna ventana en blanco.
3. Si la proyección falla, el aviso contiene el código y la descripción reales del error.
4. Un error interno de la propia aplicación no produce un diálogo crítico vacío.
5. Abrir dos proyectos seguidos en la misma sesión muestra el segundo contenido y no el primero.

## Pendientes

- Recuento de controles: el inventario fija 16 y el `.form.txt` de este commit contiene 15 bloques `Begin <Tipo>` (`Form_FormWeb.form.txt:5, 139-356`); la diferencia no se reproduce con ninguna regla de recuento simple sobre el fichero. Se cita la cifra del inventario para no romper la tabla, pero conviene reconciliarla.
- El `.form.txt` sólo declara nombre propio para `Navegador` y `Detalle`: no se puede precisar el nombre por defecto de los controles restantes, sólo su tipo y su línea.
- Identidad y versión exacta del `CustomControl` de `.form.txt:356` (control de explorador web de Microsoft): no consta en `src`, y su disponibilidad depende de la configuración de Access del equipo.
- Qué hace el usuario con la ventana: no hay control de cierre, así que la única salida es la apertura del siguiente `FormWeb` o el cierre de la aplicación.