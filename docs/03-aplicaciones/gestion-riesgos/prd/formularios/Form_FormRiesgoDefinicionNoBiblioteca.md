# Form_FormRiesgoDefinicionNoBiblioteca

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Fuente:** `src/forms/Form_FormRiesgoDefinicionNoBiblioteca.cls` (318 líneas) y `src/forms/FormRiesgoDefinicionNoBiblioteca.form.txt` (1281 líneas) · **PRD:** [README](../README.md)

Variante de la pestaña de definición para proyectos que **no** usan riesgo de biblioteca. Es el mismo formulario sin lista de riesgos tipo: los seis campos del riesgo se escriben a mano. El contenedor elige entre las dos variantes al montar el subformulario. Las citas son `Form_FormRiesgoDefinicionNoBiblioteca.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormRiesgo` | Subformulario: el contenedor fija `Me.NavGeneral.NavigationTargetName = "FormRiesgoDefinicionNoBiblioteca"` y lo monta en `Me.FrmDetalle` | `Form_FormRiesgo.cls:289`, `:307` |

No hay ningún `DoCmd.OpenForm` hacia él en el clon. A diferencia de su gemelo, **el contenedor no lo refresca** en su botón de actualizar: el despacho de `ComandoActualizar_Click` solo contempla `FormRiesgoDefinicion`, `FormRiesgoPlazoCosteCalidad`, `FormRiesgoVulnerabilidad`, `FormRiesgoMitigacion`, `FormRiesgoRetirado` y `FormRiesgoMaterializado` (`Form_FormRiesgo.cls:47-61`). No abre ningún formulario y no se cierra a sí mismo.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `Form_Open` | Refresca la pantalla | `:54-80`, `.form.txt:41` |
| `Descripcion` | cuadro de texto | `OnExit` → `Descripcion_Exit` | Descripción del riesgo | `:15-22`, `.form.txt:598-599` |
| `EntidadDetecta` | lista desplegable | `BeforeUpdate` → `EntidadDetecta_BeforeUpdate` | Entidad que detecta el riesgo y puebla «detectado por» | `:34-41`, `.form.txt:637-641` |
| `DetectadoPor` | lista desplegable | `BeforeUpdate` → `DetectadoPor_BeforeUpdate` | Persona o empresa que detecta | `:24-30`, `.form.txt:680-684` |
| `FechaDetectado` | cuadro de texto | `BeforeUpdate` → `FechaDetectado_BeforeUpdate` | Fecha de detección | `:44-50`, `.form.txt:716-718` |
| `Origen` | lista desplegable | `BeforeUpdate` → `Origen_BeforeUpdate` | Origen del riesgo | `:242-248`, `.form.txt:756-765` |
| `Priorizacion` | cuadro de texto | `OnChange` → `Priorizacion_Change` | Valida la priorización contra el resto de riesgos | `:253-317`, `.form.txt:797-799` |

El diseño conserva cinco etiquetas sin código (`Etiqueta232`, `Etiqueta34`, `Etiqueta36` a `Etiqueta38`; `.form.txt:826-946`). Los bloques `Begin` sin nombre de `.form.txt:415-552` son valores por defecto de propiedades. No hay control `CausaRaiz`: es la diferencia de conjunto con la variante de biblioteca.

## Flujos

### Apertura (`Form_Open` → `EstablecerDatos`, `:54-175`)

1. `Form_Open` pone el reloj y llama a `EstablecerDatos` (`:61`), propagando el error (`:63`).
2. `EstablecerDatos` exige edición activa (`:94-96`) y, si el permiso del contenedor está `Empty`, lo **calcula y escribe en él** (`:98-107`).
3. Llena el combo de origen con `EstablecerComboOrigen` (`Funciones Generales.bas:6973`) (`:114`).
4. Fija `Me.AllowEdits` según el permiso (`:120-123`).
5. Vuelca los seis campos del riesgo activo (`:127`, `:131`, `:134`, `:141`, `:144`, `:147`), comprueba la priorización con `Constructor.EsPriorizacionCorrecta` (`Constructor.bas:6984`) (`:150`, `:153`), ajusta el control (`:158-160`) y marca los bordes obligatorios (`:164`).

### Normal: priorización no válida (`Priorizacion_Change`, `:253-317`)

1. Pone el reloj y devuelve el foco a la fecha y a la propia priorización (`:263-264`).
2. Valida por edición si es alta (`:268`) y por riesgo si es edición (`:271`).
3. Colorea las pestañas (`:275`) y sale si la priorización está vacía (`:276-279`).
4. Si es incorrecta y hay copia inicial con valor numérico, restaura la priorización anterior en el control y en el objeto (`:284-289`); si no, la vacía (`:293`).
5. Avisa «Esa priorización no es correcta…» y lanza el error controlado (`:297-298`).

### Alternativo: elegir entidad

`EntidadDetecta_BeforeUpdate` llama a `RellenarComboDetectadoPor` (`:36`), que solo puebla la lista si la entidad es «Persona» (`:202-215`) o «Empresa» (`:218-228`): para «Persona» usa `Proyecto.CadenaNombreAutorizados` y propaga su error (`:204-205`); para «Empresa» recorre los suministradores de la edición (`:219-226`).

### De error

`Form_Open` (`:71-78`) y `Priorizacion_Change` (`:308-315`) llevan el patrón común: quitar el reloj y, si `Err.Number <> 1000`, componer el mensaje, escribirlo al administrador (`Funciones Generales.bas:964`) y avisar con diálogo crítico; si es 1000, solo advertencia. Ninguno cierra. Los cinco manejadores que solo pintan bordes y colores no tienen manejador (`:16`, `:25`, `:35`, `:45`, `:243`).

## Reglas, validaciones y cálculos

- **La elección entre las dos variantes la hace el contenedor** según `Proyecto.RequiereRiesgoDeBibliotecaCalculado` (`Proyecto.cls:667`): `FormRiesgoDefinicion` (`:286-287`) o esta (`:288-289`).
- **Los campos se escriben a mano.** No hay lista de riesgos tipo ni `CausaRaiz`: el volcado del contenedor copia seis campos aquí frente a siete en la variante de biblioteca (`Form_FormRiesgo.cls:502-509` y `:510-516`).
- **La priorización se valida contra el resto de riesgos** de la edición con `Constructor.EsPriorizacionCorrecta` (`Constructor.bas:6984`), al cambiar el campo (`:268`, `:271`) y también al cargar (`:150`, `:153`).
- **El permiso de edición se globaliza.** Aquí se usa `Me.AllowEdits` (`:120-123`), mientras que la variante de biblioteca deshabilita campo a campo (`Form_FormRiesgoDefinicion.cls:332-339` y `:410-419`).
- **Sin acceso directo a datos:** sin `getdb()`, `Recordset`, `Open`, `Execute`, SQL, `DLookup` ni transacciones; todo pasa por `Constructor` y por las clases.

## Lecturas y escrituras

| Operación | Origen | Evidencia |
|---|---|---|
| Riesgo sobre el que se opera | `m_ObjRiesgoActivo` (`Variables Globales.bas:80`) | `:126-153`, `:219`, `:287` |
| Edición activa | `m_ObjEdicionActiva` (`Variables Globales.bas:79`) | `:94`, `:100`, `:204`, `:265` |
| Copia para restaurar la priorización | `m_ObjRiesgoAlInicio` (`Variables Globales.bas:83`), **leída pero nunca escrita en este fichero** | `:284-287` |
| Permiso de edición | `Form_FormRiesgo.m_PermitidoEditar` (`Form_FormRiesgo.cls:22`), **escrito aquí** | `:98`, `:101-107`, `:120` |
| Validación de priorización | `Constructor.EsPriorizacionCorrecta` (`Constructor.bas:6984`) | `:150`, `:153`, `:268`, `:271` |
| Personas autorizadas | `m_ObjEdicionActiva.Proyecto.CadenaNombreAutorizados` | `:204` |
| Suministradores de la edición | `m_ObjRiesgoActivo.Edicion.ColSuministradores` | `:219` |
| Combo de origen | `EstablecerComboOrigen` (`Funciones Generales.bas:6973`) | `:114` |
| Escritura directa en el objeto | `.Priorizacion` en el rollback (`:287`) | `:287` |

Este formulario **no escribe en la base de datos**: deja el objeto riesgo en memoria y el contenedor lo registra. El volcado definitivo lo hace `RellenarObjeto` (`Form_FormRiesgo.cls:510-516`).

## Coexistencia

Pendiente de `../coexistencia.md`. Lo relevante aquí es que hoy la misma pantalla existe en dos variantes que el usuario no elige: las elige el proyecto. En web conviene un único formulario de definición con la parte de biblioteca activa o inactiva según el proyecto.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-GR-164.** El contenedor **no refresca esta variante** en su botón de actualizar: el despacho de `ComandoActualizar_Click` (`Form_FormRiesgo.cls:47-61`) no contempla `FormRiesgoDefinicionNoBiblioteca`. Actualizar en un proyecto sin riesgo de biblioteca recarga el riesgo y el icono, pero deja la pestaña con los valores de antes.
- **H-GR-165.** `Priorizacion_Change` lee `m_ObjRiesgoAlInicio` para restaurar la priorización anterior (`:284-287`), pero **este formulario nunca lo escribe**. Solo lo inicializan el contenedor (`:42`, `:252`) y la variante de biblioteca (`Form_FormRiesgoDefinicion.cls:370`, `:380`, `:399`). Si esta variante es la primera en ejecutarse, la restauración no existe y la priorización se vacía (`:293`).
- **H-GR-166.** Los cinco manejadores de `BeforeUpdate` y `Exit` no tienen la guarda `If m_ObjEdicionActiva.EsActivo = Sí` que sí tienen los de la variante de biblioteca (aquí `:19`, `:28`, `:39`, `:48`, `:246`; allí `Form_FormRiesgoDefinicion.cls:24`, `:83`, `:92`, `:104`, `:114`, `:591`). Sobre una edición cerrada, esta pestaña sigue marcando bordes y recoloreando pestañas.
- **H-GR-167.** `Priorizacion_Change` valida y colorea sin la guarda de edición activa (`:271`, `:275`), a diferencia de su gemelo (`Form_FormRiesgoDefinicion.cls:649`, `:655`).
- **H-GR-168.** `EstablecerDatos` **escribe el permiso del contenedor** (`:101-107`), igual que la otra variante (`Form_FormRiesgoDefinicion.cls:349-355`). Un subformulario modifica el estado de su padre y el permiso final depende del orden de carga de las pestañas.
- **H-GR-169.** `RellenarComboDetectadoPor` está **duplicada literalmente** entre las dos variantes (`Form_FormRiesgoDefinicion.cls:524-585` y aquí `:177-238`): mismas ramas, mismos mensajes y mismo propagación del error de `Proyecto`. Son dos copias que van a divergir.
- **H-GR-170.** Los cinco manejadores sin `On Error` silencian los fallos al pintar bordes y colores (`:16`, `:25`, `:35`, `:45`, `:243`), igual que ocurre en la otra variante (`Form_FormRiesgoDefinicion.cls:21`, `:81`, `:90`, `:101`, `:112`, `:589`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: un único formulario de definición con la parte de biblioteca condicionada por el proyecto; permiso calculado una sola vez en el contenedor; validación de priorización en el servidor, común a las dos variantes; lista de personas y empresas como recurso único y compartido.

## Criterios de aceptación

Propuestos, no aprobados:

1. En un proyecto sin riesgo de biblioteca, el contenedor monta esta variante y muestra los seis campos.
2. Actualizar desde el contenedor repuebla la pestaña en las dos variantes.
3. Sin permiso de edición, los campos no son editables.
4. Una priorización incorrecta no se guarda y el usuario recibe el aviso con el campo corregido.
5. Elegir entidad «Persona» o «Empresa» repuebla «detectado por» con los valores correctos.

## Pendientes

- Si el diseño de las dos variantes se mantiene sincronizado a mano o por herramienta, porque comparten seis de sus campos y difieren solo en `CausaRaiz` y en la lista de riesgos tipo.
- Qué escribe `m_ObjRiesgoAlInicio` cuando el proyecto no usa riesgo de biblioteca y el contenedor no pasa por la otra variante (**H-GR-165**).
- Si `RellenarBordeCamposObligatorios` (`Funciones Generales.bas:7095`) y `EstablecerColorPestañasRiesgo` (`Funciones Generales.bas:7372`) actúan sobre este subformulario o sobre el formulario activo, porque se les pasa `Me` desde un subformulario (**H-GR-170**).