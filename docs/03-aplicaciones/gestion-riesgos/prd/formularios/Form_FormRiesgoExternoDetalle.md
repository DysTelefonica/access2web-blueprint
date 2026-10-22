# Form_FormRiesgoExternoDetalle

> **Commit fijado:** `36fcf60161fa5ec84c3956b1308537ddf2579256` · **Fuente:** `src/forms/Form_FormRiesgoExternoDetalle.cls` (416 líneas) y `src/forms/FormRiesgoExternoDetalle.form.txt` (1970 líneas) · **PRD:** [README](../README.md)

Detalle editable de un riesgo externo: those are risks that come from outside the application (un proveedor, un pedido o un organismo) and that here only can be translated, described and linked to a risk type from the library. Es una pantalla de formulario normal, no una pestaña del contenedor `Form_FormRiesgo`. Las citas son `Form_FormRiesgoExternoDetalle.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_FormGestionRiesgosRiesgosOferta` | `DoCmd.OpenForm "FormRiesgoExternoDetalle"` desde el botón de detalle, cerrando antes la instancia previa | `Form_FormGestionRiesgosRiesgosOferta.cls:169-172` |

El contenedor se declara `WithEvents` y toma la instancia abierta (`:18`, `:174`), y escucha su evento `Registrado` para refrescar (`Form_FormGestionRiesgosRiesgosOferta.cls:716-736`). No hay ninguna referencia a `Form_FormRiesgo` en este fichero: el riesgo externo no vive en el contenedor de riesgo sino en la gestión de riesgos por oferta.

Abre `FormRiesgosBibliotecaGestion` (`:178`) y **habilita su botón de elegir desde fuera** (`:182-183`). Se cierra a sí mismo en tres puntos (`:133`, `:153`, `:274`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `Form_Open` | Ajusta el tamaño, valida el riesgo externo y decide el permiso | `:204-275`, `.form.txt:41` |
| `ComandoGrabar` | botón | `OnClick` → `ComandoGrabar_Click` | Valida, registra y cierra | `:108-147`, `.form.txt:1529-1531` |
| `ComandoSalir` | botón | `OnClick` → `ComandoSalir_Click` | Cierra sin guardar | `:149-165`, `.form.txt:1454-1456` |
| `ComandoAyuda` | botón | `OnClick` → `ComandoAyuda_Click` | Abre la ayuda de la pantalla | `:80-106`, `.form.txt:920-922` |
| `ComandoSeleccionarRiesgoDeBiblioteca` | botón | `OnClick` → `ComandoSeleccionarRiesgoDeBiblioteca_Click` | Abre la biblioteca para elegir un riesgo tipo | `:167-202`, `.form.txt:1407-1409` |
| `Trasladar` | lista desplegable | `OnExit` → `Trasladar_Exit` | Indica si el riesgo se traslada al application's risk catalogue | `:323-337`, `.form.txt:1038-1043` |
| `CodRiesgo` | cuadro de texto | — | Código con el que llega el riesgo externo | `:1002`, `:283` |
| `Descripcion` | cuadro de texto | — | Descripción del riesgo; bloqueada si el proyecto usa biblioteca | `:1076`, `:35`, `:236-246` |
| `MotivoNoIntegrado` | cuadro de texto | — | Motivo por el que no se integra | `:1139`, `:39-40`, `:312-319` |
| `CodRiesgoBiblioteca` | cuadro de texto | — | Código del riesgo tipo elegido | `:1170`, `:43`, `:286`, `:387` |
| `CausaRaiz` | cuadro de texto | — | Causa raíz; solo en proyectos con riesgo de biblioteca | `:1199`, `:37`, `:238-246` |

El manejador `m_FormRiesgosBibliotecaGestion_RiesgoDeBibliotecaElegido` (`:380-416`) no es un evento de diseño: se ata al evento de la gestión de biblioteca en `Form_FormRiesgosBibliotecaGestion.cls:18`. El diseño conserva seis etiquetas sin código (`.form.txt:1227-1377`) y un cuadro de texto, `CodigoRiesgoEnEdicion` (`.form.txt:1108`), que el código no utiliza.

## Flujos

### Apertura (`Form_Open`, `:204-275`)

1. Reloj de arena, `m_Error = ""` y `AjustarTamaño Me` (`:208-211`); el título de la ventana se copia de la etiqueta (`:213`).
2. Si `m_ObjRiesgoExtActivo Is Nothing`, avisa «No se ha podido determinar el riesgo externo activo para editar» y aborta (`:215-217`).
3. Recupera el riesgo con `Constructor.getRiesgoExterno` (`Constructor.bas:2844`) (`:220`).
4. Con `UsuarioAutorizado = Sí` habilita edición y guardado (`:224-232`); si no, los bloquea (`:251-255`).
5. Si el proyecto requiere riesgo de biblioteca, habilita el botón de biblioteca, bloquea la descripción y habilita la causa raíz (`:233-238`); si no, hace lo contrario (`:240-245`); si el indicador no es sí ni no, avisa y aborta (`:247-249`).
6. `EstablecerDatos` (`:257`).

### Normal: guardar (`ComandoGrabar_Click`, `:108-147`)

1. `HaHabidoCambios` (`:119`); si no hay cambios, avisa «Sin cambios que guardar» y aborta (`:124-125`).
2. `Registrar m_Error` (`:128`), que vuelca los controles en el objeto y llama a la clase (`:34-44`).
3. Cierra (`:133`).

### Alternativo: elegir riesgo tipo (`:380-416`)

1. Escribe el código del riesgo tipo elegido (`:387`).
2. Si el riesgo no está en retipificación, copia su descripción y habilita descripción y causa raíz (`:389-394`); si lo está, vacía la descripción y bloquea ambos (`:395-400`).

### Alternativo: trasladar o no

`Trasladar_Exit` (`:323-337`) bloquea y vacía `MotivoNoIntegrado` cuando `Trasladar = "Sí"` (`:328-331`), y lo restaura cuando no (`:334-335`). `Trasladar_AfterUpdate` (`:306-321`) tiene el mismo cuerpo.

### De error

Los manejadores siguen el patrón común: quitar el reloj y, si `Err.Number <> 1000`, componer el mensaje, escribirlo al administrador (`Funciones Generales.bas:964`) y avisar con diálogo crítico; si es 1000, solo advertencia (`:99-104`, `:139-145`, `:156-164`, `:193-201`, `:266-272`). **`Form_Open` cierra siempre**, también en error (`:274`). `EstablecerDatos` (`:299-302`), `HaHabidoCambios` (`:374-377`) y `Registrar` (`:74-77`) solo rellenan su argumento de error.

## Reglas, validaciones y cálculos

- **El riesgo externo depende de una global.** `m_ObjRiesgoExtActivo` (`Variables Globales.bas:86`) es el único punto de entrada: si falta, la pantalla no abre (`:215-217`).
- **El proyecto decide qué campos se rellenan a mano.** Con `RequiereRiesgoDeBibliotecaCalculado = Sí`, la descripción viene del riesgo tipo y la causa raíz se escribe (`:233-238`); sin él, al revés (`:240-245`).
- **El motivo solo tiene sentido si no se traslada.** `Trasladar = "Sí"` lo bloquea y lo vacía (`:312-314`).
- **Solo se guarda si algo cambió.** `HaHabidoCambios` compara `Trasladar`, `Descripcion`, `MotivoNoIntegrado` y `CodRiesgoBiblioteca` (`:348-361`), y `CausaRaiz` solo si el proyecto usa biblioteca (`:364-366`).
- **La fecha de motivo se rellena sola.** Si el motivo no es una fecha, `Registrar` pone la fecha de hoy (`:40-41`).
- **Sin acceso directo a datos** en el `.cls`: sin `getdb()`, `Recordset`, `Open`, `Execute`, SQL, `DLookup` ni transacciones. La escritura la hace `RiesgoExterno.Registrar` (`RiesgoExterno.cls:799`), que internamente usa `getdb().OpenRecordset` (`:838`) y `getdb().Execute` (`:1073`, `:1189`, `:1232`, `:1338`).

## Lecturas y escrituras

| Operación | Origen | Evidencia |
|---|---|---|
| Riesgo externo sobre el que se opera | `m_ObjRiesgoExtActivo` (`Variables Globales.bas:86`), con copia local `m_RiesgoExtAlinicio` | `:32`, `:44`, `:215`, `:220`, `:343` |
| Riesgo del application's catalogue | `m_ObjRiesgoActivo` (`Variables Globales.bas:80`), **sustituido aquí** | `:52` |
| Permiso de escritura | `blnPermitidoEscribir` (`Variables Globales.bas:121`), escrito en `:254` | `:254`, `:308`, `:325` |
| Recarga del riesgo externo | `Constructor.getRiesgoExterno` (`Constructor.bas:2844`) | `:220` |
| Lectura del riesgo asociado | `RiesgoExterno.riesgo` (`RiesgoExterno.cls:269`), que llama a `Constructor.getRiesgo` (`RiesgoExterno.cls:281`) | `:52` |
| Guardado | `RiesgoExterno.Registrar` (`RiesgoExterno.cls:799`) | `:44` |
| Aviso de retipificación | `EnviarCorreoRiesgoRequiereRetipificacion` (`Funciones Generales.bas:3006`) | `:55` |
| Elección desde la biblioteca | evento `RiesgoDeBibliotecaElegido` de `Form_FormRiesgosBibliotecaGestion.cls:143` y `:336` | `:380-416` |

Escritura efectiva: una llamada a `RiesgoExterno.Registrar`, más la sustitución de la global del riesgo y un correo si procede.

## Coexistencia

Pendiente de `../coexistencia.md`. Esta pantalla es la que más se parece al patrón de la web: un recurso riesgo externo con su transición de estado (trasladar sí o no) y su enlace opcional a un riesgo tipo. Lo que hay que decidir es si la transición se hace con un formulario de confirmación o con una acción idempotente y auditada.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- **H-GR-201.** `Form_Open` cierra el formulario **siempre**, también en error (`:274`). Un fallo puntual al cargar deja al usuario fuera de la pantalla con el aviso en un diálogo ya sin ventana que lo respalde.
- **H-GR-202.** `Trasladar_AfterUpdate` (`:306-321`) **no tiene evento en el diseño**: el diseño solo declara `OnExit` para `Trasladar` (`.form.txt:1043`). El manejador existe y duplica el de `Trasladar_Exit` (`:323-337`), pero no se ejecuta; conviene confirmar que no se lose efecto en algún momento de la edición.
- **H-GR-203.** Con `Trasladar = "Sí"`, `MotivoNoIntegrado` se bloquea **y se vacía** (`:313-314`), sin aviso. Si el motivo estaba escrito y el usuario cambia de opinión, lo pierde en silencio.
- **H-GR-204.** `Registrar` asigna `FechaMotivo = Date` cuando el motivo escrito no es una fecha (`:40-41`). La base de datos recibe la fecha de hoy como si el usuario la hubiera indicado.
- **H-GR-205.** `Registrar` sustituye la global `m_ObjRiesgoActivo` por el riesgo asociado (`:52`). Un formulario de riesgo externo reescribe el riesgo activo del contenedor de riesgo, de modo que la pestaña que se abra después muestra otro riesgo.
- **H-GR-206.** `ComandoSeleccionarRiesgoDeBiblioteca_Click` abre la gestión de biblioteca y le modifica sus controles desde fuera (`.ComandoElegir.Visible` y `.Enabled`, `:182-183`). La pantalla deja a la otra en un estado que su propio código no pone.
- **H-GR-207.** El diseño conserva un cuadro de texto, `CodigoRiesgoEnEdicion` (`.form.txt:1108`), que el código no lee ni escribe en ningún punto.
- **H-GR-208.** `HaHabidoCambios` solo compara `CausaRaiz` cuando el proyecto usa riesgo de biblioteca (`:364-366`). Si el indicador cambia con la pantalla abierta, un cambio real en la causa raíz no se detecta y el guardado se rechaza como «Sin cambios que guardar».

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidatas: transición de estado explícita y auditada en lugar de sobrescribir una global; el motivo y su fecha como datos opcionales y coherentes entre sí; elección del riesgo tipo mediante una consulta en lugar de manipulating el formulario vecino; ninguna escritura silenciosa de la fecha del sistema.

## Criterios de aceptación

Propuestos, no aprobados:

1. Abrir la pantalla muestra el riesgo externo con su código, descripción y estado de traslado.
2. En un proyecto con riesgo de biblioteca, la descripción es de solo lectura y la causa raíz se puede escribir.
3. Marcar «trasladar» bloquea el motivo y avisa si había algo escrito.
4. Guardar sin cambios avisa y no escribe nada.
5. Elegir un riesgo tipo de la biblioteca actualiza el código y la descripción según el modo del proyecto.
6. Tras registrar, el contenedor que abrió la pantalla se refresca.

## Pendientes

- Qué tablas escribe `RiesgoExterno.Registrar` (`RiesgoExterno.cls:799`) y si su uso de `getdb().Execute` (`:1073`, `:1189`, `:1232`, `:1338`) implica operaciones fuera de transacción.
- Qué criterio usa `EnviarCorreoRiesgoRequiereRetipificacion` (`Funciones Generales.bas:3006`) y a quién avisa (**H-GR-205**).
- Si existe otro punto que rellene `MotivoNoIntegrado` con una fecha válida, dado que aquí se fuerza la fecha del sistema (**H-GR-204**).
- Si el evento `Registrado` (`:69`) lo consume alguien más además de `Form_FormGestionRiesgosRiesgosOferta`.