# Form_Form0BDInformes

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_Form0BDInformes.cls` (107 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Submenú «INFORMES» de Brass: una columna de etiquetas que actúan como enlaces y abren los formularios de informes, partes y SLA. Es un formulario **no enlazado** (el `.form.txt` no declara `RecordSource`) y sin lógica de datos: cada etiqueta ejecuta `DoCmd.OpenForm` y cierra este formulario. Las citas son `Form_Form0BDInformes.cls:línea` (abreviadas `:N`) salvo indicación; `FG.bas` = `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_Form0BDOpciones` | Etiqueta `lblInformes`: `DoCmd.OpenForm "Form0BDInformes"` sin argumentos | `Form_Form0BDOpciones.cls:104-107` |

Salidas: cada etiqueta abre un formulario (tabla siguiente) y a continuación cierra éste con `DoCmd.Close acForm, Me.Name` (`:35, :48, :73, :80, :90, :101`). `cmdSalir` lo cierra sin guardar (`:6-9`).

## Controles y eventos

Controles: 13 (20 `Begin` en el `.form.txt`: 1 `Form`, 5 plantillas de estilo por defecto —`Label`, `Rectangle`, `Line`, `Image`, `CommandButton`—, 1 `Section` y 13 controles con nombre).

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `lblGestionDePartes` («- Gestión de Informes de Avería») | etiqueta | `Click` → `lblGestionDePartes_Click`; `MouseMove` | Abre `FormParteGestion` y cierra | CAP-10 | `:45-49, :53-57`; `Form_Form0BDInformes.form.txt:516-518` |
| `lblConsultarHorasTecnico` («- Consultar Horas Técnico») | etiqueta | `Click`; `MouseMove` | Abre `FormTecnicoGestion` y cierra | CAP-9 | `:32-36, :38-42`; `Form_Form0BDInformes.form.txt:550-552` |
| `lblInformeParaRAC` («- Para RAC») | etiqueta | `Click`; `MouseMove` | Abre `FormInformesParaRAC` y cierra | CAP-10 | `:70-74, :65-69`; `Form_Form0BDInformes.form.txt:829-831` |
| `lblInformeParaRACPorFactura` («- Por Factura») | etiqueta | `Click`; `MouseMove` | Abre `FormInformesPorFactura` y cierra | CAP-10 | `:77-81, :82-86`; `Form_Form0BDInformes.form.txt:863-865` |
| `lblInformeParaRACDesdeLista` («- Desde Lista») | etiqueta | `Click`; `MouseMove` | Abre `FormInformesDesdeLista` y cierra | CAP-10 | `:87-91, :92-96`; `Form_Form0BDInformes.form.txt:897-899` |
| `lblInformeSLA` («- Informes de Cumplimento SLA») | etiqueta | `Click`; `MouseMove` | Abre `FormInformeSLA` y cierra | CAP-10 | `:98-102, :103-107`; `Form_Form0BDInformes.form.txt:931-933` |
| `lblInformeParaRAC_General` («- Informes de Partes de Trabajo») | etiqueta | — (sin `OnClick` ni handler) | No hace nada | CAP-10 | `Form_Form0BDInformes.form.txt:584-585` |
| `cmdSalir` | botón | `Click` → `cmdSalir_Click` | Cierra sin guardar | CAP-10 | `:6-9`; `Form_Form0BDInformes.form.txt:799-801` |
| `lbTitulo`, `lblInformes` | etiquetas | — | Rótulos «INFORMES»; `Me.Caption` toma el de `lbTitulo` | CAP-10 | `:26`; `Form_Form0BDInformes.form.txt:639-640, 667-668` |
| `CuadroSuperior`, `Cuadro2`, `OLEIndependiente12` | 2 rectángulos y 1 imagen | — | Decoración | sin asignar | `Form_Form0BDInformes.form.txt:614, 778, 691` |
| Formulario | — | `Detalle_MouseMove` | `DesactivaLabel Me`: restaura el color de las etiquetas `lbl*` | CAP-10 | `:12-15`; `FG.bas:157-169` |
| Formulario | — | `Form_Open` | `AjustarTamaño Me`, repinta y fija `Caption` | CAP-10 | `:19-27`; `FG.bas:261-285`; `Form_Form0BDInformes.form.txt:40` |

El `MouseMove` de cada etiqueta cambia `ForeColor` a la constante literal `26367` si es distinta (`:40, :55, :67, :84, :94, :105`); `DesactivaLabel` lo devuelve a `lngColorEtiquetaMenuSinPulsar` al pasar por el detalle (`FG.bas:162-165`).

## Flujos

- **Normal:** el usuario pulsa una etiqueta → se abre el formulario destino → este formulario se cierra (p. ej. `:70-74`).
- **Alternativo:** `cmdSalir` cierra sin abrir nada (`:6-9`). La etiqueta `lblInformeParaRAC_General` no responde.
- **Errores:** todos los handlers usan `On Error Resume Next` (`:7, :21, :33, :46, :71, :78, :88, :99`): si `OpenForm` falla, el error se ignora y `Close` se ejecuta igualmente.

## Reglas, validaciones y cálculos

No hay validaciones, cálculos ni control de permisos en el formulario. El acceso a la aplicación se decide antes, en `EVE` (`VG.bas:343-347`: solo administradores y técnicos).

## Lecturas y escrituras

Ninguna: el formulario no abre recordsets ni ejecuta SQL (`:1-107`).

## Coexistencia

Sin SQL ni tipos de datos que adaptar. A preservar: la navegación a los seis destinos y el cierre del submenú tras elegir. Los destinos concentran el trabajo de PostgreSQL vinculado.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `lblInformeParaRAC_General` («- Informes de Partes de Trabajo») no tiene evento: la opción es visible y no responde (`Form_Form0BDInformes.form.txt:584-585`).
- `Close` tras `OpenForm` con `On Error Resume Next` (`:34-35`, `:47-48`, `:72-73`, `:79-80`, `:89-90`, `:100-101`): si el destino no se abre, el usuario pierde el submenú sin aviso.
- `PieDelFormulario_MouseMove` (`:59-62`) es un handler para un control que no figura entre los `Begin` con nombre del `.form.txt` (solo existe la sección `Detalle`, `Form_Form0BDInformes.form.txt:495`): código huérfano.
- Rótulo «Cumplimento» en lugar de «Cumplimiento» (`Form_Form0BDInformes.form.txt:932`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: sustituir el submenú por una agrupación en el menú principal y resolver qué hace la opción «Informes de Partes de Trabajo».

## Criterios de aceptación

Propuestos, no aprobados:

1. Existen seis entradas activas hacia Gestión de Informes de Avería, Consultar Horas Técnico, Para RAC, Por Factura, Desde Lista y Cumplimiento SLA (`:32-102`).
2. La entrada «Informes de Partes de Trabajo» queda definida (funcional o retirada) y no se muestra inerte (`Form_Form0BDInformes.form.txt:584`).
3. Salir devuelve al menú principal sin guardar nada (`:6-9`).
4. Si el destino no se puede abrir, se informa al usuario y se mantiene el submenú (hoy no ocurre, `:34-35`).

## Pendientes

- Existencia y contenido de los formularios destino, que se fichan por separado.
- Contenido de `DesactivaLabel` y `AjustarTamaño` más allá de lo citado (`FG.bas:157-169, 261-285`).
