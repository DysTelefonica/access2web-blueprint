# Form_Form0BDOpcionesOtras

> **Commit fijado:** `3f77d586814096c04f24bd2b6c229831cfbfcd34` · **Fuente:** `src/forms/Form_Form0BDOpcionesOtras.cls` (144 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Submenú «OTRAS OPCIONES» de Brass: trece etiquetas que abren los formularios de catálogos y estructura (nodos, BUI, subsistemas, ubicaciones, originadores, tipos), equipos, calibraciones, equipos de medida, mantenimientos planificados y anexos no alcanzables. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`) y sin lógica de datos. Las citas son `Form_Form0BDOpcionesOtras.cls:línea` (abreviadas `:N`) salvo indicación; `FG.bas` = `modules/Funciones Generales.bas`.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| `Form_Form0BDOpciones` | Etiqueta `lblOtrasOpciones`: `DoCmd.OpenForm "Form0BDOpcionesOtras"` | `Form_Form0BDOpciones.cls:160-163` |

Salidas: cada etiqueta abre un formulario de gestión y **no** cierra este (a diferencia del submenú de informes). `cmdSalir` lo cierra sin guardar (`:4-7`).

## Controles y eventos

Controles: 30 (39 `Begin` en el `.form.txt`: 1 `Form`, 7 plantillas de estilo por defecto, 1 `Section` y 30 controles con nombre).

| Control | Tipo | Evento y handler | Qué hace | Capacidad | Evidencia |
|---|---|---|---|---|---|
| `lblGestionNodos` («Nodos») | etiqueta | `Click`; `MouseMove` | Abre `FormNodoGestion` | CAP-12 | `:85-92`; `Form_Form0BDOpcionesOtras.form.txt:295-297` |
| `lblGestionBUI` («BUI'S») | etiqueta | `Click`; `MouseMove` | Abre `FormBUIGestion` | CAP-12 | `:69-76`; `Form_Form0BDOpcionesOtras.form.txt:328-330` |
| `lblGestionSubSistemas` («SubSistemas») | etiqueta | `Click`; `MouseMove` | Abre `FormSubSistemaGestion` | CAP-12 | `:101-108`; `Form_Form0BDOpcionesOtras.form.txt:362-364` |
| `lblGestionIDEquipos` («Equipos») | etiqueta | `Click`; `MouseMove` | Abre `FormEquipoGestion` | CAP-5 | `:77-84`; `Form_Form0BDOpcionesOtras.form.txt:396-398` |
| `lblGestionOriginadores` («Originadores») | etiqueta | `Click`; `MouseMove` | Abre `FormOriginadorGestion` | CAP-12 | `:93-100`; `Form_Form0BDOpcionesOtras.form.txt:430-432` |
| `lblGestionTipoEventos` («Tipo Eventos») | etiqueta | `Click`; `MouseMove` | Abre `FormEventoTipoGestion` | CAP-3 | `:118-125`; `Form_Form0BDOpcionesOtras.form.txt:464-466` |
| `lblRelacionNodoBUI` («Relación Nodo-BUI») | etiqueta | `Click`; `MouseMove` | Abre `FormNodoBUIGestion` | CAP-12 | `:135-142`; `Form_Form0BDOpcionesOtras.form.txt:498-500` |
| `lblGestionUbicaciones` («Ubicaciones») | etiqueta | `Click`; `MouseMove` | Abre `FormUbicacionGestion` | CAP-12 | `:127-134`; `Form_Form0BDOpcionesOtras.form.txt:532-534` |
| `lblGestionTipoActividad` («Tipo Actividad») | etiqueta | `Click`; `MouseMove` | Abre `FormActividadTipoGestion` | CAP-4 | `:110-117`; `Form_Form0BDOpcionesOtras.form.txt:565-567` |
| `lblCalibraciones` («Calibraciones») | etiqueta | `Click`; `MouseMove` | Abre `FormCalibracionEquipoGestion` | CAP-5 | `:43-50`; `Form_Form0BDOpcionesOtras.form.txt:599-601` |
| `lblMtoProg` («MNTOS. PLANIFICADOS») | etiqueta | `Click`; `MouseMove` | Abre `FormPlanificacionGestion` | CAP-6 | `:34-41`; `Form_Form0BDOpcionesOtras.form.txt:1623-1625` |
| `lblEquiposMedida` («Equipos Medida») | etiqueta | `Click`; `MouseMove` | Abre `FormEquiposMedidaGestion` | CAP-5 | `:24-32`; `Form_Form0BDOpcionesOtras.form.txt:1657-1659` |
| `lblAnexosNoAlcanzables` («Anexos No Alcanzables») | etiqueta | `Click` (sin `MouseMove`) | Abre `FormAnexosNoAlcanzables` | CAP-11 | `:19-22`; `Form_Form0BDOpcionesOtras.form.txt:1782-1784` |
| `cmdSalir` («Salir») | botón | `Click` → `cmdSalir_Click` | Cierra sin guardar | CAP-2 | `:4-7`; `Form_Form0BDOpcionesOtras.form.txt:629-631` |
| `lbTitulo` | etiqueta | — | Rótulo «OTRAS OPCIONES»; su `Caption` pasa a `Me.Caption` | CAP-2 | `:61`; `Form_Form0BDOpcionesOtras.form.txt:682-683` |
| `Cuadro1`, `CuadroSuperior` | 2 rectángulos | — | Decoración | sin asignar | `Form_Form0BDOpcionesOtras.form.txt:270, 657` |
| 13 imágenes (`OLEIndependiente0`, `3`, `5`…`10`, `12`, `13`, `19`, `22`, `Imagen14`) | imagen | — | Iconos | sin asignar | `Form_Form0BDOpcionesOtras.form.txt:705-1813` |
| Formulario | — | `Form_Open` | Hourglass, `AjustarTamaño`, `Caption`, repintado | CAP-2 | `:53-68` |
| Formulario | — | `Detalle_DblClick` | Muestra el ribbon | CAP-2 | `:13-17` |
| Formulario | — | `Detalle_MouseMove` | `DesactivaLabel Me` | CAP-2 | `:8-11`; `FG.bas:157-169` |

El `MouseMove` de las etiquetas pone `ForeColor = longColorEtiquetaMenuPulsado` (`:31, :40, :49, :75, …`); `DesactivaLabel` lo restaura (`FG.bas:162-165`).

## Flujos

- **Normal:** el usuario pulsa una etiqueta → `DoCmd.OpenForm` del destino; el submenú queda abierto (p. ej. `:85-88`).
- **Apertura:** `Form_Open` con `On Error Resume Next` (`:54`): reloj de arena, `AjustarTamaño Me`, título y repintado (`:55-67`).
- **Alternativo:** `cmdSalir` cierra el submenú (`:4-7`).
- **Errores:** todos los handlers ignoran errores, salvo `Detalle_DblClick`, que no tiene tratamiento (`:13-17`); un fallo de `OpenForm` pasa sin aviso.

## Reglas, validaciones y cálculos

No hay validaciones, cálculos ni control de permisos en el formulario; el acceso general se decide en `EVE` (`VG.bas:343-347`) y los permisos finos, si existen, estarán en los formularios destino.

## Lecturas y escrituras

Ninguna: el formulario no abre recordsets ni ejecuta SQL (`:1-144`).

## Coexistencia

Sin SQL ni tipos que adaptar. A preservar: la lista de destinos y su agrupación. Cualquier adaptación a PostgreSQL vinculado se resuelve en los formularios destino.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- `lblGestionIDEquipos` y la etiqueta «Equipos» del menú principal abren el mismo formulario `FormEquipoGestion` (`:77-80`; `Form_Form0BDOpciones.cls:136-139`): duplicidad de entradas.
- `lblAnexosNoAlcanzables` no tiene `MouseMove`, por lo que no se resalta como las demás (`:19-22`).
- `Detalle_DblClick` muestra el ribbon sin comprobar el rol ni `On Error` (`:13-17`).
- Valor fijo `26367` y colores definidos en `EVE`: si `EVE` no se ejecutó, `longColorEtiquetaMenuPulsado` valdría 0 (`VG.bas:357`; a verificar el orden de arranque).
- Rótulo «BUI'S» con apóstrofo (`Form_Form0BDOpcionesOtras.form.txt:329`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: integrar los catálogos en un único menú de administración con visibilidad por rol.

## Criterios de aceptación

Propuestos, no aprobados:

1. Existen trece entradas, cada una hacia su pantalla de gestión (`:19-142`).
2. Abrir una entrada no cierra el submenú (`:85-88`).
3. Salir cierra el submenú y no guarda nada (`:4-7`).
4. Si el destino falla, se informa al usuario (hoy se ignora, `:86`).

## Pendientes

- Existencia y comportamiento de los formularios destino, que se fichan por separado.
- Quién tiene derecho a ver este submenú: no se evalúa en este formulario.
- Contenido de `AjustarTamaño` más allá de lo citado (`FG.bas:261-285`).
