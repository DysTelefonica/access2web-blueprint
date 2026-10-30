# Form_frm0OtrosAdmin

> **Commit fijado:** `c88347ef61d965fff20f17fd37df3ec64bd7ae4a` · **Fuente:** `src/forms/Form_frm0OtrosAdmin.cls` (64 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Submenú de administración: búsqueda de expedientes y consulta de los registros de cambios y de errores. Es **no enlazado** y no lee ni escribe datos por sí mismo. Las citas son `Form_frm0OtrosAdmin.cls:línea` salvo indicación.

## Entrada y salida

- **Se abre desde:** Form_frm0Ppal, botón `cmdAdminOpciones`, visible solo con rol Administrador (`Form_frm0Ppal.cls:38-40, 159`).
- **Abre:** frmBuscarExpediente (`:25`), frmVerLogCambios (`:38`) y frmVerLogErrores (`:49`).
- **Permisos:** el formulario no comprueba el rol; depende de la visibilidad del botón en Form_frm0Ppal.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| Formulario | — | `OnLoad` | `Ajustar Me` y título «Apartado adminstracion» | `:9-20`, `.form.txt:222` |
| `cmdBuscarExpedientes` (etiqueta «Expedientes») | botón | `OnClick` | Abre frmBuscarExpediente sin argumentos | `:22-31`, `.form.txt:1220-1222, 1598-1599` |
| `cmdVerLogCambios` (etiqueta «Ver Logs cambios») | botón | `OnClick` | Abre frmVerLogCambios | `:35-44`, `.form.txt:968-970, 1566-1567` |
| `cmdVerLogErrores` (etiqueta «Ver Logs Errores») | botón | `OnClick` | Abre frmVerLogErrores | `:46-55`, `.form.txt:2043-2045, 2297-2298` |
| `cmdVolver` | botón | `OnClick` | Cierra el formulario | `:57-65`, `.form.txt:2335-2337` |
| `ComandoAyuda` | botón | — | Sin evento declarado ni handler | `.form.txt:853-854` |
| `lblBloqueo`, `lblVersion` | etiquetas | — | Sin uso en el código; la versión conserva el valor de diseño | `.form.txt:933-944, 1535-1536` |

## Flujos

- **Normal:** al cargar ajusta el tamaño y fija el título (`:12-13`; `modules/FUNCIONES UTILES.bas:325-339`); cada botón abre su formulario con `DoCmd.OpenForm` y este queda abierto detrás (`:25, 38, 49`); `cmdVolver` lo cierra (`:59`).
- **Alternativos:** ninguno; no hay ramas por rol ni por datos.
- **Errores:** todos los handlers delegan en `ManejadorErroresFormulario`, que muestra el mensaje y registra en `tbLogErrores` (`:17-19, 28-30, 41-43, 52-54, 62-64`; `modules/FUNCIONES UTILES.bas:588-676, 792-798`).

## Lecturas y escrituras

Ninguna directa. Los formularios que abre leen `tbLogCambios` y `tbLogErrores` (`Form_frmVerLogCambios.cls:17, 21`; `Form_frmVerLogErrores.cls:18, 23`); su detalle va en sus fichas.

## Coexistencia

Sin SQL propio. Los formularios de destino rellenan sus combos con `UNION` sobre `MSysObjects` (`Form_frmVerLogCambios.cls:17`; `Form_frmVerLogErrores.cls:18`), tabla de sistema de Access sin equivalente en PostgreSQL.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba.

**Posibles defectos (sin corregir, a verificar):**

- Sin control de rol propio: quien abra el formulario por nombre accede a los registros (`:9-20`).
- `ComandoAyuda` no hace nada al pulsarlo (`.form.txt:853-854`).
- `lblVersion` muestra siempre el texto de diseño (`.form.txt:1536`); los menús principales lo fijan desde el INI (`Form_frm0Ppal.cls:160`).
- Errata en el título: «adminstracion» (`:13`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: proteger la sección de administración por rol en el servidor.

## Criterios de aceptación

Propuestos, no aprobados:

1. Solo el rol Administrador accede a la sección.
2. Desde la sección se llega a la búsqueda de expedientes, al registro de cambios y al registro de errores.
3. Volver regresa al menú principal.

## Pendientes

- Intención de `ComandoAyuda` (etiqueta «Salir» en el diseño, `.form.txt:854`): ayuda o salida.
