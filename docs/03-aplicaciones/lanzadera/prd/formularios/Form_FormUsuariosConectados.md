# Form_FormUsuariosConectados

> **Commit fijado:** `1474e8e8c2a8c352599ffa8b846223c6eb6e0f17` · **Fuente:** `src/forms/Form_FormUsuariosConectados.cls` (27 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Ventana informativa que muestra un texto recibido por `OpenArgs`, presumiblemente la lista de usuarios conectados a una aplicación. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`) y no accede a datos. Las citas son `Form_FormUsuariosConectados.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| No localizado | Ningún `DoCmd.OpenForm` ni `SourceObject` de `src` nombra `FormUsuariosConectados`; el único `OpenForm` con nombre en variable abre el menú principal tras el login | `Variables Globales.bas:724-736` |

No abre otros formularios ni lanza procesos.

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `OnOpen` → `Form_Open` | Copia `OpenArgs` (o vacío) en `UsuariosConectados` | `:7-25`, `.form.txt:34` |
| `UsuariosConectados` | texto bloqueado, `Tag = "DATO"` | — | Muestra el texto recibido; no editable | `:13`, `.form.txt:577-590` |

## Flujos

- **Normal:** al abrir, `Me.UsuariosConectados = Nz(Me.OpenArgs(), "")` (`:13`).
- **Errores:** `CorreoAlAdministrador` salvo error 1000 y `MsgBox` crítico (`:16-23`).

## Reglas, validaciones y cálculos

Ninguna. El texto llega ya formateado por quien abre el formulario.

## Lecturas y escrituras

Ninguna.

## Coexistencia

Sin acceso a datos: no requiere adaptación (ver `../coexistencia.md`).

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** muestra el texto recibido.

**Posibles defectos (sin corregir, a verificar):**

- Formulario posiblemente huérfano: no se ha encontrado quien lo abra en `src`. La misma información se muestra en `Form_FormAplicacionesGestion` con `CadenaCorreoUsuariosConectados` (`Form_FormAplicacionesGestion.cls:111`; `classes/Aplicacion.cls:1662`).
- El texto de ayuda del control dice «Nombre Completo de Usuario (Obligatorio)» (`.form.txt:590`), heredado de otro formulario; no corresponde a su contenido.

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: no migrarlo como pantalla propia e integrar la lista en el detalle de la aplicación.

## Criterios de aceptación

Propuestos, no aprobados:

1. Si se conserva, abrir la vista con un texto lo muestra sin permitir editarlo; sin texto, la muestra vacía.

## Pendientes

- Confirmar en el binario (macros, menús o cinta) si algo abre `FormUsuariosConectados`.
- Qué contenido se le pasaría por `OpenArgs` (candidatas: `CadenaNombreUsuariosConectados`, `classes/Aplicacion.cls:1625`, y `CadenaCorreoUsuariosConectados`, `classes/Aplicacion.cls:1662`).
