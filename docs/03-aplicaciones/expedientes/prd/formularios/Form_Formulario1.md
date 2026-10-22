# Form_Formulario1

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/forms/Form_Formulario1.cls` (89 líneas) y `.form.txt` · **PRD:** [README](../README.md)

Prototipo de notificación en tiempo real con Firebase incrustado en un control de navegador. Es un formulario **no enlazado** (sin `RecordSource` en el `.form.txt`) y **sin punto de entrada**: ningún módulo, clase ni formulario de `src` lo nombra (búsqueda de `Formulario1` fuera de sus dos ficheros, sin resultados). No lee ni escribe en la base de datos. Las citas son `Form_Formulario1.cls:línea` salvo indicación.

## Entrada y salida

| Desde | Cómo | Evidencia |
|---|---|---|
| — | Sin `OpenForm` ni referencia en `src`; solo se abriría a mano desde el panel de navegación | búsqueda sin resultados |

No abre formularios ni informes. Escribe un fichero HTML en la carpeta temporal del usuario y lo carga en `WebBrowser0` (`:37-70`).

## Controles y eventos

| Control | Tipo | Evento y handler | Qué hace | Evidencia |
|---|---|---|---|---|
| (formulario) | — | `OnOpen` → `Form_Open` | Llama a `InitializeFirebase` | `:19-22`, `.form.txt:31` |
| `WebBrowser0` | control de navegador Edge (`Begin Edge`) | — | Carga la página HTML generada | `:70`, `.form.txt:226-241` |
| `txtInput` | texto | `OnChange` → `txtInput_Change` | En cada pulsación envía el texto (`.Text`) al nodo `notificaesto` ejecutando JavaScript en el navegador | `:78-89`, `.form.txt:250-251` |
| `txtRealtime` | texto | — | Recibe el valor del nodo vía `UpdateRealtime` | `:73-76`, `.form.txt:269` |
| `btnSend` | — (no existe en el `.form.txt`) | `btnSend_Click` | Igual que `txtInput_Change`, pero con `.value` | `:23-33` |

El `CommandButton` de `.form.txt:146` es la plantilla de controles por defecto del formulario, no un control.

## Flujos

### Apertura (`Form_Open`, `:19-22`)

1. Compone una página HTML que carga el SDK de Firebase desde un CDN público (`:41-43`) con la configuración del proyecto escrita en el código (`:45-53`).
2. Se suscribe a cambios del nodo `notificaEsto` y los reenvía a `window.external.UpdateRealtime` (`:56-58`).
3. Escribe la página en `%TEMP%\firebaseRealtime.html` (`:38, 63-67`) y la carga en `WebBrowser0` (`:70`).

### Envío (`txtInput_Change`, `:78-89`)

1. Toma el texto en edición (`:81`).
2. Lo concatena en `firebase.database().ref('notificaesto').set('<texto>')` (`:85`).
3. Lo ejecuta con `Document.parentWindow.execScript` (`:88`).

### Errores

Ningún procedimiento tiene `On Error`; cualquier fallo (escritura del fichero, navegador, `execScript`) llega sin tratar al usuario.

## Reglas, validaciones y cálculos

- No hay validaciones del texto enviado (`:81-85`).
- `ShowWindow` y las constantes `SW_*` se declaran pero no se usan (`:5-17`).

## Lecturas y escrituras

| Operación | Destino | Tipo | Evidencia |
|---|---|---|---|
| Fichero HTML | `%TEMP%\firebaseRealtime.html` | E/S de fichero | `:63-67` |
| Lectura y escritura remota | Firebase Realtime Database, nodos `notificaEsto` / `notificaesto` | JavaScript en el navegador | `:56, 85` |

Sin acceso a tablas ni consultas de Access.

## Coexistencia

No aplica: no toca la base de datos. Ver `../coexistencia.md`.

## Comportamiento actual · Posibles defectos · Propuestas web pendientes de decisión

**Comportamiento actual:** lo descrito arriba; formulario huérfano sin punto de entrada.

**Posibles defectos (sin corregir, a verificar):**

- Credenciales y configuración de un proyecto Firebase en claro en el código (`:46-52`); deberían tratarse como expuestas y rotarse.
- Inyección de JavaScript: el texto del usuario se concatena entre comillas simples sin escapar (`:30, 85`).
- Lee del nodo `notificaEsto` y escribe en `notificaesto` (`:56` frente a `:30, 85`); las claves de Firebase distinguen mayúsculas, así que el eco no llegaría.
- `WebBrowser0` es un control Edge (`.form.txt:226`), pero el código usa `Object.Document.parentWindow.execScript` y `window.external` (`:32, 58, 88`), propios del control de Internet Explorer.
- La página solo carga `firebase-app.js` (`:43`) y llama a `firebase.database()` (`:55`); en la versión 9 eso requiere además el módulo de base de datos.
- Handler `btnSend_Click` huérfano: no hay control `btnSend` en el `.form.txt` (`:23-33`).

**Propuestas web pendientes de decisión:** ninguna aprobada. Candidata: no migrar el formulario y retirarlo del legacy tras confirmar que nadie lo usa.

## Criterios de aceptación

Propuestos, no aprobados:

1. La versión web no incluye este formulario salvo decisión expresa del responsable de producto.
2. Ninguna credencial de servicios externos figura en el código fuente de la versión web.

## Pendientes

- Confirmar que no se abre desde macros, la cinta o el panel de navegación (las macros no están exportadas en `src`).
- Confirmar si el proyecto Firebase sigue activo y quién es su propietario.
