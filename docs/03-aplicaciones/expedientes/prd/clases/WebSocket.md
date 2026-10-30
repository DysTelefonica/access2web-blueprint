# WebSocket

> **Commit fijado:** `ec6a8c39cb3a6db3ab428ae1c5373c2b555cc95f` · **Fuente:** `src/classes/WebSocket.cls` (39 líneas) · **PRD:** [README](../README.md)

Puente con un navegador embebido (`WebBrowserControl`) para recibir notificaciones: carga `about:blank`, espera a que el control termine (`EsperarACargar`, `:34-37`) y añade un `<script>` con `https://www.gstatic.com/firebasejs/4.12.1/firebase.js` (`:24-30`), con los eventos del documento y de un `TextBox` de notificación (`:13-22`).

**Sin uso en RC1:** ningún fichero de `src` cita la clase. Depende de un servicio externo (Firebase, versión 4.12.1). La clave y la configuración de otro proyecto Firebase aparecen en claro en `Form_Formulario1`: ver la escalación al final de [issues-propuestas](../issues-propuestas.md).

## Comportamiento que la migración debe conservar

Ninguno verificable: no hay flujo que la use. Si el negocio espera notificaciones en tiempo real, es una capacidad nueva en web, no una migración.

## Pendientes

- Confirmar con negocio si alguna vez se usó y si se espera algo parecido ([P-EXP-12](../issues-propuestas.md)).
- Confirmar que no se invoca desde el binario, fuera de la lectura estática de `src`.
