type: patch

`adoption generate-doc` y `manifest seal` escriben con LF en todas las plataformas (#1408). En Windows el modo texto escribía CRLF y el tramo 2 de una ronda contaba el documento operativo entero como reescrito.
