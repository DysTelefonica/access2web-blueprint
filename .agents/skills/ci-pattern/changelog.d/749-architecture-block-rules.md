type: minor

El bloque `architecture` admite tres claves opcionales (solo Python) para sustituir un verificador de capas propio sin perder reglas: `independent` (módulos que no se importan entre sí, importaciones directas), `pure` (capas que no importan paquetes externos prohibidos, con recorrido AST y `fichero:línea`) y `tolerated` (importaciones heredadas enumeradas que solo pueden encoger; una entrada obsoleta es `STALE`). Sin las claves, la configuración derivada no cambia (#749).
