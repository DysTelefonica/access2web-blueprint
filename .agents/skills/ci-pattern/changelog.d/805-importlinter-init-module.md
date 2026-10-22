type: patch

HR-68 (#805): la configuración de import-linter derivada del bloque `architecture` nombraba un `__init__.py` como el módulo `pkg.__init__`, que import-linter rechaza («Module does not exist»). Ahora un `__init__.py` se nombra como su paquete (`pkg`). Los repos que no clasifican `__init__.py` no cambian.
