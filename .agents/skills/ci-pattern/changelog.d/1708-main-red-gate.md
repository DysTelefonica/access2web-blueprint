type: minor

La parada de línea del merge gobernado (`stop_the_line.py`, #1627) gana una segunda señal, independiente de la nocturna: mientras haya una issue abierta con la etiqueta `prerelease:main-red`, un PR pasa solo si enlaza una de esas issues del mismo repositorio; si no, sale con 1 y da el desbloqueo («arregle main en el PR que enlaza #N; el siguiente run verde de main la cierra»). Se activa por defecto en la llamada que ya hace el consumer, sin flag nuevo, y falla cerrado (exit 2) ante una lectura fallida. Cuando no hay issue abierta, la señal pasa. Refs #1708.
