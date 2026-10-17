type: patch

El stub HTTP de los tests de runner-controls (`_github_stub.py`) ya no copia el `Host` de la petición a la cabecera `Link`: la URL base sale de la dirección del servidor, lo que cierra el aviso de CodeQL «HTTP Response Splitting». Un consumer que vendoriza el stub debe re-vendorizarlo.
