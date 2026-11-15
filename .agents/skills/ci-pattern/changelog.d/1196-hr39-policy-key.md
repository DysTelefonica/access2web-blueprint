type: patch

El gate de required jobs rechazaba la clave `state_dependent_gates` por desconocida (fail-closed), pero es la misma clave que el gate de política de workflows (HR-39) lee del mismo fichero de política: un consumer que declarase HR-39 se ponía el job `required` en rojo sin poder evitarlo. La clave pasa a admitirse entre las opcionales, con test y con la fila en el README del asset. Declarar HR-39 no es una relajación (no quita jobs ni añade skips), así que no necesita campo `relaxation` (#1196, #1197).
