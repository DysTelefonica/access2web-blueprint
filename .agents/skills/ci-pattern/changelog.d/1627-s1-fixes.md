type: patch

Corrige la declaración del perfil pre-release para que acepte el fichero real de APAP_WEB (`{"enabled": true}`), con `schema_version` opcional, y el validador sin trazas ante ficheros no UTF-8 o UTF-16, con BOM aceptado y claves duplicadas como error. El partial web aclara que la regla dura rige ya y añade las reglas del dueño que faltaban. (#1627)
