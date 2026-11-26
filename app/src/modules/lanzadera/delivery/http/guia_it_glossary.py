"""Glossary of the admin IT handover guide (Citrix + Linux).

Every technical term the guide links to is defined here, so a reader
who does not know a word (for example, bitness) can look it up.
"""

from __future__ import annotations

from typing import NamedTuple


class Term(NamedTuple):
    key: str
    term: str
    definition: str


GLOSSARY: tuple[Term, ...] = (
    Term(
        "acl",
        "ACL",
        "Lista de permisos de una carpeta o archivo: quién puede leer, quién escribir y quién "
        "no tiene acceso.",
    ),
    Term(
        "accde",
        "ACCDB y ACCDE",
        "Formatos de archivo de Access. ACCDB es editable; ACCDE es la versión compilada que se "
        "entrega a los usuarios, sin código modificable. Un ACCDE solo abre en un Access de su "
        "misma bitness.",
    ),
    Term(
        "runtime",
        "Access Runtime",
        "Versión gratuita de Access que ejecuta aplicaciones pero no permite editarlas; evita "
        "instalar Office completo en el servidor.",
    ),
    Term(
        "ad",
        "Active Directory (AD)",
        "Directorio corporativo de usuarios, grupos y equipos. Da el inicio de sesión y los "
        "permisos de red.",
    ),
    Term(
        "anexos",
        "ANEXOS",
        "Carpeta de cada aplicación con los documentos adjuntos guardados como archivos.",
    ),
    Term(
        "ini",
        "Archivo .ini de versión",
        "Archivo de texto junto al front publicado que indica su versión. La Lanzadera lo "
        "compara con la versión de la copia local del usuario para decidir si copia el front "
        "nuevo.",
    ),
    Term(
        "backend",
        "Backend",
        "Archivo Access que guarda las tablas con los datos; lo comparten todos los usuarios "
        "de la aplicación.",
    ),
    Term(
        "bitness",
        "Bitness (x86 / x64)",
        "Si un programa es de 32 bits (x86) o de 64 bits (x64). Access u Office, los ACCDE y el "
        "controlador ODBC tienen que ser de la misma: un Access de 32 bits solo ve controladores "
        "ODBC de 32 bits, y un ACCDE compilado en 32 bits no abre en un Access de 64 bits (ni al "
        "revés). Se consulta en Access, en Archivo › Cuenta › Acerca de Access.",
    ),
    Term(
        "ca",
        "CA y certificado",
        "Un certificado acredita la identidad de un servidor; la CA (autoridad de certificación) "
        "es quien lo firma. La validación falla si el nombre del servidor o la hora del equipo "
        "no cuadran.",
    ),
    Term(
        "cabeceras",
        "Cabeceras reenviadas",
        "Datos que el proxy añade a cada petición (IP original, protocolo). La aplicación solo "
        "debe creerlas si vienen del proxy real; si no, cualquiera podría falsearlas.",
    ),
    Term(
        "citrix",
        "Citrix",
        "Plataforma que ejecuta las aplicaciones en servidores centrales y muestra la pantalla "
        "al usuario. El puesto solo necesita el cliente Citrix.",
    ),
    Term(
        "upm",
        "Citrix UPM y FSLogix",
        "Dos formas de conservar los perfiles de usuario en Citrix. UPM copia las carpetas del "
        "perfil al entrar y al salir; FSLogix monta el perfil completo desde un disco virtual. "
        "Se elige una de las dos.",
    ),
    Term(
        "coexistencia",
        "Coexistencia",
        "Periodo en que conviven la Lanzadera Access y la web sobre los mismos datos. Cada "
        "aplicación está en un solo lado a la vez.",
    ),
    Term(
        "docker",
        "Contenedor, Docker y docker-compose",
        "Un contenedor empaqueta una aplicación con todo lo que necesita para ejecutarse. "
        "docker-compose arranca varios contenedores juntos; el del repositorio es solo para "
        "desarrollo, no para producción.",
    ),
    Term(
        "offline",
        "Copia fuera de línea",
        "Copia guardada en otra máquina, sin acceso desde el servidor original, para que "
        "sobreviva a un borrado o a un ataque.",
    ),
    Term(
        "d93",
        "D93",
        "Decisión de arquitectura: cada aplicación tiene en PostgreSQL un esquema <app>_legacy "
        "con las tablas compartidas con Access y otro <app> con lo que añade la web.",
    ),
    Term(
        "dns",
        "DNS",
        "Servicio que traduce nombres de servidor a direcciones IP. Los certificados se validan "
        "contra el nombre, no contra la IP.",
    ),
    Term(
        "esquema",
        "Esquema",
        "Agrupación de tablas dentro de una base de datos PostgreSQL, parecida a una carpeta.",
    ),
    Term(
        "front",
        "Front",
        "Parte de una aplicación Access que ve el usuario: formularios e informes. Cada usuario "
        "ejecuta su propia copia.",
    ),
    Term(
        "hipotesis",
        "H-3 y H-4",
        "Hipótesis pendientes de ensayo. H-3: la copia del front por usuario funciona en "
        "Citrix. H-4: cómo accede la web a los ANEXOS.",
    ),
    Term("https", "HTTPS (puerto 443)", "Acceso web cifrado con TLS."),
    Term(
        "digest",
        "Imagen fijada por digest",
        "La imagen es el paquete de un contenedor; el digest es su huella única (sha256:…). "
        "Fijarlo garantiza desplegar exactamente lo revisado.",
    ),
    Term(
        "json",
        "JSON de traspaso",
        "Archivo de texto con las marcas y notas de la lista de entrega, para llevarlas a otro "
        "equipo o adjuntarlas a la incidencia.",
    ),
    Term(
        "lanzadera",
        "Lanzadera",
        "Aplicación de entrada que muestra a cada usuario las aplicaciones a las que tiene "
        "permiso. La de Access abre aplicaciones Access; la web, solo aplicaciones web.",
    ),
    Term(
        "migraciones",
        "Migraciones",
        "Scripts versionados que crean o cambian las tablas de la base de datos, aplicados en "
        "orden.",
    ),
    Term(
        "odbc",
        "ODBC",
        "Estándar de Windows para que un programa, aquí Access, se conecte a una base de datos "
        "externa mediante un controlador.",
    ),
    Term(
        "openshift",
        "OpenShift",
        "Plataforma corporativa de contenedores a la que se prevé llevar la web al terminar la "
        "convivencia.",
    ),
    Term(
        "appdata",
        "%APPDATA%",
        "Carpeta del perfil del usuario (Roaming) que viaja con el perfil. La Lanzadera guarda "
        "aquí la copia del front y su versión.",
    ),
    Term(
        "temp",
        "%TEMP% y %LOCALAPPDATA%",
        "Carpetas locales de la máquina: normalmente no viajan con el perfil y pueden vaciarse "
        "al cerrar sesión. Por eso no sirven para la copia de la Lanzadera.",
    ),
    Term(
        "perfil",
        "Perfil persistente",
        "Carpeta con la configuración y los archivos de cada usuario que se guarda al cerrar "
        "sesión y se recupera al entrar, en cualquier VDA.",
    ),
    Term(
        "pitr",
        "PITR",
        "Recuperación de la base de datos a un instante concreto, combinando una copia con el "
        "registro de cambios de PostgreSQL.",
    ),
    Term(
        "postgresql",
        "PostgreSQL",
        "Base de datos que sustituirá a los backends Access; la usarán a la vez Access, "
        "mediante ODBC, y la web.",
    ),
    Term(
        "preproduccion",
        "Preproducción o ensayo",
        "Entorno con copias de producción para probar sin riesgo; nunca toca datos reales.",
    ),
    Term(
        "macros",
        "Protección de macros",
        "Seguridad de Office que bloquea código VBA de origen no confiable. Se autorizan rutas "
        "concretas (ubicaciones de confianza) en lugar de desactivarla.",
    ),
    Term(
        "proxy",
        "Proxy inverso",
        "Servidor web situado delante de la aplicación: recibe el HTTPS y reenvía las "
        "peticiones al puerto interno, de modo que la aplicación no queda expuesta.",
    ),
    Term(
        "psqlodbc",
        "psqlODBC",
        "Controlador ODBC oficial de PostgreSQL. Se usa la variante Unicode, que trata bien "
        "acentos y eñes, y de la misma bitness que Access.",
    ),
    Term(
        "puerto",
        "Puerto TCP",
        "Número que identifica un servicio en una máquina: 443 web, 5432 PostgreSQL, 445 "
        "carpetas compartidas. Abrirlo solo a orígenes concretos limita quién puede conectarse.",
    ),
    Term(
        "rpo",
        "RPO y RTO",
        "RPO: cuántos datos se pueden perder como máximo (por ejemplo, la última hora). RTO: "
        "cuánto puede tardar la recuperación.",
    ),
    Term(
        "smb",
        "SMB (puerto 445)",
        r"Protocolo de Windows para carpetas compartidas en red, como \\datoste\APLICACIONES_DYS.",
    ),
    Term(
        "health",
        "Sonda de salud (/health)",
        "Dirección que responde si la aplicación está en marcha. No prueba que los datos ni las "
        "conexiones funcionen.",
    ),
    Term("tls", "TLS", "Cifrado de las conexiones de red; impide leer o alterar el tráfico."),
    Term(
        "traslado",
        "Traslado único",
        "Paso de todos los datos de Access a PostgreSQL de una sola vez en producción. Es "
        "irreversible.",
    ),
    Term(
        "uvicorn",
        "Uvicorn (puerto 8000)",
        "Servidor que ejecuta la aplicación web Python dentro del contenedor. Ese puerto solo "
        "debe ser accesible desde el proxy.",
    ),
    Term(
        "variables",
        "Variables de entorno y secretos",
        "Configuración que recibe la aplicación al arrancar. Los secretos (contraseñas, claves) "
        "llegan desde un almacén aprobado y nunca se escriben en documentos ni registros.",
    ),
    Term(
        "verify-full",
        "verify-full",
        "Modo de conexión de PostgreSQL que cifra con TLS y además comprueba que el certificado "
        "lo firma una CA de confianza y corresponde al nombre del servidor. Evita suplantaciones.",
    ),
    Term(
        "vda",
        "VDA",
        "Máquina Windows donde Citrix ejecuta las sesiones. Hay varias, y un usuario puede "
        "entrar cada día en una distinta.",
    ),
    Term(
        "multisesion",
        "Windows multisesión (RDS)",
        "Windows Server configurado para que varios usuarios trabajen a la vez en la misma "
        "máquina.",
    ),
)

GLOSSARY_TERMS: dict[str, str] = {entry.key: entry.term for entry in GLOSSARY}

__all__ = ["GLOSSARY", "GLOSSARY_TERMS", "Term"]
