"""Static content of the admin IT handover guide (Citrix + Linux).

Facts come from docs/07-migracion/epica-coexistencia-postgresql.md
(decisions 10, 18-20, 28; H-3, H-4) and docs/architecture.md (D93).
The guide prepares an isolated rehearsal; it never authorises the
production data move.
"""

from __future__ import annotations

from typing import NamedTuple

from app.src.modules.lanzadera.delivery.http.guia_it_glossary import (
    GLOSSARY,
    GLOSSARY_TERMS,
)


class Scenario(NamedTuple):
    key: str
    label: str
    title: str
    summary: str
    tag: str
    terms: tuple[str, ...]


class Flow(NamedTuple):
    route: str
    transport: str
    condition: str
    terms: tuple[str, ...]


class Check(NamedTuple):
    key: str
    label: str


class Command(NamedTuple):
    key: str
    label: str
    text: str


SCENARIOS: tuple[Scenario, ...] = (
    Scenario(
        "actual",
        "Actual",
        "01 · Producción actual",
        "La Lanzadera Access abre únicamente aplicaciones Access. Backends por aplicación, "
        "ANEXOS y recursos (.accde, .ini, iconos) residen en "
        r"\\datoste\APLICACIONES_DYS\Aplicaciones PpD\. El ensayo no modifica esta ruta "
        "ni sus datos.",
        "Sin cambios",
        ("lanzadera", "backend", "anexos", "accde", "ini"),
    ),
    Scenario(
        "ensayo",
        "Ensayo aislado",
        "02 · Ensayo aislado",
        "Citrix Windows publica la Lanzadera Access y un navegador; una máquina Linux de la "
        "misma red aloja la web y PostgreSQL de ensayo. URL, recursos SMB, bases, "
        "almacenamiento y cuentas son de preproducción y distintos de los reales. Las copias "
        "llegan por el procedimiento aprobado y no se conectan a producción.",
        "Preparar hoy",
        ("citrix", "postgresql", "smb", "preproduccion"),
    ),
    Scenario(
        "futuro",
        "Coexistencia futura",
        "03 · Coexistencia propuesta, tras aprobación",
        "Access abre solo las aplicaciones legacy y el navegador publicado solo las web; "
        "ambas trabajan sobre los mismos datos en PostgreSQL, sin sincronizar copias. "
        "Secuencia: saneamiento del legacy, ensayo de todos los backends, traslado único y "
        "autorizado de todos los datos, y después cambio de interfaz por aplicación. Volver "
        "a habilitar un botón de Access no revierte el traslado de datos, que es irreversible. "
        "Al terminar la convivencia, la web se prevé en OpenShift.",
        "Propuesta",
        ("coexistencia", "postgresql", "traslado", "openshift"),
    ),
)

FLOWS: tuple[Flow, ...] = (
    Flow(
        "Cliente Citrix → pasarela corporativa",
        "Según IT",
        "Publicación y acceso corporativo autorizados.",
        ("citrix",),
    ),
    Flow(
        "Windows → recurso de ficheros",
        "SMB 445",
        "Solo rutas y cuentas aprobadas; ACL por aplicación.",
        ("smb", "acl"),
    ),
    Flow(
        "Windows ODBC → PostgreSQL",
        "TCP 5432",
        "Solo hosts Windows autorizados; TLS verify-full con CA y nombre DNS correctos.",
        ("odbc", "puerto", "tls", "verify-full", "ca", "dns"),
    ),
    Flow(
        "Navegador publicado → proxy web",
        "HTTPS 443",
        "Certificado válido y URL del entorno correcto.",
        ("proxy", "https"),
    ),
    Flow(
        "Proxy → aplicación",
        "Puerto interno (8000 en app/Dockerfile)",
        "No publicar 8000; el proxy es el único origen de cabeceras reenviadas confiables.",
        ("uvicorn", "cabeceras"),
    ),
    Flow(
        "Aplicación → PostgreSQL",
        "TCP 5432 interno",
        "Solo el servicio autorizado, con conexión protegida según política.",
        ("postgresql", "puerto"),
    ),
)

PROFILE_TESTS: tuple[str, ...] = (
    "Dos usuarios distintos no comparten carpeta ni versión.",
    "El mismo usuario en otro VDA reutiliza su copia.",
    "Versión igual: no se vuelve a copiar; versión nueva en el .ini: se actualiza.",
    "Cierre de sesión, reinicio del host y nueva entrada conservan la copia.",
    "Actualización publicada con la aplicación abierta: comportamiento registrado.",
    "Caída de red durante la copia: sin escrituras obsoletas ni sobrescritura parcial.",
)

CHECKS: tuple[Check, ...] = (
    Check("aislamiento", "Entornos, cuentas y almacenamiento separados y comprobados"),
    Check("red", "Red, DNS, certificados, hora y licencias comprobados"),
    Check("perfil", "Perfil persistente y pruebas multisesión documentadas"),
    Check("anexos", "Acceso y ACL de ANEXOS ensayados, o H-4 registrada como pendiente"),
    Check("recuperacion", "Copias, RPO/RTO y restauración acordados y ensayados"),
    Check("artefacto", "Entregables de desarrollo recibidos y revisados"),
    Check("ensayo", "Ensayo completo y resultados registrados, sin alterar producción"),
)

COMMANDS: tuple[Command, ...] = (
    Command("dns", "Copiar ejemplo DNS", "nslookup <nombre-dns-del-entorno>"),
    Command(
        "tcp",
        "Copiar ejemplo PowerShell",
        "Test-NetConnection <servidor-postgresql-del-entorno> -Port 5432",
    ),
)

SOURCES: tuple[str, ...] = (
    "docs/07-migracion/epica-coexistencia-postgresql.md: decisiones 10, 18-20 y 28; H-3 y H-4.",
    "docs/architecture.md: persistencia D93 y gestión de secretos.",
    "app/Dockerfile: puerto y proxies; docker-compose.yml: solo desarrollo local.",
    "docs/codebase/interfaces.md: /health.",
)

GUIDE_CONTEXT: dict[str, object] = {
    "glossary": GLOSSARY,
    "glossary_terms": GLOSSARY_TERMS,
    "scenarios": SCENARIOS,
    "flows": FLOWS,
    "profile_tests": PROFILE_TESTS,
    "checks": CHECKS,
    "commands": COMMANDS,
    "sources": SOURCES,
}

__all__ = ["GUIDE_CONTEXT"]
