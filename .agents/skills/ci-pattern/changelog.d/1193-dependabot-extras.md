type: minor

La declaración `dependabot` del contrato de adopción acepta, por ecosystem, `directories` (lista), `target-branch`, `ignore` y `groups`, además del `directory` que ya admitía (#1182). Así un consumer puede declarar su configuración real sin que el generador la pierda: el caso de prueba es el grupo `github/codeql-action/*` de access2web-blueprint, que el render reproduce byte a byte desde la declaración. Cualquier otra clave es un hallazgo de HR-75 (fail-closed, HR-3) en vez de descartarse en silencio. El esquema del contrato admite la forma con objeto (#1193).
