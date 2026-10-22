type: minor

HR-68 (#714): `docs-architecture` impone también las reglas de imports en TypeScript con dependency-cruiser (`stack: typescript`): deriva la configuración del bloque `architecture`, falla si diverge, si el linter no está instalado o no analiza ningún módulo, y nombra `fichero:línea` y la regla de cada import prohibido o local roto.
