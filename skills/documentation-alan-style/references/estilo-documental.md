# Estilo documental — detalle

Material de consulta de la skill `documentation-alan-style`. Las reglas
obligatorias están en su `SKILL.md`; aquí está el detalle que no cabe allí.

## Tono y voz

Castellano peninsular formal. La segunda persona es **usted** en todas sus
formas, en singular y en plural.

| Evitar | Usar |
|---|---|
| `vos`, `tú` | `usted` |
| `ejecutá`, `corré`, `mirá`, `fijate` | `ejecute`, `corra`, `mire`, `fíjese` |
| `che`, `dale`, `listo`, `bárbaro` | `de acuerdo`, `correcto`, `terminado` |
| `vamos a`, `hicimos` | `el sistema`, `el equipo` (3.ª persona) |
| `genial`, `guay`, `mola` | `correcto`, `válido`, `adecuado` |

### Anglicismos que no se traducen

`FTS5`, `scope`, `topic`, `upsert`, `soft-delete`, `hard-delete`, `drift`,
`blast radius`, `kill switch`, `ratchet`, `squash-merge`, `tenant`, `WAL`,
`PITR`, `kebab-case`, `spec`, `protocol`, `commit`, `merge`, `frontend`,
`backend`, `template`, `placeholder`, `walkthrough`, `digest`, `preflight`.

### Tipografía

- Citas en prosa con comillas angulares: «...». No use comillas inglesas.
- Dentro de código, comillas rectas o acentos graves.
- Mayúscula tras `?` sólo si la frase anterior terminó.

## Estructura raíz del repositorio

| Archivo | Propósito | Audiencia | ¿Obligatorio? |
|---|---|---|---|
| `README.md` | Visión del producto, legible en cinco minutos | Externos | Sí |
| `AGENTS.md` | Alcance, reglas duras e índice de skills | Agentes | Sí |
| `DOCS.md` | Índice navegable y referencia técnica | Técnicos | Si es API-product |
| `CODEBASE-GUIDE.md` | Propiedad, flujos y guardarraíles | Mantenedores | Recomendado |
| `CONTRIBUTING.md` | Flujo de contribución y presupuesto de PR | Contribuidores | Sí |
| `CHANGELOG.md` | Cambios por versión, con «Unreleased» | Usuarios | Sí |
| `SECURITY.md` | Proceso de divulgación de vulnerabilidades | Externos | Si hay datos sensibles |

## Taxonomía de `docs/`

Un documento va suelto en `docs/` cuando trata un tema único y estable. Va a una
subcarpeta cuando pertenece a una familia que crece con el tiempo.

| Subcarpeta | Contiene | Nombre |
|---|---|---|
| `architecture/` | Decisiones y modelos que gobiernan el diseño | `kebab-case.md` |
| `codebase/` | Recorridos del código, interfaces, playbooks | `kebab-case.md` |
| `testing/` | Estrategia y guías de prueba | `kebab-case.md` |
| `audits/` | Auditorías con fecha, inmutables | `YYYY-MM-DD-asunto.md` |
| `releases/` | Cierres de versión | `vX.Y.Z-asunto.md` |
| `uat/` | Actas de aceptación | `uat-<audiencia>-YYYY-MM-DD.html` |
| `assets/` | Imágenes y material de marca | `kebab-case` |

Las carpetas con fecha en el nombre son históricas: **no se editan** una vez
publicadas. Si algo cambió, se escribe un documento nuevo.

Cuando el proyecto numera sus documentos por orden de lectura
(`00-alcance.md`, `01-inventario.md`), respete la numeración existente y continúe
la serie. No renumere lo ya publicado: rompe cualquier enlace externo.

## Nombres de fichero

- Raíz: `MAYUSCULAS.md`.
- Subcarpetas: `minusculas-con-guiones.md`.
- Diagramas y assets: `kebab-case`.
- Prohibido el guion bajo (`readme_v2.md`) y la numeración improvisada al margen
  de una serie ya establecida.

## Anti-patrones

- Índice de contenidos generado por la herramienta de build.
- Emojis decorativos en encabezados (`## 🚀 Inicio rápido`).
- Párrafos de más de doscientos caracteres sin punto y aparte.
- «Se recomienda», «sería bueno», «podría» en lugar del imperativo.
- Marketing: «potente», «sencillo», «de primer nivel», «revolucionario».
- Duplicar información entre `README.md`, `DOCS.md` y `CODEBASE-GUIDE.md` en
  lugar de referenciar.
- Sección de instalación al final. Va al principio.
- Más de seis enlaces externos en un mismo archivo.
- Mayúsculas sostenidas fuera de nombres propios y siglas (`HTTP`, `SQL`, `MCP`).
- Redactar a mano en `openspec/specs/`. Esa capa se genera al archivar.

## Lista de verificación

Antes de publicar, compruebe:

- [ ] Un único `H1`, y ningún `H4`.
- [ ] Ningún párrafo supera los doscientos caracteres.
- [ ] Cero coincidencias de `vos`, `ejecutá`, `corré`, `che`, `dale`, `listo` en
      la prosa. Este documento las contiene a propósito, entre backticks y dentro
      de sus tablas de prohibiciones: son el objeto de la regla, no una
      infracción. Al comprobar, excluya lo que va entre backticks.
- [ ] Ningún dato duplicado en otro documento sin referencia cruzada.
- [ ] Todo bloque de código declara su lenguaje.
- [ ] El documento está en la capa correcta y en la carpeta correcta.
