---
name: documentation-alan-style
description: "Trigger: redactar o revisar README, AGENTS, DOCS, CODEBASE-GUIDE, CONTRIBUTING, CHANGELOG, docs/, épicas o walkthroughs JSON. Patrón documental del proyecto en Castellano peninsular formal."
license: MIT
metadata:
  author: ardelperal
  version: "2.1"
---

## Activation Contract

Cargue esta skill antes de escribir o revisar:

- Un documento raíz: `README.md`, `AGENTS.md`, `DOCS.md`, `CODEBASE-GUIDE.md`,
  `CONTRIBUTING.md`, `CHANGELOG.md`.
- Cualquier documento de `docs/`, una `epic.md` o un `walkthrough-*.json`. Las
  dos últimas son extensiones opt-in; el contrato base no las requiere.
- Un PR que cambie comportamiento ya documentado en alguno de ellos.

No la cargue para decidir arquitectura, diseño visual, ni el contenido técnico de
un endpoint. Eso es el spec, no su documentación.

## Hard Rules

1. **Castellano peninsular formal, tratando de usted.** Sin voseo (`vos`,
   `ejecutá`, `corré`), sin tuteo coloquial, sin regionalismos (`che`, `dale`,
   `listo`, `bárbaro`).
2. **Una sola fuente de verdad.** Si dos documentos necesitan el mismo dato, uno
   lo tiene y el otro lo referencia.
3. **Un único `H1` por documento.** `H2` para secciones, `H3` para subsecciones.
   Nunca `H4` ni más profundo. No aplica a los `SKILL.md`, que empiezan por su
   frontmatter y siguen la estructura LLM-first sin `H1`.
4. **Párrafos por debajo de doscientos caracteres.** Una idea por frase, voz
   activa: «Cargue la skill», no «la skill debería ser cargada».
5. **Imperativo directo.** Nunca «se recomienda», «sería bueno» ni «podría».
6. **Sin emojis decorativos** en encabezados ni cuerpo. Única excepción: la tabla
   de labels de `CONTRIBUTING.md`.
7. **Sin marketing.** Nada de «potente», «sencillo», «de primer nivel».
8. **Sin índice generado** por la herramienta de build.
9. **Encabezados en castellano en sentence case**; en inglés, Title Case.

## Decision Gates

### Dónde vive cada cosa

La documentación tiene tres capas y confundirlas es el error más caro.

| Capa | Dónde | Quién la actualiza |
|---|---|---|
| Narrativa del producto | Documentos raíz y `docs/` | Una persona, cuando cambia el enfoque |
| Contrato vivo de capacidades | `openspec/specs/` | La fase `archive`, mecánicamente |
| Historia del porqué | `openspec/changes/archive/` | Nadie: es inmutable |

Nunca redacte a mano en `openspec/specs/`. Esa capa se genera al archivar un
change; escribirla directamente rompe su trazabilidad.

### Qué formato para qué contenido

| Contenido | Formato |
|---|---|
| Endpoints, variables de entorno, labels, hallazgos | Tabla |
| Flujo paso a paso donde el orden importa | Lista numerada |
| Comandos, peticiones y respuestas | Bloque con lenguaje (`bash`, `json`, `text`) |
| Diagrama o jerarquía | Bloque `text` con ASCII, máximo treinta líneas |
| Explicación de un concepto | Párrafo de dos o tres frases |

Detalle completo de tono, glosario, anglicismos, estructura raíz, taxonomía de
`docs/`, nombres de fichero y anti-patrones: `references/estilo-documental.md`.

## Execution Steps

1. Identifique el documento y su audiencia. Si no sabe para quién escribe, no
   escriba todavía.
2. Determine la capa: narrativa, contrato vivo o historia. Si es contrato vivo,
   pare: eso se genera archivando un change, no redactando.
3. Compruebe si el dato ya existe en otro documento. Si existe, referencie.
4. Elija el formato con la tabla anterior antes de escribir el primer párrafo.
5. Redacte aplicando las Hard Rules.
6. Verifique con la lista de `references/estilo-documental.md`.

## Output Contract

Devuelva:

- El documento escrito o revisado, con su ruta.
- Qué capa ocupa y por qué.
- Las referencias cruzadas añadidas para evitar duplicación.
- Las reglas incumplidas que ha corregido, si revisaba un documento existente.

## References

- `references/estilo-documental.md` — tono, glosario, estructura, nombres y
  anti-patrones.
- `../../CONTRIBUTING.md` — convenciones de commit, rama y PR.
- `../../AGENTS.md` — alcance del repositorio y skills disponibles.
- Caso verificado: `Gentleman-Programming/gentle-ai` (CLI Go + Bubbletea TUI;
  https://github.com/Gentleman-Programming/gentle-ai). El patrón se observa en
  `docs/CODEBASE-GUIDE.md` y las páginas radiales bajo `docs/codebase/`.
