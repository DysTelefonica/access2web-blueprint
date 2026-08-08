# access2web-blueprint — Guía de contribución

## Bienvenida

Consulte esta guía antes de abrir un issue o un PR. Mantenga cada cambio centrado en una única idea revisable.

## Cómo contribuir

1. Abra un issue que describa el problema, el alcance y el resultado esperado.
2. Cree una branch específica a partir de `main`.
3. Aplique el cambio y documente las decisiones relevantes.
4. Verifique el contenido con los checks definidos por el repositorio.
5. Abra un PR pequeño con una descripción verificable.
6. Responda a la revisión y mantenga actualizada la branch.

## Conventional Commits

| Tipo | Ejemplo | Uso |
|---|---|---|
| `docs` | `docs(app): documentar reglas de autorización` | Cambios de documentación |
| `fix` | `fix(app): corregir referencia de formulario` | Correcciones de comportamiento o contenido |
| `feat` | `feat(app): añadir capacidad de migración` | Nueva capacidad documentada |
| `chore` | `chore(repo): actualizar configuración de herramientas` | Mantenimiento sin cambio funcional |
| `refactor` | `refactor(app): reorganizar la especificación` | Reorganización sin cambio de alcance |

## Label system

| Label | Color | Use for |
|---|---|---|
| `🐛 bug` | `#d73a4a` | Defectos verificables |
| `📚 documentation` | `#0075ca` | Cambios documentales |
| `✨ enhancement` | `#a2eeef` | Mejoras de alcance existente |
| `🔍 research` | `#7057ff` | Investigación y evidencia |
| `⚙️ chore` | `#e4e669` | Mantenimiento técnico |

## Reglas específicas de este repositorio

- Mantenga los PRs pequeños y centrados en una única idea.
- Use `--squash --delete-branch` al integrar un PR.
- No añada `Co-Authored-By` ni atribución de IA a los commits.
- Use el prefijo `docs(app):` para las épicas.
- Mantenga los walkthrough JSONs en commits separados.

[← Back to README](README.md) · [Next: CHANGELOG.md →](CHANGELOG.md)
