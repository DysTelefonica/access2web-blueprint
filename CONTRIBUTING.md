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

## Convención multi-app

Este repositorio gobierna la migración a web de ocho aplicaciones legadas. Cada issue y cada commit declara a qué aplicación pertenece, o se marca como transversal. La convención permite que varias IAs trabajen en paralelo sin pisarse y habilita el filtrado por aplicación.

### Etiquetas por aplicación

Aplique la etiqueta `app/<slug>` a los issues que pertenezcan a una sola aplicación. El slug coincide con el directorio canónico `docs/03-aplicaciones/<slug>/`.

| Etiqueta | Aplicación |
|---|---|
| `app/condor` | Condor |
| `app/hps` | HPS |
| `app/hps-solicitudes` | HPS Solicitudes |
| `app/brass` | Brass |
| `app/gestion-riesgos` | Gestión de Riesgos |
| `app/no-conformidades` | No Conformidades |
| `app/lanzadera` | Lanzadera |
| `app/expedientes` | Expedientes |

### Etiqueta transversal

Aplique `cross-cutting` a los issues que afecten a varias aplicaciones simultáneamente. Ejemplos: el patrón de secret manager (D93), la migración de booleanos `Text(2)` a `BOOLEAN` (D102), la regla de gitignore para `.accdb` (D92). El label `cross-cutting` coexiste con cualquier `app/<slug>` cuando una sola app origina la iniciativa.

### Prefijo en el título del issue

Anteponga al título del issue el código de la aplicación entre corchetes, en mayúsculas. Use `XCUT` para transversal.

```text
[EXP] Definir catálogos web
[LANZ] Scaffolding del MVP
[XCUT] Estandarizar booleanos a BOOLEAN
```

### Scope en commits

Cierre el placeholder `(app)` de los conventional commits con el slug de la aplicación, o con `platform` para transversal.

```text
docs(expedientes): documentar matriz de migración
feat(lanzadera): portar formulario de login
fix(platform): migrar Text(2) a BOOLEAN
```

### Codificación de prefijos

| Código | Aplicación |
|---|---|
| `[COND]` | Condor |
| `[HPS]` | HPS |
| `[HPSS]` | HPS Solicitudes |
| `[BRASS]` | Brass |
| `[GR]` | Gestión de Riesgos |
| `[NC]` | No Conformidades |
| `[LANZ]` | Lanzadera |
| `[EXP]` | Expedientes |
| `[XCUT]` | Transversal |

## Tamaño de los PRs

El presupuesto es de **400 líneas** (`additions + deletions`), comprobado por
`scripts/check_pr_size.py`. No es una cifra arbitraria: es el límite a partir
del cual una revisión deja de ser atenta y pasa a ser un vistazo.

Cuando el cambio no cabe, el orden de preferencia es este:

1. **Partir por unidad de trabajo.** Si el cambio contiene dos ideas
   separables, son dos PRs. Es la opción por defecto y casi siempre la correcta.
2. **PRs encadenados.** Cuando las partes dependen unas de otras y no pueden
   entrar por separado, cada rama parte de la anterior en lugar de `main`, y
   cada PR apunta a la rama que le precede. Se integran en orden. Así cada
   revisión ve una porción entendible en lugar del muro entero.
3. **`size:exception`.** Último recurso, y sólo cuando el diff grande es
   inevitable: código generado, dependencias, migraciones. Requiere añadir al
   cuerpo del PR una línea `size-exception-reason: <motivo>`.

Los commits por unidad de trabajo son los ladrillos; los PRs encadenados son los
tramos de muro. Un cambio grande se trocea al planificarlo, no al final, cuando
el gate ya ha rebotado.

Usar `size:exception` para ahorrarse un corte deja el gate sin sentido. Si se
recurre a ella con frecuencia, el problema no es el presupuesto: es que el
trabajo se está planificando en piezas demasiado grandes.

## Reglas específicas de este repositorio

- Mantenga los PRs pequeños y centrados en una única idea.
- Integre con `--squash`. **No borre la rama remota.** Limpie el worktree local
  con `git worktree remove <ruta>`.
- No añada `Co-Authored-By` ni atribución de IA a los commits.
- Use el prefijo `docs(app):` para las épicas.
- Mantenga los walkthrough JSONs en commits separados.

### Por qué se conserva la rama remota

La integración es `--squash`, de modo que `main` recibe un único commit por PR.
Los commits por unidad de trabajo —los que separan una entrega en pasos
revisables— existen únicamente en la rama.

Borrar la rama los deja sin referencia y el recolector de basura acaba
eliminándolos. Se perdería justo la granularidad que el trabajo por unidades
pretende crear: `main` conserva el qué, y la rama, el cómo se llegó.

El coste es una lista de ramas larga. La convención de nombre
`<tipo>/<nº issue>-<slug>` la mantiene navegable.

[← Back to README](README.md) · [Next: CHANGELOG.md →](CHANGELOG.md)
