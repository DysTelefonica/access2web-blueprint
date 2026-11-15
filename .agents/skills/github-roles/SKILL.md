---
name: github-roles
description: "Trigger: «actúa como coordinador», «actúa como colaborador», «como coordinador», «como colaborador» — la identidad de GitHub con la que opera la IA: coordinador (dys-orquestador, fusiona) o colaborador (dys-agentes, abre PRs y no las fusiona). Sin rol dicho, la IA es colaborador."
license: Apache-2.0
metadata:
  author: ardelperal
  version: "1.1"
  last_verified: 2026-10-29
  scope: ['universal']
  auto_invoke: ['actúa como coordinador', 'actúa como colaborador', 'como coordinador', 'como colaborador']
  tiers: ['universal']
---

# Roles de GitHub para las IAs: coordinador y colaborador

Dos roles, dos identidades, una sola regla: **la IA que abre una PR no la fusiona**.

| Rol | Identidad | Puede | No puede |
| --- | --- | --- | --- |
| **colaborador** | App `dys-agentes` (login `dys-agentes[bot]`) | crear issues, empujar ramas, abrir PRs, escribir código, comentar | **fusionar** |
| **coordinador** | App `dys-orquestador` (login `dys-orquestador[bot]`) | todo lo del colaborador **y** fusionar | usar la cuenta personal del dueño |

El login es lo que aparece como autor de la PR y como fusionador en la API; es lo que el control post-merge compara.

**Sin rol dicho, la IA es colaborador.** La capacidad de fusionar se otorga, no se supone: el default es la identidad que no fusiona.

## Activation Contract

Esta skill se activa cuando la petición nombra un rol: «actúa como coordinador», «actúa como colaborador», «como coordinador» o «como colaborador». También se activa cuando la conversación ya está en curso en uno de los dos roles, para no perder la identidad a mitad de una tarea.

**No se activa** —y la IA se queda como colaborador— cuando:

- la petición no dice ningún rol;
- dice un rol que no es ninguno de los dos;
- pide «delega», «pide a otro agente que lo haga» o cualquier fórmula equivalente.

El rol se declara **una vez por tarea**. Cambiar de rol a mitad de una tarea exige decirlo explícitamente.

## Hard Rules

- **La cuenta personal del dueño no se usa nunca.** Ni para fusionar, ni para aprobar, ni para abrir una issue. Las dos identidades son las Apps: `dys-agentes` y `dys-orquestador`. Razón: mientras la IA no toque el `gh` del dueño, el dueño conserva la capacidad de fusionar y de aprobar, que es donde tiene que estar la decisión humana.
- **El colaborador no fusiona.** Abre la PR y se detiene ahí. El merge lo firma el coordinador o el dueño.
- **Al delegar, se delega siempre con la identidad de colaborador.** Un agente o subagente delegado recibe el PATH, el token y el credential helper de `dys-agentes`; **nunca** los de `dys-orquestador`. Delegar no transfiere la capacidad de fusionar: el coordinador conserva esa capacidad para sí y delega el trabajo, no el permiso.
- **El coordinador puede hacer el trabajo él mismo** o delegarlo en otros agentes; en los dos casos, el merge lo firma él.
- **Si el repo tiene merge gobernado, el coordinador fusiona por él, nunca con `gh pr merge`.** El merge gobernado es el workflow `runner-merge.yml` o `governed-merge.yml` del repo: comprueba el contrato del host (checks requeridos, etiqueta `type:*`, presupuesto de tamaño, frescura con `merge_freshness`) antes de fusionar. `gh pr merge` se salta ese contrato. Si el workflow bloquea la PR, el coordinador corrige lo que nombra el hallazgo o lo informa; no recurre a `gh pr merge` para esquivarlo. Solo en un repo sin ese workflow se fusiona con `gh pr merge`.
- **El rol no se deduce del alcance de la tarea.** Que la tarea sea grande no convierte a la IA en coordinador; el rol lo dice la persona.

## Decision Gates

1. ¿La petición nombra un rol? → úsalo. Si no lo nombra → **colaborador**.
2. ¿La tarea incluye fusionar? → hace falta el rol de coordinador. Si la IA está en rol de colaborador, abre la PR, avisa y para: no pide el `gh` del dueño para poder fusionar.
3. ¿El repo tiene `runner-merge.yml` o `governed-merge.yml`? → el coordinador fusiona lanzando ese workflow (ver Execution Steps). Si no tiene ninguno → `gh pr merge`.
4. ¿La tarea se delega? → el delegado hereda **colaborador**, siempre. Si el delegado necesita fusionar, vuelve al coordinador.
5. ¿Hay duda sobre qué identidad se está usando? → comprobarlo (`gh api user` no sirve para Apps; mira el autor de la última PR o el `gh auth status` del wrapper) antes de actuar.

## Execution Steps

Cambiar de identidad es **cambiar el `gh` que hay en el PATH y el credential helper de git** para que las dos cosas (la API y el push) usen la misma App:

```bash
# Colaborador (por defecto)
W="$HOME/bin/app-gh/agentes"
export PATH="$W:$PATH"
export GH_TOKEN="$("$W/gh" auth token)"
export GIT_CONFIG_COUNT=4
export GIT_CONFIG_KEY_1=credential.https://github.com.helper
export GIT_CONFIG_VALUE_1="!$W/gh auth git-credential"

# Coordinador (solo con el rolCoordinator concedido)
W="$HOME/bin/app-gh/orquestador"
```

En la caja del operador, `~/bin/merge-tools/agent-as-app.sh` ya deja el PATH y el credential helper como **colaborador**; el rol de coordinador es el mismo guion con `~/bin/app-gh/orquestador`.

Cada wrapper minte su propio token de instalación de la App (con caché de 50 minutos) y deduce el owner del repo o del comando: `ardelperal` si el comando o el repo lo nombran, `DysTelefonica` si no.

Antes de empujar o de abrir una PR, comprobar con qué identidad se va a hacer:

```bash
gh api "repos/OWNER/REPO/pulls" --jq '.[0].user.login'   # quien firma
```

### Fusionar como coordinador

El coordinador busca el workflow de merge gobernado del repo y, si existe, fusiona lanzándolo. La App `dys-orquestador` tiene `actions: write` para poder hacerlo. El merge lo hace el propio workflow, así que la PR queda fusionada por `github-actions`.

```bash
WF=$(gh api "repos/OWNER/REPO/actions/workflows" \
  --jq '[.workflows[].path | select(test("/(runner|governed)-merge[.]yml$"))][0] // empty')
if [ -n "$WF" ]; then
  gh workflow run "${WF##*/}" --repo OWNER/REPO -f pr=N
  # esperar la ejecución y leer su veredicto
  RUN=$(gh run list --repo OWNER/REPO --workflow "${WF##*/}" --limit 1 --json databaseId --jq '.[0].databaseId')
  gh run watch "$RUN" --repo OWNER/REPO --exit-status \
    || gh run view "$RUN" --repo OWNER/REPO --log | grep BLOCKED
else
  gh pr merge N --repo OWNER/REPO --merge   # repo sin merge gobernado
fi
```

Un `BLOCKED` es el veredicto del contrato: nombra lo que falta (un check, la etiqueta, la rama atrasada respecto de su base). Se corrige y se vuelve a lanzar el workflow.

## Output Contract

- El rol con el que se está operando y la identidad efectiva de GitHub.
- Si la tarea quedó en manos del coordinador por una razón de merge: el motivo y lo que falta para el merge (nada de «fusiona tú con tu cuenta»).
- Si se delegó: qué agente, con qué identidad (**colaborador**) y qué queda pendiente para el coordinador.
- El SHA, la rama y el estado de la PR, como en cualquier entrega.

## Anti-patterns

| Síntoma | Por qué está mal | Qué hacer en su lugar |
| --- | --- | --- |
| La IA fusiona porque el repo «no tiene protección» | La protección no es la identidad: la App tiene `contents: write` y fusiona igual | El rol colaborador no fusiona aunque el host lo permita |
| La IA usa el `gh` del dueño «porque hace falta» | Quita al dueño la única capacidad que le queda | La identidad de la App; si hace falta el coordinador, se pide el rol |
| Se delega con el token del coordinador | El delegado hereda el permiso de fusionar | Delegar siempre con `dys-agentes` |
| Se deduce el rol por el tamaño de la tarea | El tamaño no otorga autoridad | El rol lo dice la persona |
| El cambio de rol a mitad de tarea, sin decirlo | La identidad cambia sin que nadie lo decida | Declarar el cambio en voz alta |
| El coordinador usa `gh pr merge` en un repo con `runner-merge.yml` | Se salta el contrato del host: un PR atrasado se fusiona aunque el contrato pida `merge_freshness: "required"` | Lanzar el workflow de merge gobernado |
| El workflow bloquea y el coordinador fusiona con `gh pr merge` | Convierte el veredicto del contrato en un consejo | Corregir lo que nombra el `BLOCKED` y relanzar el workflow |

## References

- `scripts/check-main-pr-only.sh` — el control post-merge que falla si un PR lo fusionó una identidad que no puede fusionar (`dys-agentes[bot]`), y falla cerrado si no se puede leer quién lo fusionó.
- `testing/suites/main-pr-only/test-main-pr-only.sh` — los casos de ese control, incluido el de la identidad prohibida.
- `~/bin/app-gh/agentes` y `~/bin/app-gh/orquestador` — los wrappers de identidad en la caja del operador.
- Team-skills#1095 — la issue que introduce estos roles.
