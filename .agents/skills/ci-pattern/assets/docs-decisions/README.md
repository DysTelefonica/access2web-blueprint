# docs-decisions — registro de decisiones por consumer

Asset de `ci-pattern` (HR-67, DysTelefonica/team-skills#713). Un fichero por
decisión en `docs/decisions/` (`AAAA-MM-DD-<slug>.md`) y un índice `README.md`.
Estados y reglas, tomados del libro de cierre de gentle-ai (cerradas, sustituidas,
aplazadas con motivo, rechazadas):

| Campo | Regla |
|---|---|
| `estado` | `adoptada`, `rechazada`, `sustituida` o `aplazada` |
| `fecha` | ISO `AAAA-MM-DD` |
| `decide` | Rol o persona que decide |
| `motivo` | Por qué; no vacío |
| `sustituida-por` | Obligatorio si `sustituida`: slug de una decisión que existe, sin ciclos ni autorreferencia |
| `revision` | Obligatorio si `aplazada`: condición o fecha de revisión |

Cada línea tiene la forma `- campo: valor`. Toda cita `decisión: <slug>` de
`README.md`, `AGENTS.md` o `docs/` apunta a una decisión que existe, y el índice
lista exactamente las decisiones del directorio. Un directorio sin decisiones
falla (HR-3).

Activación: `docs_decisions` en el contrato de adopción (`dir`, `cite_roots`, `job`);
sin declaración, `adoption check --phase 5` avisa por `stderr` y sale 0. Si declara
`job`, la fase 5 exige que un workflow lo defina y que figure en `required_jobs`
de `.github/required-jobs-policy.json`. Plantillas en `assets/templates/docs/decisions/`.
