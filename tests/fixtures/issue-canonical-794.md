### Problema y contexto

El lock de dependencias del proyecto (`app/uv.lock`) está atrás de su propia resolución en dos paquetes transitivos que ya tienen versión con arreglo publicado. Hoy no se nota porque el venv de auditoría de `pip-audit` instala una **resolución fresca** (pip elige la última permitida), así que el job sale verde; pero en cuanto el venv de auditoría se instala desde el lock —que es lo que los gates de la fase 4 usan para instalar la toolchain— el job sale rojo con avisos reales. El riesgo es que la auditoría esté midiendo un conjunto de paquetes distinto del que el CI instala.

### Evidencia verificable

Salida literal del paso `pip-audit --strict` con el venv instalado desde `app/requirements-dev.lock` (corrida `37663233316`, job `112935794618`, PR #793):

```text
dependency scan: auditing 71 installed packages
Found 9 known vulnerabilities in 3 packages
Name      Version ID              Fix Versions
--------- ------- --------------- ------------
starlette 0.50.0  PYSEC-2026-161  1.0.1
starlette 0.50.0  PYSEC-2026-249  1.3.1
starlette 0.50.0  PYSEC-2026-248  1.3.0
starlette 0.50.0  PYSEC-2026-2281 1.1.0
starlette 0.50.0  PYSEC-2026-2280 1.1.0
urllib3   2.7.0   PYSEC-2026-4177 2.8.0
urllib3   2.7.0   PYSEC-2026-4176 2.8.0
urllib3   2.7.0   PYSEC-2026-4175 2.8.0
```

Los dos paquetes del alcance, con su versión fijada hoy y su versión con arreglo:

| Paquete | Versión en `app/uv.lock` | Arreglo | Avisos |
|---|---|---|---|
| `starlette` | `0.50.0` (transitiva de `fastapi`) | `1.0.1` / `1.1.0` / `1.3.0` / `1.3.1` | `PYSEC-2026-161`, `PYSEC-2026-2280`, `PYSEC-2026-2281`, `PYSEC-2026-248`, `PYSEC-2026-249` |
| `urllib3` | `2.7.0` (transitiva) | `2.8.0` | `PYSEC-2026-4175`, `PYSEC-2026-4176`, `PYSEC-2026-4177` |

### Alcance y no objetivos

En alcance: actualizar el lock para que `starlette` y `urllib3` queden en una versión sin avisos, y decidir si el venv de auditoría de `security.yml` se instala desde el mismo lock que usan los gates (para que audite exactamente lo que el CI instala).

No objetivos: el mismo escaneo encontró un aviso en `pytest 8.4.2` (`PYSEC-2026-1845`, arreglo `9.0.3`), pero ese es un pin directo y deliberado en `app/pyproject.toml`: subirlo de mayor es una decisión aparte. Tampoco se toca el gate `pip-audit --strict` ni se añaden exenciones.

### Criterios de aceptación

- [ ] `app/uv.lock` deja `starlette` en una versión con arreglo (o documenta por qué no se puede dentro de la restricción de `fastapi`) y `urllib3` en `>= 2.8.0`.
- [ ] `pip-audit --strict` sobre un venv instalado desde `app/requirements-dev.lock` sale 0.
- [ ] La suite sigue en verde con las versiones nuevas (`make verify` en local y el job `quality` en CI).
- [ ] Queda escrito si el venv de auditoría pasa a instalarse desde el lock (un solo conjunto auditado e instalado) o si se mantiene la resolución fresca con su motivo.

### Plan de validación

```bash
uv lock --upgrade-package starlette --upgrade-package urllib3 --project app
uv export --project app --frozen --no-emit-project --extra dev --format requirements-txt
pip-audit --strict            # sobre un venv instalado desde el lock
make verify
```

### Dependencias y riesgos

- Los arreglos de `starlette` son **saltos de versión mayor**; `fastapi` fija un rango sobre `starlette`, así que la actualización puede exigir subir `fastapi` también. Si el rango lo impide, la salida es declarar el bloqueo y su motivo, no una exención en el escáner.
- Antecedente: #712 (cerrada) actualizó FastAPI para resolver avisos de Starlette; este es un caso nuevo sobre avisos nuevos.
- El canal de actualizaciones del repositorio es Dependabot (#555); si preferís que lo haga ese canal, esta issue queda como su justificación y su criterio de cierre.

