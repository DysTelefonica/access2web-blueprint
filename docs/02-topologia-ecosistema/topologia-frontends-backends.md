# 02 · Topología del ecosistema — Topología de frontends y backends

## Propósito

Inventario físico de los binarios frontend y backend observables en cada `00_main`, sin reinterpretarlos. Aporta la base para discutir rutas, entornos y contingency sin entrar todavía en diseño de arquitectura web.

## Qué va aquí

- Lista de binarios frontend y backend por aplicación, según `exploration.md`.
- Aviso de binarios backend no presentes en `00_main` (Condor, Brass).
- Aviso de configuraciones Dysflow ausentes o con `path-mismatch` (Condor, Brass, Expedientes).
- Mapeo de rutas relativas observadas (no se resuelven rutas de red en esta fase).

## Estado del contenido

**Placeholder.** No incluye tablas con binarios todavía; su redacción se apoya en `01-inventario-aplicaciones.md` y en los resultados de `get_capabilities` de cada checkout.

## Fuentes de autoridad

1. Checkout `00_main` de cada aplicación bajo `C:\00repos\codigo`.
2. Resultado de `dysflow.get_capabilities` y `dysflow.list_objects` (solo lectura).
3. `C:\00repos\documentacion\OPENSPEC\` por aplicación.

## Reglas de evidencia

- Un binario se lista solo si está físicamente presente en el checkout o si Dysflow lo resuelve en modo lectura.
- Las rutas se citan tal cual aparecen; no se normalizan a una convención única en esta fase.
- Los backends no localizados se marcan como `pendiente de ruta` sin asumir la ruta.
- La configuración `.dysflow/project.json` se respeta; no se corrige en esta fase.

## Checklist

- [ ] Cada binario lleva ruta relativa al `00_main` y nombre exacto.
- [ ] Los huecos se justifican con el motivo (`no localizado`, `path-mismatch`, etc.).
- [ ] Ningún binario de APAP o APAP_WEB aparece aquí.

## Siguiente paso

Cruzar este documento con `04-integraciones-y-operacion/tablas-vinculadas-y-backends.md` al poblar ambos.
