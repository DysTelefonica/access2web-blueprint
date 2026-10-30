type: patch

`workflow-policy` (HR-39) casa ahora las rutas absolutas que pasa el adoption check con las rutas relativas declaradas en `pr_workflow` y `scheduled_workflow` (`_find_workflow`, por componentes de ruta, nunca por subcadena).

Antes, `adoption check` resolvía los `@ruta` del contrato contra el repo (rutas absolutas) y el gate reportaba `pr_workflow ... no esta entre los inputs` aunque el workflow estuviera declarado y presente: fase 4 en rojo sobre `access2web-blueprint` con la política verde en CI, que pasa rutas relativas.

- Una ruta declarada relativa sigue sin casar con un workflow ajeno: el criterio es igualdad de componentes finales, no prefijo.
- Las rutas relativas y la CI no cambian.
