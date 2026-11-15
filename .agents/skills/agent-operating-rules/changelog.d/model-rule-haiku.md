type: major

`agent-operating-rules` 1.4: HR-12 sigue la regla del dueño del 2026-11-02, que sustituye a la del 1 de noviembre. El worker trabaja primero con MiniMax. Si MiniMax se agota y el coordinador tiene Claude, cambia en el sitio a Haiku por `claude-bridge` y vuelve a MiniMax en el primer punto seguro de cada ventana. Si se agota también el Claude del dueño, todo se aparca y los coordinadores cierran en `gpt-6.1-terra` con un handoff. Se actualizan el gate «MiniMax se ha agotado», el paso 6 y el antipatrón de Haiku. Infra lo implementó en ardelperal/infra#41.

Migración: un worker aparcado por falta de MiniMax con Claude disponible en su coordinador se retoma ya con Haiku; no espera a la ventana.
