type: minor

HR-63 (#709): el mapa del código hub+spokes (`docs/CODEBASE-GUIDE.md` + `docs/codebase/`) se exige por declaración `docs_map` en el contrato de adopción. La fase 5 de `adoption check` valida el hub, los spokes enlazados en los dos sentidos, la sincronía mapa↔código de `docs-codeguide` y el job requerido; sin declaración avisa por stderr y sale 0. Plantillas de hub y spoke en `assets/templates/docs/`.
