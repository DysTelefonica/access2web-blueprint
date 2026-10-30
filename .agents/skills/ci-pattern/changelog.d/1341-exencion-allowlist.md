type: patch

Una ampliación legítima del allowlist se declara como exención con identidad (DysTelefonica/team-skills#1341). Con `--seal-baseline`, el allowlist no puede autorizar de más que el de la base; ahora, la ampliación que el propio allowlist declara como exención —la superficie, el elemento y quién la autoriza (`actor` o `issue`, el patrón HR-33)— se admite, y una exención sin identidad no exime de nada: nadie puede responder por ella. La forma se documenta en el molde del allowlist (`assets/env-isolation/allowlist.example.json`).
