type: patch

Registro (#776): `ardelperal/APAP_WEB` declara `pr_size_excludes_vendored: true` (su `pr-size` excluye `.agents/skills/` y `skills/`, ardelperal/APAP_WEB#2112). La propagación a APAP agrupa así el tramo vendorizado en pocos PRs en lugar de trocearlo a 400 líneas. Una prueba exige que el campo sea booleano en cualquier consumer que lo declare.
