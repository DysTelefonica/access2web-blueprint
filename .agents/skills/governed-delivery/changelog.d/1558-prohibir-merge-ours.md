type: minor

`governed-delivery` 1.6: la regla 1 prohíbe `git merge -s ours` y `-X ours` en una rama que va a un PR. Las historias se unen con un `git merge origin/<base>` normal; `-s ours` descarta los cambios de la base y el merge gobernado puede revertirlos en silencio (ardelperal/APAP_WEB, 2026-11; #1558).
