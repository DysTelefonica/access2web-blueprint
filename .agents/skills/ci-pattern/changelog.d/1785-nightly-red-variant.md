type: minor

El parrafo "Regla dura del dueno" del partial web (aparcar issues y congelar merges por noche roja) va entre marcas `nightly-red` (team-skills#1785). La clave opcional `nightly_red_policy` de la entrada del registry elige la variante: sin clave, el texto de hoy byte a byte; `priority-issue` (declarada solo para ardelperal/APAP_WEB) deja una noche roja con issue prioritaria y relanzamiento el mismo dia, sin aparcar ni congelar. Un valor distinto aborta antes de propagar. `prerelease-profile.md` documenta ambas variantes.
