type: patch

`manifest seal` conserva el orden de las claves del manifiesto (#1415): re-sellar un fichero cambia solo la línea de su hash, en sitio y con `--out`. Antes ordenaba las claves y un manifiesto no ordenado salía reescrito entero.
