type: minor

El bloque `architecture` con `modules` admite `covered_roots` y `exempt`: todo fichero fuente bajo un paquete cubierto debe estar en un módulo de una capa o en `exempt`, y uno sin capa es rojo y se nombra. Un patrón de `exempt` que no casa nada es `STALE`. Cierra el hueco por el que un fichero fuera de los módulos descritos pasaba sin medirse (#771).
