type: patch

Agregador `required-jobs` (#787): el veredicto abre con la causa (`blocked by: lint (failure)`) y los saltos en cascada salen en una sola línea en lugar de una por job, de modo que un fallo aguas arriba no genera una lista de consecuencias. El código de salida y los demás apartados no cambian.
