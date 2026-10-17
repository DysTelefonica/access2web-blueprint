type: patch

La fixture `allow_network` del snippet de pytest-socket ahora hace `yield` del manejador y llama a `restore()` en el teardown: antes la red se quedaba abierta tras el test. El guard de vulture reconoce como uso los métodos `do_<MÉTODO>` (despacho por nombre de `http.server`) y `DEAD_GRANDFATHERED` desaparece: ya no queda ningún símbolo muerto (#663).
