type: patch

Los autotests `test-rc-aggregation.sh` y `test-suite-timing.sh` del nightly dejan de depender de que `actionlint` esté en el `PATH` del host: su repositorio falso incluye un stub propio y lo antepone al `PATH` del runner real, igual que ya hacía `test-bash-parallel.sh` (#1236, #1237). Así el paso de autotests puede correr antes de instalar la herramienta, sin convertir su ausencia en un falso `FAIL actionlint` (#1360).
