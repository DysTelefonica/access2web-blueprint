type: patch

El canon corre actionlint sobre sus workflows: el gate ve un `if:` que lee un `steps.<id>` de un paso sin `id:` —la suite del propagador estuvo muerta nueve días por eso— y sale en rojo. La instalación va pineada a 1.7.12 con el SHA256 verificado para `linux_amd64` y `linux_arm64` (`scripts/install-actionlint.sh`), y el runner local ejecuta el mismo comando para que la paridad local/CI lo vea (HR-4). La política de paridad del canon declara la provisión del binario como exención, con su motivo (#1090).
