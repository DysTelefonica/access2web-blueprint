type: minor

`issue-spec` y `pr-contract` (#787): «Closes y Refs a la misma issue» pasa de fallo a aviso (`issue-spec: aviso: …`, gana `Closes`) y `pr-contract` ya no lo repite; una referencia a una PR es un aviso si el cuerpo enlaza además una issue válida (sin ella sigue siendo fallo); en un eslabón `chain:partial` la issue no necesita `status:approved` (sí un `type:*`), la aprobación se exige en la punta de cadena. Los avisos salen por stderr y no cambian el código de salida. HR-54 refleja el cambio.
