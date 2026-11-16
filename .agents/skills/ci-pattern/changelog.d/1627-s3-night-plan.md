type: minor

Asset nuevo `assets/prerelease/nightly_pipeline.py plan` (#1627, slice 3): decide DEPLOY o HOLD para la canalización nocturna. Despliega solo si el perfil `.github/prerelease.json` está activo, todos los jobs requeridos están en verde sobre la cabeza de main y no hay migraciones pendientes; cualquier lectura fallida retiene. Con `--caller` acepta el prefijo `<caller> / ` de workflow_call; un fallo bajo otro prefijo se ignora, porque ese prefijo no es el de la nocturna. Refs #1627.
