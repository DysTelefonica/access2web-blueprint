type: minor

Plantilla del workflow nocturno pre-release (#1627, slice 3), `assets/prerelease/nightly-pipeline.template.yml`: schedule, workflow_dispatch y workflow_call con `dry_run`; los jobs son marcadores que fallan en cerrado hasta que el proyecto pone su comando; el gate ejecuta `nightly_pipeline.py plan` y el deploy solo corre tras un gate verde y fuera de ensayo. El README documenta el asset. Refs #1627.
