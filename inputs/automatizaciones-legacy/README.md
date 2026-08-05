# Automatizaciones legacy (VBS)

Sanitized overview of the 11 `.vbs` automation scripts that drove the legacy Access ecosystem.
Raw evidence (the original `.vbs` files) lives in `raw/` and is Git-ignored — it stays local only.

## What this folder is

- `README.md` (this file) — sanitized, versioned.
- `raw/` — **unversioned**, Git-ignored, exact original bytes of the 11 scripts.

Do not commit, push, paste, or otherwise reproduce the contents of `raw/`.

## Entry flow

`raw/entrypoint/script.vbs` is a Windows-script daemon (`Do While True ... WScript.Sleep`)
that orchestrates the daily batch and the mailer cycles. Cycle length depends on
"is laborable day?" (Festivos file on the internal UNC) and "is night?" (hour band).

Daily orchestrated batch (run sequentially on laborable days after 06:00):

1. `raw/orchestrated-tasks/NoConformidades_bat.vbs` (sleep ~30 s)
2. `raw/orchestrated-tasks/GestionRiesgos_bat.vbs` (sleep ~60 s)
3. `raw/orchestrated-tasks/BRASS.vbs` (sleep ~30 s)
4. `raw/orchestrated-tasks/TareaExpedientes.vbs` (sleep ~30 s)
5. `raw/orchestrated-tasks/Tareas.VBS` (sleep ~120 s)
6. `raw/orchestrated-tasks/HPS_SOLICITUDES.vbs` (sleep ~120 s)

Every cycle (laborable + festivo):

- `raw/mailers/EnviarCorreoTareas.vbs` — mailer: tareas.
- `raw/mailers/EnviarCorreoNoEnviado.vbs` — mailer: pendientes.
- `EnviarCorreosHPS.vbs` — invoked via `cscript.exe` from an external UNC path; **not part of `raw/`**.

The two scripts marked `huérfano` in the mapping below are not reachable from the entrypoint.

## Application mapping

| Script | Application (blueprint) |
|---|---|
| `script.vbs` | Daemon orquestador (entrypoint) |
| `Tareas.VBS` | Tareas (daily HTML digest) |
| `TareaExpedientes.vbs` | Lanzadera / Expedientes (ExpDiario) |
| `HPS_SOLICITUDES.vbs` | Lanzadera / HPS |
| `GestionRiesgos_bat.vbs` | Riesgos (semanal técnico + mensual calidad) |
| `NoConformidades_bat.vbs` | No Conformidades (diario técnico + semanal calidad) |
| `BRASS.vbs` | Equipos de medida BRASS |
| `EnviarCorreoTareas.vbs` | Mailer: Tareas |
| `EnviarCorreoNoEnviado.vbs` | Mailer: pendientes |
| `EnviarCorreoNOTAREAS.vbs` | Mailer alternativo (huérfano) |
| `CopiaDatosTE.vbs` | Utilidad NAS (huérfano) |

## Security warning

The raw scripts in `raw/` contain **hardcoded sensitive material** that has **not** been
removed from the originals:

- A shared Jet-OLEDB database password (one constant, repeated across all 11 files).
- Hardcoded internal SMTP server address.
- Hardcoded internal UNC hostnames and absolute paths.
- Hardcoded internal distribution lists and `From` addresses.

Treat `raw/` as **read-only evidence** for migration analysis only. Do not execute, edit,
copy, or re-paste from `raw/`. Migration to web must externalize every secret listed above
into a configuration source not embedded in code.

## Provenance

Each file in `raw/` is byte-identical to its source under `scripts_legacy/`. SHA-256 hashes
were captured before the move and verified after; see the move record in Engram topic
`discovery/legacy-automation-scripts` for the exact hash set.

The original `scripts_legacy/` directory was removed once the move was verified.