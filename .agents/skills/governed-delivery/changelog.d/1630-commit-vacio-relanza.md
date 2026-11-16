type: patch

Regla 3: un eslabón de cadena que necesita un run lo consigue solo con un commit vacío; cerrar y reabrir el PR deja de ser una vía. Bajo el perfil pre-release (#1627, regla 5), re-apuntar sin cambiar el head reutiliza el run verde que ya existe sobre ese SHA. Un run cancelado o caído lo relanza el coordinador con `rerun_run.py` (#1630).
