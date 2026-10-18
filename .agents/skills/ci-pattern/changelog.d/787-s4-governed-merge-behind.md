type: minor

Merge gobernado (#787): un PR atrasado respecto de su base se bloquea antes de fusionar (exit 1) con la acción a tomar (actualizar la rama y repetir), porque el árbol fusionado ya no sería el que revisó el CI y la evidencia de deploy lo rechaza («main moved»). El bloqueo por solape de workflows (#783) se conserva con su mensaje propio. Un compare sin `behind_by` válido falla cerrado (exit 2).
