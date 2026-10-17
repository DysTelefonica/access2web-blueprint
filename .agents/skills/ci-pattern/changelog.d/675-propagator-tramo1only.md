type: minor

Propagador (#675): nuevo parámetro `-Tramo1Only`, que abre solo el PR del tramo 1 (slice, `.team-skills.yaml`, sello de gobierno) y deja el tramo vendorizado sin commitear; y el propagador falla cerrado (`unmanaged_paths_blocked`, salida 2) si hay algún path staged fuera de su conjunto gestionado, en lugar de commitearlo en el tramo 2.
