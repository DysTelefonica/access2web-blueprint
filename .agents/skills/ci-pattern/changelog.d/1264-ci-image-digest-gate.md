type: patch

Gate genérico de imágenes de CI fijadas por digest: `container.image`,
`services.<id>.image`, `driver-opts: image=` de buildx, `docker run|pull|build` y
`uses: docker://` llevan `@sha256:<64 hex>`; opt-in por `P56_ci_images_gate`
(`off` por defecto) y excepciones con motivo en `P57_ci_image_exceptions` (#1264).
