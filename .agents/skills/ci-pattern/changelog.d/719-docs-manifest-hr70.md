type: minor

HR-70 (#719): manifiesto canónico de documentos (`assets/docs-manifest/required-docs.json`) y su gate. Por declaración `docs_manifest` en el contrato, la fase 5 de `adoption check` exige cada documento activado y cada sección obligatoria no vacía, y nombra fichero y sección; el consumer solo puede añadir, nunca quitar. Sin declaración avisa por stderr y sale 0.
