type: minor

HR-63 (#737): `check_docs_codeguide.py` acepta como citas válidas un glob que casa con al menos un fichero, un segmento marcador `<nombre>` con prefijo existente y las alternancias `{a,b}` cuyas expansiones lo son; un glob sin coincidencias o un prefijo inexistente siguen en rojo y se nombran. Antes obligaban al consumer a reescribir su documentación.
