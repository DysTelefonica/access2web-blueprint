type: patch

HR-63 (#725): el extractor de rutas de `check_docs_codeguide.py` ignora los bloques de código cercados (comillas invertidas y tildes); antes invertían el emparejamiento del código en línea posterior, con falsos positivos y rutas rotas sin detectar.
