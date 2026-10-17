type: patch

#753: `check_pr_contract.py` cumple el ratchet de ruff extendido de los consumers sin supresiones: las excepciones tienen clase propia con el mensaje dentro (`EventFileMissingError`, `PolicyFileMissingError`, `PullRequestMissingError`, `UnknownChecksError`, …), un tipo inesperado lanza `TypeError`, y el mínimo de padres de un merge es una constante. `testing/suites/vendorable-ruff` ejecuta el conjunto extendido (`S,ERA,ARG,FAST,N,C901,PLR,SIM,RET,TRY,PTH`) sobre los activos que APAP copia byte a byte. Mensajes y códigos de salida no cambian.
