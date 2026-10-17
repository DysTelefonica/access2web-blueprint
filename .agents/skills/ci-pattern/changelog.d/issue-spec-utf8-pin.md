type: patch

issue-spec fija stdout y stderr a UTF-8 con _pin_output_encoding al inicio de main(), igual que los gates de APAP_WEB (issue #488). Sin el pin, un mensaje con caracteres fuera del locale (cp1252) lanzaba UnicodeEncodeError. Lo exige tests/test_gate_output_encoding.py de APAP_WEB.
