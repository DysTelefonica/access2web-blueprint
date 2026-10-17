#!/usr/bin/env python3
# test_markdown_fences.py — detección CommonMark de bloques cercados (#712, #725)
"""El módulo compartido ``markdown_fences``: tildes, delimitadores más largos,
bloque anidado y bloque sin cerrar. Solo stdlib."""

import importlib.util
import unittest
from pathlib import Path

_PATH = Path(__file__).resolve().parents[2] / "markdown_fences.py"
_SPEC = importlib.util.spec_from_file_location("markdown_fences", _PATH)
FENCES = importlib.util.module_from_spec(_SPEC)
_SPEC.loader.exec_module(FENCES)


def outside(text):
    return [line.strip() for _n, line in FENCES.unfenced_lines(text)]


class MarkdownFencesTests(unittest.TestCase):
    def test_un_bloque_de_tildes_se_ignora(self):
        self.assertEqual(outside("a\n~~~\n[x](ghost.md)\n~~~\nb\n"), ["a", "b"])

    def test_un_bloque_de_cuatro_comillas_que_contiene_uno_de_tres_no_se_rompe(self):
        text = "a\n````md\n```sh\n[x](ghost.md)\n```\n[y](ghost.md)\n````\n[z](real.md)\n"
        self.assertEqual(outside(text), ["a", "[z](real.md)"])

    def test_no_cierra_con_menos_comillas_ni_con_otro_caracter(self):
        self.assertEqual(outside("````\n```\n~~~~\nx\n"), [])

    def test_un_bloque_sin_cerrar_llega_hasta_el_final(self):
        self.assertEqual(outside("a\n```\nb\nc\n"), ["a"])

    def test_los_numeros_de_linea_no_se_mueven(self):
        self.assertEqual([n for n, _l in FENCES.unfenced_lines("a\n```\nb\n```\nc\n")], [1, 5])


    def test_fenced_blocks_devuelve_informacion_cuerpo_y_linea(self):
        text = "a\n```json x\n{}\n```\nb\n~~~\nraw\n"
        self.assertEqual(list(FENCES.fenced_blocks(text)), [("json x", "{}\n", 2), ("", "raw\n", 6)])

    def test_fenced_blocks_respeta_el_bloque_anidado(self):
        text = "````md\n```sh\nls\n```\n````\n"
        self.assertEqual(list(FENCES.fenced_blocks(text)), [("md", "```sh\nls\n```\n", 1)])


if __name__ == "__main__":
    unittest.main()
