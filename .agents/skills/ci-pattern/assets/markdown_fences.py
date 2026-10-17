"""Bloques de código cercados de Markdown, compartido por los gates de documentación.

Sigue CommonMark: un bloque lo abren tres o más comillas invertidas o tildes
(una apertura con comillas invertidas no lleva comillas en su texto informativo)
y lo cierra una línea sin texto con el mismo carácter y al menos la misma
longitud. Un bloque sin cerrar llega hasta el final. Un único módulo para que
los gates no tengan detectores que puedan divergir (como ``host_plan``).
"""

import re

_FENCE_RE = re.compile(r"^\s*(`{3,}|~{3,})(.*)$")


def _opens(match):
    """Una línea abre un bloque: comillas invertidas sin comillas en la información, o tildes."""
    return bool(match) and not (match.group(1)[0] == "`" and "`" in match.group(2))


def _closes(match, fence):
    """Una línea cierra el bloque abierto con ``fence``: mismo carácter, longitud mayor o igual, sin texto."""
    return bool(match) and match.group(1)[0] == fence[0] and len(match.group(1)) >= len(fence) \
        and not match.group(2).strip()


def classify_lines(text):
    """``(número, línea, cercada)``; las líneas delimitadoras cuentan como cercadas."""
    fence = ""
    for number, line in enumerate(text.splitlines(keepends=True), 1):
        match = _FENCE_RE.match(line)
        if not fence:
            opens = _opens(match)
            if opens:
                fence = match.group(1)
            yield number, line, opens
            continue
        if _closes(match, fence):
            fence = ""
        yield number, line, True


def unfenced_lines(text):
    """``(número, línea)`` de las líneas fuera de bloques cercados."""
    return ((number, line) for number, line, fenced in classify_lines(text) if not fenced)


def fenced_blocks(text):
    """``(información, cuerpo, línea de apertura)`` de cada bloque cercado.

    El cuerpo es el texto entre los delimitadores; uno sin cerrar llega al final."""
    fence, info, body, start = "", "", [], 0
    for number, line in enumerate(text.splitlines(keepends=True), 1):
        match = _FENCE_RE.match(line)
        if not fence:
            if _opens(match):
                fence, info, body, start = match.group(1), match.group(2).strip(), [], number
            continue
        if _closes(match, fence):
            yield info, "".join(body), start
            fence = ""
        else:
            body.append(line)
    if fence:
        yield info, "".join(body), start
