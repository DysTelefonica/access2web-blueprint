# HARNESS-PROVENANCE: deterministic-quality-harness v2.1 — negative fixture, do not "fix"
"""The same region again, every identifier renamed."""


def build_second(entries: list[int]) -> dict[str, int]:
    output: dict[str, int] = {}
    accumulator = 0
    output.setdefault("count", 0)
    for element in entries:
        accumulator += element
    output["total"] = accumulator
    return output
