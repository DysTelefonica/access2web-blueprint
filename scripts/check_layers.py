

def _check_direction(
    origin: tuple[str, str], target: tuple[str, str], file: str, line: int
) -> Violation | None:
    _, origin_layer = origin
    _, target_layer = target
    if target_layer in ALLOWED_IMPORTS[origin_layer]:
        return None
    return Violation(
        key=f"direction:{origin_layer}->{target_layer}",
        detail=f"{origin_layer} may not import {target_layer}",
        file=file,
        line=line,
    )


def _check_slice(
    origin: tuple[str, str], target: tuple[str, str], file: str, line: int
) -> Violation | None:
    origin_module, _ = origin
    target_module, _ = target
    if target_module == origin_module or target_module in CROSS_CUTTING_MODULES:
        return None
    return Violation(
        key=f"slice:{origin_module}->{target_module}",
        detail=f"module {origin_module} reaches across into {target_module}",
        file=file,
        line=line,
    )


def _check_purity(
    origin: tuple[str, str], imported: str, file: str, line: int
) -> Violation | None:
    _, origin_layer = origin
    if origin_layer not in PURE_LAYERS:
        return None
    distribution = imported.split(".")[0]
    if distribution not in FORBIDDEN_IN_PURE_LAYERS:
        return None
    return Violation(
        key=f"purity:{origin_layer}",
        detail=f"{origin_layer} imports framework {distribution}",
        file=file,
        line=line,
    )


def collect_violations(root: Path) -> tuple[list[Violation], int]:
    """Return ``(violations, skipped_count)``.

    ``skipped_count`` is the number of files that were intentionally outside
    the layered lattice (composition roots, empty `__init__.py` markers).
    Hard Rule 18 still holds for files the gate cannot classify for an
    unintentional reason ÔÇö those are violations, not skips.
    """
    violations: list[Violation] = []
    skipped = 0
    for path in _iter_source_files(root):
        display = str(path.relative_to(root)).replace("\\", "/")
        origin = classify_file(path, root)
        if origin == SKIP_SENTINEL:
            # Composition root or empty `__init__.py` ÔÇö intentionally outside
            # the module lattice. Silent skip, not a violation.
            skipped += 1
            continue
        if origin is None:
            # Hard Rule 18: a file this gate cannot classify is a file this gate did not check.
            # Skipping it silently reports "clean" for code nobody looked at ÔÇö which is how a
            # layout mismatch turns a whole codebase invisible while CI stays green. Either
            # teach classify_file the layout, or record the file in BASELINE deliberately.
            violations.append(
                Violation(
                    key="unclassified",
                    detail="not classifiable into (module, layer); this file was NOT checked",
                    file=display,
                    line=0,
                )
            )
            continue
        try:
            tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        except SyntaxError as exc:  # a file that cannot be parsed is a failure, not a skip
            violations.append(
                Violation(
                    key="unparseable",
                    detail=f"syntax error: {exc.msg}",
                    file=display,
                    line=exc.lineno or 0,
                )
            )
            continue
        for imported, line in _imported_names(tree, path, root):
            purity = _check_purity(origin, imported, display, line)
            if purity:
                violations.append(purity)
            target = classify_dotted(imported)
            if target is None:
                continue
            direction = _check_direction(origin, target, display, line)
            if direction:
                violations.append(direction)
            slicing = _check_slice(origin, target, display, line)
            if slicing:
                violations.append(slicing)
    return violations, skipped


def evaluate(violations: list[Violation], today: date) -> tuple[int, list[str]]:
    """Compare observed violations against BASELINE. Returns ``(exit_code, report_lines)``."""
    observed: dict[str, list[Violation]] = {}
    for violation in violations:
        observed.setdefault(violation.key, []).append(violation)

    lines: list[str] = []
    failed = False

    for key in sorted(observed):
        entries = observed[key]
        allowance = BASELINE.get(key)
        if allowance is None:
            failed = True
            lines.append(f"FAIL  {key}: {len(entries)} occurrence(s), not in BASELINE")
        elif len(entries) > allowance.count:
            failed = True
            lines.append(
                f"FAIL  {key}: {len(entries)} occurrence(s), BASELINE allows {allowance.count}"
            )
        elif today.isoformat() > allowance.target_date and len(entries) > allowance.target:
            failed = True
            lines.append(
                f"FAIL  {key}: BASELINE expired on {allowance.target_date} with "
                f"{len(entries)} occurrence(s), target was {allowance.target}"
            )
        elif len(entries) < allowance.count:
            lines.append(
                f"NOTE  {key}: {len(entries)} occurrence(s), below BASELINE {allowance.count}; "
                f"lower the BASELINE to lock the gain in"
            )
        for entry in entries:
            lines.append(f"        {entry.file}:{entry.line}  {entry.detail}")

    for key in sorted(BASELINE):
        if key in observed:
            continue
        lines.append(f"NOTE  {key}: no longer occurs; remove it from BASELINE")

    if not failed:
        lines.append("OK    layer gate clean")
    return (1 if failed else 0), lines


def build_report(
    violations: list[Violation], status: str, files_seen: int = 0, skipped: int = 0
) -> dict:
    unclassified = sum(1 for violation in violations if violation.key == "unclassified")
    return {
        "gate": "layers",
        "status": status,
        "indicators": {
            "violations": len(violations),
            "violation_classes": len({violation.key for violation in violations}),
            # Coverage of the gate itself: how much of the package it actually inspected.
            "files_checked": files_seen - unclassified - skipped,
            "files_unclassified": unclassified,
            "files_skipped": skipped,
        },
        "ceilings": {"violations": 0, "violation_classes": 0, "files_unclassified": 0},
        "findings": [
            {"file": violation.file, "line": violation.line, "detail": violation.detail}
            for violation in sorted(violations, key=lambda item: (item.file, item.line, item.key))
        ],
    }


def _pin_output_encoding() -> None:
    """Pin stdout/stderr to UTF-8.

    Python picks the output encoding from the platform locale, so the same gate emits different
    bytes on a Windows workstation (cp1252) and a Linux runner (utf-8) ÔÇö and a non-encodable
    character crashes the write outright. A harness that claims determinism cannot let its own
    output depend on where it ran.
    """
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            reconfigure(encoding="utf-8")


def main(argv: list[str] | None = None) -> int:
    _pin_output_encoding()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--root",
        type=Path,
        default=Path.cwd(),
        help="repository root containing the package directory (default: cwd)",
    )
    parser.add_argument("--json", action="store_true", help="emit the indicator envelope")
    args = parser.parse_args(argv)

    root = args.root.resolve()
    package_dir = _package_dir(root)
    if not package_dir.is_dir():
        message = f"root package '{ROOT_PACKAGE}' not found under {root} (looked in '{package_dir}')"
        if args.json:
            print(json.dumps({"gate": "layers", "status": "error", "detail": message}))
        else:
            print(f"FAIL  {message}", file=sys.stderr)
        return 1

    violations, skipped = collect_violations(root)
    files_seen = len(_iter_source_files(root))
    exit_code, lines = evaluate(violations, date.today())

    if args.json:
        status = "pass" if exit_code == 0 else "fail"
        print(json.dumps(build_report(violations, status, files_seen, skipped), indent=2))
    else:
        for line in lines:
            print(line)
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
