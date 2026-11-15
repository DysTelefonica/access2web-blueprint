#!/usr/bin/env python3
# ci-pattern asset — image provenance gate for images built outside CI (#1389)
"""Image provenance: release manifest, and the gate that compares an image with it.

A web consumer whose registry is unreachable from CI builds and deploys its
images by hand. Nothing then ties the deployed image to a release tag: in
DysTelefonica/cadete production ran an image built from a commit that was
never pushed while its tag pointed elsewhere (cadete#748, cadete#755).

Subcommands:

- ``manifest``: run by the release workflow after the release gates. Writes
  the sha256 of every file of the deployable tree AT THE TAG COMMIT, read from
  the git objects (``git ls-tree`` + ``git cat-file``), never from a checkout.
- ``gate``: run before ``oc set image``. Reads the image archive, compares
  its deployable tree with the manifest of its release tag and checks that
  the image revision is that tag's commit, on the default branch.

Normalisation (explicit, #1389): NONE. The manifest holds the sha256 of the
git blob bytes; the image must carry exactly those bytes. A file that
differs only in line endings is still a finding, reported as "line endings
only" so the cause is visible, never hidden. File modes and owners are not
compared: the build may chmod/chown, the content may not change. Symlinks
are recorded by target.

Exit codes: 0 clean · 1 findings · 2 unmeasurable or invalid input (fail
closed). stdlib only.
"""

from __future__ import annotations

import argparse
import fnmatch
import hashlib
import json
import subprocess
import sys
import tarfile
import tempfile
from pathlib import Path

MANIFEST_SCHEMA = "ci-pattern.image-provenance.manifest/v1"
NORMALISATION = ("none: sha256 of the git blob bytes at the tag commit; "
                 "line-ending differences are findings")
REQUIRED_CONFIG = ("tree", "image_path", "image", "github_repo")
MANIFEST_ASSET = "image-provenance-manifest.json"


class GateError(Exception):
    """A condition the gate refuses to interpret; maps to exit 2."""


def sha256(data: bytes) -> str:
    return "sha256:" + hashlib.sha256(data).hexdigest()


def lf_sha256(data: bytes) -> str:
    """Digest of the content with CRLF turned into LF (diagnosis only)."""
    return sha256(data.replace(b"\r\n", b"\n"))


def load_config(path: str) -> dict:
    try:
        cfg = json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        msg = f"cannot read config {path!r}: {exc}"
        raise GateError(msg) from exc
    if not isinstance(cfg, dict):
        msg = f"config {path!r} is not a JSON object"
        raise GateError(msg)
    missing = [k for k in REQUIRED_CONFIG if not cfg.get(k)]
    if missing:
        msg = f"config {path!r} lacks required key(s): {', '.join(missing)}"
        raise GateError(msg)
    exclude = cfg.setdefault("exclude", [])
    if not isinstance(exclude, list) or not all(isinstance(p, str) and p for p in exclude):
        msg = f"config {path!r}: 'exclude' must be a list of non-empty globs"
        raise GateError(msg)
    cfg["tree"] = cfg["tree"].strip("/")
    cfg.setdefault("default_branch", "main")
    return cfg


def is_excluded(rel: str, patterns: list[str]) -> bool:
    """A pattern excludes the path it names, everything below it, or its glob."""
    for pat in patterns:
        base = pat.rstrip("/")
        if base.endswith("/**"):
            base = base[:-3]
        if rel == base or rel.startswith(base + "/") or fnmatch.fnmatchcase(rel, pat):
            return True
    return False


def execute(cmd: list[str], *, input_bytes: bytes | None = None) -> subprocess.CompletedProcess:
    """The single process boundary: argument lists built here, never a shell;
    git, podman, oc and gh are resolved on the operator's PATH on purpose."""
    return subprocess.run(cmd, input=input_bytes, capture_output=True, check=False)  # noqa: S603


def git_status(repo, *args) -> tuple[int, str]:
    """``(returncode, stdout)`` of a git query whose failure is an answer, not an error."""
    r = execute(["git", "-C", str(repo), *args])
    return r.returncode, r.stdout.decode("utf-8", "replace").strip()


def git(repo, *args, input_bytes: bytes | None = None) -> bytes:
    r = execute(["git", "-C", str(repo), *args], input_bytes=input_bytes)
    if r.returncode != 0:
        msg = f"git {' '.join(args)} failed: {r.stderr.decode('utf-8', 'replace').strip()}"
        raise GateError(msg)
    return r.stdout


def git_tree_entries(repo, commit: str, tree: str, exclude: list[str]) -> dict[str, str]:
    """``{path relative to tree: "sha256:…" | "symlink:<target>"}`` at ``commit``."""
    listing = git(repo, "ls-tree", "-r", "-z", "--full-tree", commit, "--", tree)
    wanted = []
    prefix = tree + "/"
    for raw in listing.split(b"\0"):
        meta, _, name = raw.partition(b"\t")
        path = name.decode("utf-8")
        rel = path[len(prefix):]
        if not path.startswith(prefix) or is_excluded(rel, exclude):
            continue
        mode, kind, obj = meta.decode().split(" ")
        if kind != "blob":
            msg = f"{path}: {kind} entries (submodules) are not supported"
            raise GateError(msg)
        wanted.append((rel, mode, obj))
    if not wanted:
        msg = f"tree {tree!r} has no files at {commit[:12]} (empty subject)"
        raise GateError(msg)
    out = git(repo, "cat-file", "--batch",
              input_bytes="".join(f"{obj}\n" for _, _, obj in wanted).encode())
    entries, pos = {}, 0
    for rel, mode, _ in wanted:
        header_end = out.index(b"\n", pos)
        size = int(out[pos:header_end].split(b" ")[2])
        data = out[header_end + 1:header_end + 1 + size]
        pos = header_end + 1 + size + 1
        entries[rel] = ("symlink:" + data.decode("utf-8")) if mode == "120000" else sha256(data)
    return entries


def build_manifest(repo, tag: str, cfg: dict) -> dict:
    try:
        commit = git(repo, "rev-parse", "--verify", f"refs/tags/{tag}^{{commit}}").decode().strip()
    except GateError as exc:
        msg = f"tag {tag!r} does not exist in {repo!r} (empty subject)"
        raise GateError(msg) from exc
    return {
        "schema": MANIFEST_SCHEMA,
        "tag": tag,
        "commit": commit,
        "tree": cfg["tree"],
        "exclude": cfg["exclude"],
        "normalisation": NORMALISATION,
        "files": git_tree_entries(repo, commit, cfg["tree"], cfg["exclude"]),
    }


def load_manifest(path: str) -> dict:
    try:
        man = json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        msg = f"cannot read manifest {path!r}: {exc}"
        raise GateError(msg) from exc
    if not isinstance(man, dict) or man.get("schema") != MANIFEST_SCHEMA:
        msg = f"manifest {path!r} is not a {MANIFEST_SCHEMA} document"
        raise GateError(msg)
    if not isinstance(man.get("files"), dict) or not man["files"]:
        msg = f"manifest {path!r} lists no files (empty subject)"
        raise GateError(msg)
    return man


def compare(expected: dict[str, str], actual: dict[str, str],
            actual_lf: dict[str, str]) -> list[str]:
    """Findings of ``actual`` (image) against ``expected`` (manifest), sorted by path."""
    findings = []
    for path in sorted(expected.keys() | actual.keys()):
        want, got = expected.get(path), actual.get(path)
        if got is None:
            findings.append(f"missing in image: {path}")
        elif want is None:
            findings.append(f"extra in image: {path}")
        elif want != got:
            if actual_lf.get(path) == want:
                findings.append(f"changed in line endings only (CRLF in image): {path}")
            else:
                findings.append(f"changed: {path}")
    return findings


def _norm(name: str) -> str:
    """A tar member name as a plain relative POSIX path."""
    while name.startswith("./"):
        name = name[2:]
    return name.strip("/")


class ImageTree:
    """Labels and deployable files of one image archive (docker-archive or OCI)."""

    REVISION = "org.opencontainers.image.revision"
    VERSION = "org.opencontainers.image.version"

    def __init__(self, archive: str, image_path: str, exclude: list[str]):
        self.archive = archive
        self.prefix = image_path.strip("/") + "/"
        self.exclude = exclude
        self.files: dict[str, str] = {}
        self.lf: dict[str, str] = {}
        try:
            with tarfile.open(archive) as outer:
                config, layers = self._index(outer)
                self.labels = (config.get("config") or {}).get("Labels") or {}
                for name in layers:
                    with tarfile.open(fileobj=outer.extractfile(name), mode="r:*") as layer:
                        members = layer.getmembers()
                        self._whiteouts(members)
                        for member in members:
                            self._add(layer, member)
        except (OSError, tarfile.TarError, ValueError, KeyError, IndexError) as exc:
            msg = f"cannot read image archive {archive!r}: {exc}"
            raise GateError(msg) from exc
        if not self.files:
            msg = f"image archive {archive!r} has no files under {image_path!r} (empty subject)"
            raise GateError(msg)

    @staticmethod
    def _json(outer, name: str):
        return json.loads(outer.extractfile(name).read())

    def _index(self, outer) -> tuple[dict, list[str]]:
        names = set(outer.getnames())
        if "manifest.json" in names:  # docker-archive (docker save, podman save)
            entry = self._json(outer, "manifest.json")[0]
            return self._json(outer, entry["Config"]), list(entry["Layers"])
        if "index.json" in names:  # oci-archive
            def blob(digest: str) -> str:
                return "blobs/" + digest.replace(":", "/", 1)
            index = self._json(outer, "index.json")
            manifest = self._json(outer, blob(index["manifests"][0]["digest"]))
            return (self._json(outer, blob(manifest["config"]["digest"])),
                    [blob(layer["digest"]) for layer in manifest["layers"]])
        msg = "neither manifest.json (docker-archive) nor index.json (OCI)"
        raise ValueError(msg)

    def _rel(self, name: str) -> str | None:
        name = _norm(name)
        if not name.startswith(self.prefix):
            return None
        rel = name[len(self.prefix):].rstrip("/")
        return rel if rel and not is_excluded(rel, self.exclude) else None

    def _drop(self, rel: str) -> None:
        for path in [p for p in self.files if p == rel or p.startswith(rel + "/")]:
            del self.files[path]
            self.lf.pop(path, None)

    def _whiteouts(self, members) -> None:
        """Whiteouts hide lower-layer content before the layer adds its own."""
        for m in members:
            parent, _, leaf = _norm(m.name).rpartition("/")
            if leaf == ".wh..wh..opq":
                self._drop_dir(parent)
            elif leaf.startswith(".wh."):
                rel = self._rel(f"{parent}/{leaf[4:]}")
                if rel:
                    self._drop(rel)

    def _add(self, layer, m) -> None:
        rel = self._rel(m.name)
        if rel is None or m.isdir() or m.name.rsplit("/", 1)[-1].startswith(".wh."):
            return
        self._drop(rel)
        if m.issym():
            self.files[rel] = "symlink:" + m.linkname
        elif m.islnk():
            target = self._rel(m.linkname)
            self.files[rel] = self.files.get(target, "hardlink:" + m.linkname)
            if target in self.lf:
                self.lf[rel] = self.lf[target]
        elif m.isfile():
            data = layer.extractfile(m).read()
            self.files[rel], self.lf[rel] = sha256(data), lf_sha256(data)

    def _drop_dir(self, directory: str) -> None:
        """Opaque whiteout: the directory's lower-layer content disappears."""
        if directory + "/" == self.prefix or self.prefix.startswith(directory + "/"):
            self.files.clear()
            self.lf.clear()
        elif directory.startswith(self.prefix):
            self._drop(directory[len(self.prefix):])


def _pin_output_encoding() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")


def cmd_manifest(args) -> int:
    cfg = load_config(args.config)
    man = build_manifest(args.repo, args.tag, cfg)
    Path(args.out).write_text(json.dumps(man, indent=2, sort_keys=True) + "\n",
                              encoding="utf-8")
    print(f"MANIFEST OK: {len(man['files'])} file(s) of {cfg['tree']!r} at tag "
          f"{args.tag!r} ({man['commit'][:12]}) -> {args.out}")
    return 0


def run_cmd(cmd: list[str], what: str) -> str:
    try:
        r = execute(cmd)
    except OSError as exc:
        msg = f"{what}: cannot run {cmd[0]!r}: {exc}"
        raise GateError(msg) from exc
    if r.returncode != 0:
        err = r.stderr.decode("utf-8", "replace").strip()
        msg = f"{what}: {' '.join(cmd)} exited {r.returncode}: {err}"
        raise GateError(msg)
    return r.stdout.decode("utf-8", "replace").strip()


def image_tree(cfg: dict, exclude: list[str], source: tuple[str | None, str | None],
               workdir: Path, label: str) -> ImageTree:
    """Read ``source`` = (archive, reference); a reference is ``<engine> save``-d first."""
    archive, ref = source
    if archive is None:
        archive = str(workdir / f"{label}.tar")
        run_cmd([*cfg.get("engine_cmd", ["podman"]), "save", "-o", archive, ref],
                f"saving {label} image {ref!r}")
    return ImageTree(archive, cfg["image_path"], exclude)


def remote_facts(remote: str, tag: str, revision: str | None, branch: str,
                 workdir: Path) -> tuple[str | None, list[str]]:
    """The tag commit on the remote and the provenance findings, from a fresh clone."""
    bare = workdir / "remote.git"
    git(workdir, "clone", "--quiet", "--bare", "--filter=blob:none", "--", remote, str(bare))
    if git_status(bare, "rev-parse", "-q", "--verify", f"refs/heads/{branch}")[0] != 0:
        msg = f"default branch {branch!r} does not exist on {remote!r}"
        raise GateError(msg)
    findings = []
    rc, out = git_status(bare, "rev-parse", "-q", "--verify", f"refs/tags/{tag}^{{commit}}")
    tag_commit = out if rc == 0 else None
    if tag_commit is None:
        findings.append(f"provenance: release tag {tag!r} does not exist on the remote")
    else:
        rc, _ = git_status(bare, "merge-base", "--is-ancestor", tag_commit, f"refs/heads/{branch}")
        if rc not in (0, 1):
            msg = f"git merge-base failed on {remote!r}"
            raise GateError(msg)
        if rc == 1:
            findings.append(f"provenance: tag {tag!r} ({tag_commit[:12]}) is not on {branch!r}")
    rc, resolved = git_status(bare, "rev-parse", "-q", "--verify", f"{revision}^{{commit}}") \
        if revision else (1, "")
    if not revision:
        findings.append(f"provenance: image has no {ImageTree.REVISION} label "
                        "(not built by the release build script)")
    elif rc != 0:
        findings.append(f"provenance: image revision {revision} is not a commit of the "
                        "remote (built from a commit that was never pushed)")
    elif tag_commit and resolved != tag_commit:
        findings.append(f"provenance: image revision {revision[:12]} is not the commit of "
                        f"tag {tag!r} ({tag_commit[:12]})")
    return tag_commit, findings


def fetch_manifest(cfg: dict, tag: str, workdir: Path) -> str:
    """Download the manifest asset of the tag's GitHub Release."""
    run_cmd([*cfg.get("gh_cmd", ["gh"]), "release", "download", tag, "-R", cfg["github_repo"],
             "-p", MANIFEST_ASSET, "-D", str(workdir), "--clobber"],
            f"downloading {MANIFEST_ASSET!r} of release {tag!r}")
    return str(workdir / MANIFEST_ASSET)


def current_image_ref(cfg: dict) -> str:
    """The image the cluster deployment runs now (``oc get deployment``)."""
    shift = cfg.get("openshift") or {}
    if not shift.get("namespace") or not shift.get("deployment"):
        msg = "config 'openshift' needs 'namespace' and 'deployment'"
        raise GateError(msg)
    container = shift.get("container")
    path = (f'{{.spec.template.spec.containers[?(@.name=="{container}")].image}}'
            if container else "{.spec.template.spec.containers[0].image}")
    cmd = [*cfg.get("oc_cmd", ["oc"])]
    if shift.get("context"):
        cmd += ["--context", shift["context"]]
    ref = run_cmd([*cmd, "get", f"deployment/{shift['deployment']}", "-n",
                   shift["namespace"], "-o", f"jsonpath={path}"], "reading the running image")
    if not ref:
        msg = "the deployment reports no image for that container"
        raise GateError(msg)
    return ref


def diff_against_current(new: ImageTree, current: ImageTree) -> list[str]:
    lines = []
    for path in sorted(new.files.keys() | current.files.keys()):
        a, b = current.files.get(path), new.files.get(path)
        if b is None:
            lines.append(f"only in the running image (disappears): {path}")
        elif a is None:
            lines.append(f"only in the new image: {path}")
        elif a != b:
            lines.append(f"changed: {path}")
    return lines


def content_findings(cfg: dict, image: ImageTree, man: dict, tag: str,
                     tag_commit: str | None) -> tuple[ImageTree, list[str]]:
    """The image against its manifest, with the exclusion the manifest recorded."""
    findings = []
    if man.get("tag") != tag or man.get("commit") != tag_commit:
        findings.append(f"provenance: manifest is for {man.get('tag')!r} "
                        f"({str(man.get('commit'))[:12]}), not for tag {tag!r}")
    if man.get("exclude") != cfg["exclude"]:
        image = ImageTree(image.archive, cfg["image_path"], man.get("exclude") or [])
    findings += [f"content: {f}" for f in compare(man["files"], image.files, image.lf)]
    return image, findings


def print_current_diff(args, cfg: dict, image: ImageTree, work: Path) -> None:
    ref = args.current_image or (current_image_ref(cfg) if args.current_from_cluster else None)
    current = image_tree(cfg, image.exclude, (args.current_archive, ref), work, "current")
    diff = diff_against_current(image, current)
    print(f"DIFF AGAINST THE RUNNING IMAGE ({ref or args.current_archive}): "
          f"{len(diff)} difference(s)")
    for line in diff:
        print(f"  {line}")


def report_failure(name: str, tag: str, revision: str | None, findings: list[str]) -> int:
    print(f"IMAGE-PROVENANCE FAIL: {len(findings)} finding(s) for {name} "
          f"(release tag {tag!r}, revision {revision or 'none'}); deployment refused")
    kinds: dict[str, int] = {}
    for line in findings:
        kind = line.split(":", 2)[1].split(",")[0].strip() \
            if line.startswith("content:") else "provenance"
        kinds[kind] = kinds.get(kind, 0) + 1
    print("  summary: " + ", ".join(f"{n} {k}" for k, n in kinds.items()))
    for line in findings:
        print(f"  - {line}")
    return 1


def cmd_gate(args) -> int:
    cfg = load_config(args.config)
    with tempfile.TemporaryDirectory(prefix="image-provenance-") as tmp:
        work = Path(tmp)
        image = image_tree(cfg, cfg["exclude"], (args.image_archive, args.image), work, "new")
        tag = args.tag or image.labels.get(ImageTree.VERSION)
        if not tag:
            msg = f"no release tag: pass --tag or build the image with the {ImageTree.VERSION} label"
            raise GateError(msg)
        revision = image.labels.get(ImageTree.REVISION)
        remote = args.git_remote or f"https://github.com/{cfg['github_repo']}.git"
        tag_commit, findings = remote_facts(remote, tag, revision, cfg["default_branch"], work)
        try:
            man = load_manifest(args.manifest or fetch_manifest(cfg, tag, work))
        except GateError as exc:
            if not findings:
                raise
            findings.append(f"content: not compared, no release manifest ({exc})")
            man = None
        if man is not None:
            image, more = content_findings(cfg, image, man, tag, tag_commit)
            findings += more
        if args.current_archive or args.current_image or args.current_from_cluster:
            print_current_diff(args, cfg, image, work)
    name = args.image or args.image_archive
    if findings:
        return report_failure(name, tag, revision, findings)
    print(f"IMAGE-PROVENANCE OK: {name} is release {tag!r} ({tag_commit[:12]}) on "
          f"{cfg['default_branch']!r}; {len(man['files'])} file(s) match its manifest.")
    return 0


def parser() -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(description="Image provenance (#1389).")
    sub = ap.add_subparsers(dest="cmd", required=True)
    m = sub.add_parser("manifest", help="write the release manifest of a tag")
    m.add_argument("--config", required=True)
    m.add_argument("--repo", default=".")
    m.add_argument("--tag", required=True)
    m.add_argument("--out", default=MANIFEST_ASSET,
                   help=f"output file (default: {MANIFEST_ASSET}, the release asset name)")
    m.set_defaults(func=cmd_manifest)
    g = sub.add_parser("gate", help="refuse an image that is not its release, byte for byte")
    g.add_argument("--config", required=True)
    src = g.add_mutually_exclusive_group(required=True)
    src.add_argument("--image", help="image reference, saved with the configured engine")
    src.add_argument("--image-archive", help="docker-archive or oci-archive tarball")
    g.add_argument("--tag", help="release tag (default: the image's version label)")
    g.add_argument("--manifest", help="local manifest instead of the release asset")
    g.add_argument("--git-remote", help="remote URL (default: the GitHub repository)")
    cur = g.add_mutually_exclusive_group()
    cur.add_argument("--current-image", help="running image reference to diff against")
    cur.add_argument("--current-archive", help="running image tarball to diff against")
    cur.add_argument("--current-from-cluster", action="store_true",
                     help="diff against the image the configured deployment runs")
    g.set_defaults(func=cmd_gate)
    return ap


def main(argv) -> int:
    _pin_output_encoding()
    args = parser().parse_args(argv)
    try:
        return args.func(args)
    except GateError as exc:
        print(f"IMAGE-PROVENANCE ERROR: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
