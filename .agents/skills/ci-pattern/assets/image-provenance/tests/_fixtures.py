# ci-pattern asset tests — shared fixtures for image-provenance (#1389)
"""Throwaway git repositories and image archives, local only (no network).
Git runs without global or system configuration and takes its identity from
``GIT_*`` variables, never from ``git config``."""
from __future__ import annotations

import atexit
import hashlib
import io
import json
import os
import subprocess
import sys
import tarfile
import tempfile
from pathlib import Path

ASSET = Path(__file__).resolve().parent.parent
TOOL = ASSET / "image_provenance.py"

_fd, _EMPTY_CONFIG = tempfile.mkstemp(prefix="ci-pattern-1389-gitconfig-")
os.close(_fd)
atexit.register(os.unlink, _EMPTY_CONFIG)
GIT_ENV = {
    **os.environ,
    "GIT_CONFIG_GLOBAL": _EMPTY_CONFIG,
    "GIT_CONFIG_NOSYSTEM": "1",
    "GIT_AUTHOR_NAME": "fixture", "GIT_AUTHOR_EMAIL": "fixture@example.test",
    "GIT_COMMITTER_NAME": "fixture", "GIT_COMMITTER_EMAIL": "fixture@example.test",
}


def sha(data: bytes) -> str:
    return "sha256:" + hashlib.sha256(data).hexdigest()


def run_tool(*args, cwd=None) -> subprocess.CompletedProcess:
    return subprocess.run([sys.executable, str(TOOL), *map(str, args)], cwd=cwd,
                          capture_output=True, text=True, encoding="utf-8", env=GIT_ENV)


def git(cwd, *args) -> str:
    return subprocess.run(["git", "-C", str(cwd), *map(str, args)], capture_output=True,
                          text=True, check=True, env=GIT_ENV).stdout.strip()


class Repo:
    """A local repository whose files are written as exact bytes."""

    def __init__(self, d):
        self.d = Path(d)
        self.d.mkdir(parents=True, exist_ok=True)
        self.links: dict[str, str] = {}
        git(self.d, "init", "-q", "-b", "main")

    def write(self, rel: str, data: bytes) -> None:
        p = self.d / rel
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_bytes(data)

    def symlink(self, rel: str, target: str) -> None:
        """Commit a symlink without touching the filesystem (portable on Windows)."""
        self.links[rel] = target

    def commit(self, msg="c") -> str:
        git(self.d, "add", "-A")
        for rel, target in self.links.items():
            obj = subprocess.run(["git", "-C", str(self.d), "hash-object", "-w", "--stdin"],
                                 input=target, capture_output=True, text=True, check=True,
                                 env=GIT_ENV).stdout.strip()
            git(self.d, "update-index", "--add", "--cacheinfo", f"120000,{obj},{rel}")
        git(self.d, "commit", "-q", "--allow-empty", "-m", msg)
        return git(self.d, "rev-parse", "HEAD")

    def tag(self, name, ref="HEAD") -> None:
        git(self.d, "tag", "-a", name, "-m", f"release {name}", ref)


def write_config(path: Path, **over) -> Path:
    cfg = {"tree": "www", "image_path": "/var/www/html", "exclude": [],
           "image": "registry.example.test/team/app", "github_repo": "org/app"}
    cfg.update(over)
    path.write_text(json.dumps(cfg), encoding="utf-8")
    return path


def _tar_bytes(members: list[tuple[tarfile.TarInfo, bytes | None]]) -> bytes:
    buf = io.BytesIO()
    with tarfile.open(fileobj=buf, mode="w") as tar:
        for info, data in members:
            tar.addfile(info, io.BytesIO(data) if data is not None else None)
    return buf.getvalue()


def layer(files: dict[str, bytes | str | None]) -> bytes:
    """A layer tar: bytes = regular file, ``"->target"`` = symlink, None = directory."""
    members = []
    for name, data in files.items():
        info = tarfile.TarInfo(name)
        if data is None:
            info.type = tarfile.DIRTYPE
        elif isinstance(data, str):
            info.type, info.linkname = tarfile.SYMTYPE, data[2:]
            data = None
        else:
            info.size = len(data)
        members.append((info, data))
    return _tar_bytes(members)


def docker_archive(path: Path, layers: list[bytes], labels: dict | None) -> Path:
    """A ``docker save`` / ``podman save`` (docker-archive) image tarball."""
    config = json.dumps({"config": {"Labels": labels}}).encode()
    members = [(_info("cfg.json", config), config)]
    names = []
    for i, data in enumerate(layers):
        names.append(f"l{i}/layer.tar")
        members.append((_info(names[-1], data), data))
    manifest = json.dumps([{"Config": "cfg.json", "Layers": names}]).encode()
    members.append((_info("manifest.json", manifest), manifest))
    path.write_bytes(_tar_bytes(members))
    return path


def oci_archive(path: Path, layers: list[bytes], labels: dict | None) -> Path:
    """An ``oci-archive`` image tarball (OCI image layout)."""
    blobs = {}

    def blob(data: bytes) -> str:
        digest = "sha256:" + hashlib.sha256(data).hexdigest()
        blobs[digest] = data
        return digest

    config = blob(json.dumps({"config": {"Labels": labels}}).encode())
    manifest = blob(json.dumps({"config": {"digest": config},
                                "layers": [{"digest": blob(d)} for d in layers]}).encode())
    index = json.dumps({"manifests": [{"digest": manifest}]}).encode()
    members = [(_info("index.json", index), index)]
    for digest, data in blobs.items():
        members.append((_info("blobs/sha256/" + digest.split(":")[1], data), data))
    path.write_bytes(_tar_bytes(members))
    return path


def _info(name: str, data: bytes) -> tarfile.TarInfo:
    info = tarfile.TarInfo(name)
    info.size = len(data)
    return info
