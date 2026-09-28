"""Portable fingerprints for shared rules and the exact analysis input."""

import hashlib
import json
import subprocess
from pathlib import Path


def provenance(rows, threshold, root=None, mode="wall-openings", discipline=None):
    root = Path(root or Path(__file__).resolve().parents[1]).resolve()
    files = [root / "VERSION"]
    for directory in ("src", "rules", "knowledge", "skills"):
        files.extend(p for p in (root / directory).rglob("*")
                     if p.is_file() and p.suffix in {".py", ".md", ".yaml"})
    digest = hashlib.sha256()
    for path in sorted(files, key=lambda p: p.relative_to(root).as_posix()):
        relative = path.relative_to(root).as_posix()
        content = path.read_text(encoding="utf-8-sig").replace("\r\n", "\n")
        digest.update(relative.encode("utf-8") + b"\0" + content.encode("utf-8") + b"\0")

    def git(*args):
        result = subprocess.run(["git", "-C", str(root), *args], capture_output=True,
                                text=True, timeout=10, check=True)
        return result.stdout.strip()

    commit = dirty = None
    try:
        if root.samefile(Path(git("rev-parse", "--show-toplevel"))):
            dirty = bool(git("status", "--porcelain", "--untracked-files=all"))
            commit = git("rev-parse", "HEAD")
    except (OSError, subprocess.SubprocessError):
        pass

    input_bytes = json.dumps(rows, ensure_ascii=True, sort_keys=True,
                            separators=(",", ":"), allow_nan=False).encode("utf-8")
    return {
        "ruleset_version": (root / "VERSION").read_text(encoding="utf-8").strip(),
        "ruleset_commit": commit,
        "ruleset_dirty": dirty,
        "ruleset_sha256": digest.hexdigest(),
        "input_sha256": hashlib.sha256(input_bytes).hexdigest(),
        "parameters": {"mode": mode, "opening_min_edge_mm": threshold,
                       "mep_discipline": discipline},
    }
