"""Create an offline evidence record; does not connect to or change Revit."""
import argparse
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
import subprocess

ROOT = Path(__file__).resolve().parents[1]


def digest(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def provenance(root: Path = ROOT) -> dict:
    files = [root / "VERSION"]
    for name in ("rules", "src", "knowledge", "skills"):
        files.extend(p for p in (root / name).rglob("*") if p.is_file() and p.suffix in (".md", ".py", ".yaml"))
    normalized = []
    for path in sorted(files):
        normalized.append(path.relative_to(root).as_posix().encode() + b"\0" + path.read_bytes().replace(b"\r\n", b"\n").replace(b"\r", b"\n") + b"\0")

    def git(*args):
        try:
            result = subprocess.run(["git", "-C", str(root), *args], capture_output=True, text=True, timeout=5, check=True)
            return result.stdout.strip()
        except (OSError, subprocess.SubprocessError):
            return None

    # Do not claim a parent workspace's Git HEAD as this knowledge repository.
    top = git("rev-parse", "--show-toplevel")
    own_repo = top is not None and Path(top).resolve() == root.resolve()
    status = git("status", "--porcelain") if own_repo else None
    return {"version": (root / "VERSION").read_text().strip(), "git_head": git("rev-parse", "HEAD") if own_repo else None,
            "git_dirty": None if status is None else bool(status), "fingerprint_sha256": digest(b"".join(normalized))}


def create_record(input_path: Path, context: dict, root: Path = ROOT) -> dict:
    raw = input_path.read_bytes()
    data = json.loads(raw)
    return {"schema_version": 1, "created_utc": datetime.now(timezone.utc).isoformat(), "knowledge": provenance(root),
            "input_sha256": digest(raw), "project_context": context, "evidence": data,
            "decisions": [], "overrides": [], "applied_operations": []}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path)
    parser.add_argument("--context", required=True, type=Path, help="JSON: project, requirements/sources, session_id, document_id, revision")
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args()
    context = json.loads(args.context.read_text(encoding="utf-8-sig"))
    if not isinstance(context, dict): parser.error("Context must be a JSON object")
    result = create_record(args.input, context)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    # Never silently replace a previous review record.
    with args.output.open("x", encoding="utf-8") as stream:
        json.dump(result, stream, ensure_ascii=False, indent=2, allow_nan=False)


if __name__ == "__main__": main()
