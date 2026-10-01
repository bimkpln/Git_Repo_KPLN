"""Register only mcp_servers.revit in Codex, preserving unrelated configuration."""
import argparse
from datetime import datetime
import json
import os
from pathlib import Path
import subprocess
import tomllib


def install(config: Path, python: Path, server: Path):
    if not python.is_file() or not server.is_file():
        raise ValueError("Python executable or MCP server file not found")
    subprocess.run([str(python), "-B", "-c", "import mcp, httpx"], check=True)
    original = config.read_text(encoding="utf-8-sig") if config.exists() else ""
    parsed = tomllib.loads(original)
    desired = {"command": str(python), "args": ["-B", str(server)]}
    existing = parsed.get("mcp_servers", {}).get("revit")
    if existing is not None:
        if existing.get("command") == desired["command"] and existing.get("args") == desired["args"]:
            print("Already registered:", config)
            return
        raise ValueError("mcp_servers.revit already exists with another command; review it before replacing")
    quote = lambda value: json.dumps(value, ensure_ascii=False)
    updated = original.rstrip() + "\n\n[mcp_servers.revit]\ncommand = " + quote(str(python)) + "\nargs = [\"-B\", " + quote(str(server)) + "]\nstartup_timeout_sec = 20\ntool_timeout_sec = 60\n"
    tomllib.loads(updated)
    config.parent.mkdir(parents=True, exist_ok=True)
    if config.exists():
        backup = config.with_name(config.name + ".backup-revit-" + datetime.now().strftime("%Y%m%d-%H%M%S-%f"))
        backup.write_bytes(config.read_bytes())
        print("Configuration backup:", backup)
        if config.read_text(encoding="utf-8-sig") != original:
            raise RuntimeError("Configuration changed during installation; retry after checking it")
    temp = config.with_name(config.name + ".revit.tmp")
    with temp.open("x", encoding="utf-8", newline="\n") as stream: stream.write(updated)
    os.replace(temp, config)
    print("Registered MCP revit:", config)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--python", type=Path, required=True, help="Python 3.11+ environment with requirements.txt installed")
    parser.add_argument("--server", type=Path, default=Path(__file__).resolve().parents[1] / "server/revit_mcp_server.py")
    parser.add_argument("--config", type=Path, default=Path(os.environ.get("CODEX_HOME", str(Path.home() / ".codex"))) / "config.toml")
    args = parser.parse_args()
    install(args.config.resolve(), args.python.resolve(), args.server.resolve())


if __name__ == "__main__": main()
