"""Local transport only. Engineering rules live in KPLN_RevitRules."""
from __future__ import annotations

import json
import os
import re
import uuid
from pathlib import Path
from urllib.parse import urlsplit

import httpx


class BridgeError(RuntimeError):
    pass


class BridgeClient:
    def __init__(self, sessions_dir: Path | None = None, transport=None):
        self.sessions_dir = sessions_dir or Path(os.environ.get("LOCALAPPDATA", str(Path.home() / "AppData/Local"))) / "KPLN/RevitMcpBridge/sessions"
        self.transport = transport

    def _descriptor(self, path: Path) -> dict:
        if path.stat().st_size > 8192:
            raise ValueError("Oversized session descriptor")
        data = json.loads(path.read_text(encoding="utf-8-sig"))
        session_id = data["session_id"]
        if not isinstance(session_id, str) or not re.fullmatch(r"[a-f0-9]{32}", session_id) or path.stem != session_id:
            raise ValueError("Invalid session id")
        url = urlsplit(data["url"])
        if url.scheme != "http" or url.hostname != "127.0.0.1" or url.username or url.password or url.path not in ("", "/") or url.query or url.fragment or not (18765 <= (url.port or 0) < 18785):
            raise ValueError("Session URL must be a local Revit bridge")
        if data.get("protocol") != 1 or not isinstance(data.get("token"), str) or len(data["token"]) != 44:
            raise ValueError("Invalid session protocol/token")
        return data

    def _request(self, session: dict, method: str, path: str, payload=None, timeout=40) -> dict:
        # Disable proxy environment variables and redirects: credentials stay on loopback.
        with httpx.Client(transport=self.transport, trust_env=False, follow_redirects=False, timeout=timeout) as client:
            try:
                response = client.request(method, session["url"] + path, json=payload, headers={"Authorization": "Bearer " + session["token"]})
                if response.is_redirect:
                    raise BridgeError("Локальный мост неожиданно вернул перенаправление.")
                data = response.json()
                if response.is_error or not isinstance(data, dict) or data.get("ok") is not True:
                    error = data.get("error", {}) if isinstance(data, dict) else {}
                    raise BridgeError(f'{error.get("code", response.status_code)}: {error.get("message", "Ошибка ответа Revit")}')
                return data
            except httpx.TimeoutException as exc:
                operation_id = payload.get("operation_id") if payload else None
                raise BridgeError(f"Таймаут транспорта. Результат не известен; не повторяйте запись. get_operation: session_id={session['session_id']}, operation_id={operation_id}") from exc
            except httpx.RequestError as exc:
                raise BridgeError("Мост Revit недоступен. Проверьте кнопку KPLN → BIM → Codex Мост. При потере связи во время записи проверьте модель перед повтором.") from exc
            except (ValueError, UnicodeError) as exc:
                raise BridgeError("Мост вернул некорректный JSON.") from exc

    def sessions(self) -> dict:
        live, unavailable = [], []
        for path in sorted(self.sessions_dir.glob("*.json")):
            try:
                session = self._descriptor(path)
                health = self._request(session, "GET", "/health", timeout=1.5)
                if health.get("session_id") != session["session_id"] or health.get("protocol") != 1:
                    raise BridgeError("Сессия по этому адресу сменилась.")
                live.append({key: session[key] for key in ("session_id", "revit_version", "pid", "url")})
            except (OSError, ValueError, KeyError, TypeError, BridgeError) as exc:
                unavailable.append({"session_file": path.name, "reason": str(exc)})
        return {"sessions": live, "unavailable": unavailable}

    def _session(self, session_id: str | None) -> dict:
        if session_id:
            if not re.fullmatch(r"[a-f0-9]{32}", session_id):
                raise BridgeError("Некорректный session_id; используйте list_revit_sessions.")
            try:
                return self._descriptor(self.sessions_dir / (session_id + ".json"))
            except (OSError, ValueError, KeyError, TypeError) as exc:
                raise BridgeError("Сессия не найдена; используйте list_revit_sessions.") from exc
        sessions = self.sessions()["sessions"]
        if len(sessions) != 1:
            raise BridgeError("Укажите session_id из list_revit_sessions: доступно сессий " + str(len(sessions)))
        return self._session(sessions[0]["session_id"])

    def health(self, session_id=None):
        session = self._session(session_id)
        result = self._request(session, "GET", "/health")
        if result.get("session_id") != session["session_id"]:
            raise BridgeError("Сессия Revit сменилась.")
        return result

    def command(self, command: str, session_id=None, **args):
        session = self._session(session_id)
        operation_id = uuid.uuid4().hex
        payload = dict(args, command=command, session_id=session["session_id"], operation_id=operation_id)
        try:
            result = self._request(session, "POST", "/command", payload)
        except BridgeError as exc:
            raise BridgeError(f"{exc}\nЗапрос: session_id={session['session_id']}, operation_id={operation_id}") from exc
        return dict(result, session_id=session["session_id"], operation_id=operation_id)

    def operation(self, operation_id: str, session_id: str):
        try:
            operation_id = uuid.UUID(operation_id).hex
        except ValueError as exc:
            raise BridgeError("operation_id должен быть UUID.") from exc
        return self._request(self._session(session_id), "GET", "/operations/" + operation_id)
