"""
config.py — 读取 .env，产出 CadBridge / MCP 服务需要的参数。
"""
from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv

# .env 与本文件同目录
_ENV_PATH = Path(__file__).resolve().parent / ".env"
load_dotenv(_ENV_PATH)


def _int(name: str, default: int) -> int:
    raw = os.getenv(name, "").strip()
    if not raw:
        return default
    try:
        return int(raw)
    except ValueError:
        return default


def _bool(name: str, default: bool = False) -> bool:
    raw = os.getenv(name, "").strip().lower()
    if not raw:
        return default
    return raw in {"1", "true", "yes", "on"}


@dataclass(frozen=True)
class Settings:
    prog_id: str
    timeout_ms: int
    retry_times: int
    retry_delay_ms: int
    transport: str
    http_host: str
    http_port: int
    log_commands: bool
    default_layer: str

    @classmethod
    def load(cls) -> "Settings":
        return cls(
            prog_id=os.getenv("CAD_PROG_ID", "AutoCAD.Application.24").strip() or "AutoCAD.Application.24",
            timeout_ms=_int("CAD_TIMEOUT", 15000),
            retry_times=_int("CAD_RETRY_TIMES", 2),
            retry_delay_ms=_int("CAD_RETRY_DELAY", 800),
            transport=(os.getenv("MCP_TRANSPORT", "stdio").strip().lower() or "stdio"),
            http_host=os.getenv("MCP_HTTP_HOST", "127.0.0.1").strip() or "127.0.0.1",
            http_port=_int("MCP_HTTP_PORT", 8770),
            log_commands=_bool("CAD_LOG_COMMANDS", True),
            default_layer=(os.getenv("CAD_DEFAULT_LAYER", "").strip() or ""),
        )
