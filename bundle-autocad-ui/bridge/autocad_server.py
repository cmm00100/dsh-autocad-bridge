"""
autocad_server.py — AutoCAD MCP Server（DSH ↔ AutoCAD 桥）

这一层只做一件事：把 cad_com.CadBridge 的能力包装成 MCP 工具，
让 DSH 里的 Agent 能通过 MCP 协议调用。

设计要点：
  1. 一个全局单例 bridge —— MCP 是长驻进程，没必要每次调用都重连 COM；
  2. 所有工具都返回 dict，出错也返回 {"ok": false, "error": "..."}，
     不让异常穿透到 MCP 层（否则 Agent 只看到一句 "Tool execution failed"）；
  3. 工具名用 mcp__autocad__<name> 形式，在 DSH 里就是这个名字；
  4. stdio 模式下所有日志必须走 stderr —— stdout 是 MCP 协议通道，写脏话就断连。

启动：
  .venv\\Scripts\\python.exe autocad_server.py            # stdio（DSH 用这个）
  set MCP_TRANSPORT=http && .venv\\Scripts\\python.exe autocad_server.py   # http 调试
"""
from __future__ import annotations

import sys
import traceback
from typing import Any

from mcp.server.mcpserver import MCPServer

from cad_com import CadBridge, CadError
from config import Settings

settings = Settings.load()

# ---------------------------------------------------------------------------
# 日志：一律 stderr
# ---------------------------------------------------------------------------

def log(msg: str) -> None:
    print(f"[autocad-bridge] {msg}", file=sys.stderr, flush=True)


# ---------------------------------------------------------------------------
# 单例桥
# ---------------------------------------------------------------------------

_bridge: CadBridge | None = None


def bridge() -> CadBridge:
    """懒连接：第一次调用工具时才连 CAD，避免 MCP 启动就被 CAD 阻塞。"""
    global _bridge
    if _bridge is None:
        _bridge = CadBridge(
            prog_id=settings.prog_id,
            timeout_ms=settings.timeout_ms,
            retry_times=settings.retry_times,
            retry_delay_ms=settings.retry_delay_ms,
        )
        log(f"创建桥实例 prog_id={settings.prog_id}")
    return _bridge


def _ensure_connected() -> CadBridge:
    b = bridge()
    if not b.is_connected():
        log("尚未连接，尝试 connect…")
        b.connect()
    return b


def _fail(exc: Exception, action: str) -> dict[str, Any]:
    log(f"{action} 失败：{exc}")
    traceback.print_exc(file=sys.stderr)
    return {"ok": False, "action": action, "error": str(exc)}


# ---------------------------------------------------------------------------
# MCP Server
# ---------------------------------------------------------------------------

mcp = MCPServer(
    name="autocad",
    title="AutoCAD 2024 Bridge",
    description="通过 COM 操控本机 AutoCAD 2024：画线/圆/矩形/多段线/文字/圆弧、图层管理、发送原生命令、读取图元、保存 DWG。",
    instructions=(
        "这是连接本机 AutoCAD 2024 的桥。首次调用任何绘图工具时会自动连接（或启动）AutoCAD。\n"
        "约定：\n"
        "- 所有坐标都是 WCS 世界坐标，单位与图纸一致（一般为毫米）。\n"
        "- 图层不存在会自动创建；layer 参数可省略，省略则用 CAD 当前图层。\n"
        "- AutoCAD 同一时刻只能执行一条命令，遇忙会自动重试；仍失败请让用户在 CAD 里按 ESC 关闭弹窗。\n"
        "- send_command 是异步发送，返回 sent=true 只代表已发出，不代表已执行完。\n"
        "- 有多个 DWG 打开时，用 doc_name 指定目标文档，否则操作当前活动文档。"
    ),
)


# ---- 连接与状态 -----------------------------------------------------------

@mcp.tool(name="cad_connect", description="连接本机 AutoCAD 2024（未运行时通过 COM 启动）。返回版本、已打开文档列表。")
def cad_connect() -> dict[str, Any]:
    try:
        return _ensure_connected().connect()
    except Exception as exc:  # noqa: BLE001
        return _fail(exc, "cad_connect")


@mcp.tool(name="cad_status", description="查询 AutoCAD 状态：disconnected / idle / busy / error，含活动文档与版本。")
def cad_status() -> dict[str, Any]:
    try:
        return bridge().status()
    except Exception as exc:  # noqa: BLE001
        return _fail(exc, "cad_status")


@mcp.tool(name="cad_escape", description="尝试取消 CAD 当前命令并关闭模态弹窗（PostCommand 发 ESC）。CAD 卡死时用。")
def cad_escape() -> dict[str, Any]:
    try:
        return bridge().send_escape()
    except Exception as exc:  # noqa: BLE001
        return _fail(exc, "cad_escape")


@mcp.tool(name="cad_list_progids", description="列出本机注册表中真实存在的 AutoCAD ProgID，连不上时用来确认版本号。")
def cad_list_progids() -> dict[str, Any]:
    return {"ok": True, "prog_ids": CadBridge.list_available_prog_ids(), "current": settings.prog_id}


# ---- 文档 -----------------------------------------------------------------

@mcp.tool(name="cad_list_documents", description="列出当前打开的所有 DWG 文档，标出哪个是活动文档。")
def cad_list_documents() -> dict[str, Any]:
    try:
        return {"ok": True, "documents": _ensure_connected().list_documents()}
    except Exception as exc:  # noqa: BLE001
        return _fail(exc, "cad_list_documents")


@mcp.tool(name="cad_new_document", description="新建一张空白图纸。template 为可选的 dwt 模板绝对路径。")
def cad_new_document(template: str = "") -> dict[str, Any]:
    try:
        return _ensure_connected().new_document(template or None)
    except Exception as exc:  # noqa: BLE001
        return _fail(exc, "cad_new_document")


@mcp.tool(name="cad_save", description="保存当前图纸。不给 path 则就地保存；给 path 则另存为（自动补 .dwg）。")
def cad_save(path: str = "", doc_name: str = "") -> dict[str, Any]:
    try:
        return _ensure_connected().save(path or None, doc_name or None)
    except Exception as exc:  # noqa: BLE001
        return _fail(exc, "cad_save")


# ---- 图层 -----------------------------------------------------------------

@mcp.tool(name="cad_ensure_layer", description="创建图层（已存在则复用）。color 支持颜色名 red/yellow/green/cyan/blue/magenta/white 或 ACI 索引 1-255。")
def cad_ensure_layer(name: str, color: str = "", doc_name: str = "") -> dict[str, Any]:
    try:
        return _ensure_connected().ensure_layer(name, color or None, doc_name or None)
    except Exception as exc:  # noqa: BLE001
        return _fail(exc, "cad_ensure_layer")


@mcp.tool(name="cad_list_layers", description="列出图纸中的所有图层及颜色索引。")
def cad_list_layers(doc_name: str = "") -> dict[str, Any]:
    try:
        return {"ok": True, "layers": _ensure_connected().list_layers(doc_name or None)}
    except Exception as exc:  # noqa: BLE001
        return _fail(exc, "cad_list_layers")


# ---- 绘图 -----------------------------------------------------------------

@mcp.tool(name="cad_add_line", description="画直线。start/end 为 [x,y] 或 [x,y,z]，单位与图纸一致（一般毫米）。")
def cad_add_line(start: list[float], end: list[float], layer: str = "", doc_name: str = "") -> dict[str, Any]:
    try:
        return _ensure_connected().add_line(start, end, layer or None, doc_name or None)
    except Exception as exc:  # noqa: BLE001
        return _fail(exc, "cad_add_line")


@mcp.tool(name="cad_add_circle", description="画圆。center 为 [x,y]，radius 为半径（正数）。")
def cad_add_circle(center: list[float], radius: float, layer: str = "", doc_name: str = "") -> dict[str, Any]:
    try:
        return _ensure_connected().add_circle(center, radius, layer or None, doc_name or None)
    except Exception as exc:  # noqa: BLE001
        return _fail(exc, "cad_add_circle")


@mcp.tool(name="cad_add_rectangle", description="画矩形。corner1/corner2 为对角两个点 [x,y]。")
def cad_add_rectangle(corner1: list[float], corner2: list[float], layer: str = "", doc_name: str = "") -> dict[str, Any]:
    try:
        return _ensure_connected().add_rectangle(corner1, corner2, layer or None, doc_name or None)
    except Exception as exc:  # noqa: BLE001
        return _fail(exc, "cad_add_rectangle")


@mcp.tool(name="cad_add_polyline", description="画多段线。points 为 [[x,y],...] 至少两点，closed=true 闭合。")
def cad_add_polyline(points: list[list[float]], closed: bool = False, layer: str = "", doc_name: str = "") -> dict[str, Any]:
    try:
        return _ensure_connected().add_polyline(points, closed, layer or None, doc_name or None)
    except Exception as exc:  # noqa: BLE001
        return _fail(exc, "cad_add_polyline")


@mcp.tool(name="cad_add_text", description="写单行文字。position 为插入点，height 为字高（默认 2.5）。")
def cad_add_text(text: str, position: list[float], height: float = 2.5, layer: str = "", doc_name: str = "") -> dict[str, Any]:
    try:
        return _ensure_connected().add_text(text, position, height, layer or None, doc_name or None)
    except Exception as exc:  # noqa: BLE001
        return _fail(exc, "cad_add_text")


@mcp.tool(name="cad_add_arc", description="画圆弧。角度用度（不是弧度），从 start_angle_deg 逆时针到 end_angle_deg。")
def cad_add_arc(
    center: list[float],
    radius: float,
    start_angle_deg: float,
    end_angle_deg: float,
    layer: str = "",
    doc_name: str = "",
) -> dict[str, Any]:
    try:
        return _ensure_connected().add_arc(center, radius, start_angle_deg, end_angle_deg, layer or None, doc_name or None)
    except Exception as exc:  # noqa: BLE001
        return _fail(exc, "cad_add_arc")


# ---- 原生命令 / 读取 / 视图 -----------------------------------------------

@mcp.tool(name="cad_send_command", description="向 AutoCAD 命令行发送任意原生命令（如 '_.ZOOM\\n_E\\n'）。异步发送，不等待执行完成。")
def cad_send_command(command: str, doc_name: str = "") -> dict[str, Any]:
    try:
        return _ensure_connected().send_command(command, doc_name or None)
    except Exception as exc:  # noqa: BLE001
        return _fail(exc, "cad_send_command")


@mcp.tool(name="cad_list_entities", description="列出模型空间图元摘要（类型/图层/句柄），用于了解当前图纸里已有什么。")
def cad_list_entities(limit: int = 200, doc_name: str = "") -> dict[str, Any]:
    try:
        return _ensure_connected().list_entities(doc_name or None, limit)
    except Exception as exc:  # noqa: BLE001
        return _fail(exc, "cad_list_entities")


@mcp.tool(name="cad_zoom_extents", description="缩放到全图范围，让用户立刻看到刚绘制的内容。")
def cad_zoom_extents(doc_name: str = "") -> dict[str, Any]:
    try:
        return _ensure_connected().zoom_extents(doc_name or None)
    except Exception as exc:  # noqa: BLE001
        return _fail(exc, "cad_zoom_extents")


# ---------------------------------------------------------------------------
# 入口
# ---------------------------------------------------------------------------

def main() -> None:
    transport = settings.transport
    log(f"启动 AutoCAD MCP Server | transport={transport} | prog_id={settings.prog_id}")
    if transport == "http":
        log(f"HTTP 模式：http://{settings.http_host}:{settings.http_port}/mcp")
        mcp.run(transport="streamable-http", host=settings.http_host, port=settings.http_port)
    else:
        mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
