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
        # V1.1-2：配置里的默认图层（.env 的 CAD_DEFAULT_LAYER）作为图元自动归属兜底
        if settings.default_layer:
            _bridge.default_layer = settings.default_layer
            log(f"默认图层 = {settings.default_layer}")
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
    description="通过原生 COM 操控本机 AutoCAD 2024：绘图图元、机械标准图层自动化、国标尺寸标注体系、块与属性块、文字/MText/形位公差、命令发送、读取图元、保存 DWG。",
    instructions=(
        "这是连接本机 AutoCAD 2024 的桥。首次调用任何绘图工具时会自动连接（或启动）AutoCAD。\n"
        "约定：\n"
        "- 所有坐标都是 WCS 世界坐标，单位与图纸一致（一般为毫米）。角度参数统一用度（不是弧度）。\n"
        "- 图层不存在会自动创建；layer 省略时用 CAD 当前图层。调用 cad_setup_mechanical_environment\n"
        "  后，省略 layer 的图元会按大类自动归属到 轮廓/中心线/虚线/标注/文字/剖面线 标准图层。\n"
        "- 画正规机械零件图推荐顺序：cad_setup_mechanical_environment → 画轮廓 → 加标注 → cad_zoom_extents。\n"
        "- AutoCAD 同一时刻只能执行一条命令，遇忙会自动重试；仍失败请让用户在 CAD 里按 ESC 关闭弹窗，\n"
        "  或调用 cad_reconnect 重建 COM 连接。\n"
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


@mcp.tool(name="cad_escape", description="尝试取消 CAD 当前命令并关闭模态弹窗（文档 PostCommand 发 ESC）。CAD 卡死时用。")
def cad_escape(doc_name: str = "") -> dict[str, Any]:
    try:
        return _ensure_connected().send_escape(doc_name or None)
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
# V1.1-5 修复：超时重连
# ---------------------------------------------------------------------------

@mcp.tool(name="cad_reconnect", description="重建 COM 连接（丢弃卡死的单 STA 线程并换新的）。队列长期积压、模态弹窗锁死、或 CAD 重启后连不上时用。")
def cad_reconnect() -> dict[str, Any]:
    try:
        return bridge().reconnect()
    except Exception as exc:  # noqa: BLE001
        return _fail(exc, "cad_reconnect")


# ---------------------------------------------------------------------------
# V1.1-1 完整尺寸标注体系 + 国标标注样式
# ---------------------------------------------------------------------------

@mcp.tool(name="cad_setup_text_style", description="建/复用中文工程文字样式（默认 gbenor.shx + gbcbig.shx），否则中文显示成 ???。")
def cad_setup_text_style(name: str = "DSH-CN", font: str = "gbenor.shx", bigfont: str = "gbcbig.shx", doc_name: str = "") -> dict[str, Any]:
    try:
        return _ensure_connected().setup_text_style(name, font, bigfont, doc_name or None)
    except Exception as exc:  # noqa: BLE001
        return _fail(exc, "cad_setup_text_style")


@mcp.tool(name="cad_setup_gb_dim_style", description="建立并激活国标标注样式（GB-机械）。dim_vars 可覆盖单项 DIM* 变量。顺序已实测：先激活样式再写变量，最后 CopyFrom 固化。")
def cad_setup_gb_dim_style(text_style: str = "DSH-CN", dim_style: str = "GB-机械", dim_vars: dict[str, Any] | None = None, doc_name: str = "") -> dict[str, Any]:
    try:
        return _ensure_connected().setup_gb_dim_style(text_style, dim_style, dim_vars, doc_name or None)
    except Exception as exc:  # noqa: BLE001
        return _fail(exc, "cad_setup_gb_dim_style")


@mcp.tool(name="cad_add_dim_linear", description="线性标注。p1/p2 为两个测量点，dim_line_point 是尺寸线经过的点，angle_deg 是尺寸线角度（0=水平、90=竖直）。text_override 可改标注文字。")
def cad_add_dim_linear(p1: list[float], p2: list[float], dim_line_point: list[float], angle_deg: float = 0.0, text_override: str = "", layer: str = "", doc_name: str = "") -> dict[str, Any]:
    try:
        return _ensure_connected().add_dim_linear(p1, p2, dim_line_point, angle_deg, text_override, layer or None, doc_name or None)
    except Exception as exc:  # noqa: BLE001
        return _fail(exc, "cad_add_dim_linear")


@mcp.tool(name="cad_add_dim_aligned", description="对齐标注：尺寸线平行于 p1-p2 连线，用于标注斜边长度。")
def cad_add_dim_aligned(p1: list[float], p2: list[float], dim_line_point: list[float], text_override: str = "", layer: str = "", doc_name: str = "") -> dict[str, Any]:
    try:
        return _ensure_connected().add_dim_aligned(p1, p2, dim_line_point, text_override, layer or None, doc_name or None)
    except Exception as exc:  # noqa: BLE001
        return _fail(exc, "cad_add_dim_aligned")


@mcp.tool(name="cad_add_dim_radial", description="半径标注。center 为圆心，chord_point 为圆周取点，leader_length 为引线长度。")
def cad_add_dim_radial(center: list[float], chord_point: list[float], leader_length: float = 10.0, text_override: str = "", layer: str = "", doc_name: str = "") -> dict[str, Any]:
    try:
        return _ensure_connected().add_dim_radial(center, chord_point, leader_length, text_override, layer or None, doc_name or None)
    except Exception as exc:  # noqa: BLE001
        return _fail(exc, "cad_add_dim_radial")


@mcp.tool(name="cad_add_dim_diametric", description="直径标注。chord1/chord2 是直径两端点，leader_length 为引线长度。")
def cad_add_dim_diametric(chord1: list[float], chord2: list[float], leader_length: float = 10.0, text_override: str = "", layer: str = "", doc_name: str = "") -> dict[str, Any]:
    try:
        return _ensure_connected().add_dim_diametric(chord1, chord2, leader_length, text_override, layer or None, doc_name or None)
    except Exception as exc:  # noqa: BLE001
        return _fail(exc, "cad_add_dim_diametric")


@mcp.tool(name="cad_add_dim_angular", description="角度标注。vertex 是角顶点，p1/p2 分别在两条边上，text_point 是文字位置。")
def cad_add_dim_angular(vertex: list[float], p1: list[float], p2: list[float], text_point: list[float], text_override: str = "", layer: str = "", doc_name: str = "") -> dict[str, Any]:
    try:
        return _ensure_connected().add_dim_angular(vertex, p1, p2, text_point, text_override, layer or None, doc_name or None)
    except Exception as exc:  # noqa: BLE001
        return _fail(exc, "cad_add_dim_angular")


@mcp.tool(name="cad_add_dim_ordinate", description="坐标标注。use_x_axis=True 标 X 坐标，False 标 Y 坐标。definition_point 是被标点，leader_end_point 是引线端点。")
def cad_add_dim_ordinate(definition_point: list[float], leader_end_point: list[float], use_x_axis: bool = True, text_override: str = "", layer: str = "", doc_name: str = "") -> dict[str, Any]:
    try:
        return _ensure_connected().add_dim_ordinate(definition_point, leader_end_point, use_x_axis, text_override, layer or None, doc_name or None)
    except Exception as exc:  # noqa: BLE001
        return _fail(exc, "cad_add_dim_ordinate")


# ---------------------------------------------------------------------------
# V1.1-2 机械标准图层自动化
# ---------------------------------------------------------------------------

@mcp.tool(name="cad_setup_standard_layers", description="一次性建好机械标准图层：轮廓/中心线/虚线/标注/文字/剖面线（含颜色、线宽、线型）。幂等。")
def cad_setup_standard_layers(doc_name: str = "") -> dict[str, Any]:
    try:
        return _ensure_connected().setup_standard_layers(doc_name or None)
    except Exception as exc:  # noqa: BLE001
        return _fail(exc, "cad_setup_standard_layers")


@mcp.tool(name="cad_set_layer_state", description="批量改图层状态：锁定 lock / 冻结 freeze / 开关 layer_on / 颜色 color / 线宽 lineweight / 设为当前层 active。不传的项保持原样。")
def cad_set_layer_state(
    name: str,
    lock: bool | None = None,
    freeze: bool | None = None,
    layer_on: bool | None = None,
    color: str = "",
    lineweight: int | None = None,
    active: bool = False,
    doc_name: str = "",
) -> dict[str, Any]:
    try:
        return _ensure_connected().set_layer_state(name, lock, freeze, layer_on, color or None, lineweight, active, doc_name or None)
    except Exception as exc:  # noqa: BLE001
        return _fail(exc, "cad_set_layer_state")


@mcp.tool(name="cad_set_active_layer", description="把某个图层设为当前层。")
def cad_set_active_layer(name: str, doc_name: str = "") -> dict[str, Any]:
    try:
        return _ensure_connected().set_active_layer(name, doc_name or None)
    except Exception as exc:  # noqa: BLE001
        return _fail(exc, "cad_set_active_layer")


@mcp.tool(name="cad_setup_mechanical_environment", description="一句话完成分层规范：一次调用建好全部标准图层 + 中文文字样式 + 国标标注样式 + mm/线宽/关联标注文档变量，并开启图元按大类自动归属图层。画正规机械零件图前先调它。")
def cad_setup_mechanical_environment(doc_name: str = "") -> dict[str, Any]:
    try:
        return _ensure_connected().setup_mechanical_environment(doc_name or None)
    except Exception as exc:  # noqa: BLE001
        return _fail(exc, "cad_setup_mechanical_environment")


# ---------------------------------------------------------------------------
# V1.1-3 块与属性块
# ---------------------------------------------------------------------------

@mcp.tool(name="cad_list_blocks", description="列出图纸中已定义的块名（过滤内建 *Model_Space / *Paper_Space）。")
def cad_list_blocks(doc_name: str = "") -> dict[str, Any]:
    try:
        return _ensure_connected().list_blocks(doc_name or None)
    except Exception as exc:  # noqa: BLE001
        return _fail(exc, "cad_list_blocks")


@mcp.tool(name="cad_create_block", description="新建块定义。entities 是 [{'type': 'line'|'circle'|'arc'|'text'|'polyline', ...}]；attributes 是 [{'tag','value','prompt','height','position','mode'}]。块已存在则直接返回不覆盖。")
def cad_create_block(
    name: str,
    base_point: list[float] | None = None,
    entities: list[dict[str, Any]] | None = None,
    attributes: list[dict[str, Any]] | None = None,
    doc_name: str = "",
) -> dict[str, Any]:
    try:
        return _ensure_connected().create_block(name, base_point or [0.0, 0.0], entities, attributes, doc_name or None)
    except Exception as exc:  # noqa: BLE001
        return _fail(exc, "cad_create_block")


@mcp.tool(name="cad_insert_block", description="把块插入模型空间，可按 tag 覆盖属性值。position 是插入点，rotation_deg 是旋转角（度），attributes 是 {'tag':'新值'}。返回块参照句柄。")
def cad_insert_block(
    name: str,
    position: list[float],
    xscale: float = 1.0,
    yscale: float = 1.0,
    zscale: float = 1.0,
    rotation_deg: float = 0.0,
    attributes: dict[str, str] | None = None,
    layer: str = "",
    doc_name: str = "",
) -> dict[str, Any]:
    try:
        return _ensure_connected().insert_block(name, position, xscale, yscale, zscale, rotation_deg, attributes, layer or None, doc_name or None)
    except Exception as exc:  # noqa: BLE001
        return _fail(exc, "cad_insert_block")


@mcp.tool(name="cad_list_block_attributes", description="读某个块参照的全部属性 (tag, value, prompt)。handle 由 cad_insert_block 返回。")
def cad_list_block_attributes(handle: str, doc_name: str = "") -> dict[str, Any]:
    try:
        return _ensure_connected().list_block_attributes(handle, doc_name or None)
    except Exception as exc:  # noqa: BLE001
        return _fail(exc, "cad_list_block_attributes")


@mcp.tool(name="cad_set_block_attribute", description="改某个块参照里指定 tag 的属性值。")
def cad_set_block_attribute(handle: str, tag: str, value: str, doc_name: str = "") -> dict[str, Any]:
    try:
        return _ensure_connected().set_block_attribute(handle, tag, value, doc_name or None)
    except Exception as exc:  # noqa: BLE001
        return _fail(exc, "cad_set_block_attribute")


@mcp.tool(name="cad_replace_block", description="批量替换：把所有 old_name 的块参照换成 new_name，保留插入点/比例/旋转。目标块必须已定义。")
def cad_replace_block(old_name: str, new_name: str, doc_name: str = "") -> dict[str, Any]:
    try:
        return _ensure_connected().replace_block(old_name, new_name, doc_name or None)
    except Exception as exc:  # noqa: BLE001
        return _fail(exc, "cad_replace_block")


# ---------------------------------------------------------------------------
# V1.1-4 文字 / MText / 公差
# ---------------------------------------------------------------------------

@mcp.tool(name="cad_add_mtext", description="写多行文字（MText）。\\P 换行；配合 DSH-CN 样式可写中文技术要求。width 是文本框宽度，attachment_point 是对齐点（1=左上）。")
def cad_add_mtext(text: str, position: list[float], width: float = 100.0, height: float = 3.5, style: str = "", attachment_point: int = 1, layer: str = "", doc_name: str = "") -> dict[str, Any]:
    try:
        return _ensure_connected().add_mtext(text, position, width, height, style, attachment_point, layer or None, doc_name or None)
    except Exception as exc:  # noqa: BLE001
        return _fail(exc, "cad_add_mtext")


@mcp.tool(name="cad_add_tolerance", description="写形位公差框格（AcDbFcf）。text 用 AutoCAD 控制码，如 %%c40%%p0.01；direction 是框格方向向量，默认水平。")
def cad_add_tolerance(text: str, position: list[float], direction: list[float] | None = None, layer: str = "", doc_name: str = "") -> dict[str, Any]:
    try:
        return _ensure_connected().add_tolerance(text, position, direction or [1.0, 0.0, 0.0], layer or None, doc_name or None)
    except Exception as exc:  # noqa: BLE001
        return _fail(exc, "cad_add_tolerance")


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
