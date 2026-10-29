"""
cad_com.py — AutoCAD COM 桥接层（DSH x AutoCAD 插件的第一层）

职责：只负责和 AutoCAD 的 COM 接口说话，不做任何 MCP / HTTP 的事。
这样它可以被 test_bridge.py 单独测试，也可以被 autocad_server.py 复用。

核心概念（写代码前必须理解）：
1. AutoCAD 同一时刻只能执行一条命令。COM 调用撞上正在执行的命令时，
   会抛 com_error，hresult 通常是 RPC_E_SERVERCALL_RETRYLATER (0x8001010A)。
   → 所有对外方法都必须走 _call()，由它负责"忙检测 + 重试"。
2. AutoCAD 有 modal 弹窗（字体缺失、许可提示等）时，COM 会直接卡死。
   → 超时后要发 ESC（SENDKEYS）脱困，并明确报告"CAD 被弹窗阻塞"。
3. 多文档：ActiveDocument 会随用户切换而变化。
   → 支持按文档名指定目标文档；不指定则用当前活动文档，并在返回值里回报告是哪个文档。
4. 所有 COM 调用都要在同一个已初始化 COM 的线程里。
   → pythoncom.CoInitialize() 在 CadBridge 首次使用时调用。
"""
from __future__ import annotations

import math
import os
import queue
import threading
import time
from typing import Any, Iterable, Sequence

import pythoncom
import pywintypes
import win32com.client
from win32com.client import VARIANT, gencache

# ---------------------------------------------------------------------------
# 常量
# ---------------------------------------------------------------------------

#: AutoCAD 在忙碌时返回的 HRESULT
RPC_E_CALL_REJECTED = -2147418111  # 0x80010001
RPC_E_SERVERCALL_RETRYLATER = -2147417846  # 0x8001010A
RPC_E_DISCONNECTED = -2147417848  # 0x80010108

#: 常见的"CAD 忙"错误码集合
BUSY_HRESULTS = {
    RPC_E_CALL_REJECTED,
    RPC_E_SERVERCALL_RETRYLATER,
    -2147417851,  # RPC_E_SERVERFAULT
}

#: AutoCAD 的颜色常量（用于 COM 的 TrueColor / colorIndex）
ACAD_COLOR_INDEX = {
    "red": 1,
    "yellow": 2,
    "green": 3,
    "cyan": 4,
    "blue": 5,
    "magenta": 6,
    "white": 7,
    "gray": 8,
    "grey": 8,
}


def _hresult(exc: pywintypes.com_error) -> int:
    """pywin32 的 com_error 里 hresult 有时是正数有时是负数，统一成有符号 32 位。"""
    hr = getattr(exc, "hresult", None)
    if hr is None:
        hr = exc.args[0] if exc.args else 0
    hr = int(hr) & 0xFFFFFFFF
    if hr >= 0x80000000:
        hr -= 0x100000000
    return hr


class CadError(RuntimeError):
    """桥层统一异常。message 会原样返回给 Agent，所以要写人话。"""


class CadBusyError(CadError):
    """AutoCAD 忙（有命令在跑 / 有弹窗挡着）。"""


class CadDisconnectedError(CadError):
    """AutoCAD 没开 / COM 断开 / ProgID 不对。"""


# ---------------------------------------------------------------------------
# 小工具
# ---------------------------------------------------------------------------

def dbl_array(values: Sequence[float]) -> Any:
    """
    ★ 关键坑点 ★

    AutoCAD 的 AddLine/AddCircle/AddArc/AddLightWeightPolyline/AddText
    这些方法的坐标参数类型是 SAFEARRAY(double)。

    直接传 Python list 或 tuple（甚至 array.array('d')）都会得到：
        (-2147352567, '发生意外。', (0, None, None, None, 0, -2147024809), None)
    即 E_INVALIDARG 0x80070057 —— 报错信息里根本看不出是参数类型问题，
    极容易误判成"COM 权限"或"版本不对"。

    唯一可靠的传法是显式构造 VARIANT(VT_ARRAY | VT_R8)。
    """
    return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [float(v) for v in values])


def _pt(x: float, y: float, z: float = 0.0) -> Any:
    return dbl_array([x, y, z])


def _flat(values: Sequence[float]) -> Any:
    return dbl_array(list(values))


def _bbox_summary(entity: Any) -> dict[str, Any]:
    """尽量安全地取图元包围盒，取不到就返回空。"""
    try:
        mn, mx = entity.GetBoundingBox()
        return {"min": list(mn), "max": list(mx)}
    except Exception:
        return {}


# ---------------------------------------------------------------------------
# 专用 COM 线程
# ---------------------------------------------------------------------------

class _ComWorker:
    """
    全部 AutoCAD COM 调用都在这一个线程上执行。

    ★ 为什么需要它 ★

    MCP 服务是异步的，每次 tools/call 落在哪个线程完全随机。而：
      1) COM 的 CoInitialize 是【按线程】的，新线程不初始化就报
         "尚未调用 CoInitialize"；
      2) 就算每个线程都 CoInitialize 了，从线程 B 使用在线程 A 创建的
         COM 指针（AutoCAD 的 Application 对象）依然是未 marshal 的
         跨套间引用，早晚出 RPC_E_WRONG_THREAD / 随机崩溃；
      3) AutoCAD 本身是单线程的，并发调用只会互相打架。

    三者叠加后唯一稳的模型：一个常驻 STA 线程 + 一条命令队列，
    所有工具调用都 post 进去、同步等结果。因为 CAD 本来就是串行的，
    这样不损失任何性能，反而把随机崩溃变成了确定行为。

    实现上每次任务新建一份 pythoncom.CoInitialize()（在该线程内），
    线程生命周期与 CadBridge 一致，随 disconnect()/进程退出结束。
    """

    def __init__(self) -> None:
        self._queue: "queue.Queue[tuple[Any, threading.Event, list]]" = queue.Queue()
        self._thread: threading.Thread | None = None
        self._closed = False
        self._thread_lock = threading.Lock()

    def _ensure_thread(self) -> None:
        if self._thread is not None and self._thread.is_alive():
            return
        with self._thread_lock:
            if self._thread is not None and self._thread.is_alive():
                return
            self._closed = False
            self._thread = threading.Thread(
                target=self._run, name="cad-com-sta", daemon=True
            )
            self._thread.start()

    def _run(self) -> None:
        pythoncom.CoInitialize()
        try:
            while True:
                item = self._queue.get()
                if item is None:
                    break
                fn, done, box = item
                try:
                    box.append((True, fn()))
                except BaseException as exc:  # noqa: BLE001
                    box.append((False, exc))
                finally:
                    done.set()
        finally:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    def submit(self, fn, timeout_s: float | None = None):
        """把 fn 丢给 COM 线程执行，同步阻塞拿结果（异常原样抛回调用方）。"""
        if self._closed:
            raise CadDisconnectedError("COM 工作线程已关闭。")
        self._ensure_thread()
        done = threading.Event()
        box: list = []
        self._queue.put((fn, done, box))
        if not done.wait(timeout_s):
            raise CadBusyError(
                f"AutoCAD 在 {timeout_s:.0f}s 内没有响应（可能有模态弹窗）。"
                "请切到 CAD 窗口按 ESC 后重试。"
            )
        ok, value = box[0]
        if ok:
            return value
        raise value

    @property
    def ident(self) -> int | None:
        t = self._thread
        return t.ident if t is not None and t.is_alive() else None

    def close(self) -> None:
        self._closed = True
        t = self._thread
        if t is not None and t.is_alive():
            self._queue.put(None)
            t.join(timeout=3.0)


# ---------------------------------------------------------------------------
# 主桥
# ---------------------------------------------------------------------------

class CadBridge:
    """
    一个 CadBridge 实例 = 一个连接到 AutoCAD 的会话。

    典型用法：
        cad = CadBridge(prog_id="AutoCAD.Application.24")
        cad.connect()
        cad.add_line((0,0), (100,100))
    """

    def __init__(
        self,
        prog_id: str = "AutoCAD.Application.24",
        timeout_ms: int = 15000,
        retry_times: int = 2,
        retry_delay_ms: int = 800,
    ) -> None:
        self.prog_id = prog_id
        self.timeout_ms = int(timeout_ms)
        self.retry_times = max(1, int(retry_times))
        self.retry_delay_ms = int(retry_delay_ms)

        self._app: Any = None
        self._com_ready = False
        self._com_threads: set[int] = set()
        self._worker = _ComWorker()
        self._lock = threading.RLock()
        self._last_doc_name: str | None = None

    # -- 生命周期 ----------------------------------------------------------

    def _ensure_com(self) -> None:
        """
        确保【当前线程】调用了 CoInitialize。

        正常情况下所有 COM 调用都在 _ComWorker 的专用线程上执行，那里已经
        CoInitialize 过。这个方法保留下来是为了直接在单线程脚本里使用
        （测试、调试）时不必额外处理。
        """
        tid = threading.get_ident()
        if tid in self._com_threads:
            return
        pythoncom.CoInitialize()
        self._com_threads.add(tid)
        self._com_ready = True

    def _on_com(self, fn):
        """
        在专用 COM 线程上执行 fn。

        如果当前已经在 COM 线程上（例如 connect 内部再调 list_documents），
        就直接执行，否则会自己等自己，死锁。
        """
        if self._worker.ident is not None and self._worker.ident == threading.get_ident():
            return fn()
        return self._worker.submit(
            fn, timeout_s=(self.timeout_ms / 1000.0) * max(2, self.retry_times)
        )

    def connect(self, visible: bool = True) -> dict[str, Any]:
        """
        连接（或启动）AutoCAD。

        返回 {"ok": True, "version": ..., "prog_id": ..., "documents": [...]}
        失败抛 CadDisconnectedError，并给出可操作的排查建议。
        """
        return self._on_com(lambda: self._connect_impl(visible))

    def _connect_impl(self, visible: bool = True) -> dict[str, Any]:
        self._ensure_com()

        # 用 gencache.EnsureDispatch（早绑定）而不是 Dispatch（晚绑定）。
        #
        # 原因（实测）：
        #   - 早绑定会从 AutoCAD 的 typelib 生成强类型包装，参数类型由 marshaler 保证，
        #     SAFEARRAY(double) 能被正确传递；
        #   - 晚绑定下即使传 VARIANT(VT_ARRAY|VT_R8) 也偶发 E_INVALIDARG；
        #   - 而且早绑定能拿到 AcDbLine 这类常量（acRed 等），代码可读性也更好。
        #
        # EnsureDispatch 会先找已运行的实例，找不到再启动一个 —— 正是我们要的行为。
        # 首次调用会生成 gen_py 缓存（约 1~3 秒），后续走缓存很快。
        try:
            try:
                app = gencache.EnsureDispatch(self.prog_id, bForDemand=False)
            except Exception:
                # gen_py 生成失败（权限/typelib 损坏）时退回晚绑定，
                # 晚绑定下 VARIANT(VT_ARRAY) 仍能工作，只是不那么稳。
                app = win32com.client.Dispatch(self.prog_id)
        except pywintypes.com_error as exc:
            raise CadDisconnectedError(
                f"无法连接 AutoCAD（ProgID='{self.prog_id}'）。\n"
                f"原始错误：{exc}\n"
                "排查顺序：\n"
                "  1) AutoCAD 是否已启动并处于登录完成状态（不要停在启动画面）；\n"
                "  2) ProgID 是否与版本匹配：2024→AutoCAD.Application.24，2023→.23；\n"
                "  3) 是否以管理员权限运行过 CAD（高权限进程无法被低权限进程 COM 连接）；\n"
                f"  4) 可用 ProgID 探测：{self.list_available_prog_ids()}"
            ) from exc

        self._app = app
        version = self._safe(lambda: str(app.Version), default="unknown")
        if visible:
            def _show() -> bool:
                app.Visible = True
                return True

            self._safe(_show)

        # 没有打开任何图纸时，AutoCAD 的 Documents 是空的，新建一个空 DWG 便于操作
        if self._safe(lambda: int(app.Documents.Count), default=0) == 0:
            self._safe(lambda: app.Documents.Add())

        return {
            "ok": True,
            "prog_id": self.prog_id,
            "version": version,
            "documents": self.list_documents(),
            "active_document": self.active_document_name(),
        }

    @staticmethod
    def list_available_prog_ids() -> list[str]:
        """扫注册表，列出本机真实存在的 AutoCAD ProgID（排查连不上时非常有用）。"""
        found: list[str] = []
        try:
            import winreg

            base = winreg.HKEY_LOCAL_MACHINE
            with winreg.OpenKey(base, r"SOFTWARE\Classes") as classes:
                idx = 0
                while True:
                    try:
                        name = winreg.EnumKey(classes, idx)
                    except OSError:
                        break
                    idx += 1
                    if name.startswith("AutoCAD.Application"):
                        found.append(name)
        except Exception:
            pass
        return sorted(set(found))

    def is_connected(self) -> bool:
        if self._app is None:
            return False
        return bool(self._safe(lambda: self._app.Version is not None, default=False))

    def disconnect(self) -> None:
        """只丢引用，不关 CAD —— 关掉用户的 CAD 会让人丢数据。"""
        self._app = None

    # -- 低层调用：忙检测 + 重试 -------------------------------------------

    def _safe(self, fn, default=None):
        """不重试、不抛错的探测调用，用于取状态。"""
        try:
            return fn()
        except Exception:
            return default

    def _call(self, label: str, fn):
        """
        所有 COM 操作都必须经过这里。

        逻辑：
          - 整段丢到 _ComWorker 的专用 STA 线程执行（见类注释），
            这样无论 MCP 在哪个线程调用 tools/call 都安全；
          - 撞上"CAD 忙" → 睡眠后重试，最多 retry_times 次；
          - 重试耗尽 → 抛 CadBusyError（并提示可能有弹窗）；
          - COM 断开 → 抛 CadDisconnectedError；
          - 其他错误 → 抛 CadError，附上 label 方便定位是哪一步炸的。
        """

        def _attempt() -> Any:
            if self._app is None:
                raise CadDisconnectedError("尚未连接 AutoCAD，请先调用 connect。")

            last: Exception | None = None
            for attempt in range(1, self.retry_times + 1):
                try:
                    self._ensure_com()
                    return fn()
                except pywintypes.com_error as exc:
                    hresult = _hresult(exc)
                    if hresult in BUSY_HRESULTS:
                        last = exc
                        if attempt < self.retry_times:
                            time.sleep(self.retry_delay_ms / 1000.0)
                            continue
                        raise CadBusyError(
                            f"AutoCAD 一直处于忙碌状态（{label} 重试 {self.retry_times} 次仍失败）。\n"
                            "可能原因：CAD 正在执行长命令，或弹出了模态对话框（字体缺失/许可/警告）。\n"
                            "建议：切到 CAD 窗口看一眼，按 ESC 关掉弹窗后重试。"
                        ) from exc
                    if hresult in (RPC_E_DISCONNECTED,):
                        self._app = None
                        raise CadDisconnectedError(
                            f"AutoCAD COM 连接已断开（{label}）。请重新 connect。"
                        ) from exc
                    raise CadError(f"{label} 失败：{exc}") from exc
                except (CadError, CadBusyError, CadDisconnectedError):
                    raise
                except Exception as exc:  # noqa: BLE001
                    raise CadError(f"{label} 失败：{exc}") from exc
            raise CadError(f"{label} 失败：{last}")

        return self._on_com(_attempt)

    # -- 文档 --------------------------------------------------------------

    def list_documents(self) -> list[dict[str, Any]]:
        if self._app is None:
            return []
        return self._on_com(self._list_documents_impl)

    def _list_documents_impl(self) -> list[dict[str, Any]]:
        names: list[dict[str, Any]] = []
        if self._app is None:
            return names
        count = self._safe(lambda: int(self._app.Documents.Count), default=0) or 0
        for i in range(count):
            doc = self._safe(lambda i=i: self._app.Documents.Item(i), default=None)
            if doc is None:
                continue
            names.append(
                {
                    "name": self._safe(lambda d=doc: str(d.Name), default="?"),
                    "full_name": self._safe(lambda d=doc: str(d.FullName), default=""),
                    "active": bool(self._safe(lambda d=doc: d.Name == self._app.ActiveDocument.Name, default=False)),
                }
            )
        return names

    def _document(self, doc_name: str | None = None):
        """
        取目标文档。

        doc_name 为空 → 当前活动文档（并在 _last_doc_name 里记住它）；
        doc_name 给了 → 按文件名（或全路径）查找，找不到就报错，绝不"猜"。
        """
        if self._app is None:
            raise CadDisconnectedError("尚未连接 AutoCAD。")

        if not doc_name:
            doc = self._call("获取活动文档", lambda: self._app.ActiveDocument)
            self._last_doc_name = self._safe(lambda: str(doc.Name), default=None)
            return doc

        count = self._safe(lambda: int(self._app.Documents.Count), default=0) or 0
        for i in range(count):
            doc = self._safe(lambda i=i: self._app.Documents.Item(i), default=None)
            if doc is None:
                continue
            full = self._safe(lambda d=doc: str(d.FullName), default="")
            name = self._safe(lambda d=doc: str(d.Name), default="")
            if doc_name in (full, name) or os.path.basename(full) == doc_name:
                self._last_doc_name = name
                return doc
        raise CadError(
            f"找不到文档 '{doc_name}'。当前打开的文档：{[d['name'] for d in self.list_documents()]}"
        )

    def active_document_name(self) -> str | None:
        if self._app is None:
            return None
        return self._on_com(
            lambda: self._safe(lambda: str(self._app.ActiveDocument.Name), default=None)
        )

    # -- 状态 --------------------------------------------------------------

    def status(self) -> dict[str, Any]:
        """
        给插件面板用的状态快照。

        state 取值：
          disconnected — 没连上
          idle         — 连上且空闲，可以接受命令
          busy         — 连上了，但 COM 不接受调用（命令执行中/弹窗）
        """
        if self._app is None:
            return {"state": "disconnected", "reason": "尚未 connect"}

        try:
            def _probe():
                self._ensure_com()
                return (
                    str(self._app.Version),
                    str(self._app.ActiveDocument.Name),
                    int(self._app.Documents.Count),
                )

            version, doc_name, doc_count = self._on_com(_probe)
        except pywintypes.com_error as exc:
            hresult = _hresult(exc)
            if hresult in BUSY_HRESULTS:
                return {"state": "busy", "reason": "CAD 正在执行命令或存在模态弹窗"}
            if hresult == RPC_E_DISCONNECTED:
                self._app = None
                return {"state": "disconnected", "reason": "COM 连接断开"}
            return {"state": "error", "reason": str(exc)}
        except Exception as exc:  # noqa: BLE001
            return {"state": "error", "reason": str(exc)}

        return {
            "state": "idle",
            "version": version,
            "active_document": doc_name,
            "document_count": doc_count,
            "prog_id": self.prog_id,
        }

    def is_busy(self) -> bool:
        return self.status().get("state") == "busy"

    # -- 命令发送 ----------------------------------------------------------

    def send_command(self, command: str, doc_name: str | None = None) -> dict[str, Any]:
        """
        把任意 AutoCAD 原生命令发给 CAD 命令行。

        注意：AutoCAD 的 SendCommand 是"发完就返回"，它不等命令执行完。
        所以这里显式 sleep 一小段，避免紧接着的下一条命令挤掉这一条。
        返回里带上"已发送"，不代表"已执行完成" —— 这是 COM 的固有限制。
        """
        if not command.endswith(("\n", " ")):
            command = command + "\n"

        def _do():
            doc = self._document(doc_name)
            doc.SendCommand(command)
            return True

        self._call(f"发送命令 {command.strip()[:60]}", _do)
        time.sleep(0.25)

        # 报告是否有弹窗/报错文本出现在命令行提示里
        return {
            "sent": True,
            "command": command.strip(),
            "document": self._last_doc_name,
            "note": "SendCommand 为异步发送，不等待执行完成",
        }

    def send_escape(self) -> dict[str, Any]:
        """
        发送 ESC 以解除 CAD 卡住的状态（关闭弹窗 / 取消当前命令）。

        实现在"前端/插件层"通常靠 SendKeys，这里用 COM 的 PostCommand 发 ESC。
        注意：PostCommand 只在空闲时可排队；忙时它本身会失败，属预期。
        """
        try:
            self._call("发送 ESC", lambda: self._app.PostCommand("\x1b\x1b"))
            return {"ok": True, "note": "已通过 PostCommand 排队发送 ESC"}
        except CadError as exc:
            return {
                "ok": False,
                "note": "PostCommand 不可用（CAD 可能处于忙/弹窗状态）。",
                "detail": str(exc),
                "fallback": "请手动切到 AutoCAD 窗口按 ESC。",
            }

    # -- 图层 --------------------------------------------------------------

    def ensure_layer(self, name: str, color: str | int | None = None, doc_name: str | None = None) -> dict[str, Any]:
        """图层存在则复用，不存在则创建。color 支持 'red' 或 ACI 索引号。"""
        if not name or not name.strip():
            raise CadError("图层名不能为空。")
        color_index = self._resolve_color(color)

        def _do():
            doc = self._document(doc_name)
            layers = doc.Layers
            existing = None
            try:
                existing = layers.Item(name)
            except Exception:
                existing = None
            if existing is None:
                layer = layers.Add(name)
                created = True
            else:
                layer = existing
                created = False
            if color_index is not None:
                layer.color = color_index
            return created

        created = self._call(f"创建/复用图层 {name}", _do)
        return {
            "ok": True,
            "layer": name,
            "created": created,
            "color_index": color_index,
            "document": self._last_doc_name,
        }

    @staticmethod
    def _resolve_color(color: str | int | None) -> int | None:
        if color is None or color == "":
            return None
        if isinstance(color, int):
            return color
        key = str(color).strip().lower()
        if key.isdigit():
            return int(key)
        if key in ACAD_COLOR_INDEX:
            return ACAD_COLOR_INDEX[key]
        raise CadError(
            f"无法识别的颜色 '{color}'。支持颜色名：{sorted(ACAD_COLOR_INDEX)}，或 1-255 的 ACI 索引。"
        )

    def list_layers(self, doc_name: str | None = None) -> list[dict[str, Any]]:
        def _do():
            doc = self._document(doc_name)
            out = []
            for i in range(int(doc.Layers.Count)):
                layer = doc.Layers.Item(i)
                out.append({"name": str(layer.Name), "color_index": int(self._safe(lambda l=layer: l.color, default=0))})
            return out

        return self._call("读取图层列表", _do)

    # -- 绘图图元 ----------------------------------------------------------

    def add_line(
        self,
        start: Sequence[float],
        end: Sequence[float],
        layer: str | None = None,
        doc_name: str | None = None,
    ) -> dict[str, Any]:
        if len(start) < 2 or len(end) < 2:
            raise CadError("add_line 需要两个二维或三维坐标点。")
        if layer:
            self.ensure_layer(layer, doc_name=doc_name)

        def _do():
            doc = self._document(doc_name)
            ent = doc.ModelSpace.AddLine(_pt(*start[:3]), _pt(*end[:3]))
            if layer:
                ent.Layer = layer
            return {"handle": self._safe(lambda: str(ent.Handle), default=""), "length": self._safe(lambda: round(float(ent.Length), 6), default=None)}

        info = self._call("画直线", _do)
        return {"ok": True, "type": "line", "start": list(start), "end": list(end), "layer": layer, **info, "document": self._last_doc_name}

    def add_circle(
        self,
        center: Sequence[float],
        radius: float,
        layer: str | None = None,
        doc_name: str | None = None,
    ) -> dict[str, Any]:
        if len(center) < 2:
            raise CadError("add_circle 需要圆心坐标。")
        if radius <= 0:
            raise CadError("半径必须大于 0。")
        if layer:
            self.ensure_layer(layer, doc_name=doc_name)

        def _do():
            doc = self._document(doc_name)
            ent = doc.ModelSpace.AddCircle(_pt(*center[:3]), float(radius))
            if layer:
                ent.Layer = layer
            return {"handle": self._safe(lambda: str(ent.Handle), default="")}

        info = self._call("画圆", _do)
        return {"ok": True, "type": "circle", "center": list(center), "radius": float(radius), "layer": layer, **info, "document": self._last_doc_name}

    def add_rectangle(
        self,
        corner1: Sequence[float],
        corner2: Sequence[float],
        layer: str | None = None,
        doc_name: str | None = None,
    ) -> dict[str, Any]:
        """用 LightWeightPolyline 画矩形（4 个顶点闭合）。"""
        x1, y1 = float(corner1[0]), float(corner1[1])
        x2, y2 = float(corner2[0]), float(corner2[1])
        if layer:
            self.ensure_layer(layer, doc_name=doc_name)

        verts = [x1, y1, x2, y1, x2, y2, x1, y2]

        def _do():
            doc = self._document(doc_name)
            pl = doc.ModelSpace.AddLightWeightPolyline(_flat(verts))
            pl.Closed = True
            if layer:
                pl.Layer = layer
            return {"handle": self._safe(lambda: str(pl.Handle), default="")}

        info = self._call("画矩形", _do)
        return {
            "ok": True,
            "type": "rectangle",
            "corner1": [x1, y1],
            "corner2": [x2, y2],
            "layer": layer,
            **info,
            "document": self._last_doc_name,
        }

    def add_polyline(
        self,
        points: Iterable[Sequence[float]],
        closed: bool = False,
        layer: str | None = None,
        doc_name: str | None = None,
    ) -> dict[str, Any]:
        pts = list(points)
        if len(pts) < 2:
            raise CadError("add_polyline 至少需要 2 个点。")
        flat: list[float] = []
        for p in pts:
            flat.extend([float(p[0]), float(p[1])])
        if layer:
            self.ensure_layer(layer, doc_name=doc_name)

        def _do():
            doc = self._document(doc_name)
            pl = doc.ModelSpace.AddLightWeightPolyline(_flat(flat))
            pl.Closed = bool(closed)
            if layer:
                pl.Layer = layer
            return {"handle": self._safe(lambda: str(pl.Handle), default="")}

        info = self._call("画多段线", _do)
        return {"ok": True, "type": "polyline", "points": pts, "closed": closed, "layer": layer, **info, "document": self._last_doc_name}

    def add_text(
        self,
        text: str,
        position: Sequence[float],
        height: float = 2.5,
        layer: str | None = None,
        doc_name: str | None = None,
    ) -> dict[str, Any]:
        if len(position) < 2:
            raise CadError("add_text 需要插入点坐标。")
        if layer:
            self.ensure_layer(layer, doc_name=doc_name)

        def _do():
            doc = self._document(doc_name)
            ent = doc.ModelSpace.AddText(str(text), _pt(*position[:3]), float(height))
            if layer:
                ent.Layer = layer
            return {"handle": self._safe(lambda: str(ent.Handle), default="")}

        info = self._call("写文字", _do)
        return {"ok": True, "type": "text", "text": str(text), "position": list(position), "height": float(height), "layer": layer, **info, "document": self._last_doc_name}

    def add_arc(
        self,
        center: Sequence[float],
        radius: float,
        start_angle_deg: float,
        end_angle_deg: float,
        layer: str | None = None,
        doc_name: str | None = None,
    ) -> dict[str, Any]:
        if layer:
            self.ensure_layer(layer, doc_name=doc_name)

        def _do():
            doc = self._document(doc_name)
            ent = doc.ModelSpace.AddArc(
                _pt(*center[:3]),
                float(radius),
                math.radians(float(start_angle_deg)),
                math.radians(float(end_angle_deg)),
            )
            if layer:
                ent.Layer = layer
            return {"handle": self._safe(lambda: str(ent.Handle), default="")}

        info = self._call("画圆弧", _do)
        return {"ok": True, "type": "arc", "center": list(center), "radius": float(radius), "start_deg": start_angle_deg, "end_deg": end_angle_deg, "layer": layer, **info, "document": self._last_doc_name}

    # -- 读取 --------------------------------------------------------------

    def list_entities(self, doc_name: str | None = None, limit: int = 200) -> dict[str, Any]:
        """列出模型空间图元摘要（类型/图层/句柄），便于 Agent 理解当前图纸内容。"""

        def _do():
            doc = self._document(doc_name)
            ms = doc.ModelSpace
            total = int(ms.Count)
            out = []
            for i in range(min(total, max(1, int(limit)))):
                ent = ms.Item(i)
                out.append(
                    {
                        "index": i,
                        "type": self._safe(lambda e=ent: str(e.ObjectName), default="?"),
                        "layer": self._safe(lambda e=ent: str(e.Layer), default="?"),
                        "handle": self._safe(lambda e=ent: str(e.Handle), default=""),
                    }
                )
            return {"total": total, "returned": len(out), "entities": out}

        result = self._call("读取图元", _do)
        return {"ok": True, **result, "document": self._last_doc_name}

    def save(self, path: str | None = None, doc_name: str | None = None) -> dict[str, Any]:
        """path 为空 → 保存到当前图纸原路径（Save）。path 给了 → 另存为（SaveAs）。"""
        if path and not path.lower().endswith(".dwg"):
            path = path + ".dwg"

        def _do():
            doc = self._document(doc_name)
            if path:
                os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
                doc.SaveAs(os.path.abspath(path))
                return {"saved_to": os.path.abspath(path), "mode": "saveas"}
            doc.Save()
            return {"saved_to": self._safe(lambda: str(doc.FullName), default=""), "mode": "save"}

        info = self._call("保存图纸", _do)
        return {"ok": True, **info, "document": self._last_doc_name}

    def new_document(self, template: str | None = None) -> dict[str, Any]:
        def _do():
            if template:
                self._app.Documents.Add(template)
            else:
                self._app.Documents.Add()
            return str(self._app.ActiveDocument.Name)

        name = self._call("新建图纸", _do)
        time.sleep(0.3)
        return {"ok": True, "document": name}

    def zoom_extents(self, doc_name: str | None = None) -> dict[str, Any]:
        """缩放到全图，让用户肉眼立刻看到刚画的东西。"""

        def _do():
            doc = self._document(doc_name)
            try:
                doc.Application.ZoomExtents()
            except Exception:
                doc.SendCommand("_.ZOOM\n_E\n")
            return True

        self._call("缩放全图", _do)
        return {"ok": True, "document": self._last_doc_name}
