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
from win32com.client import VARIANT, dynamic, gencache

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

# ---------------------------------------------------------------------------
# V1.1 工业基础：机械制图国标（GB）默认图层 / 标注样式
# ---------------------------------------------------------------------------

#: 机械制图常用标准图层（GB/T 14689 等）。lineweight 单位 1/100 mm，
#: linetype 用 AutoCAD 自带线型名（CENTER / DASHED 需先从 acad.lin 载入）。
GB_LAYERS: dict[str, dict[str, Any]] = {
    "轮廓": {"color": 7, "lineweight": 50, "linetype": "Continuous"},
    "中心线": {"color": 1, "lineweight": 18, "linetype": "CENTER"},
    "虚线": {"color": 2, "lineweight": 18, "linetype": "DASHED"},
    "标注": {"color": 3, "lineweight": 18, "linetype": "Continuous"},
    "文字": {"color": 7, "lineweight": 25, "linetype": "Continuous"},
    "剖面线": {"color": 8, "lineweight": 18, "linetype": "Continuous"},
}

#: 国标标注样式默认尺寸变量（在激活 DimStyle 之后设置，再 CopyFrom 固化）。
#: 角度一律是"度"由调用方转换；DIMLFAC 默认 1（几何单位=标注单位）。
GB_DIM_VARS: dict[str, Any] = {
    "DIMTXT": 3.5,      # 标注文字高
    "DIMASZ": 2.5,      # 箭头大小
    "DIMEXE": 1.5,      # 尺寸界线超出量
    "DIMEXO": 1.0,      # 尺寸界线偏移
    "DIMGAP": 1.0,      # 文字与尺寸线间隙
    "DIMDEC": 0,        # 小数位
    "DIMZIN": 0,        # 不消零
    "DIMSCALE": 1.0,    # 全局比例
    "DIMTAD": 1,        # 文字在尺寸线上方
    "DIMTIH": 0,        # 文字不强制水平（对齐标注）
    "DIMTOH": 0,
    "DIMTIX": 0,        # 文字可放在尺寸线之间
    "DIMSOXD": 0,       # 允许画外侧尺寸线
    "DIMLFAC": 1.0,     # 测量单位比例
    "DIMASSOC": 2,      # 关联标注
    "DIMLUNIT": 2,      # 十进制
    "DIMCLRD": 7,       # 尺寸线颜色
    "DIMCLRE": 7,       # 尺寸界线颜色
    "DIMCLRT": 7,       # 文字颜色
    "DIMBLK": ".",      # 默认实心箭头
    "DIMLDRBLK": ".",   # 引线箭头
    "DIMAUNIT": 0,      # 角度：十进制度
    "DIMFRAC": 0,
}

#: 中文工程字：西文 gbenor.shx + 中文大字体 gbcbig.shx（AutoCAD 2024 实测可用）
CN_FONT = "gbenor.shx"
CN_BIGFONT = "gbcbig.shx"

#: 图元大类 → 标准图层（setup_standard_layers 之后启用"自动归属"）
DEFAULT_LAYER_BY_KIND: dict[str, str] = {
    "outline": "轮廓",
    "center": "中心线",
    "hidden": "虚线",
    "dim": "标注",
    "text": "文字",
    "hatch": "剖面线",
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
        # V1.1-5：可观测的队列/线程状态。worker 一旦卡在 COM 调用里（模态弹窗），
        # 后续所有 submit 都会排队超时 —— 必须能看出来，才能「超时重连」。
        self._state_lock = threading.Lock()
        self._pending = 0
        self._busy_since: float | None = None
        self._last_error: str | None = None

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
                with self._state_lock:
                    self._busy_since = time.monotonic()
                try:
                    box.append((True, fn()))
                except BaseException as exc:  # noqa: BLE001
                    box.append((False, exc))
                finally:
                    with self._state_lock:
                        self._pending = max(0, self._pending - 1)
                        if self._pending == 0:
                            self._busy_since = None
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
        with self._state_lock:
            self._pending += 1
        self._queue.put((fn, done, box))
        if not done.wait(timeout_s):
            with self._state_lock:
                self._last_error = f"last submit timed out after {timeout_s:.0f}s"
                busy_for = (time.monotonic() - self._busy_since) if self._busy_since else None
            raise CadBusyError(
                f"AutoCAD 在 {timeout_s:.0f}s 内没有响应（可能有模态弹窗）。"
                f"当前 COM 线程已忙 {busy_for:.0f}s，队列积压 {self._pending} 个调用。"
                "请切到 CAD 窗口按 ESC 关掉弹窗；若长期卡死可调用 cad_reconnect 重建连接。"
            )
        with self._state_lock:
            self._last_error = None
        ok, value = box[0]
        if ok:
            return value
        raise value

    def health(self) -> dict[str, Any]:
        """V1.1-5：队列/线程健康快照，供 status() 与故障排查用。"""
        with self._state_lock:
            busy_since = self._busy_since
            return {
                "alive": bool(self._thread is not None and self._thread.is_alive()),
                "pending": self._pending,
                "busy_seconds": round(time.monotonic() - busy_since, 1) if busy_since else 0.0,
                "last_error": self._last_error,
            }

    def abandon(self) -> None:
        """
        V1.1-5：丢弃这个 worker（不 join —— 它可能正卡在模态弹窗里）。

        Python 无法强杀线程，但旧线程是 daemon，且它持有的 COM 指针会随
        CadBridge 换用新 worker 而被 GC。发哨兵让它在脱困后自行退出。
        """
        self._closed = True
        t = self._thread
        if t is not None and t.is_alive():
            try:
                self._queue.put(None)
            except Exception:
                pass

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
        # V1.1-2：默认图层（由 autocad_server 从 settings.default_layer 注入）。
        # 图元/标注/文字未显式指定 layer 时自动归属到这里，实现"自动分层"。
        self.default_layer: str = ""
        # V1.1-2：按图元大类自动分层（outline/center/hidden/dim/text/hatch）。
        # 默认关闭 —— 保持「省略 layer 就用 CAD 当前图层」的向下兼容行为；
        # setup_mechanical_environment() 会把它打开，实现"一句话分层规范"。
        self.auto_layer_by_kind: bool = False

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

    def _on_com(self, fn, timeout_s: float | None = None):
        """
        在专用 COM 线程上执行 fn。

        如果当前已经在 COM 线程上（例如 connect 内部再调 list_documents），
        就直接执行，否则会自己等自己，死锁。

        timeout_s 可覆盖默认超时；connect() 的启动等待会用更长的预算。
        """
        if self._worker.ident is not None and self._worker.ident == threading.get_ident():
            return fn()
        if timeout_s is None:
            timeout_s = (self.timeout_ms / 1000.0) * max(2, self.retry_times)
        return self._worker.submit(fn, timeout_s=timeout_s)

    def connect(self, visible: bool = True) -> dict[str, Any]:
        """
        连接（或启动）AutoCAD。

        V1.1-5 修复：CAD 未启动时 EnsureDispatch 会把它拉起来，但启动期
        COM 会连着抛 RPC_E_CALL_REJECTED，pywin32 甚至会包成
            TypeError: This COM object can not automate the makepy process
        这不是"连不上"，而是"CAD 还在启动"。所以：
          - 用独立的、更长的超时预算（默认 60s，足够冷启动）；
          - 重试循环同时捕获 com_error / TypeError / 通用异常；
          - 每次重试之间等待，让 CAD 完成初始化；
          - 早绑定持续失败才退回晚绑定 Dispatch，最终再报错。
        """
        # 冷启动可能需要几十秒，给这条调用单独的超时预算
        startup_timeout = max(60.0, (self.timeout_ms / 1000.0) * max(2, self.retry_times))
        return self._on_com(lambda: self._connect_impl(visible), timeout_s=startup_timeout)

    def _connect_impl(self, visible: bool = True) -> dict[str, Any]:
        self._ensure_com()

        # ★ 用 gencache.EnsureDispatch（早绑定）而不是 Dispatch（晚绑定）。
        # 原因（实测）：早绑定生成强类型包装，SAFEARRAY(double) 能被正确传递；
        # 晚绑定下即使传 VARIANT(VT_ARRAY|VT_R8) 也偶发 E_INVALIDARG。
        # EnsureDispatch 会先找已运行实例，找不到再启动 —— 正是"自动唤醒"所需。
        #
        # V1.1-5：启动期 CAD 尚未就绪时，EnsureDispatch 内部 makepy 会抛
        #   com_error(RPC_E_CALL_REJECTED) 或 TypeError(can not automate makepy)
        # 所以这里做有界重试，直到拿到可用 Application 或超时。
        app = None
        last_exc: Exception | None = None
        deadline = time.monotonic() + 55.0
        attempt = 0
        while time.monotonic() < deadline:
            attempt += 1
            try:
                app = gencache.EnsureDispatch(self.prog_id, bForDemand=False)
                # 拿到对象不等于就绪：Version 在启动期也可能被拒，主动探一次
                _ = str(app.Version)
                break
            except pywintypes.com_error as exc:
                last_exc = exc
                hr = _hresult(exc)
                if hr not in BUSY_HRESULTS:
                    # 非"忙碌"类错误（如 ProgID 不存在）没必要死等
                    break
                app = None
                time.sleep(0.8)
            except Exception as exc:  # noqa: BLE001  TypeError 等启动期包装异常
                last_exc = exc
                app = None
                time.sleep(0.8)

        if app is None:
            # 早绑定始终失败 → 退回避重连的晚绑定，再不行才报错
            try:
                app = win32com.client.Dispatch(self.prog_id)
                _ = str(app.Version)
            except Exception as exc:  # noqa: BLE001
                last_exc = exc
                app = None

        if app is None:
            raise CadDisconnectedError(
                f"无法连接 AutoCAD（ProgID='{self.prog_id}'，重试 {attempt} 次）。\n"
                f"原始错误：{last_exc}\n"
                "排查顺序：\n"
                "  1) AutoCAD 是否已完成启动（不要停在启动画面/许可弹窗）；\n"
                "  2) ProgID 是否与版本匹配：2024→AutoCAD.Application.24，2023→.23；\n"
                "  3) 是否以管理员权限运行过 CAD（高权限进程无法被低权限进程 COM 连接）；\n"
                f"  4) 可用 ProgID 探测：{self.list_available_prog_ids()}"
            )

        self._app = app
        version = self._safe(lambda: str(app.Version), default="unknown")
        if visible:
            def _show() -> bool:
                app.Visible = True
                return True

            self._safe(_show)

        # 没有打开任何图纸时，Documents 是空的，新建一个空 DWG 便于操作。
        # V1.1-5：Documents.Add() 非幂等（忙时报错但图已建好），按 Count 核对重试。
        if self._safe(lambda: int(app.Documents.Count), default=0) == 0:
            doc_deadline = time.monotonic() + 20.0
            while time.monotonic() < doc_deadline:
                try:
                    app.Documents.Add()
                except Exception:  # noqa: BLE001  忙也可能已建好，下一轮按 Count 核对
                    pass
                if self._safe(lambda: int(app.Documents.Count), default=0) > 0:
                    break
                time.sleep(0.6)

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

    def reconnect(self) -> dict[str, Any]:
        """
        V1.1-5「超时重连」：队列卡死 / COM 断开后重建 worker + 重连 CAD。

        不主动关闭用户的 AutoCAD，只丢弃旧 COM 引用与可能卡死的专用线程，
        再走一次 connect()。这样 modal 弹窗关掉后无需重启整个 MCP 进程。
        """
        old_worker = self._worker
        self._app = None
        self._com_ready = False
        self._com_threads.clear()
        self._worker = _ComWorker()
        old_worker.abandon()
        return self.connect()

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
            "default_layer": self.default_layer or None,
            "queue": self._worker.health(),
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

    def send_escape(self, doc_name: str | None = None) -> dict[str, Any]:
        """
        发送 ESC 以解除 CAD 卡住的状态（关闭弹窗 / 取消当前命令）。

        ★ 实测：PostCommand 是 IAcadDocument 的方法，不是 IAcadApplication 的
        （早绑定下 app.PostCommand 会报 "object has no attribute 'PostCommand'"）。
        且 ESC 必须用 PostCommand 排队，用 SendCommand 发 "\x1b" 会报「输入无效」。
        注意：PostCommand 只在空闲时可排队；忙时它本身会失败，属预期。
        """
        def _do():
            doc = self._document(doc_name)
            doc.PostCommand("\x1b\x1b")
            return True

        try:
            self._call("发送 ESC", _do)
            return {"ok": True, "note": "已通过文档 PostCommand 排队发送 ESC"}
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
        spec = GB_LAYERS.get(name)
        lw = spec["lineweight"] if spec else None
        lt = spec["linetype"] if spec else None

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
            elif spec is not None:
                layer.color = spec["color"]
            # V1.1-2：标准图层自动带线宽/线型（动态 Dispatch 兜住 narrow 化）
            if lw is not None or lt:
                d = dynamic.Dispatch(layer)
                if lw is not None:
                    try:
                        d.Lineweight = lw
                    except Exception:
                        pass
                if lt:
                    try:
                        doc.Linetypes.Load(lt, "acad.lin")
                    except Exception:
                        pass  # 已加载 / 线型文件缺失时忽略
                    try:
                        d.Linetype = lt
                    except Exception:
                        pass
            return created

        created = self._call(f"创建/复用图层 {name}", _do)
        return {
            "ok": True,
            "layer": name,
            "created": created,
            "color_index": color_index,
            "document": self._last_doc_name,
        }

    def _layer_for(self, layer: str | None, kind: str | None = None) -> str | None:
        """
        V1.1-2 自动归属，优先级：显式 layer > 配置 default_layer > 按大类。

        「按大类」只在 auto_layer_by_kind=True 时生效（由
        setup_mechanical_environment() 打开）。默认关闭是刻意的向下兼容：
        老行为是「省略 layer 就用 CAD 当前图层」，不能悄悄改掉。
        """
        if layer:
            return layer
        if self.default_layer:
            return self.default_layer
        if self.auto_layer_by_kind and kind:
            return DEFAULT_LAYER_BY_KIND.get(kind)
        return None

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
        layer = self._layer_for(layer, "outline")
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
        layer = self._layer_for(layer, "outline")
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
        layer = self._layer_for(layer, "outline")
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
        layer = self._layer_for(layer, "outline")
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
        layer = self._layer_for(layer, "text")
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
        layer = self._layer_for(layer, "outline")
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

    # -----------------------------------------------------------------------
    # V1.1-1 完整尺寸标注体系
    #
    # 全部经 probe_v11.py / probe_v11b.py / probe_v11c.py 在 AutoCAD 2024 实测：
    #   - 角度一律是【弧度】（AddDimRotated 第 4 参数、InsertBlock 旋转角）
    #   - 坐标一律 VARIANT(VT_ARRAY|VT_R8)
    #   - 标注样式变量必须「先建并激活 DimStyle → 再 SetVariable → CopyFrom 固化」
    #     否则变量会落到旧样式上（probe_v11b 序列 B 读回 0.18/Standard）
    #   - 读属性/读文字一律 dynamic.Dispatch（gen_py 会窄化成 IAcadEntity）
    # -----------------------------------------------------------------------

    def _finish_dim(self, ent: Any, layer: str | None, text_override: str | None) -> dict[str, Any]:
        """给新建标注统一收尾：归属图层、文字替代、回报测量值。在 COM 线程内调用。"""
        d = dynamic.Dispatch(ent)
        if layer:
            d.Layer = layer
        if text_override:
            d.TextOverride = str(text_override)
        return {
            "handle": self._safe(lambda: str(ent.Handle), default=""),
            "object_name": self._safe(lambda: str(ent.ObjectName), default=""),
            "measurement": self._safe(lambda: round(float(d.Measurement), 6), default=None),
        }

    def setup_text_style(
        self,
        name: str = "DSH-CN",
        font: str = CN_FONT,
        bigfont: str = CN_BIGFONT,
        doc_name: str | None = None,
    ) -> dict[str, Any]:
        """建/复用中文工程文字样式（gbenor.shx + gbcbig.shx），否则中文会显示成 ???。"""
        if not name or not name.strip():
            raise CadError("文字样式名不能为空。")

        def _do():
            doc = self._document(doc_name)
            styles = doc.TextStyles
            existing = None
            try:
                existing = styles.Item(name)
            except Exception:
                existing = None
            created = existing is None
            style = existing if existing is not None else styles.Add(name)
            # 早绑定对象没有 fontFile/bigFontFile 属性，必须走动态 Dispatch
            d = dynamic.Dispatch(style)
            d.fontFile = font
            d.bigFontFile = bigfont
            return {"created": created}

        info = self._call(f"建文字样式 {name}", _do)
        return {
            "ok": True, "style": name, "font": font, "bigfont": bigfont,
            **info, "document": self._last_doc_name,
        }

    def setup_gb_dim_style(
        self,
        text_style: str = "DSH-CN",
        dim_style: str = "GB-机械",
        dim_vars: dict[str, Any] | None = None,
        doc_name: str | None = None,
    ) -> dict[str, Any]:
        """
        建立并激活国标标注样式。

        顺序是实测定稿的（probe_v11b）：
          建中文文字样式 → 建/激活 DimStyle → 设 DIM* 变量 → DimStyle.CopyFrom(doc)
        反序会让变量落到旧样式，尺寸文字仍是 0.18 / Standard。
        """
        # 文字样式必须先存在，DIMTXSTY 才设得上
        self.setup_text_style(text_style, doc_name=doc_name)
        vars_ = dict(GB_DIM_VARS)
        if dim_vars:
            vars_.update(dim_vars)
        vars_["DIMTXSTY"] = text_style

        def _do():
            doc = self._document(doc_name)
            styles = doc.DimStyles
            existing = None
            try:
                existing = styles.Item(dim_style)
            except Exception:
                existing = None
            created = existing is None
            ds = existing if existing is not None else styles.Add(dim_style)
            # ① 先激活目标样式
            doc.ActiveDimStyle = ds
            # ② 再写变量（此时写进的是刚激活的样式）
            failed = []
            for key, value in vars_.items():
                try:
                    doc.SetVariable(key, value)
                except Exception:  # noqa: BLE001  个别变量在部分版本不存在
                    failed.append(key)
            # ③ 固化到样式记录里，切走再切回也不会丢
            try:
                ds.CopyFrom(doc)
            except Exception as exc:  # noqa: BLE001
                failed.append(f"CopyFrom:{exc}")
            return {"created": created, "vars_set": len(vars_), "failed": failed}

        info = self._call(f"建标注样式 {dim_style}", _do)
        return {
            "ok": True, "dim_style": dim_style, "text_style": text_style,
            **info, "document": self._last_doc_name,
        }

    def add_dim_linear(
        self,
        p1: Sequence[float],
        p2: Sequence[float],
        dim_line_point: Sequence[float],
        angle_deg: float = 0.0,
        text_override: str = "",
        layer: str | None = None,
        doc_name: str | None = None,
    ) -> dict[str, Any]:
        """线性标注（水平/垂直/倾斜）。angle_deg 是尺寸线角度（度），0=水平、90=竖直。"""
        layer = self._layer_for(layer, "dim")
        if layer:
            self.ensure_layer(layer, doc_name=doc_name)

        def _do():
            doc = self._document(doc_name)
            ent = doc.ModelSpace.AddDimRotated(
                _pt(*p1[:3]), _pt(*p2[:3]), _pt(*dim_line_point[:3]),
                math.radians(float(angle_deg)),
            )
            return self._finish_dim(ent, layer, text_override)

        info = self._call("线性标注", _do)
        return {"ok": True, "type": "linear", "angle_deg": angle_deg,
                "layer": layer, **info, "document": self._last_doc_name}

    def add_dim_aligned(
        self,
        p1: Sequence[float],
        p2: Sequence[float],
        dim_line_point: Sequence[float],
        text_override: str = "",
        layer: str | None = None,
        doc_name: str | None = None,
    ) -> dict[str, Any]:
        """对齐标注：尺寸线平行于 p1-p2 连线（用于斜边）。"""
        layer = self._layer_for(layer, "dim")
        if layer:
            self.ensure_layer(layer, doc_name=doc_name)

        def _do():
            doc = self._document(doc_name)
            ent = doc.ModelSpace.AddDimAligned(
                _pt(*p1[:3]), _pt(*p2[:3]), _pt(*dim_line_point[:3]),
            )
            return self._finish_dim(ent, layer, text_override)

        info = self._call("对齐标注", _do)
        return {"ok": True, "type": "aligned", "layer": layer,
                **info, "document": self._last_doc_name}

    def add_dim_radial(
        self,
        center: Sequence[float],
        chord_point: Sequence[float],
        leader_length: float = 10.0,
        text_override: str = "",
        layer: str | None = None,
        doc_name: str | None = None,
    ) -> dict[str, Any]:
        """半径标注。chord_point 是圆周上的取点，leader_length 是引线长度。"""
        layer = self._layer_for(layer, "dim")
        if layer:
            self.ensure_layer(layer, doc_name=doc_name)

        def _do():
            doc = self._document(doc_name)
            ent = doc.ModelSpace.AddDimRadial(
                _pt(*center[:3]), _pt(*chord_point[:3]), float(leader_length),
            )
            return self._finish_dim(ent, layer, text_override)

        info = self._call("半径标注", _do)
        return {"ok": True, "type": "radial", "layer": layer,
                **info, "document": self._last_doc_name}

    def add_dim_diametric(
        self,
        chord1: Sequence[float],
        chord2: Sequence[float],
        leader_length: float = 10.0,
        text_override: str = "",
        layer: str | None = None,
        doc_name: str | None = None,
    ) -> dict[str, Any]:
        """直径标注。chord1/chord2 是直径两端点。"""
        layer = self._layer_for(layer, "dim")
        if layer:
            self.ensure_layer(layer, doc_name=doc_name)

        def _do():
            doc = self._document(doc_name)
            ent = doc.ModelSpace.AddDimDiametric(
                _pt(*chord1[:3]), _pt(*chord2[:3]), float(leader_length),
            )
            return self._finish_dim(ent, layer, text_override)

        info = self._call("直径标注", _do)
        return {"ok": True, "type": "diametric", "layer": layer,
                **info, "document": self._last_doc_name}

    def add_dim_angular(
        self,
        vertex: Sequence[float],
        p1: Sequence[float],
        p2: Sequence[float],
        text_point: Sequence[float],
        text_override: str = "",
        layer: str | None = None,
        doc_name: str | None = None,
    ) -> dict[str, Any]:
        """角度标注（两条线之间的夹角）。"""
        layer = self._layer_for(layer, "dim")
        if layer:
            self.ensure_layer(layer, doc_name=doc_name)

        def _do():
            doc = self._document(doc_name)
            ent = doc.ModelSpace.AddDimAngular(
                _pt(*vertex[:3]), _pt(*p1[:3]), _pt(*p2[:3]), _pt(*text_point[:3]),
            )
            return self._finish_dim(ent, layer, text_override)

        info = self._call("角度标注", _do)
        return {"ok": True, "type": "angular", "layer": layer,
                **info, "document": self._last_doc_name}

    def add_dim_ordinate(
        self,
        definition_point: Sequence[float],
        leader_end_point: Sequence[float],
        use_x_axis: bool = True,
        text_override: str = "",
        layer: str | None = None,
        doc_name: str | None = None,
    ) -> dict[str, Any]:
        """坐标标注。use_x_axis=True 标 X 坐标，False 标 Y 坐标。"""
        layer = self._layer_for(layer, "dim")
        if layer:
            self.ensure_layer(layer, doc_name=doc_name)

        def _do():
            doc = self._document(doc_name)
            ent = doc.ModelSpace.AddDimOrdinate(
                _pt(*definition_point[:3]), _pt(*leader_end_point[:3]), bool(use_x_axis),
            )
            return self._finish_dim(ent, layer, text_override)

        info = self._call("坐标标注", _do)
        return {"ok": True, "type": "ordinate", "use_x_axis": bool(use_x_axis),
                "layer": layer, **info, "document": self._last_doc_name}

    # -----------------------------------------------------------------------
    # V1.1-2 机械标准图层自动化
    # -----------------------------------------------------------------------

    def setup_standard_layers(self, doc_name: str | None = None) -> dict[str, Any]:
        """一次性建好 GB_LAYERS 里全部机械标准图层（含颜色/线宽/线型）。幂等。"""
        created = []
        for name, spec in GB_LAYERS.items():
            self.ensure_layer(name, color=spec["color"], doc_name=doc_name)
            created.append(name)
        return {"ok": True, "layers": created, "count": len(created),
                "document": self._last_doc_name}

    def set_layer_state(
        self,
        name: str,
        lock: bool | None = None,
        freeze: bool | None = None,
        layer_on: bool | None = None,
        color: str | int | None = None,
        lineweight: int | None = None,
        active: bool = False,
        doc_name: str | None = None,
    ) -> dict[str, Any]:
        """批量改图层状态：锁定/冻结/开关/颜色/线宽/设为当前层。"""
        color_index = self._resolve_color(color) if color is not None else None

        def _do():
            doc = self._document(doc_name)
            try:
                lay = doc.Layers.Item(name)
            except Exception as exc:  # noqa: BLE001
                raise CadError(f"找不到图层 '{name}'：{exc}") from exc
            d = dynamic.Dispatch(lay)
            applied: dict[str, Any] = {}
            failed: dict[str, str] = {}

            def _set(key, prop, value):
                try:
                    setattr(d, prop, value)
                    applied[key] = value
                except Exception as exc:  # noqa: BLE001
                    failed[key] = str(exc)

            if lock is not None:
                _set("lock", "Lock", bool(lock))
            if freeze is not None:
                _set("freeze", "Freeze", bool(freeze))
            if layer_on is not None:
                _set("layer_on", "LayerOn", bool(layer_on))
            if color_index is not None:
                _set("color", "color", color_index)
            if lineweight is not None:
                _set("lineweight", "Lineweight", int(lineweight))
            if active:
                try:
                    doc.ActiveLayer = lay
                    applied["active"] = True
                except Exception as exc:  # noqa: BLE001  当前层不能冻结/关闭
                    failed["active"] = str(exc)
            return {"applied": applied, "failed": failed}

        info = self._call(f"设置图层状态 {name}", _do)
        return {"ok": True, "layer": name, **info, "document": self._last_doc_name}

    def set_active_layer(self, name: str, doc_name: str | None = None) -> dict[str, Any]:
        """把某个图层设为当前层（后续不指定 layer 的图元自动归属到这里）。"""
        return self.set_layer_state(name, active=True, doc_name=doc_name)

    def setup_mechanical_environment(self, doc_name: str | None = None) -> dict[str, Any]:
        """
        V1.1 交付标准「一句话完成分层规范」：一次调用建好全部标准图层 +
        中文文字样式 + 国标标注样式 + 常用文档变量（mm / 线宽显示 / 关联标注）。
        """
        layers = self.setup_standard_layers(doc_name)
        text = self.setup_text_style(doc_name=doc_name)
        dim = self.setup_gb_dim_style(doc_name=doc_name)

        def _do():
            doc = self._document(doc_name)
            for key, value in {
                "INSUNITS": 4,      # 毫米
                "LWDISPLAY": 1,     # 显示线宽
                "MEASUREMENT": 1,   # 公制
                "DIMASSOC": 2,      # 关联标注
                "LTSCALE": 1.0,
            }.items():
                try:
                    doc.SetVariable(key, value)
                except Exception:  # noqa: BLE001
                    pass
            return True

        self._call("设置文档变量", _do)
        # 交付标准「自动分层规范」：环境建好后，后续省略 layer 的调用按大类自动归属。
        self.auto_layer_by_kind = True
        return {
            "ok": True,
            "layers": layers["layers"],
            "text_style": text["style"],
            "dim_style": dim["dim_style"],
            "auto_layer_by_kind": True,
            "document": self._last_doc_name,
        }

    # -----------------------------------------------------------------------
    # V1.1-3 块与属性块
    # -----------------------------------------------------------------------

    @staticmethod
    def _find_by_handle(doc: Any, handle: str) -> Any | None:
        """在模型空间按句柄找图元；返回 dynamic.Dispatch 包装（避免 IAcadEntity 窄化）。"""
        for i in range(int(doc.ModelSpace.Count)):
            ent = doc.ModelSpace.Item(i)
            try:
                if str(ent.Handle) == handle:
                    return dynamic.Dispatch(ent)
            except Exception:
                continue
        return None

    def list_blocks(self, doc_name: str | None = None) -> dict[str, Any]:
        """列出图纸中已定义的块名（过滤掉 *Model_Space / *Paper_Space 内建块）。"""

        def _do():
            doc = self._document(doc_name)
            out = []
            for i in range(int(doc.Blocks.Count)):
                blk = doc.Blocks.Item(i)
                name = self._safe(lambda b=blk: str(b.Name), default="")
                if name and not name.startswith("*"):
                    out.append(name)
            return {"count": len(out), "blocks": out}

        result = self._call("读取块定义列表", _do)
        return {"ok": True, **result, "document": self._last_doc_name}

    def create_block(
        self,
        name: str,
        base_point: Sequence[float] = (0.0, 0.0),
        entities: list[dict[str, Any]] | None = None,
        attributes: list[dict[str, Any]] | None = None,
        doc_name: str | None = None,
    ) -> dict[str, Any]:
        """
        新建块定义。entities 支持 line/circle/arc/text/polyline；
        attributes 是 [{"tag","value","prompt","height","position","mode"}]。
        已存在则直接返回（幂等，不覆盖用户已有块）。
        """
        if not name or not name.strip():
            raise CadError("块名不能为空。")
        ents = list(entities or [])
        attrs = list(attributes or [])

        def _do():
            doc = self._document(doc_name)
            try:
                doc.Blocks.Item(name)
                return {"created": False, "reason": "块已存在，未改动"}
            except Exception:
                pass
            blk = doc.Blocks.Add(_pt(*base_point[:3]), str(name))
            added: list[str] = []
            for e in ents:
                kind = str(e.get("type", "")).strip().lower()
                if kind == "line":
                    blk.AddLine(_pt(*e["start"][:3]), _pt(*e["end"][:3]))
                elif kind == "circle":
                    blk.AddCircle(_pt(*e["center"][:3]), float(e["radius"]))
                elif kind == "arc":
                    blk.AddArc(
                        _pt(*e["center"][:3]), float(e["radius"]),
                        math.radians(float(e["start_angle_deg"])),
                        math.radians(float(e["end_angle_deg"])),
                    )
                elif kind == "text":
                    blk.AddText(str(e["text"]), _pt(*e["position"][:3]),
                                float(e.get("height", 2.5)))
                elif kind == "polyline":
                    flat: list[float] = []
                    for p in e["points"]:
                        flat.extend([float(p[0]), float(p[1])])
                    pl = blk.AddLightWeightPolyline(_flat(flat))
                    pl.Closed = bool(e.get("closed", False))
                else:
                    continue
                added.append(kind)
            for a in attrs:
                blk.AddAttribute(
                    float(a.get("height", 2.5)),
                    int(a.get("mode", 0)),
                    str(a.get("prompt", "")),
                    _pt(*a["position"][:3]),
                    str(a["tag"]),
                    str(a.get("value", "")),
                )
            return {"created": True, "entities_added": added, "attributes_added": len(attrs)}

        info = self._call(f"新建块 {name}", _do)
        return {"ok": True, "block": name, **info, "document": self._last_doc_name}

    def insert_block(
        self,
        name: str,
        position: Sequence[float],
        xscale: float = 1.0,
        yscale: float = 1.0,
        zscale: float = 1.0,
        rotation_deg: float = 0.0,
        attributes: dict[str, str] | None = None,
        layer: str | None = None,
        doc_name: str | None = None,
    ) -> dict[str, Any]:
        """把块插入模型空间，并可按 tag 覆盖属性值（属性块核心用法）。"""
        layer = self._layer_for(layer)
        if layer:
            self.ensure_layer(layer, doc_name=doc_name)
        attr_map = {str(k): str(v) for k, v in (attributes or {}).items()}

        def _do():
            doc = self._document(doc_name)
            try:
                doc.Blocks.Item(name)
            except Exception as exc:  # noqa: BLE001
                raise CadError(f"块定义 '{name}' 不存在：{exc}") from exc
            ref = doc.ModelSpace.InsertBlock(
                _pt(*position[:3]), str(name),
                float(xscale), float(yscale), float(zscale),
                math.radians(float(rotation_deg)),
            )
            d = dynamic.Dispatch(ref)
            if layer:
                d.Layer = layer
            changed = []
            if attr_map:
                for a in d.GetAttributes():
                    ad = dynamic.Dispatch(a)
                    tag = str(ad.TagString)
                    if tag in attr_map:
                        ad.TextString = attr_map[tag]
                        changed.append(tag)
            return {
                "handle": self._safe(lambda: str(ref.Handle), default=""),
                "effective_name": self._safe(lambda: str(d.EffectiveName), default=name),
                "attributes_changed": changed,
            }

        info = self._call(f"插入块 {name}", _do)
        return {"ok": True, "block": name, "layer": layer,
                **info, "document": self._last_doc_name}

    def list_block_attributes(self, handle: str, doc_name: str | None = None) -> dict[str, Any]:
        """读某个块参照的全部属性 (tag, value)。"""

        def _do():
            doc = self._document(doc_name)
            ref = self._find_by_handle(doc, handle)
            if ref is None:
                raise CadError(f"找不到句柄 '{handle}' 对应的图元。")
            out = []
            for a in ref.GetAttributes():
                ad = dynamic.Dispatch(a)
                out.append({
                    "tag": self._safe(lambda x=ad: str(x.TagString), default=""),
                    "value": self._safe(lambda x=ad: str(x.TextString), default=""),
                    "prompt": self._safe(lambda x=ad: str(x.PromptString), default=""),
                })
            return {"effective_name": self._safe(lambda: str(ref.EffectiveName), default=""),
                    "attributes": out, "count": len(out)}

        result = self._call(f"读取块属性 {handle}", _do)
        return {"ok": True, "handle": handle, **result, "document": self._last_doc_name}

    def set_block_attribute(
        self,
        handle: str,
        tag: str,
        value: str,
        doc_name: str | None = None,
    ) -> dict[str, Any]:
        """改某个块参照里指定 tag 的属性值。"""

        def _do():
            doc = self._document(doc_name)
            ref = self._find_by_handle(doc, handle)
            if ref is None:
                raise CadError(f"找不到句柄 '{handle}' 对应的图元。")
            changed = []
            for a in ref.GetAttributes():
                ad = dynamic.Dispatch(a)
                if str(ad.TagString) == str(tag):
                    old = self._safe(lambda x=ad: str(x.TextString), default="")
                    ad.TextString = str(value)
                    changed.append({"tag": str(tag), "old": old, "new": str(value)})
            if not changed:
                raise CadError(f"图元 {handle} 上没有标签 '{tag}' 的属性。")
            return {"changed": changed}

        info = self._call(f"改块属性 {handle}.{tag}", _do)
        return {"ok": True, "handle": handle, **info, "document": self._last_doc_name}

    def replace_block(
        self,
        old_name: str,
        new_name: str,
        doc_name: str | None = None,
    ) -> dict[str, Any]:
        """批量替换：把所有 old_name 的块参照换成 new_name，保留插入点/比例/旋转。"""

        def _do():
            doc = self._document(doc_name)
            try:
                doc.Blocks.Item(new_name)
            except Exception as exc:  # noqa: BLE001
                raise CadError(f"目标块 '{new_name}' 不存在：{exc}") from exc
            ms = doc.ModelSpace
            targets = []
            for i in range(int(ms.Count)):
                ent = ms.Item(i)
                objname = self._safe(lambda e=ent: str(e.ObjectName), default="")
                if objname != "AcDbBlockReference":
                    continue
                r = dynamic.Dispatch(ent)
                if self._safe(lambda: str(r.EffectiveName), default="") == old_name:
                    targets.append(r)
            replaced = 0
            for r in targets:
                ins = list(self._safe(lambda: r.InsertionPoint, default=(0.0, 0.0, 0.0)))
                xs = float(self._safe(lambda: r.XScaleFactor, default=1.0))
                ys = float(self._safe(lambda: r.YScaleFactor, default=1.0))
                zs = float(self._safe(lambda: r.ZScaleFactor, default=1.0))
                rot = float(self._safe(lambda: r.Rotation, default=0.0))
                r.Delete()
                ms.InsertBlock(_pt(*ins[:3]), new_name, xs, ys, zs, rot)
                replaced += 1
            return {"matched": len(targets), "replaced": replaced}

        info = self._call(f"批量替换块 {old_name}→{new_name}", _do)
        return {"ok": True, "old_name": old_name, "new_name": new_name,
                **info, "document": self._last_doc_name}

    # -----------------------------------------------------------------------
    # V1.1-4 文字 / MText / 公差
    # -----------------------------------------------------------------------

    def add_mtext(
        self,
        text: str,
        position: Sequence[float],
        width: float = 100.0,
        height: float = 3.5,
        style: str = "",
        attachment_point: int = 1,
        layer: str | None = None,
        doc_name: str | None = None,
    ) -> dict[str, Any]:
        """多行文字（MText）。\\P 换行；配合 DSH-CN 样式可写中文技术要求。"""  # noqa: E501
        layer = self._layer_for(layer, "text")
        if layer:
            self.ensure_layer(layer, doc_name=doc_name)

        def _do():
            doc = self._document(doc_name)
            ent = doc.ModelSpace.AddMText(_pt(*position[:3]), float(width), str(text))
            d = dynamic.Dispatch(ent)
            d.Height = float(height)
            if attachment_point is not None:
                d.AttachmentPoint = int(attachment_point)
            if style:
                try:
                    d.StyleName = style
                except Exception:  # noqa: BLE001  样式不存在时忽略
                    pass
            if layer:
                d.Layer = layer
            return {"handle": self._safe(lambda: str(ent.Handle), default=""),
                    "object_name": self._safe(lambda: str(ent.ObjectName), default="")}

        info = self._call("写多行文字", _do)
        return {"ok": True, "type": "mtext", "text": str(text), "width": float(width),
                "height": float(height), "style": style, "layer": layer,
                **info, "document": self._last_doc_name}

    def add_tolerance(
        self,
        text: str,
        position: Sequence[float],
        direction: Sequence[float] = (1.0, 0.0, 0.0),
        layer: str | None = None,
        doc_name: str | None = None,
    ) -> dict[str, Any]:
        """形位公差框格（AcDbFcf）。text 用 AutoCAD 控制码，如 %%c40%%p0.01。"""
        layer = self._layer_for(layer, "dim")
        if layer:
            self.ensure_layer(layer, doc_name=doc_name)

        def _do():
            doc = self._document(doc_name)
            ent = doc.ModelSpace.AddTolerance(
                str(text), _pt(*position[:3]), _pt(*direction[:3]),
            )
            d = dynamic.Dispatch(ent)
            if layer:
                d.Layer = layer
            return {"handle": self._safe(lambda: str(ent.Handle), default=""),
                    "object_name": self._safe(lambda: str(ent.ObjectName), default="")}

        info = self._call("写形位公差", _do)
        return {"ok": True, "type": "tolerance", "text": str(text), "layer": layer,
                **info, "document": self._last_doc_name}

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
