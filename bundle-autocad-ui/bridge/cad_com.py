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

import itertools
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
#: V1.2：批量引擎给 SelectionSet 生成全局唯一名（probe_v12f 三连过的前提）。
_SS_COUNTER = itertools.count(1)

DEFAULT_LAYER_BY_KIND: dict[str, str] = {
    "outline": "轮廓",
    "center": "中心线",
    "hidden": "虚线",
    "dim": "标注",
    "text": "文字",
    "hatch": "剖面线",
}

#: V2.0-④：工程活动日志环形缓冲上限。_call() 是所有 COM 操作的唯一通道，
#: 在此单点埋点即可覆盖全部 MCP 工具（包括未来新增的），无需逐个工具改代码。
ACTIVITY_LOG_CAPACITY = 2000

#: V2.0-②：AutoCAD 另存为 .dwt 模板的 SaveAsType 枚举值（ac2000_dwt）。
#: probe_v20d.py 实测：SaveAs(path, 14) 能生成带自定义图层/文字样式的 .dwt，
#: 再 Documents.Add(该 dwt) 后自定义规范确实随模板存活。
DWT_SAVEAS_TYPE = 14


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

        # V2.0-④：工程活动日志（内存环形缓冲）。_call() 是所有 COM 操作的唯一
        # 通道，因此在那里单点埋点即可回溯「每一步绘图」，无需逐工具改代码。
        # 用独立锁，避免与 _lock（保护 COM 生命周期）互相争用。
        self._activity_lock = threading.Lock()
        self._activity_seq: int = 0
        self._activity_log: list[dict[str, Any]] = []

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

    def _call(
        self,
        label: str,
        fn,
        *,
        retries: int | None = None,
        retry_delay_s: float | None = None,
        timeout_s: float | None = None,
    ):
        """
        所有 COM 操作都必须经过这里。

        逻辑：
          - 整段丢到 _ComWorker 的专用 STA 线程执行（见类注释），
            这样无论 MCP 在哪个线程调用 tools/call 都安全；
          - 撞上"CAD 忙" → 睡眠后重试，最多 retry_times 次；
          - 重试耗尽 → 抛 CadBusyError（并提示可能有弹窗）；
          - COM 断开 → 抛 CadDisconnectedError；
          - 其他错误 → 抛 CadError，附上 label 方便定位是哪一步炸的。

        V1.2：retries / retry_delay_s / timeout_s 三个可选覆盖只给批量引擎用。
        实测（live_v12_test）：连续 DXF/Wblock/PDF 导出后，CAD 后台收尾会忙
        2~5 秒，默认的 retry_times=2 × 0.8s 窗口不够，导致第 3 张图纸偶发失败。
        不传时行为与 V1.1 完全一致（向下兼容）。
        """
        n = int(retries) if retries is not None else self.retry_times
        delay = float(retry_delay_s) if retry_delay_s is not None else self.retry_delay_ms / 1000.0

        def _attempt() -> Any:
            if self._app is None:
                raise CadDisconnectedError("尚未连接 AutoCAD，请先调用 connect。")

            last: Exception | None = None
            for attempt in range(1, n + 1):
                try:
                    self._ensure_com()
                    return fn()
                except pywintypes.com_error as exc:
                    hresult = _hresult(exc)
                    if hresult in BUSY_HRESULTS:
                        last = exc
                        if attempt < n:
                            time.sleep(delay)
                            continue
                        raise CadBusyError(
                            f"AutoCAD 一直处于忙碌状态（{label} 重试 {n} 次仍失败）。\n"
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

        # V2.0-④：单点埋点。所有 COM 操作（含未来新增工具）都经过这里，
        # 因此一次记录即可让「每一步绘图可回溯 / 出错精准定位」覆盖全部能力。
        t0 = time.time()
        mono0 = time.monotonic()
        ok = False
        err: str | None = None
        try:
            result = self._on_com(_attempt, timeout_s=timeout_s)
            ok = True
            return result
        except Exception as exc:  # noqa: BLE001
            err = f"{type(exc).__name__}: {exc}"
            raise
        finally:
            self._record_activity(
                label=label,
                ok=ok,
                error=err,
                started_at=t0,
                duration_ms=round((time.monotonic() - mono0) * 1000.0, 2),
            )

    # -- V2.0-④ 工程活动日志 ----------------------------------------------

    def _record_activity(
        self,
        *,
        label: str,
        ok: bool,
        error: str | None,
        started_at: float,
        duration_ms: float,
    ) -> None:
        """把一次 _call 记入内存环形缓冲。任何异常都不允许影响主流程。"""
        try:
            with self._activity_lock:
                self._activity_seq += 1
                rec: dict[str, Any] = {
                    "seq": self._activity_seq,
                    "action": label,
                    "ok": ok,
                    "error": error,
                    "ts": round(started_at, 3),
                    "time": time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(started_at)),
                    "duration_ms": duration_ms,
                    "thread": threading.get_ident(),
                }
                self._activity_log.append(rec)
                if len(self._activity_log) > ACTIVITY_LOG_CAPACITY:
                    del self._activity_log[: len(self._activity_log) - ACTIVITY_LOG_CAPACITY]
        except Exception:  # noqa: BLE001
            pass

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

    # -----------------------------------------------------------------------
    # V1.2 非标自动化：机械辅助 / 外部参照 / BOM / 批量引擎
    #
    # 所有 COM 操作仍走 _call() 的 STA 单线程队列；坐标一律 _pt()/dbl_array()；
    # 角度对外用「度」，进 COM 前 math.radians()。下列签名全部来自本机
    # AutoCAD 2024 的 live probe（bridge/probe_v12*.py），不靠猜。
    # -----------------------------------------------------------------------

    @staticmethod
    def _fsize(path: str) -> int:
        try:
            return os.path.getsize(path)
        except OSError:
            return -1

    @staticmethod
    def _select_all(doc: Any) -> Any:
        """建一个全局唯一名的 SelectionSet 并全选。

        实测（probe_v12f）：必须用 col.Add() 返回的「裸对象」调 Select(5)；
        dynamic.Dispatch(ss) 会丢掉 Select。唯一名是连续批量 3 连过的前提。
        """
        col = doc.SelectionSets
        name = f"DSH_BATCH_{next(_SS_COUNTER)}"
        s = col.Add(name)
        try:
            s.Select(5)  # acSelectionSetAll
        except Exception:
            dynamic.Dispatch(s).Select(5)
        return s

    @staticmethod
    def _drawing_bounds(doc: Any):
        """遍历图元 GetBoundingBox 求全图范围（比 EXTMIN/EXTMAX 稳，后者需先 Zoom）。"""
        lo = [float("inf")] * 3
        hi = [float("-inf")] * 3
        found = False
        ms = doc.ModelSpace
        for i in range(int(ms.Count)):
            try:
                e = dynamic.Dispatch(ms.Item(i))
                mn, mx = e.GetBoundingBox()
                for k in range(min(3, len(mn), len(mx))):
                    a, b = float(mn[k]), float(mx[k])
                    lo[k] = min(lo[k], a, b)
                    hi[k] = max(hi[k], a, b)
                found = True
            except Exception:
                continue
        return (lo, hi) if found else None

    @staticmethod
    def _set_cell(table: Any, row: int, col: int, text: str) -> None:
        s = str(text)
        try:
            table.SetText(int(row), int(col), s)
        except Exception:
            table.SetTextString(int(row), int(col), 0, s)

    @staticmethod
    def _xref_block(doc: Any, name: str) -> Any:
        for i in range(int(doc.Blocks.Count)):
            b = doc.Blocks.Item(i)
            try:
                if str(b.Name) == str(name):
                    return dynamic.Dispatch(b)
            except Exception:
                continue
        raise CadError(f"找不到块/外部参照 '{name}'。")

    # -- V1.2-4 机械常用辅助功能 -------------------------------------------

    def add_centerlines(
        self,
        corner1: Sequence[float],
        corner2: Sequence[float],
        extension: float = 5.0,
        layer: str | None = None,
        doc_name: str | None = None,
    ) -> dict[str, Any]:
        """按矩形两角自动生成 H/V 两条居中中心线（超出轮廓 extension）。"""
        layer = self._layer_for(layer, "center")
        if layer:
            self.ensure_layer(layer, doc_name=doc_name)

        def _do():
            doc = self._document(doc_name)
            ms = doc.ModelSpace
            x1, y1 = float(corner1[0]), float(corner1[1])
            x2, y2 = float(corner2[0]), float(corner2[1])
            minx, maxx = min(x1, x2), max(x1, x2)
            miny, maxy = min(y1, y2), max(y1, y2)
            cx, cy = (minx + maxx) / 2.0, (miny + maxy) / 2.0
            e = float(extension)
            h = ms.AddLine(_pt(minx - e, cy), _pt(maxx + e, cy))
            v = ms.AddLine(_pt(cx, miny - e), _pt(cx, maxy + e))
            for ent in (h, v):
                if layer:
                    try:
                        dynamic.Dispatch(ent).Layer = layer
                    except Exception:
                        pass
            return {"center": [cx, cy],
                    "handles": [self._safe(lambda o=h: str(o.Handle), default=""),
                                self._safe(lambda o=v: str(o.Handle), default="")]}

        info = self._call("生成中心线", _do)
        return {"ok": True, "layer": layer, **info, "document": self._last_doc_name}

    def mirror_entities(
        self,
        handles: Sequence[str],
        axis_point1: Sequence[float],
        axis_point2: Sequence[float],
        keep_source: bool = True,
        layer: str | None = None,
        doc_name: str | None = None,
    ) -> dict[str, Any]:
        """按镜像轴补全图形；keep_source=False 可只留镜像结果。"""
        layer = self._layer_for(layer, "outline")
        if layer:
            self.ensure_layer(layer, doc_name=doc_name)

        def _do():
            doc = self._document(doc_name)
            p1 = _pt(*axis_point1[:3])
            p2 = _pt(*axis_point2[:3])
            created, errors = [], []
            for h in handles:
                try:
                    e = dynamic.Dispatch(doc.HandleToObject(str(h)))
                    new = e.Mirror(p1, p2)
                    if layer:
                        try:
                            dynamic.Dispatch(new).Layer = layer
                        except Exception:
                            pass
                    created.append(self._safe(lambda o=new: str(o.Handle), default=""))
                    if not keep_source:
                        try:
                            e.Delete()
                        except Exception:
                            pass
                except Exception as exc:
                    errors.append(f"{h}: {exc}")
            return {"created": created, "errors": errors}

        info = self._call("镜像图元", _do)
        return {"ok": True, "keep_source": bool(keep_source), "layer": layer,
                **info, "document": self._last_doc_name}

    def offset_entity(
        self,
        handle: str,
        distance: float,
        layer: str | None = None,
        doc_name: str | None = None,
    ) -> dict[str, Any]:
        """等距偏移一个图元（距离可为负，表示反向）。"""
        layer = self._layer_for(layer, "outline")
        if layer:
            self.ensure_layer(layer, doc_name=doc_name)

        def _do():
            doc = self._document(doc_name)
            e = dynamic.Dispatch(doc.HandleToObject(str(handle)))
            r = e.Offset(float(distance))
            out = []
            for item in (r if isinstance(r, (tuple, list)) else (r,)):
                if layer:
                    try:
                        dynamic.Dispatch(item).Layer = layer
                    except Exception:
                        pass
                out.append(self._safe(lambda o=item: str(o.Handle), default=""))
            return {"created": out, "distance": float(distance)}

        info = self._call("偏移图元", _do)
        return {"ok": True, "layer": layer, **info, "document": self._last_doc_name}

    def array_rectangular(
        self,
        handle: str,
        rows: int,
        cols: int,
        dist_rows: float,
        dist_cols: float,
        levels: int = 1,
        dist_levels: float = 0.0,
        layer: str | None = None,
        doc_name: str | None = None,
    ) -> dict[str, Any]:
        """矩形阵列；返回新生成图元句柄（不含原图元）。"""
        layer = self._layer_for(layer, "outline")
        if layer:
            self.ensure_layer(layer, doc_name=doc_name)

        def _do():
            doc = self._document(doc_name)
            e = dynamic.Dispatch(doc.HandleToObject(str(handle)))
            r = e.ArrayRectangular(int(rows), int(cols), int(levels),
                                   float(dist_rows), float(dist_cols), float(dist_levels))
            out = []
            for item in (r if isinstance(r, (tuple, list)) else (r,)):
                if layer:
                    try:
                        dynamic.Dispatch(item).Layer = layer
                    except Exception:
                        pass
                out.append(self._safe(lambda o=item: str(o.Handle), default=""))
            return {"created": out, "rows": int(rows), "cols": int(cols)}

        info = self._call("矩形阵列", _do)
        return {"ok": True, "layer": layer, **info, "document": self._last_doc_name}

    def array_polar(
        self,
        handle: str,
        count: int,
        angle_deg: float = 360.0,
        center: Sequence[float] = (0.0, 0.0),
        layer: str | None = None,
        doc_name: str | None = None,
    ) -> dict[str, Any]:
        """环形阵列；angle_deg 用度（内部转弧度）。"""
        layer = self._layer_for(layer, "outline")
        if layer:
            self.ensure_layer(layer, doc_name=doc_name)

        def _do():
            doc = self._document(doc_name)
            e = dynamic.Dispatch(doc.HandleToObject(str(handle)))
            r = e.ArrayPolar(int(count), math.radians(float(angle_deg)),
                             _pt(*center[:3]))
            out = []
            for item in (r if isinstance(r, (tuple, list)) else (r,)):
                if layer:
                    try:
                        dynamic.Dispatch(item).Layer = layer
                    except Exception:
                        pass
                out.append(self._safe(lambda o=item: str(o.Handle), default=""))
            return {"created": out, "count": int(count), "angle_deg": float(angle_deg)}

        info = self._call("环形阵列", _do)
        return {"ok": True, "layer": layer, **info, "document": self._last_doc_name}

    def trim_at_points(
        self,
        points: Sequence[Sequence[float]],
        doc_name: str | None = None,
    ) -> dict[str, Any]:
        """批量修剪：SendCommand 跑 _.TRIM，接受全部剪切边，依次点选要剪掉的部分。

        实测（probe_v12d）：COM 没有 Trim/Extend/Break 实体方法，只能走命令；
        第一个 \\n\\n = 回车接受全部作为剪切边，末尾 \\n = 回车结束。
        """
        picks = "".join(f"{float(p[0])},{float(p[1])}\n" for p in (points or []))

        def _do():
            doc = self._document(doc_name)
            doc.SendCommand("_.TRIM\n\n" + picks + "\n")
            time.sleep(0.6)
            return {"picked": len(points or [])}

        info = self._call("批量修剪", _do)
        return {"ok": True, **info, "document": self._last_doc_name}

    # -- V1.2-3 外部参照管理 -----------------------------------------------

    def list_xrefs(self, doc_name: str | None = None) -> dict[str, Any]:
        """列出当前图纸里的外部参照块名与路径。"""
        def _do():
            doc = self._document(doc_name)
            out = []
            for i in range(int(doc.Blocks.Count)):
                try:
                    d = dynamic.Dispatch(doc.Blocks.Item(i))
                    if bool(d.IsXRef):
                        out.append({"name": str(d.Name), "path": str(d.Path)})
                except Exception:
                    continue
            return {"xrefs": out, "count": len(out)}

        info = self._call("列出外部参照", _do)
        return {"ok": True, **info, "document": self._last_doc_name}

    def attach_xref(
        self,
        path: str,
        name: str,
        position: Sequence[float] = (0.0, 0.0),
        xscale: float = 1.0,
        yscale: float = 1.0,
        zscale: float = 1.0,
        rotation_deg: float = 0.0,
        overlay: bool = False,
        doc_name: str | None = None,
    ) -> dict[str, Any]:
        """附着一个 DWG 外部参照。"""
        abs_path = os.path.abspath(str(path))

        def _do():
            doc = self._document(doc_name)
            if not os.path.isfile(abs_path):
                raise CadError(f"外部参照文件不存在：{abs_path}")
            ref = doc.ModelSpace.AttachExternalReference(
                abs_path, str(name), _pt(*position[:3]),
                float(xscale), float(yscale), float(zscale),
                math.radians(float(rotation_deg)), bool(overlay),
            )
            return {"handle": self._safe(lambda: str(ref.Handle), default=""),
                    "object_name": self._safe(lambda: str(ref.ObjectName), default=""),
                    "name": str(name), "path": abs_path}

        info = self._call("附着外部参照", _do)
        return {"ok": True, **info, "document": self._last_doc_name}

    def xref_action(
        self,
        name: str,
        action: str = "reload",
        new_path: str | None = None,
        doc_name: str | None = None,
    ) -> dict[str, Any]:
        """对外部参照执行 bind / unload / reload / detach / repair。

        实测（probe_v12c）：生命周期方法在 BLOCK 定义（doc.Blocks.Item(name)）
        上，不在 AcDbBlockReference 实体上。
        """
        act = str(action).lower()

        def _do():
            doc = self._document(doc_name)
            blk = self._xref_block(doc, name)
            if act == "bind":
                blk.Bind(False)
            elif act == "unload":
                blk.Unload()
            elif act == "reload":
                blk.Reload()
            elif act == "detach":
                blk.Detach()
            elif act in ("repair", "path", "set_path"):
                if not new_path:
                    raise CadError("repair 需要 new_path。")
                blk.Path = os.path.abspath(str(new_path))
            else:
                raise CadError(f"不支持的 xref action：{act}（可用 bind/unload/reload/detach/repair）")
            return {"name": str(name), "action": act,
                    "is_xref": self._safe(lambda: bool(blk.IsXRef), default=None),
                    "path": self._safe(lambda: str(blk.Path), default="")}

        info = self._call("外部参照操作", _do)
        return {"ok": True, **info, "document": self._last_doc_name}

    def repair_xref_paths(
        self,
        old_prefix: str,
        new_prefix: str,
        doc_name: str | None = None,
    ) -> dict[str, Any]:
        """批量路径修复：把以 old_prefix 开头的参照路径换成 new_prefix。"""
        def _do():
            doc = self._document(doc_name)
            changed = []
            for i in range(int(doc.Blocks.Count)):
                try:
                    d = dynamic.Dispatch(doc.Blocks.Item(i))
                    if not bool(d.IsXRef):
                        continue
                    old = str(d.Path)
                    if old.startswith(str(old_prefix)):
                        new = str(new_prefix) + old[len(str(old_prefix)):]
                        d.Path = new
                        changed.append({"name": str(d.Name), "old": old, "new": new})
                except Exception:
                    continue
            return {"changed": changed, "count": len(changed)}

        info = self._call("批量修复参照路径", _do)
        return {"ok": True, "old_prefix": str(old_prefix), "new_prefix": str(new_prefix),
                **info, "document": self._last_doc_name}

    # -- V1.2-1 明细表（BOM） ----------------------------------------------

    def count_blocks(self, doc_name: str | None = None) -> dict[str, Any]:
        """统计模型空间里的块参照数量，并取一份属性样例（BOM 数据源）。"""
        def _do():
            doc = self._document(doc_name)
            ms = doc.ModelSpace
            agg: dict[str, dict[str, Any]] = {}
            total = 0
            for i in range(int(ms.Count)):
                try:
                    e = dynamic.Dispatch(ms.Item(i))
                    if str(e.ObjectName) != "AcDbBlockReference":
                        continue
                    nm = str(e.EffectiveName)
                except Exception:
                    continue
                total += 1
                item = agg.setdefault(nm, {"name": nm, "count": 0, "attributes": {}})
                item["count"] += 1
                if not item["attributes"]:
                    try:
                        for a in e.GetAttributes():
                            ad = dynamic.Dispatch(a)
                            item["attributes"][str(ad.TagString)] = str(ad.TextString)
                    except Exception:
                        pass
            return {"total_refs": total, "distinct": len(agg),
                    "items": list(agg.values())}

        info = self._call("统计块", _do)
        return {"ok": True, **info, "document": self._last_doc_name}

    def generate_bom(
        self,
        position: Sequence[float] | None = None,
        columns: Sequence[str] | None = None,
        row_height: float = 8.0,
        col_width: float = 30.0,
        margin: float = 10.0,
        layer: str | None = None,
        doc_name: str | None = None,
    ) -> dict[str, Any]:
        """自动统计块 → 生成明细表表格；不给 position 时对齐到图纸右下角。"""
        layer = self._layer_for(layer, "text")
        if layer:
            self.ensure_layer(layer, doc_name=doc_name)
        cols = [str(c) for c in (columns or ["序号", "名称", "数量", "备注"])]

        def _do():
            doc = self._document(doc_name)
            ms = doc.ModelSpace
            agg: dict[str, dict[str, Any]] = {}
            total = 0
            for i in range(int(ms.Count)):
                try:
                    e = dynamic.Dispatch(ms.Item(i))
                    if str(e.ObjectName) != "AcDbBlockReference":
                        continue
                    nm = str(e.EffectiveName)
                except Exception:
                    continue
                total += 1
                item = agg.setdefault(nm, {"name": nm, "count": 0, "attributes": {}})
                item["count"] += 1
                if not item["attributes"]:
                    try:
                        for a in e.GetAttributes():
                            ad = dynamic.Dispatch(a)
                            item["attributes"][str(ad.TagString)] = str(ad.TextString)
                    except Exception:
                        pass
            items = list(agg.values())
            rows = 1 + len(items)
            if position:
                x, y = float(position[0]), float(position[1])
            else:
                bounds = self._drawing_bounds(doc)
                if bounds:
                    lo, hi = bounds
                    x = float(hi[0]) - len(cols) * float(col_width) - float(margin)
                    y = float(lo[1]) + rows * float(row_height) + float(margin)
                else:
                    x, y = 0.0, 0.0
            tbl = ms.AddTable(_pt(x, y), rows, len(cols),
                              float(row_height), float(col_width))
            t = dynamic.Dispatch(doc.HandleToObject(str(tbl.Handle)))
            for c, name in enumerate(cols):
                self._set_cell(t, 0, c, name)
            for r, item in enumerate(items, start=1):
                note = " ".join(f"{k}={v}" for k, v in item["attributes"].items())[:60]
                vals = [str(r), item["name"], str(item["count"]), note]
                for c in range(len(cols)):
                    self._set_cell(t, r, c, vals[c] if c < len(vals) else "")
            if layer:
                try:
                    t.Layer = layer
                except Exception:
                    pass
            return {"handle": self._safe(lambda: str(tbl.Handle), default=""),
                    "rows": rows, "columns": len(cols), "position": [x, y],
                    "total_refs": total, "items": items}

        info = self._call("生成明细表", _do)
        return {"ok": True, "layer": layer, **info, "document": self._last_doc_name}

    # -- V1.2-2 图纸批量处理引擎 -------------------------------------------

    def _unify_in_doc(self, doc: Any, style: str = "", layer: str = "") -> dict[str, Any]:
        if style:
            try:
                doc.TextStyles.Item(style)
            except Exception:
                s = doc.TextStyles.Add(style)
                ds = dynamic.Dispatch(s)
                try:
                    ds.fontFile = CN_FONT
                except Exception:
                    pass
                try:
                    ds.bigFontFile = CN_BIGFONT
                except Exception:
                    pass
        ms = doc.ModelSpace
        lc = sc = failed = 0
        for i in range(int(ms.Count)):
            try:
                e = dynamic.Dispatch(ms.Item(i))
                if layer:
                    e.Layer = layer
                    lc += 1
                if style and str(e.ObjectName) in ("AcDbText", "AcDbMText"):
                    e.StyleName = style
                    sc += 1
            except Exception:
                failed += 1
        return {"layer_changed": lc, "style_changed": sc, "failed": failed}

    def unify_styles(
        self,
        style: str = "",
        layer: str = "",
        doc_name: str | None = None,
    ) -> dict[str, Any]:
        """批量统一字体样式 / 图层（样式不存在时自动建中文大字体）。"""
        def _do():
            doc = self._document(doc_name)
            return self._unify_in_doc(doc, str(style or ""), str(layer or ""))

        info = self._call("统一字体/图层", _do)
        return {"ok": True, "style": style, "layer": layer,
                **info, "document": self._last_doc_name}

    def export_dxf(self, path: str, doc_name: str | None = None) -> dict[str, Any]:
        """当前图纸导出 DXF。实测 FileName 必须不带扩展名，AutoCAD 自己补 .dxf。"""
        base = os.path.abspath(str(path))
        if base.lower().endswith(".dxf"):
            base = base[:-4]
        os.makedirs(os.path.dirname(base), exist_ok=True)

        def _do():
            doc = self._document(doc_name)
            s = self._select_all(doc)
            try:
                doc.Export(base, "DXF", s)
            finally:
                try:
                    s.Delete()
                except Exception:
                    pass
            return True

        self._call("导出 DXF", _do)
        final = base + ".dxf"
        return {"ok": True, "format": "dxf", "path": final,
                "size": self._fsize(final), "document": self._last_doc_name}

    def export_dwg(self, path: str, doc_name: str | None = None) -> dict[str, Any]:
        """当前图纸用 Wblock 导出 DWG（全选图元）。"""
        final = os.path.abspath(str(path))
        if not final.lower().endswith(".dwg"):
            final += ".dwg"
        os.makedirs(os.path.dirname(final), exist_ok=True)

        def _do():
            doc = self._document(doc_name)
            s = self._select_all(doc)
            try:
                doc.Wblock(final, s)
            finally:
                try:
                    s.Delete()
                except Exception:
                    pass
            return True

        self._call("导出 DWG", _do)
        return {"ok": True, "format": "dwg", "path": final,
                "size": self._fsize(final), "document": self._last_doc_name}

    def export_pdf(self, path: str, doc_name: str | None = None) -> dict[str, Any]:
        """当前图纸高清打印成 PDF（DWG To PDF.pc3 + 按范围）。"""
        final = os.path.abspath(str(path))
        if not final.lower().endswith(".pdf"):
            final += ".pdf"
        os.makedirs(os.path.dirname(final), exist_ok=True)

        def _do():
            doc = self._document(doc_name)
            lay = dynamic.Dispatch(doc.ActiveLayout)
            old = self._safe(lambda: str(lay.ConfigName), default="")
            lay.ConfigName = "DWG To PDF.pc3"
            doc.SetVariable("BACKGROUNDPLOT", 0)
            lay.PlotType = 1  # acExtents
            doc.Plot.PlotToFile(final)  # 实测不传第二个 config 参数
            return {"previous_config": old, "config": "DWG To PDF.pc3"}

        info = self._call("导出 PDF", _do)
        return {"ok": True, "format": "pdf", "path": final,
                "size": self._fsize(final), **info,
                "document": self._last_doc_name}

    def _convert_one(
        self,
        src: str,
        out_dir: str,
        formats: Sequence[str],
        unify: bool,
        close_after: bool,
    ) -> dict[str, Any]:
        app = self._app
        existing = None
        for i in range(int(app.Documents.Count)):
            try:
                d = app.Documents.Item(i)
                if str(d.FullName).lower() == str(src).lower():
                    existing = d
                    break
            except Exception:
                continue
        # 实测（live_v12_batch_debug）：对已打开的文档再调 Documents.Open(src)
        # 会新建一个重复的文档视图，累积后 COM 状态混乱导致批量偶发失败。
        # 已打开就复用、不再 Open；这类文档属于用户/前序流程，结束时也不关闭。
        if existing is not None:
            doc = existing
            already = True
        else:
            doc = app.Documents.Open(src)
            already = False
        app.ActiveDocument = doc
        base = os.path.splitext(os.path.basename(src))[0]
        outputs: dict[str, Any] = {}
        try:
            if unify:
                self._unify_in_doc(doc, "DSH-CN", "")
            s = self._select_all(doc)
            try:
                for f in formats:
                    f = str(f).lower()
                    if f == "dxf":
                        p = os.path.join(out_dir, base)
                        doc.Export(p, "DXF", s)
                        outputs["dxf"] = {"path": p + ".dxf", "size": self._fsize(p + ".dxf")}
                    elif f == "dwg":
                        p = os.path.join(out_dir, base + "_wblock.dwg")
                        doc.Wblock(p, s)
                        outputs["dwg"] = {"path": p, "size": self._fsize(p)}
                    elif f == "pdf":
                        p = os.path.join(out_dir, base + ".pdf")
                        lay = dynamic.Dispatch(doc.ActiveLayout)
                        lay.ConfigName = "DWG To PDF.pc3"
                        doc.SetVariable("BACKGROUNDPLOT", 0)
                        lay.PlotType = 1
                        doc.Plot.PlotToFile(p)
                        outputs["pdf"] = {"path": p, "size": self._fsize(p)}
                    else:
                        outputs[f] = {"error": "不支持的格式（可用 dxf/dwg/pdf）"}
            finally:
                try:
                    s.Delete()
                except Exception:
                    pass
        finally:
            if close_after and not already:
                try:
                    doc.Close(False)
                except Exception:
                    pass
        return {"base": base, "outputs": outputs,
                "closed": bool(close_after and not already)}

    def batch_convert(
        self,
        paths: Sequence[str],
        out_dir: str,
        formats: Sequence[str] | None = None,
        unify: bool = False,
        close_after: bool = True,
    ) -> dict[str, Any]:
        """「一句话交付」：批量打开整套图纸 → 每种格式导出 → 关闭。

        每个文档单独一次 _call（避免单次 STA 超时）；SelectionSet 名全局唯一。
        """
        fmts = [str(f) for f in (formats or ["dxf", "dwg", "pdf"])]
        os.makedirs(os.path.abspath(out_dir), exist_ok=True)
        results = []
        for p in paths or []:
            abs_in = os.path.abspath(str(p))
            if not os.path.isfile(abs_in):
                results.append({"input": str(p), "ok": False, "error": "文件不存在"})
                continue
            try:
                r = self._call(
                    f"批量处理 {os.path.basename(abs_in)}",
                    lambda src=abs_in: self._convert_one(
                        src, os.path.abspath(out_dir), fmts, bool(unify), bool(close_after)
                    ),
                    retries=12,
                    retry_delay_s=0.6,
                    timeout_s=180.0,
                )
                results.append({"input": str(p), "ok": True, **r})
            except Exception as exc:
                results.append({"input": str(p), "ok": False, "error": str(exc)})
        ok = sum(1 for r in results if r.get("ok"))
        return {"ok": True, "out_dir": os.path.abspath(out_dir), "formats": fmts,
                "succeeded": ok, "failed": len(results) - ok, "results": results}

    # -- V1.3 3D 建模 / 2D→3D / 视图 --------------------------------------
    #
    # 全部配方都来自 live probe（bridge/probe_v13*.py），关键结论：
    #   * 坐标参数是 SAFEARRAY(double) → 一律 _pt / dbl_array。
    #   * AddRegion 的对象参数是 SAFEARRAY(VT_DISPATCH)。
    #   * 面域会被 AddExtrudedSolid/AddRevolvedSolid/AddExtrudedSolidAlongPath
    #     消耗，所以每次拉伸/旋转/扫掠都要用 _closed_region 建全新面域。
    #   * center 是几何中心（沿 Z 对称），不是底心。
    #   * 旋转/锥度/角度参数是弧度，对外一律用度，内部 math.radians。
    #   * Boolean 0=并 / 1=交 / 2=差（与文档相反）；交集为空时双方一起被删。
    #   * 视图方向改完必须 doc.ActiveViewport = vp 赋值回，否则不生效；
    #     用 GetVariable("VIEWDIR") 验证，不要用 VPOINT/VSCURRENT（不可读）。
    #   * 全库无 Chamfer/Fillet/Sweep/Loft 实体方法：倒角/圆角走 SendCommand，
    #     扫掠走 AddExtrudedSolidAlongPath。

    @staticmethod
    def _dispatch_array(objects: Sequence[Any]) -> Any:
        """SAFEARRAY(VT_DISPATCH)：AddRegion 等需要「对象数组」而非点数组。"""
        return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, list(objects))

    @staticmethod
    def _obj_by_handle(doc: Any, handle: str) -> Any:
        """按句柄取回对象并 dynamic.Dispatch（避免 IAcadEntity 窄化）。"""
        try:
            return dynamic.Dispatch(doc.HandleToObject(str(handle)))
        except Exception as exc:  # noqa: BLE001
            raise CadError(f"找不到句柄 '{handle}' 对应的对象：{exc}")

    @staticmethod
    def _view_direction(name: str) -> tuple[float, float, float]:
        table = {
            "iso": (1.0, -1.0, 1.0), "isometric": (1.0, -1.0, 1.0),
            "front": (0.0, -1.0, 0.0), "back": (0.0, 1.0, 0.0),
            "top": (0.0, 0.0, 1.0), "bottom": (0.0, 0.0, -1.0),
            "right": (1.0, 0.0, 0.0), "left": (-1.0, 0.0, 0.0),
        }
        key = str(name).lower()
        if key not in table:
            raise CadError(f"未知视图 '{name}'。可选：{', '.join(sorted(table))}")
        return table[key]

    def _closed_region(self, ms: Any, points: Sequence[Sequence[float]]) -> Any:
        """由闭合 2D 轮廓建一个全新面域（每次建模都要新建，面域会被消耗）。"""
        flat: list[float] = []
        for p in points:
            flat.extend([float(p[0]), float(p[1])])
        if len(flat) < 6:
            raise CadError("轮廓至少需要 3 个点。")
        pl = ms.AddLightWeightPolyline(dbl_array(flat))
        pl.Closed = True  # 读回初始是 False，必须显式设
        regions = ms.AddRegion(self._dispatch_array([pl]))
        try:
            return regions[0]
        except Exception:  # noqa: BLE001
            return regions

    def _apply_view(self, doc: Any, view: str, zoom: bool = True) -> dict[str, Any]:
        """在 _do() 里调用：改 ActiveViewport 并验证 VIEWDIR（纯 COM，不发裸 VPOINT）。

        只做「纯 COM 方向切换」，不碰 SendCommand —— 视觉样式由 _view_and_shade
        单独一次 _call 处理，避免与建模/改视口挤在同一 COM 调用里。
        """
        d = self._view_direction(view)
        vp = dynamic.Dispatch(doc.ActiveViewport)
        vp.Direction = dbl_array(list(d))
        doc.ActiveViewport = vp  # ★ 赋值回，否则只改了副本
        if zoom:
            self._safe(lambda: self._app.ZoomExtents(), default=None)
        vd = self._safe(
            lambda: tuple(round(float(x), 4) for x in doc.GetVariable("VIEWDIR")),
            default=None,
        )
        want = tuple(round(x, 4) for x in d)
        return {"view": str(view).lower(), "direction": list(d), "viewdir": vd,
                "verified": vd == want}

    def _view_and_shade(self, view: str, shade: str, doc_name: str | None) -> dict[str, Any]:
        """切视图 + 着色，两件事各用一次 _call 彻底隔离。

        实测坑：把「建模」或「改 ActiveViewport」与 doc.SendCommand 挤进同一个
        COM 调用会触发 (-2145386414, '对象已打开进行通知')。所以方向（纯 COM）
        一次 _call、VSCURRENT 命令另一次 _call，互不同步执行。
        """
        info = self._call(
            f"切 3D 视图 {view}",
            lambda: self._apply_view(self._document(doc_name), view, zoom=True),
        )
        shade_sent = False
        if shade:
            def _shade():
                doc = self._document(doc_name)
                doc.SendCommand(f"_.VSCURRENT\n_{shade}\n")
                time.sleep(0.4)
                return True
            self._call(f"设置视觉样式 {shade}", _shade)
            shade_sent = True
        return {**info, "shade": shade or None, "shade_sent": shade_sent}

    # ---- V1.3-1 3D 原生能力 ----------------------------------------------

    def add_3d_primitive(
        self,
        kind: str,
        center: Sequence[float] = (0.0, 0.0, 0.0),
        params: dict[str, Any] | None = None,
        layer: str | None = None,
        doc_name: str | None = None,
    ) -> dict[str, Any]:
        """7 种 3D 原语。center 是几何中心（沿 Z 对称），不是底心。

        kind / params：
          box      {length, width, height}
          cylinder {radius, height}
          sphere   {radius}
          cone     {radius, height}
          torus    {torus_radius, tube_radius}
          wedge    {length, width, height}
        """
        k = str(kind).lower()
        pr = dict(params or {})
        c = [float(center[0]), float(center[1]),
             float(center[2]) if len(center) >= 3 else 0.0]
        layer = self._layer_for(layer, "outline")
        if layer:
            self.ensure_layer(layer, doc_name=doc_name)

        def _num(*names, default=0.0):
            for n in names:
                if n in pr and pr[n] is not None:
                    return float(pr[n])
            return float(default)

        def _do():
            ms = self._document(doc_name).ModelSpace
            p = _pt(*c)
            if k == "box":
                ent = ms.AddBox(p, _num("length", "l", default=10),
                                _num("width", "w", default=10), _num("height", "h", default=10))
            elif k == "cylinder":
                ent = ms.AddCylinder(p, _num("radius", "r", default=5),
                                     _num("height", "h", default=10))
            elif k == "sphere":
                ent = ms.AddSphere(p, _num("radius", "r", default=5))
            elif k == "cone":
                ent = ms.AddCone(p, _num("radius", "r", default=5),
                                 _num("height", "h", default=10))
            elif k == "torus":
                ent = ms.AddTorus(p, _num("torus_radius", "tr", default=10),
                                  _num("tube_radius", "r", default=3))
            elif k == "wedge":
                ent = ms.AddWedge(p, _num("length", "l", default=10),
                                  _num("width", "w", default=10), _num("height", "h", default=10))
            else:
                raise CadError(f"未知 3D 原语 '{kind}'。可选：box/cylinder/sphere/cone/torus/wedge")
            d = dynamic.Dispatch(ent)
            if layer:
                d.Layer = layer
            return {"handle": str(d.Handle), "object_name": str(d.ObjectName),
                    "bbox": _bbox_summary(d), "kind": k, "center": c, "params": pr}

        info = self._call(f"建 3D 原语 {k}", _do)
        return {"ok": True, **info, "layer": layer, "document": self._last_doc_name}

    def extrude_profile(
        self,
        points: Sequence[Sequence[float]],
        height: float = 10.0,
        taper_deg: float = 0.0,
        layer: str | None = None,
        doc_name: str | None = None,
    ) -> dict[str, Any]:
        """闭合 2D 轮廓 → 面域 → 拉伸实体（同时是 V1.3-2 的自动厚度核心）。"""
        layer = self._layer_for(layer, "outline")
        if layer:
            self.ensure_layer(layer, doc_name=doc_name)

        def _do():
            ms = self._document(doc_name).ModelSpace
            region = self._closed_region(ms, points)  # 每次新建，面域会被消耗
            solid = ms.AddExtrudedSolid(region, float(height), math.radians(float(taper_deg)))
            d = dynamic.Dispatch(solid)
            if layer:
                d.Layer = layer
            return {"handle": str(d.Handle), "object_name": str(d.ObjectName),
                    "bbox": _bbox_summary(d), "height": float(height), "taper_deg": float(taper_deg)}

        info = self._call("拉伸轮廓", _do)
        return {"ok": True, "type": "extruded_solid", **info,
                "layer": layer, "document": self._last_doc_name}

    def revolve_profile(
        self,
        points: Sequence[Sequence[float]],
        axis_point: Sequence[float] = (0.0, 0.0, 0.0),
        axis_dir: Sequence[float] = (0.0, 1.0, 0.0),
        angle_deg: float = 360.0,
        layer: str | None = None,
        doc_name: str | None = None,
    ) -> dict[str, Any]:
        """闭合 2D 轮廓绕轴旋转成实体；angle_deg 用度，轴方向是向量。"""
        layer = self._layer_for(layer, "outline")
        if layer:
            self.ensure_layer(layer, doc_name=doc_name)
        ap = _pt(*[float(v) for v in axis_point[:3]])
        ad = _pt(*[float(v) for v in axis_dir[:3]])

        def _do():
            ms = self._document(doc_name).ModelSpace
            region = self._closed_region(ms, points)
            solid = ms.AddRevolvedSolid(region, ap, ad, math.radians(float(angle_deg)))
            d = dynamic.Dispatch(solid)
            if layer:
                d.Layer = layer
            return {"handle": str(d.Handle), "object_name": str(d.ObjectName),
                    "bbox": _bbox_summary(d), "angle_deg": float(angle_deg),
                    "axis_point": list(axis_point[:3]), "axis_dir": list(axis_dir[:3])}

        info = self._call("旋转轮廓", _do)
        return {"ok": True, "type": "revolved_solid", **info,
                "layer": layer, "document": self._last_doc_name}

    def sweep_profile(
        self,
        points: Sequence[Sequence[float]],
        path_start: Sequence[float] = (0.0, 0.0, 0.0),
        path_end: Sequence[float] = (0.0, 0.0, 80.0),
        layer: str | None = None,
        doc_name: str | None = None,
    ) -> dict[str, Any]:
        """闭合轮廓沿直线路径扫掠成实体（AddExtrudedSolidAlongPath）。"""
        layer = self._layer_for(layer, "outline")
        if layer:
            self.ensure_layer(layer, doc_name=doc_name)
        ps = _pt(*[float(v) for v in path_start[:3]])
        pe = _pt(*[float(v) for v in path_end[:3]])

        def _do():
            ms = self._document(doc_name).ModelSpace
            path = ms.AddLine(ps, pe)  # 实测：AcDbLine 可作扫掠路径
            region = self._closed_region(ms, points)
            solid = ms.AddExtrudedSolidAlongPath(region, path)
            d = dynamic.Dispatch(solid)
            if layer:
                d.Layer = layer
            return {"handle": str(d.Handle), "object_name": str(d.ObjectName),
                    "bbox": _bbox_summary(d), "path_start": list(path_start[:3]),
                    "path_end": list(path_end[:3])}

        info = self._call("沿路径扫掠", _do)
        return {"ok": True, "type": "swept_solid", **info,
                "layer": layer, "document": self._last_doc_name}

    def boolean_solids(
        self,
        handle1: str,
        handle2: str,
        operation: str = "union",
        doc_name: str | None = None,
    ) -> dict[str, Any]:
        """布尔运算：union(0) / intersect(1) / subtract(2)。结果留在 handle1。

        注意：交集为空时 COM 会把双方一起删掉；用返回的 bbox 是否为空判断存活。
        """
        op_map = {"union": 0, "intersect": 1, "subtract": 2,
                  "0": 0, "1": 1, "2": 2}
        key = str(operation).lower()
        if key not in op_map:
            raise CadError(f"未知布尔运算 '{operation}'。可选：union/intersect/subtract")
        op = op_map[key]

        def _do():
            doc = self._document(doc_name)
            a = self._obj_by_handle(doc, handle1)
            b = self._obj_by_handle(doc, handle2)
            a.Boolean(op, b)
            bbox = _bbox_summary(a)
            return {"handle": handle1, "operation": key, "op_code": op,
                    "bbox": bbox, "survived": bool(bbox)}

        info = self._call(f"布尔 {key}", _do)
        return {"ok": True, **info, "document": self._last_doc_name}

    def rotate_3d(
        self,
        handle: str,
        axis_point1: Sequence[float] = (0.0, 0.0, 0.0),
        axis_point2: Sequence[float] = (1.0, 0.0, 0.0),
        angle_deg: float = 90.0,
        doc_name: str | None = None,
    ) -> dict[str, Any]:
        """实体绕 3D 轴旋转；angle_deg 用度（COM 内部是弧度）。"""
        p1 = _pt(*[float(v) for v in axis_point1[:3]])
        p2 = _pt(*[float(v) for v in axis_point2[:3]])

        def _do():
            ent = self._obj_by_handle(self._document(doc_name), handle)
            ent.Rotate3D(p1, p2, math.radians(float(angle_deg)))
            return {"handle": handle, "bbox": _bbox_summary(ent), "angle_deg": float(angle_deg)}

        info = self._call("3D 旋转", _do)
        return {"ok": True, **info, "document": self._last_doc_name}

    def slice_solid(
        self,
        handle: str,
        p1: Sequence[float],
        p2: Sequence[float],
        p3: Sequence[float],
        negative: bool = False,
        doc_name: str | None = None,
    ) -> dict[str, Any]:
        """用三点定义的平面把实体切成两半，返回新实体。"""
        a, b, c = (_pt(*[float(v) for v in p[:3]]) for p in (p1, p2, p3))

        def _do():
            ent = self._obj_by_handle(self._document(doc_name), handle)
            new = ent.SliceSolid(a, b, c, bool(negative))
            d = dynamic.Dispatch(new)
            return {"handle": str(d.Handle), "object_name": str(d.ObjectName),
                    "bbox": _bbox_summary(d), "negative": bool(negative)}

        info = self._call("剖切实体", _do)
        return {"ok": True, **info, "document": self._last_doc_name}

    def section_solid(
        self,
        handle: str,
        p1: Sequence[float],
        p2: Sequence[float],
        p3: Sequence[float],
        doc_name: str | None = None,
    ) -> dict[str, Any]:
        """用三点定义的平面取实体截面，返回面域（AcDbRegion）。"""
        a, b, c = (_pt(*[float(v) for v in p[:3]]) for p in (p1, p2, p3))

        def _do():
            ent = self._obj_by_handle(self._document(doc_name), handle)
            region = ent.SectionSolid(a, b, c)
            d = dynamic.Dispatch(region)
            return {"handle": str(d.Handle), "object_name": str(d.ObjectName),
                    "bbox": _bbox_summary(d)}

        info = self._call("取实体截面", _do)
        return {"ok": True, **info, "document": self._last_doc_name}

    def fillet_edges(self, radius: float = 3.0, doc_name: str | None = None) -> dict[str, Any]:
        """实体圆角：走 _.FILLETEDGE 命令（COM 无 Fillet 方法）。

        这是 fire-and-forget：命令异步执行、无法读回验证。请在 CAD 里目视确认；
        需要保证几何时用 cad_send_command 发自己的命令串，或改用布尔建模。
        """

        def _do():
            doc = self._document(doc_name)
            doc.SendCommand(f"_.FILLETEDGE\n_Radius\n{float(radius)}\n_All\n\n")
            time.sleep(0.6)
            return {"sent": True, "radius": float(radius),
                    "note": "已发送 FILLETEDGE（异步，无法读回验证，请在 CAD 里确认）"}

        info = self._call("实体圆角", _do)
        return {"ok": True, **info, "document": self._last_doc_name}

    def chamfer_edges(self, dist1: float = 3.0, dist2: float = 3.0,
                      doc_name: str | None = None) -> dict[str, Any]:
        """实体倒角：走 _.CHAMFEREDGE 命令（COM 无 Chamfer 方法）。fire-and-forget。"""

        def _do():
            doc = self._document(doc_name)
            doc.SendCommand(f"_.CHAMFEREDGE\n_Distance\n{float(dist1)}\n{float(dist2)}\n_All\n\n")
            time.sleep(0.6)
            return {"sent": True, "dist1": float(dist1), "dist2": float(dist2),
                    "note": "已发送 CHAMFEREDGE（异步，无法读回验证，请在 CAD 里确认）"}

        info = self._call("实体倒角", _do)
        return {"ok": True, **info, "document": self._last_doc_name}

    def solid_info(self, handle: str, doc_name: str | None = None) -> dict[str, Any]:
        """读实体包围盒 / 图层 / 存活状态。MassProperties 不可用，用 bbox 代替。"""

        def _do():
            ent = self._obj_by_handle(self._document(doc_name), handle)
            d = dynamic.Dispatch(ent)
            bbox = _bbox_summary(d)
            return {"handle": handle, "object_name": str(d.ObjectName),
                    "bbox": bbox, "alive": bool(bbox),
                    "layer": self._safe(lambda: str(d.Layer), default=""),
                    "note": "MassProperties 在此 COM 接口不可用，体积/质心请用 bbox 推断"}

        info = self._call("读实体信息", _do)
        return {"ok": True, **info, "document": self._last_doc_name}

    # ---- V1.3-3 视图 ------------------------------------------------------

    def set_3d_view(
        self,
        view: str = "iso",
        shade: str = "",
        doc_name: str | None = None,
    ) -> dict[str, Any]:
        """切 3D 视图：iso/front/back/top/bottom/right/left。

        方向用纯 COM 的 ActiveViewport.Direction（必须赋值回 doc.ActiveViewport），
        用 GetVariable("VIEWDIR") 验证；shade 走 _.VSCURRENT 命令（异步，不可读回）。
        """

        info = self._view_and_shade(view, shade, doc_name)
        info["regen"] = False
        return {"ok": True, **info, "document": self._last_doc_name}

    # ---- V1.3-2 一句话：尺寸/轮廓 → 3D 模型 -------------------------------

    def dimensions_to_3d(
        self,
        length: float,
        width: float,
        height: float,
        shade: str = "Conceptual",
        layer: str | None = None,
        doc_name: str | None = None,
    ) -> dict[str, Any]:
        """「一句话」：输入零件长宽高 → box 实体 + 等轴测视图 + 着色。"""
        L, W, H = float(length), float(width), float(height)
        layer = self._layer_for(layer, "outline")
        if layer:
            self.ensure_layer(layer, doc_name=doc_name)

        def _do():
            ms = self._document(doc_name).ModelSpace
            # center 是几何中心 → 让 box 占据 [0,L]×[0,W]×[0,H]
            ent = ms.AddBox(_pt(L / 2.0, W / 2.0, H / 2.0), L, W, H)
            d = dynamic.Dispatch(ent)
            if layer:
                d.Layer = layer
            return {"handle": str(d.Handle), "object_name": str(d.ObjectName),
                    "bbox": _bbox_summary(d), "dimensions": [L, W, H]}

        geom = self._call("由尺寸生成 3D 模型", _do)
        view = self._view_and_shade("iso", shade, doc_name)
        return {"ok": True, **geom, **view, "layer": layer, "document": self._last_doc_name}

    def profile_to_3d(
        self,
        points: Sequence[Sequence[float]],
        thickness: float = 10.0,
        shade: str = "Conceptual",
        layer: str | None = None,
        doc_name: str | None = None,
    ) -> dict[str, Any]:
        """「一句话」：闭合 2D 草图 + 厚度 → 拉伸实体 + 等轴测视图 + 着色。"""
        layer = self._layer_for(layer, "outline")
        if layer:
            self.ensure_layer(layer, doc_name=doc_name)

        def _do():
            ms = self._document(doc_name).ModelSpace
            region = self._closed_region(ms, points)
            solid = ms.AddExtrudedSolid(region, float(thickness), 0.0)
            d = dynamic.Dispatch(solid)
            if layer:
                d.Layer = layer
            return {"handle": str(d.Handle), "object_name": str(d.ObjectName),
                    "bbox": _bbox_summary(d), "thickness": float(thickness)}

        geom = self._call("由草图生成 3D 模型", _do)
        view = self._view_and_shade("iso", shade, doc_name)
        return {"ok": True, **geom, **view, "layer": layer, "document": self._last_doc_name}

    # ======================================================================
    # V2.0 正式工业版
    #   ① 图纸合规自检引擎  ② 企业 .dwt 模板  ③ 图元数据反解析  ④ 工程活动日志
    #
    # 全部基于「读 DWG 图元数据」，不是视觉识图，也不联网：
    # 与 ROADMAP 的 Slogan「非图像识别 / 零幻觉 / 离线」一致。
    # 所有 COM 读操作整体塞进一个 _call 的 _do()（probe_v20b 实测：连续读
    # 全表后 CAD 会短暂忙，单次 _call 自带 BUSY 重试最稳）。
    # ======================================================================

    @staticmethod
    def _prop(obj: Any, name: str, default: Any = None) -> Any:
        """安全读一个 COM 属性。属性是 attribute 访问，不是 getter 方法。"""
        try:
            return getattr(obj, name)
        except Exception:  # noqa: BLE001
            return default

    @staticmethod
    def _vec(value: Any) -> list[float] | None:
        """把 COM SAFEARRAY/VARIANT 点或坐标转成可 JSON 化的 float 列表。"""
        if value is None:
            return None
        try:
            return [round(float(x), 6) for x in value]
        except Exception:  # noqa: BLE001
            return None

    @staticmethod
    def _jsonable(value: Any) -> Any:
        """把 GetVariable 可能返回的 tuple/VARIANT 转成纯 JSON 值。"""
        if value is None or isinstance(value, (bool, int, float, str)):
            return value
        try:
            return [CadBridge._jsonable(x) for x in value]
        except Exception:  # noqa: BLE001
            return str(value)

    @staticmethod
    def _coll_names(coll: Any) -> list[str]:
        """
        枚举图元集合里每个对象的 Name。RAW 读取即可（probe_v20b 实测）。
        BUSY com_error 原样抛出（交给 _read_drawing_standard 的只读重试），
        其它属性错误跳过。
        """
        out: list[str] = []
        try:
            for i in range(int(coll.Count)):
                item = coll.Item(i)
                name = CadBridge._prop(item, "Name")
                if name is None:
                    name = CadBridge._prop(dynamic.Dispatch(item), "Name")
                if name is not None:
                    out.append(str(name))
        except Exception as exc:  # noqa: BLE001
            if isinstance(exc, pywintypes.com_error) and _hresult(exc) in BUSY_HRESULTS:
                raise
        return out

    def _read_drawing_standard(self, doc: Any) -> dict[str, Any]:
        """
        读当前文档的制图规范（图层/文字/标注样式/系统变量）。必须在 _do() 内调用。

        ★ BUSY 自愈（V2.0 live 实测）：Documents.Add(dwt) / SaveAs 之后 AutoCAD
        会短暂拒绝呼叫（RPC_E_CALL_REJECTED = -2147418111）。这些 COM 读错误原先
        被逐集合的 `except: pass` 吞成空列表，导致 load_template 返回的 standard
        偶发缺图层 / INSUNITS（tests/test_v20_industrial.py 的 flaky FAIL 根因）。
        这里在 COM 线程内对整段「只读枚举」做有限重试，直到读到一个图层且
        INSUNITS 可读（任何合法 DWG 都有 "0" 层，故这是安全的完备性判据）。
        重试只在线程内重跑读取，绝不把 BUSY 抛回 _call —— 那会让 load_template
        的 _do() 整个重跑、触发 Documents.Add 重复建图（非幂等）。
        """
        last: dict[str, Any] = {}
        for _attempt in range(10):
            try:
                result = self._read_drawing_standard_once(doc)
            except pywintypes.com_error as exc:
                # 只在这里消化 BUSY：整段只读重跑，不惊动 _call，
                # 因而绝不会让 load_template 的 Documents.Add 重跑。
                if _hresult(exc) in BUSY_HRESULTS:
                    time.sleep(0.4)
                    continue
                raise
            last = result
            # 完备性兜底：任何合法 DWG 都至少有 "0" 层和 INSUNITS。
            if result["layers"] and result["sysvars"].get("INSUNITS") is not None:
                return result
            time.sleep(0.4)
        last["_incomplete"] = True
        last["_error"] = "AutoCAD 持续忙，制图规范读取不完整（请稍后重试）"
        return last

    def _read_drawing_standard_once(self, doc: Any) -> dict[str, Any]:
        """
        单次读取（含逐属性容错）。由 _read_drawing_standard 的 BUSY 重试循环调用。
        ★ 集合级读取若命中 BUSY com_error 必须原样抛出，交给外层循环重跑整段读取；
          绝不能像以前那样 `except: pass` 吞成空/半截列表（那正是 flaky 的根因）。
          非 BUSY 的属性错误仍按缺省值容错（不同 AutoCAD 版本属性集合略有差异）。
        """

        def _busy_guard(exc: BaseException) -> None:
            if isinstance(exc, pywintypes.com_error) and _hresult(exc) in BUSY_HRESULTS:
                raise exc

        # 就绪探针：先做两个轻量 COM 调用。刚 Documents.Add / SaveAs 之后
        # AutoCAD 常短暂拒绝呼叫，让 BUSY 在枚举之前就暴露出来（触发外层重试）。
        int(doc.Layers.Count)
        doc.GetVariable("INSUNITS")

        layers: list[dict[str, Any]] = []
        try:
            coll = doc.Layers
            for i in range(int(coll.Count)):
                lay = dynamic.Dispatch(coll.Item(i))
                layers.append({
                    "name": str(self._prop(lay, "Name", "")),
                    "color": self._jsonable(self._prop(lay, "Color")),
                    "lineweight": self._jsonable(self._prop(lay, "Lineweight")),
                    "linetype": str(self._prop(lay, "Linetype", "")),
                    "lock": bool(self._prop(lay, "Lock", False)),
                    "freeze": bool(self._prop(lay, "Freeze", False)),
                })
        except Exception as exc:  # noqa: BLE001
            _busy_guard(exc)

        text_styles: list[dict[str, Any]] = []
        try:
            coll = doc.TextStyles
            for i in range(int(coll.Count)):
                st = dynamic.Dispatch(coll.Item(i))
                text_styles.append({
                    "name": str(self._prop(st, "Name", "")),
                    "font": str(self._prop(st, "fontFile", "")),
                    "bigfont": str(self._prop(st, "bigFontFile", "")),
                })
        except Exception as exc:  # noqa: BLE001
            _busy_guard(exc)

        sysvars: dict[str, Any] = {}
        for key in ("INSUNITS", "MEASUREMENT", "LUNITS", "LWDISPLAY", "LTSCALE",
                    "DIMTXT", "DIMASZ", "DIMTXSTY", "DIMLFAC", "DIMASSOC",
                    "DWGCODEPAGE"):
            try:
                sysvars[key] = self._jsonable(doc.GetVariable(key))
            except Exception as exc:  # noqa: BLE001
                _busy_guard(exc)
                sysvars[key] = None

        active: dict[str, Any] = {}
        for prop, out in (("ActiveDimStyle", "dim_style"),
                          ("ActiveLayer", "layer"),
                          ("ActiveTextStyle", "text_style")):
            try:
                active[out] = str(getattr(doc, prop).Name)
            except Exception as exc:  # noqa: BLE001
                _busy_guard(exc)
                active[out] = None

        return {
            "layers": layers,
            "text_styles": text_styles,
            "dim_styles": self._coll_names(doc.DimStyles),
            "linetypes": self._coll_names(doc.Linetypes),
            "sysvars": sysvars,
            "active": active,
        }

    def _read_drawing_snapshot(self, doc: Any, limit: int = 5000) -> dict[str, Any]:
        """
        全表读 DWG 图元参数。必须在 _call 的 _do() 内调用（已在 COM 线程）。
        V2.0-① 合规自检与 V2.0-③ 反解析共用这一份数据，避免重复枚举。
        """
        ms = doc.ModelSpace
        total = int(ms.Count)
        n = min(total, limit) if limit and limit > 0 else total
        entities: list[dict[str, Any]] = []
        for i in range(n):
            try:
                e = dynamic.Dispatch(ms.Item(i))
                rec: dict[str, Any] = {
                    "index": i,
                    "handle": str(self._prop(e, "Handle", "")),
                    "type": str(self._prop(e, "ObjectName", "")),
                    "layer": str(self._prop(e, "Layer", "")),
                }
                t = rec["type"]
                if t == "AcDbLine":
                    rec["start"] = self._vec(self._prop(e, "StartPoint"))
                    rec["end"] = self._vec(self._prop(e, "EndPoint"))
                    rec["length"] = round(float(self._prop(e, "Length", 0) or 0), 6)
                elif "Polyline" in t:
                    rec["closed"] = bool(self._prop(e, "Closed", False))
                    rec["coordinates"] = self._vec(self._prop(e, "Coordinates"))
                    rec["length"] = round(float(self._prop(e, "Length", 0) or 0), 6)
                elif t in ("AcDbCircle", "AcDbArc"):
                    rec["center"] = self._vec(self._prop(e, "Center"))
                    rec["radius"] = round(float(self._prop(e, "Radius", 0) or 0), 6)
                    if t == "AcDbArc":
                        rec["start_angle_rad"] = round(float(self._prop(e, "StartAngle", 0) or 0), 8)
                        rec["end_angle_rad"] = round(float(self._prop(e, "EndAngle", 0) or 0), 8)
                elif t in ("AcDbText", "AcDbMText"):
                    rec["text"] = str(self._prop(e, "TextString", ""))
                    rec["height"] = round(float(self._prop(e, "Height", 0) or 0), 6)
                    rec["style"] = str(self._prop(e, "StyleName", ""))
                    rec["insertion"] = self._vec(self._prop(e, "InsertionPoint"))
                elif "Dimension" in t:
                    rec["measurement"] = round(float(self._prop(e, "Measurement", 0) or 0), 6)
                    rec["override"] = str(self._prop(e, "TextOverride", "") or "")
                    rec["style"] = str(self._prop(e, "StyleName", "") or "")
                elif t == "AcDbBlockReference":
                    rec["block"] = str(self._prop(e, "EffectiveName", ""))
                    rec["insertion"] = self._vec(self._prop(e, "InsertionPoint"))
                    rec["rotation"] = round(float(self._prop(e, "Rotation", 0) or 0), 8)
                entities.append(rec)
            except Exception as exc:  # noqa: BLE001
                entities.append({"index": i, "type": "<unreadable>", "error": str(exc)})

        histogram: dict[str, int] = {}
        for rec in entities:
            key = rec.get("type", "<unknown>")
            histogram[key] = histogram.get(key, 0) + 1

        return {
            "total": total,
            "returned": len(entities),
            "entities": entities,
            "histogram": histogram,
            "standard": self._read_drawing_standard(doc),
        }

    # -- V2.0-① 图纸合规自检引擎 -------------------------------------------

    def check_drawing_compliance(
        self,
        doc_name: str | None = None,
        min_line_length: float = 1.0,
        dim_layer: str = "标注",
        text_layer: str = "文字",
        gb_dim_style: str = "GB-机械",
    ) -> dict[str, Any]:
        """
        检查：图层错误 / 标注不规范 / 未闭合轮廓 / 重复线、碎线。
        全部按图元参数判定，不做视觉判断；返回结构化报告，不改图纸。
        """

        def _do() -> dict[str, Any]:
            doc = self._document(doc_name)
            snap = self._read_drawing_snapshot(doc)
            ents = snap["entities"]
            known_layers = {str(l["name"]) for l in snap["standard"]["layers"]}
            std = set(GB_LAYERS) | {"0", "Defpoints"} | known_layers
            issues: list[dict[str, Any]] = []

            def add(category: str, severity: str, handle: str, detail: str) -> None:
                issues.append({"category": category, "severity": severity,
                               "handle": handle, "detail": detail})

            lines: list[dict[str, Any]] = []
            for e in ents:
                t = e.get("type", "")
                if e.get("layer") and e["layer"] not in std:
                    add("图层错误", "warn", e.get("handle", ""),
                        f"图元落在非标准图层 '{e['layer']}'")
                if t == "AcDbLine":
                    lines.append(e)
                    if float(e.get("length") or 0) < float(min_line_length):
                        add("碎线", "warn", e.get("handle", ""),
                            f"直线长度 {e.get('length')} 小于阈值 {min_line_length}")
                elif "Polyline" in t and e.get("closed") is False:
                    add("未闭合轮廓", "error", e.get("handle", ""),
                        "多段线轮廓未闭合，无法生成面域/拉伸")
                elif "Dimension" in t:
                    if e.get("layer") and dim_layer and e["layer"] != dim_layer:
                        add("标注不规范", "warn", e.get("handle", ""),
                            f"尺寸在图层 '{e['layer']}'，应放在 '{dim_layer}'")
                    if e.get("override"):
                        add("标注不规范", "warn", e.get("handle", ""),
                            f"手工覆盖尺寸文字 '{e['override']}'，与实测值可能不符")
                    if gb_dim_style and snap["standard"]["dim_styles"] and \
                            gb_dim_style in snap["standard"]["dim_styles"] and \
                            e.get("style") and e["style"] != gb_dim_style:
                        add("标注不规范", "warn", e.get("handle", ""),
                            f"标注样式 '{e['style']}' 不是国标样式 '{gb_dim_style}'")
                    if float(e.get("measurement") or 0) <= 0:
                        add("标注不规范", "error", e.get("handle", ""),
                            "尺寸实测值 <= 0（可能是错误标注）")
                elif t in ("AcDbText", "AcDbMText") and e.get("layer") and \
                        text_layer and e["layer"] != text_layer:
                    add("图层错误", "warn", e.get("handle", ""),
                        f"文字在图层 '{e['layer']}'，应放在 '{text_layer}'")

            # 重复线：端点排序后做键，坐标取 4 位小数容差
            seen: dict[tuple, str] = {}
            for e in lines:
                s, en = e.get("start"), e.get("end")
                if not s or not en:
                    continue
                a = (round(s[0], 4), round(s[1], 4))
                b = (round(en[0], 4), round(en[1], 4))
                key = tuple(sorted((a, b)))
                if key in seen:
                    add("重复线", "error", e.get("handle", ""),
                        f"与句柄 {seen[key]} 的直线端点重合")
                else:
                    seen[key] = e.get("handle", "")

            by_category: dict[str, int] = {}
            errors = warnings = 0
            for it in issues:
                by_category[it["category"]] = by_category.get(it["category"], 0) + 1
                if it["severity"] == "error":
                    errors += 1
                else:
                    warnings += 1
            return {
                "checks": ["图层错误", "标注不规范", "未闭合轮廓", "重复线", "碎线"],
                "entity_count": snap["total"],
                "scanned": snap["returned"],
                "histogram": snap["histogram"],
                "issues": issues,
                "summary": {"issues": len(issues), "errors": errors,
                            "warnings": warnings, "by_category": by_category},
                "compliant": errors == 0,
            }

        info = self._call("图纸合规自检", _do, retries=6, retry_delay_s=0.5)
        return {"ok": True, **info, "document": self._last_doc_name}

    # -- V2.0-③ 图元数据反解析（非视觉识图）-------------------------------

    @staticmethod
    def _build_replay_script(snap: dict[str, Any], document: str | None) -> str:
        """从图元参数逆向生成可复刻的 Python 脚本（用同一 CadBridge 重建）。"""
        lines: list[str] = [
            "# -*- coding: utf-8 -*-",
            "# 由 DSH AutoCAD Bridge V2.0 图元反解析自动生成（读 DWG 参数，非视觉识图）",
            f"# 源文档：{document}",
            f"# 图元数：{snap['returned']}（总 {snap['total']}）",
            "# 复刻方式：读图元参数 → 用原生 COM 重建，实现「改图 / 复刻 / 迭代」闭环",
            "",
            "import os, sys, math",
            "sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))",
            "from cad_com import CadBridge",
            "",
            "b = CadBridge()",
            "b.connect()",
            "b.new_document()",
            "",
        ]
        for e in snap["entities"]:
            t = e.get("type", "")
            lay = e.get("layer") or ""
            arg = f", layer={lay!r}" if lay else ""
            try:
                if t == "AcDbLine" and e.get("start") and e.get("end"):
                    lines.append(f"b.add_line({e['start'][:2]!r}, {e['end'][:2]!r}{arg})")
                elif t == "AcDbCircle" and e.get("center"):
                    lines.append(f"b.add_circle({e['center'][:2]!r}, {e.get('radius', 0)!r}{arg})")
                elif t == "AcDbArc" and e.get("center"):
                    sa = round(math.degrees(e.get("start_angle_rad", 0)), 6)
                    ea = round(math.degrees(e.get("end_angle_rad", 0)), 6)
                    lines.append(f"b.add_arc({e['center'][:2]!r}, {e.get('radius', 0)!r}, {sa!r}, {ea!r}{arg})")
                elif "Polyline" in t and e.get("coordinates"):
                    flat = e["coordinates"]
                    pts = [flat[i:i + 2] for i in range(0, len(flat) - 1, 2)]
                    lines.append(f"b.add_polyline({pts!r}, closed={bool(e.get('closed'))!r}{arg})")
                elif t in ("AcDbText", "AcDbMText"):
                    ins = (e.get("insertion") or [0, 0, 0])[:2]
                    lines.append(f"b.add_text({e.get('text', '')!r}, {ins!r}, height={e.get('height', 2.5)!r}{arg})")
                elif t == "AcDbBlockReference" and e.get("block"):
                    ins = (e.get("insertion") or [0, 0, 0])[:2]
                    lines.append(f"b.insert_block({e['block']!r}, {ins!r}{arg})")
                else:
                    lines.append(f"# TODO 未覆盖类型 {t} (handle={e.get('handle', '')})")
            except Exception as exc:  # noqa: BLE001
                lines.append(f"# 跳过 {t} (handle={e.get('handle', '')})：{exc}")
        lines += ["", "b.zoom_extents()", "print('replay done')", ""]
        return "\n".join(lines)

    def reverse_engineer_drawing(
        self,
        doc_name: str | None = None,
        limit: int = 5000,
        include_script: bool = True,
    ) -> dict[str, Any]:
        """读取现有图纸参数 → 图元清单 + 统计 + 可复刻参数脚本（非视觉识图）。"""

        def _do() -> dict[str, Any]:
            doc = self._document(doc_name)
            snap = self._read_drawing_snapshot(doc, limit=limit)
            out: dict[str, Any] = {
                "total": snap["total"],
                "returned": snap["returned"],
                "histogram": snap["histogram"],
                "entities": snap["entities"],
                "standard": snap["standard"],
            }
            if include_script:
                script = self._build_replay_script(snap, self._last_doc_name)
                out["script"] = script
                out["script_chars"] = len(script)
            return out

        info = self._call("图元数据反解析", _do, retries=6, retry_delay_s=0.5)
        return {"ok": True, **info, "document": self._last_doc_name}

    def export_reverse_script(
        self,
        path: str,
        doc_name: str | None = None,
        limit: int = 5000,
    ) -> dict[str, Any]:
        """把反解析得到的复刻脚本写到磁盘（离线，仅本地文件）。"""
        if not path:
            raise CadError("export_reverse_script 需要 path 参数。")

        def _do() -> dict[str, Any]:
            doc = self._document(doc_name)
            snap = self._read_drawing_snapshot(doc, limit=limit)
            return {"_snap": snap}

        info = self._call("采集图元用于参数脚本", _do, retries=6, retry_delay_s=0.5)
        snap = info["_snap"]
        script = self._build_replay_script(snap, self._last_doc_name)
        target = os.path.abspath(path)
        parent = os.path.dirname(target)
        if parent:
            os.makedirs(parent, exist_ok=True)
        with open(target, "w", encoding="utf-8", newline="\n") as fh:
            fh.write(script)
        return {"ok": True, "path": target, "size": os.path.getsize(target),
                "script_chars": len(script), "entities": snap["returned"],
                "document": self._last_doc_name}

    # -- V2.0-② 企业标准模板系统 -------------------------------------------

    def _template_dir(self, doc: Any) -> str:
        """运行时解析 AutoCAD 模板目录，不写死用户名（probe_v20c 实测）。"""
        for var in ("LOCALROOTPREFIX", "ROAMABLEROOTPREFIX"):
            try:
                prefix = str(doc.GetVariable(var))
            except Exception:  # noqa: BLE001
                continue
            if not prefix:
                continue
            candidate = os.path.join(prefix, "Template")
            if os.path.isdir(candidate):
                return candidate
        # GetVariable("TEMPLATEPATH") 在本版本会报「获取系统变量时出错」，故不依赖它。
        return ""

    def list_templates(self, doc_name: str | None = None) -> dict[str, Any]:
        """列出本机可用的 .dwt 模板。"""

        def _do() -> dict[str, Any]:
            doc = self._document(doc_name)
            directory = self._template_dir(doc)
            items: list[dict[str, Any]] = []
            if directory:
                for name in sorted(os.listdir(directory)):
                    if not name.lower().endswith(".dwt"):
                        continue
                    full = os.path.join(directory, name)
                    try:
                        size = os.path.getsize(full)
                    except OSError:
                        size = 0
                    items.append({"name": name, "path": full, "size": size})
            return {"template_dir": directory, "count": len(items), "templates": items}

        info = self._call("列出 DWT 模板", _do)
        return {"ok": True, **info, "document": self._last_doc_name}

    def read_drawing_standard(self, doc_name: str | None = None) -> dict[str, Any]:
        """读取当前文档的制图规范（图层/文字/标注样式/系统变量/激活项）。"""

        def _do() -> dict[str, Any]:
            return self._read_drawing_standard(self._document(doc_name))

        info = self._call("读取制图规范", _do, retries=6, retry_delay_s=0.5)
        return {"ok": True, **info, "document": self._last_doc_name}

    def load_template(self, path: str, doc_name: str | None = None) -> dict[str, Any]:
        """
        加载企业 .dwt 模板开新图。
        注意（probe_v20c 实测）：Documents.Add(不存在的路径) 会**静默建出普通新图**，
        所以这里先 os.path.isfile 显式校验并拒绝，避免 silently 得到错的图纸。
        """

        def _do() -> dict[str, Any]:
            base = self._document(doc_name)
            candidate = os.path.abspath(path)
            if not os.path.isfile(candidate):
                directory = self._template_dir(base)
                if directory:
                    joined = os.path.join(directory, path)
                    if os.path.isfile(joined):
                        candidate = joined
            if not os.path.isfile(candidate):
                raise CadError(
                    f"找不到 .dwt 模板 '{path}'（且不在模板目录里）。"
                    "先调 cad_list_templates 查看可用模板，或传绝对路径。"
                    "注意：AutoCAD 对不存在的模板路径会静默建普通新图，此处已显式拒绝。"
                )
            before = int(self._app.Documents.Count)
            newdoc = self._app.Documents.Add(candidate)
            after = int(self._app.Documents.Count)
            self._last_doc_name = str(self._prop(newdoc, "Name", "") or self._last_doc_name)
            return {
                "template": candidate,
                "created_document": self._last_doc_name,
                "documents_before": before,
                "documents_after": after,
                "standard": self._read_drawing_standard(newdoc),
            }

        info = self._call("加载 DWT 模板", _do)
        return {"ok": True, **info, "document": self._last_doc_name}

    def save_as_template(self, path: str, doc_name: str | None = None) -> dict[str, Any]:
        """
        把当前文档另存为企业 .dwt 模板（SaveAsType=14 / ac2000_dwt，probe_v20d 实测）。
        自定义图层/文字样式/INSUNITS 会随模板存活。
        边界：本工具负责「生成并加载企业模板」，不做「按公司名自动猜规范」——
        那需要企业自己的规范库，超出离线可验证范围，见 docs/limitations.md。
        """

        def _do() -> dict[str, Any]:
            doc = self._document(doc_name)
            target = os.path.abspath(path)
            if not target.lower().endswith(".dwt"):
                target += ".dwt"
            parent = os.path.dirname(target)
            if parent:
                os.makedirs(parent, exist_ok=True)
            # examples/valve-3d.py 的教训：残留 .dwl/.dwl2 锁文件会让 SaveAs 报 -2145320861
            for suffix in (".dwl", ".dwl2"):
                lock = target + suffix
                if os.path.isfile(lock):
                    try:
                        os.remove(lock)
                    except OSError:
                        pass
            doc.SaveAs(target, DWT_SAVEAS_TYPE)
            return {"path": target,
                    "size": os.path.getsize(target) if os.path.isfile(target) else -1}

        info = self._call("另存为 DWT 模板", _do)
        return {"ok": True, **info, "document": self._last_doc_name}

    # -- V2.0-④ 工程活动日志 / 工作报告 ------------------------------------

    def get_activity_log(
        self,
        limit: int = 100,
        action: str = "",
        ok_only: bool | None = None,
    ) -> dict[str, Any]:
        """读工程活动日志（纯内存状态，不需要连接 CAD，CAD 关掉也能查）。"""
        with self._activity_lock:
            recs = [dict(r) for r in self._activity_log]
            seq = self._activity_seq
            buffered = len(self._activity_log)
        if action:
            recs = [r for r in recs if action in r["action"]]
        if ok_only is not None:
            recs = [r for r in recs if r["ok"] == bool(ok_only)]
        total = len(recs)
        if limit and limit > 0:
            recs = recs[-int(limit):]
        succeeded = sum(1 for r in recs if r["ok"])
        return {
            "ok": True,
            "seq": seq,
            "buffered": buffered,
            "capacity": ACTIVITY_LOG_CAPACITY,
            "total": total,
            "returned": len(recs),
            "succeeded": succeeded,
            "failed": len(recs) - succeeded,
            "records": recs,
        }

    def export_activity_report(
        self,
        path: str = "",
        doc_name: str | None = None,
        limit: int = 2000,
    ) -> dict[str, Any]:
        """把工程活动日志导出成 Markdown 工作报告（离线，仅本地文件）。"""
        log = self.get_activity_log(limit=limit)
        recs = log["records"]
        target = os.path.abspath(path) if path else os.path.join(
            os.environ.get("TEMP") or os.path.expanduser("~"),
            "dsh-activity-report.md",
        )
        parent = os.path.dirname(target)
        if parent:
            os.makedirs(parent, exist_ok=True)

        failed = [r for r in recs if not r["ok"]]
        out: list[str] = [
            "# DSH AutoCAD Bridge — 工程工作报告",
            "",
            f"- 生成时间：{time.strftime('%Y-%m-%d %H:%M:%S')}",
            f"- 当前文档：{self._last_doc_name or doc_name or '(未连接)'}",
            f"- 累计调用：{log['seq']}（缓冲保留 {log['buffered']}/{log['capacity']}）",
            f"- 本次报告记录：{len(recs)}，成功 {log['succeeded']}，失败 {log['failed']}",
            "",
            "## 失败明细",
            "",
        ]
        if failed:
            for r in failed:
                out.append(f"- #{r['seq']} `{r['action']}` @ {r['time']} "
                           f"({r['duration_ms']} ms)：{r['error']}")
        else:
            out.append("（无失败记录）")
        out += ["", "## 活动明细", "",
                "| # | 时间 | 动作 | 结果 | 耗时(ms) | 错误 |",
                "|---|------|------|------|----------|------|"]
        for r in recs:
            err = (r["error"] or "").replace("|", "\\|").replace("\n", " ")
            if len(err) > 160:
                err = err[:160] + "…"
            out.append(f"| {r['seq']} | {r['time']} | {r['action']} | "
                       f"{'OK' if r['ok'] else 'FAIL'} | {r['duration_ms']} | {err} |")
        out.append("")
        text = "\n".join(out)
        with open(target, "w", encoding="utf-8", newline="\n") as fh:
            fh.write(text)
        return {"ok": True, "path": target, "size": os.path.getsize(target),
                "records": len(recs), "failed": len(failed),
                "document": self._last_doc_name or doc_name}

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
