# 排错

按「现在卡在哪一步」找对应章节。

---

## 面板里的状态灯

| 灯 | 含义 | 下一步 |
|---|---|---|
| 灰 · 已停止 | 没有桥进程 | 点「启动」，或检查下方「最近错误」 |
| 黄 · 进程在但不通 | 进程活着，8770 端口没在监听 | 看日志最后几行，通常是依赖没装好或端口被占 |
| 绿 · 运行中 | 进程在、端口通 | 让 Agent 调 `cad_connect` |

---

## 装不上、起不来

### `已经安装但代码变了`

改了插件源码后重新安装，返回 `{"code":"already-installed"}` 是正常的（包管理器认为版本号没变）。要让改动生效：

在插件的行上**禁用再启用**。注意 entry id 是 `include:autocad-bridge`，不是包名——按包名操作会得到 `unknown-plugin`。

桥进程（Python）**不会热重载**。改了 `.py` 必须重启桥：面板点「重启」。

### `Cannot read properties of undefined (reading 'resolveExecutable')`

宿主的 `subprocess` 服务还没就绪。插件已经等 `ctx.inject(['connection','subprocess'])` 之后才允许操作，若仍出现，说明宿主版本不匹配。升级 DSH。

### 依赖装不上

先确认 Python 版本在 3.10–3.12。3.13 上 `pywin32` 的轮子可能缺失。

如果在系统 Python 里装过 `mcp`，可能和别的包冲突（例如 `open-interpreter` 会拽走 `anthropic` / `starlette` / `typer` 的版本）。**插件的依赖自举就是为了回避这个问题**——它建独立 venv，不碰系统环境。

---

## 连不上 AutoCAD

### 第一步：确认 AutoCAD 真的在跑

```powershell
Get-Process acad
```

插件默认**带界面**启动 AutoCAD（`visible=True`）。纯后台静默 COM 启动不稳定，这是刻意选择。

### 第二步：确认 ProgID

```powershell
# 让 Agent 调 cad_list_progids，或手动查
Get-ChildItem 'HKLM:\SOFTWARE\Classes' | Where-Object { $_.PSChildName -like 'AutoCAD.Application.*' }
```

ProgID 对应关系：

| 版本 | ProgID |
|---|---|
| 2024 | `AutoCAD.Application.24` |
| 2023 | `AutoCAD.Application.23` |
| 2022 | `AutoCAD.Application.22` |

注意同一版本可能有 `.24` / `.24.1` / `.24.2` / `.24.3` 多个，通常用最短的那个。

### 第三步：CAD 忙 / 卡死

报 `RPC_E_CALL_REJECTED`（`-2147418111`）或 `RPC_E_SERVERCALL_RETRYLATER`（`-2147417846`）说明 CAD 正忙。桥会自动重试（默认 2 次、间隔 800ms）。

如果一直忙，多半是**弹了个模态对话框**（未保存提示、许可证、插件更新）。CAD 里按下 ESC，或让 Agent 调 `cad_escape`。

V1.1 起可以先调 `cad_status` 看 `queue` 字段：`busy_seconds` 是 COM 线程已卡
多久、`pending` 是排队等着的调用数、`last_error` 是最近一次超时。确认长
期卡死后调 `cad_reconnect` ——它会丢弃旧连接、换一个新的 STA 工作线程重建
（被卡住的旧线程是 daemon，不会被等待）。

冷启动（CAD 本来没开）时 `cad_connect` 会自动拉起 `acad.exe` 并等待就绪，
约 15 秒，属正常现象；期间日志里可能出现
`TypeError: This COM object can not automate the makepy process`，桥已对此
做了重试兜底，不必手动干预。

**V1.2 批量导出场景**：连续 DXF/Wblock/PDF 后 CAD 后台收尾会额外忙 2~5 秒，
默认的 2×0.8s 窗口可能不够，报「AutoCAD 一直处于忙碌状态（批量处理 xxx
重试 2 次仍失败）」。这不是弹窗，别急着按 ESC：
- 用 `cad_batch_convert` 一次做完（它内部用 12×0.6s + 180s 的加宽窗口）；
- 或等几秒再重试，别把单文档导出工具连着猛调；
- 若刚跑完一轮测试，最干净的办法是关掉 CAD 与桥重来（同时复验冷启动唤醒）。

**V1.3 3D 建模场景**：报
`(-2145386414, '对象已打开进行通知')` 是因为「新建实体」和「发
`_.VSCURRENT` 视图命令」挤进了同一次 COM 调用——新建对象还开着通知时
CAD 不派发命令。`cad_dimensions_to_3d` / `cad_profile_to_3d` /
`cad_set_3d_view` 已把建模、改视口、发命令拆成各自独立的 `_call`，正常
调用不会遇到；若你自己写脚本把这几步塞进一次调用，拆开即可。另外
`cad_fillet_edges` / `cad_chamfer_edges` 是 fire-and-forget 命令，若发现
圆角没生效，先在 CAD 里目视确认命令有没有跑完（必要时按 ESC 结束挂起
命令），**不要发裸 `VPOINT`**（实测会让 AutoCAD 停在交互提示、COM 全被拒）。

**V2.0 模板 / 反解析场景**：`cad_load_template` / `cad_read_drawing_standard` /
`cad_check_drawing_compliance` / `cad_reverse_engineer_drawing` 都要连续读集合
与系统变量。`Documents.Add(dwt)` 之后 AutoCAD 会短暂 BUSY，早期版本会把这段
BUSY 吞成空规范（偶发 `standard.layers=[]` / `INSUNITS=None`）。当前实现已在
`_read_drawing_standard` 里做 10 次「只读」重试（完备性判据
`layers and INSUNITS is not None`），BUSY 只在该层重试、**不抛回 `_call`**
（否则整个 `_do()` 重跑会重复 `Documents.Add` 建出重复文档）。如果你自己写
脚本在 `Documents.Add` 之后立刻读规范，等 0.4 秒或加重试即可。
另一个 V2.0 坑：**不存在的 .dwt 路径不会报错，而是静默建出普通新图**——
`cad_load_template` 已先 `os.path.isfile` 校验并显式拒绝；若你手发命令加载
模板，请自己先确认文件存在。

### 第四步：`尚未调用 CoInitialize`

**这个错不应该再出现。** 如果出现了，说明桥的版本落后——当前实现用单 STA 线程 + 命令队列，所有 COM 调用都在同一个已初始化的线程上。检查 `bridge/cad_com.py` 里有没有 `_ComWorker` 类。

---

## 画出来的东西不对

### 中文变成 `???`

字体问题，见 [api-semantics.md 中文只能用 SHX 大字体](api-semantics.md#中文只能用-shx-大字体)。

速查：`fontFile = gbenor.shx` + `bigFontFile = gbcbig.shx`，并且要重新 `TextStyles.Item()` 取一次才能设属性。

### 图形跑到奇怪的位置

三个高频原因：

1. `AddCylinder` / `AddBox` 的 center 被当成底面中心（实际是几何中心）
2. `Rotate3D` 传了度数（实际要弧度）
3. `Boolean` 的 operation 用错（`0=并集 / 1=交集 / 2=差集`）

详见 [api-semantics.md](api-semantics.md)。

### 实体凭空消失

`Boolean(1, other)` 求交集结果为空时，**双方一起被删**。检查差集和交集的参数顺序。

### 尺寸标注读出垃圾数字

`AddDimRotated` 的第 4 个参数是**弧度**。传 `90.0` 会得到 50.064 这种数字。

### 保存失败

```
com_error ... '保存文档时出错' ... -2145320861
```

先关掉占用该路径的文档，再删 `.dwl` / `.dwl2` 锁文件。脚本里这两个动作通常是 `close_docs_named()` + `clear_stale_files()`。

---

## 端口问题

### 8770 被占

```powershell
Get-NetTCPConnection -LocalPort 8770 -State Listen
```

要么改面板里的 `httpPort`（同时改 `bundle-autocad-ui/cordis.patch.yml` 里 MCP 客户端的 `url`，两处必须一致），要么杀掉占用者。

### 进程树有两层

```
.venv\Scripts\python.exe   ← 父：插件直接 spawn 的
└─ D:\miniconda3\python.exe ← 孙：真正监听 8770 的
```

这是正常的：venv 的 `python.exe` 是个启动器，会再拉一个真实解释器。**看护逻辑因此不把「父进程退出」等同于「端点死亡」**，它会先探端口。

### 改了代码怎么重启

```powershell
# 找到并杀掉整棵进程树
Get-Process python | Where-Object { $_.Path -like '*DSH*' } | Stop-Process -Force
```

或者直接用面板的「重启」按钮，它会 `terminate()` + `waitForExit()`。

---

## MCP 工具看不到

Agent 侧工具名是 `mcp__autocad__<tool>`。如果全都不存在：

1. 检查 `bundle-autocad-ui/cordis.patch.yml` 里的 `autocad-mcp` 行是否启用
2. 检查 `url` 的端口和面板里的 `httpPort` 是否一致
3. 桥的 `failOnStartupError: false` 意味着**启动失败不会报错**，只会安静重连——去面板看状态灯和日志

---

## 还是不行

收集这几样东西：

- 面板「实时日志」的全部内容
- 面板「运行状态」那一栏（桥目录、Python 来源、状态文件路径）
- `%USERPROFILE%\.dsh\autocad-bridge\config.json`
- AutoCAD 版本号（`cad_status` 会返回）

开 issue 时贴上来。
