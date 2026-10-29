# DSH AutoCAD Bridge

把本机 **AutoCAD 2024** 变成 DeepSeek Harness Agent 的一双手：查状态、画图元、管图层、跑原生命令、建 3D 实体，全部通过 MCP 工具调用完成。附带一个可视化的**控制面板**，负责托管后台桥进程的启停、参数、状态灯与实时日志。

> English: [README.en.md](README.en.md)

![AutoCAD 三视图输出示例](docs/images/part-3view.png)

---

## 它解决什么问题

AutoCAD 本身是一个桌面 GUI 程序，没有对外的 HTTP API。想在 Agent 里驱动它，中间必须有一层把「命令行式的 COM 调用」翻译成「结构化的工具调用」的桥。这个项目就是那座桥，加上一个管着它的插件。

```
┌──────────────────────────────────────────────────────────────┐
│  DeepSeek Harness（Web / 桌面 Electron）                      │
│                                                              │
│   Agent ──► mcp__autocad__cad_add_circle(...)                │
│                    │                                         │
│                    ▼                                         │
│   @deepseek-ai/dsh-mcp-client  ── HTTP/SSE ──┐               │
│                                             │               │
│   @local/autocad-bridge（本仓库）◄──────────┼── 控制面板      │
│     · 拉起/看护 Python 桥进程                 │   启停·参数·日志│
│     · 参数持久化 · 崩溃自动重启               │               │
└─────────────────────────────────────────────┼───────────────┘
                                              │
                        ┌─────────────────────▼──────────────┐
                        │  bridge/autocad_server.py          │
                        │  （Python · MCP Server · stdio/HTTP）│
                        │            │                       │
                        │            ▼                       │
                        │  bridge/cad_com.py                 │
                        │  单 STA 线程 + 命令队列            │
                        └────────────┼───────────────────────┘
                                     │ COM (pywin32)
                                     ▼
                        ┌────────────────────────────────────┐
                        │  AutoCAD 2024（acad.exe，桌面进程）│
                        └────────────────────────────────────┘
```

两层设计的原因：**桥（Python）只懂 AutoCAD，插件（JS）只懂进程和界面。** 桥可以脱离 DSH 单独调试，插件可以换掉桥的实现而不动界面。

---

## 功能

### 控制面板

侧边栏新增一个图标，点开是一整块中央面板：

- **状态灯**：绿=进程在且端口通，黄=进程在但端口不通，灰=已停止
- **一键启停重启**：不用再去任务管理器杀 python
- **连接参数**：Python 路径、服务脚本、工作目录、端口、ProgID，改完即存
- **实时日志**：桥进程的 stdout/stderr 直接流进面板，600 行环形缓冲
- **崩溃看护**：进程意外退出后按 1s→2s→4s…最多 30s 的指数退避自动重启，连续 8 次失败后停手并提示
- **依赖自举**：一键在你自己的数据目录建虚拟环境、装 pywin32/mcp/python-dotenv
- **环境探测**：自动找出 Python 与桥目录，自动检测本机实际注册了哪些 AutoCAD ProgID

### 18 个 Agent 工具

| 工具 | 作用 |
|---|---|
| `cad_connect` | 连接 AutoCAD（未运行时通过 COM 启动） |
| `cad_status` | 查状态：`disconnected` / `idle` / `busy` / `error` |
| `cad_escape` | 发 ESC 取消当前命令、关掉模态弹窗（卡死时救命） |
| `cad_list_progids` | 列出本机注册表中真实存在的 ProgID |
| `cad_list_documents` | 列出所有打开的 DWG，标出活动文档 |
| `cad_new_document` | 新建空白图纸（可选 dwt 模板） |
| `cad_save` | 就地保存 / 另存为 |
| `cad_ensure_layer` | 建图层（已存在则复用），支持颜色名或 ACI 索引 |
| `cad_list_layers` | 列出全部图层与颜色 |
| `cad_add_line` | 画直线 |
| `cad_add_circle` | 画圆 |
| `cad_add_rectangle` | 画矩形 |
| `cad_add_polyline` | 画多段线（可闭合） |
| `cad_add_text` | 写单行文字 |
| `cad_add_arc` | 画圆弧（角度用度） |
| `cad_send_command` | 发任意原生命令，如 `_.ZOOM\n_E\n` |
| `cad_list_entities` | 列出模型空间全部图元（类型/图层/句柄） |
| `cad_zoom_extents` | 缩放到全图 |

3D 实体（`AddBox` / `AddCylinder` / `Boolean` / `Rotate3D`）目前通过 `cad_send_command` 配合 Python 脚本使用，见 [示例](#示例)。

---

## 安装

### 前置条件

- Windows（COM 依赖）
- AutoCAD **2024**（其他版本需改 ProgID，见 [排错](docs/troubleshooting.md#progid)）
- Python 3.10–3.12（**不要**装到系统 Python 里，插件会在自己的数据目录建 venv）
- DeepSeek Harness

### 步骤

1. 把本仓库克隆到任意位置：

   ```powershell
   git clone https://github.com/cmm00100/dsh-autocad-bridge.git
   ```

2. 在 DSH 里安装插件包（指向 `bundle-autocad-ui` 目录）：

   让 Agent 执行 `plugin_manager` 的 `install_bundle`，或直接在 profile 里装：

   ```powershell
   dsh plugin --profile web add <克隆路径>\bundle-autocad-ui
   ```

3. 打开 **设置 → 插件**，确认 `autocad-bridge` 这一行是**启用**状态。

4. 点击侧边栏的 **AutoCAD 桥接** 图标，点 **安装依赖**。插件会在 `%USERPROFILE%\.dsh\autocad-bridge\` 下建虚拟环境并装好三个依赖，然后自动启动桥进程。

5. 状态灯变绿后，让 Agent 执行 `mcp__autocad__cad_connect`，即可开始画图。

配置文件落在 `%USERPROFILE%\.dsh\autocad-bridge\config.json`，虚拟环境在同目录的 `.venv`。

---

## 配置

面板里能改的，都会写进 `config.json`：

| 字段 | 默认值 | 说明 |
|---|---|---|
| `pythonPath` | 自动探测 | 跑桥的 Python 解释器 |
| `serverScript` | 自动探测 | `autocad_server.py` 的绝对路径 |
| `workingDirectory` | 自动探测 | 桥进程的工作目录 |
| `httpPort` | `8770` | 桥的 HTTP 端口，须与 MCP 客户端的 `url` 一致 |
| `progId` | `AutoCAD.Application.24` | 见下表 |
| `autoStart` | `true` | 插件加载时是否自动拉起桥进程 |

AutoCAD 版本与 ProgID：

| 版本 | ProgID |
|---|---|
| 2024 | `AutoCAD.Application.24` |
| 2023 | `AutoCAD.Application.23` |
| 2022 | `AutoCAD.Application.22` |

不确定就跑 `cad_list_progids`，它读注册表告诉你本机到底有什么。

桥自己还有一份 `.env`（在 `bundle-autocad-ui/bridge/`），控制超时、重试、默认图层等底层参数，通常不用动。

---

## 示例

`examples/` 下有三个可以直接跑的脚本，展示这套桥能做到什么程度：

| 脚本 | 产物 |
|---|---|
| `part-3view.py` | 按一张轴测参考图画出零件三视图，A3 图框、2:1、13 个尺寸 |
| `valve-3view.py` | DN300 工业蝶式阀门三视图，A0 图框、18 个尺寸、明细表、技术要求 |
| `valve-3d.py` | 同一台阀门的 3D 多色装配模型（6 部件 / 6 图层 / 6 颜色） |

跑法：

```powershell
<venv 里的 python> examples\valve-3d.py
```

产物写到 `examples/out/`。**运行前请确认 AutoCAD 已启动并处于空闲状态**——脚本会往当前活动文档里画东西。

![DN300 蝶阀 3D 模型](docs/images/valve-3d.png)

---

## 文档

- [架构与设计取舍](docs/architecture.md)——为什么是「单 STA 线程 + 命令队列」，为什么分两层
- [掌握的工具语义](docs/api-semantics.md)——AutoCAD COM 那些和官方文档不一致的地方（这份是踩出来的）
- [排错](docs/troubleshooting.md)——连不上、卡死、中文变问号、保存失败
- [免责与已知限制](docs/limitations.md)

---

## 已知限制

- **只能 Windows**：COM 是 Windows 独有的。
- **必须先有 AutoCAD 桌面进程**：纯后台静默启动 COM 不稳定，插件默认会带界面拉起。
- **模态弹窗会锁死 COM**：AutoCAD 弹出对话框时所有调用都会卡住，用 `cad_escape` 尝试解救。
- **单实例串行**：所有 COM 调用排队经过一个 STA 线程，并发调用不会崩，但也不会更快——AutoCAD 本来就是单线程。
- **不做视觉校验**：桥只管把东西画出来。画得对不对、好不好看，得另外看图。
- **控制面板的像素级渲染未经自动化验证**：功能逻辑有 headless 测试覆盖，但界面外观需要你自己点一下确认（见 [限制](docs/limitations.md)）。

---

## 致谢与相关项目

- [DeepSeek Harness](https://github.com/deepseek-ai/dsh)——插件宿主
- [Model Context Protocol](https://modelcontextprotocol.io/)——桥与 Agent 之间的协议
- [pywin32](https://github.com/mhammond/pywin32)——COM 绑定

## 许可

[MIT](LICENSE)
