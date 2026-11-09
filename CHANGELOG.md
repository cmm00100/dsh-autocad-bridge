# Changelog

All notable changes to this project are documented in this file.
Format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).
The iteration plan lives in [docs/ROADMAP.md](docs/ROADMAP.md).

Every version must honour: native AutoCAD COM only, the single-STA-thread
command queue, automatic MCP tool registration, offline-only operation, and
full backward compatibility with tools shipped in earlier versions.

## [Unreleased]

## [1.1.0] - 2026-11-09

V1.1 「工业基础能力补全」。工具数从 18 增至 **39**，全部通过 live AutoCAD
2024 端到端实测。向下兼容：1.0.0 里所有 `mcp__autocad__*` 工具名、参数与
返回结构保持不变。

### Added — V1.1 工业基础能力补全

- **完整尺寸标注体系**（V1.1-1）：新增 6 种标注工具，角度统一用「度」，内部
  转弧度，全部自动落到「标注」层并支持文字覆盖：
  - `cad_add_dim_linear` — 线性/水平/竖直（`AddDimRotated`）
  - `cad_add_dim_aligned` — 对齐（`AddDimAligned`）
  - `cad_add_dim_radial` — 半径（`AddDimRadial`）
  - `cad_add_dim_diametric` — 直径（`AddDimDiametric`）
  - `cad_add_dim_angular` — 角度（`AddDimAngular`）
  - `cad_add_dim_ordinate` — 坐标（`AddDimOrdinate`）
  - `cad_setup_text_style` — 建中文文字样式（`gbenor.shx` + `gbcbig.shx`）
  - `cad_setup_gb_dim_style` — 国标标注样式（24 个 `DIM*` 变量），实测定稿
    的顺序为「建/激活 DimStyle → 设变量 → `DimStyle.CopyFrom(doc)` 固化 →
    才画标注」，否则变量会落在旧样式上读回 `0.18 / Standard`。
- **机械标准图层自动化**（V1.1-2）：6 个国标图层（轮廓/中心线/虚线/标注/
  文字/剖面线），各带标准颜色、线宽、线型（`CENTER` / `DASHED` 自动从
  `acad.lin` 加载）。
  - `cad_setup_standard_layers` — 一次性建齐，幂等
  - `cad_set_layer_state` — 批量改锁定/冻结/开关/颜色/线宽/设为活动层
  - `cad_set_active_layer` — 切换当前图层
  - `cad_setup_mechanical_environment` — **一句话建好整套环境**（标准图层 +
    中文样式 + GB 标注样式 + 文档变量），并打开「按大类自动归属图层」：
    之后画轮廓自动进「轮廓」、标注进「标注」、文字进「文字」。该开关默认
    关闭，刻意保持 1.0.0「省略 layer 就用 CAD 当前图层」的行为。
- **块与属性块**（V1.1-3）：
  - `cad_list_blocks` / `cad_create_block`（线/圆/弧/文字/多段线 + 属性定义，
    幂等不覆盖）/ `cad_insert_block`（可覆写属性值、可旋转缩放）/ 
    `cad_list_block_attributes` / `cad_set_block_attribute` / `cad_replace_block`
    （批量替换块参照并保留插入点/比例/旋转）。
  - 实测：读块属性必须 `dynamic.Dispatch(ref).GetAttributes()`（枚举取回的
    实体会被 gen_py 窄化成 `IAcadEntity`）。
- **文字 / MText / 形位公差**（V1.1-4）：
  - `cad_add_mtext` — 多行文字，支持 `%%c`/`%%p`、堆叠公差
    `{\H0.7x;\S+0.01^-0.02;}`
  - `cad_add_tolerance` — 形位公差框（`AcDbFcf`，`AddTolerance`）

### Fixed — V1.1 修复优化

- **CAD 未启动自动唤醒**（V1.1-5）：`connect()` 重写。冷启动时 COM 拉起
  `acad.exe` 期间 `gencache.EnsureDispatch` 会抛
  `TypeError: This COM object can not automate the makepy process`（不仅是
  `com_error`）。现在用 55s deadline 循环重试（同时捕 com_error/Exception）、
  失败退回 `Dispatch`，`Documents.Add()` 按 `Documents.Count` 核对（该调用
  非幂等）。实测强杀 `acad.exe` 后 15.4 秒自动重启并连上。
- **队列卡顿可观测 + 超时重连**：`_ComWorker` 增加 `pending / busy_seconds /
  last_error` 与 `health()` / `abandon()`；超时提示带上积压量并指引
  `cad_reconnect`。新增 `cad_reconnect` 工具与 `CadBridge.reconnect()`（换
  worker、丢弃旧 COM 指针、不 join 卡死的线程）。`cad_status` 现在返回
  `queue` 健康度。
- **`cad_escape` 缺陷**（1.0.0 遗留）：`PostCommand` 挂在 `IAcadApplication`
  上根本不存在，应挂在**文档**上（`doc.PostCommand`），且 ESC 只能用
  `PostCommand` 排队、不能用 `SendCommand`（后者报「输入无效」）。同时让
  `cad_escape` 走 `_ensure_connected()`，桥刚重启未连接时也能自动连上。

### Changed

- `MCPServer` 的 description / instructions 更新，加入「先 setup → 画轮廓 →
  加标注 → zoom」的推荐流程与角度单位约定。
- 版本号升至 1.1.0。

## [1.0.0] - 2026-10-29

### Added
- Initial open-source release. A single bundle (`bundle-autocad-ui`) that hosts
  a Python COM/MCP bridge process and exposes it as 18 `mcp__autocad__*` tools:
  connect/status/escape/progids/documents/new/save/ensure-layer/list-layers,
  line/circle/rectangle/polyline/text/arc, send-command, list-entities,
  zoom-extents.
- Sidebar control panel with process lifecycle, settings, status light and
  live logs; runtime path discovery; crash auto-restart with exponential
  backoff; dependency bootstrap into a private venv.
- Full documentation (README zh/en, architecture, API semantics,
  troubleshooting, limitations) and three worked examples.
