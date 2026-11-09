# Tests

## Tests

Two live, status-independent acceptance scripts. Both need the bridge up on
`127.0.0.1:8770` and the bridge's own virtualenv; both exit `0` only when
every assertion passes; both can be run repeatedly against one warm process.

| Script | Checks | What it locks down |
|---|---|---|
| `test_backward_compat.py` | 32 | ROADMAP「向下兼容 / 可生产使用」：四个版本工具混调 |
| `test_roadmap_acceptance.py` | 50 | ROADMAP 三条「一句话交付标准」逐版本 live 验收 |

### ROADMAP delivery-standard acceptance

`test_roadmap_acceptance.py` turns each version's 交付标准 into a repeatable
live check (see [docs/ROADMAP.md](../docs/ROADMAP.md)):

- **V1.1 生成机械零件三视图 + 自动标注 + 分层规范** — 3 orthographic
  rectangles, all six GB layers, auto-assignment to 「轮廓」, all six dimension
  kinds (linear-H/V, aligned, radial, diametric, angular, ordinate) with exact
  measurements 120/80, and Chinese MText technical requirements.
- **V1.2 批量整理整套设备图纸、自动出 BOM、自动转 PDF 交付** — attribute
  blocks → `count_blocks` → auto BOM aligned to the sheet corner → two source
  DWGs → single-doc DXF/PDF → `batch_convert` over the whole set → every DXF,
  DWG and PDF verified non-empty **on disk** (`os.path.getsize`).
- **V1.3 输入零件尺寸 → 自动 2D 三视图 + 自动 3D 模型** — 2D three-view
  rectangles + `dimensions_to_3d(44,28,40)` with iso view + shading →
  `solid_info` → front/top/right `VIEWDIR` verification → `profile_to_3d`
  thickness 12 → 2D and 3D coexist on one drawing.

```powershell
& "F:\DSH对CAD2024插件\bridge\.venv\Scripts\python.exe" `
  "F:\DSH对CAD2024插件\tests\test_roadmap_acceptance.py"
```

Expected output ends with:

```
=== ROADMAP ACCEPTANCE: PASS=50 FAIL=0 total=50 ===
  [meta]: 1/1
  [V1.1]: 20/20
  [V1.2]: 17/17
  [V1.3]: 12/12
```

Artifacts are written to `%TEMP%\dsh_roadmap_acceptance` (cleared at the start
of each run, kept afterwards for manual inspection). Verified: two consecutive
runs, 50/50 each, on fresh drawings with no bridge restart.

---

## Cross-version backward-compatibility regression

`test_backward_compat.py` is the durable regression for the ROADMAP's
"fully backward compatible / production-ready" rule. It calls tools from all
four shipped generations (V1.0 → V1.3) on one drawing, interleaved, and
re-checks that the oldest generation still works after the newest one ran —
including right after a 3D build + viewport change, where a real historical
bug (`-2145386414 对象已打开进行通知`) used to live.

### Prerequisites

- The DSH AutoCAD bridge is running and listening on `127.0.0.1:8770`
  (the plugin's watchdog starts it; check with `Get-NetTCPConnection
  -LocalPort 8770 -State Listen`).
- AutoCAD 2024 is installed and reachable over COM. The bridge will wake it
  automatically if it is not running (≈15 s cold start).
- Use the bridge's own virtualenv so `mcp` 2.3.0 is importable.

### Run

```powershell
& "F:\DSH对CAD2024插件\bridge\.venv\Scripts\python.exe" `
  "F:\DSH对CAD2024插件\tests\test_backward_compat.py"
```

Override the endpoint if the bridge is elsewhere:

```powershell
$env:AUTOCAD_BRIDGE_URL = "http://127.0.0.1:8770/mcp"
```

Expected output ends with:

```
=== BACKWARD-COMPAT REGRESSION: PASS=32 FAIL=0 total=32 ===
```

Exit code is `0` only when every assertion passes.

### Why it is repeatable on a warm process

`cad_setup_mechanical_environment` sets `auto_layer_by_kind = True` on the
process-global `CadBridge` singleton and never resets it — not even on
`cad_reconnect` or `cad_new_document`. An earlier one-shot probe therefore
had to restart the whole bridge process before it could assert V1.0's
classic "omit layer → use the current CAD layer" default.

This test instead passes an explicit `layer="0"` in Phase A. `_layer_for()`
checks `if layer: return layer` first, so an explicit layer always wins
regardless of the switch state. That removes the cold-start dependency and
makes the test runnable any number of times against the same warm bridge
(verified: two consecutive runs, 32/32 each, on two different new drawings).

### What it asserts (32 checks)

| Phase | Generation | Key assertions |
|---|---|---|
| A | V1.0 | new document, line, circle, list_entities, explicit layer `0` |
| B | V1.1 | mechanical environment, auto-layer by kind, dimension = 100, MText; Phase-A entities unchanged and still on layer `0` |
| C | V1.2 | centerlines, read-only `count_blocks` / `list_xrefs`, style-only `unify_styles` (`layer_changed=0`) |
| D | V1.3 | 3D primitive + `solid_info`, `dimensions_to_3d`, `set_3d_view front`, then **V1.0 line and V1.1 dimension again** |
| E | All | entities survive every generation, two solids at end, count grows 2→5→7→11, `zoom_extents` still works |

Other test assets live in `bridge/` (the gitignored development playground)
and in `bundle-autocad-ui/` (the shipped host/UI unit tests); see
[docs/architecture.md](../docs/architecture.md#测试策略) for the full matrix.
