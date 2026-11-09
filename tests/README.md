# Tests

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
