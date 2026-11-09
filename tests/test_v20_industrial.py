# -*- coding: utf-8 -*-
r"""
V2.0 正式工业版 live 验收（28 项，状态无关、可重复）：
  ① 图纸合规自检  ② 企业 .dwt 模板  ③ 图元数据反解析  ④ 工程活动日志

跑法：
  & "F:\DSH对CAD2024插件\bridge\.venv\Scripts\python.exe" -u "F:\DSH对CAD2024插件\tests\test_v20_industrial.py"

前置：AutoCAD 2024 在跑、桥在 8770 监听（本脚本只发工具调用，不直连 COM）。
可用 AUTOCAD_BRIDGE_URL 覆盖端点。产物写在 %TEMP%\dsh_v20_out，
跑前清空、跑完保留供人工核对（replay.py / 企业标准模板.dwt / 工程工作报告.md）。

为什么不锁死工具数：V2.0 之后还会继续加工具，所以断言「>=79 且 V2.0 九个工具都在」，
而不是等于某个数。这与 test_backward_compat.py / test_roadmap_acceptance.py 的做法一致。
"""
import asyncio
import json
import os
import shutil
import sys
import time

sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)

from mcp import ClientSession
from mcp.client.streamable_http import streamable_http_client

URL = os.environ.get("AUTOCAD_BRIDGE_URL", "http://127.0.0.1:8770/mcp")
OUT = os.path.join(os.environ.get("TEMP") or ".", "dsh_v20_out")

PASS = FAIL = 0


def unwrap(r):
    """兼容 CallToolResult 的 snake_case / camelCase 字段。"""
    d = {}
    for k in ("isError", "is_error"):
        if hasattr(r, k):
            d["is_error"] = getattr(r, k)
    sc = None
    for k in ("structuredContent", "structured_content"):
        if hasattr(r, k):
            sc = getattr(r, k)
    if isinstance(sc, dict) and "result" in sc:
        sc = sc["result"]
    if sc is not None:
        d["result"] = sc
    if d.get("result") is None:
        parts = []
        for c in getattr(r, "content", []) or []:
            t = getattr(c, "text", None)
            if t:
                parts.append(t)
        raw = "\n".join(parts)
        if raw:
            try:
                d["result"] = json.loads(raw)
            except Exception:
                d["result"] = {"raw": raw}
    d.setdefault("result", {})
    return d


def check(name, res, pred, extra=""):
    global PASS, FAIL
    ok = False
    try:
        ok = bool(pred(res))
    except Exception as exc:
        extra = f"{extra} PRED_ERR={exc}"
    tag = "OK  " if ok else "FAIL"
    print(f"[{tag}] {name} {extra}")
    if ok:
        PASS += 1
    else:
        FAIL += 1
        print("        full result:", json.dumps(res, ensure_ascii=False, default=str)[:900])
    return ok


class Bridge:
    def __init__(self, session):
        self.s = session

    async def call(self, tool_name, **args):
        r = await self.s.call_tool(tool_name, arguments=args)
        return unwrap(r)["result"]


async def main():
    global PASS, FAIL
    if os.path.isdir(OUT):
        shutil.rmtree(OUT, ignore_errors=True)
    os.makedirs(OUT, exist_ok=True)

    t0 = time.time()
    async with streamable_http_client(URL) as (rs, ws):
        async with ClientSession(rs, ws) as session:
            await session.initialize()
            b = Bridge(session)
            tools = await session.list_tools()
            names = [t.name for t in tools.tools]

            # ---------- meta ----------
            check("[meta] all generations registered (>=79 tools)", names,
                  lambda n: len(n) >= 79 and all(x in n for x in (
                      "cad_add_line", "cad_setup_mechanical_environment",
                      "cad_batch_convert", "cad_dimensions_to_3d",
                      "cad_check_drawing_compliance")),
                  f"(got {len(names)})")
            v20 = ["cad_check_drawing_compliance", "cad_reverse_engineer_drawing",
                   "cad_export_reverse_script", "cad_list_templates",
                   "cad_read_drawing_standard", "cad_load_template",
                   "cad_save_as_template", "cad_get_activity_log",
                   "cad_export_activity_report"]
            check("[meta] 9 V2.0 tools registered", names,
                  lambda n: all(x in n for x in v20))

            # ---------- ④ 工程活动日志（不依赖 CAD 连接也能查） ----------
            log0 = await b.call("cad_get_activity_log", limit=5)
            check("[log] get_activity_log basics", log0,
                  lambda r: r.get("ok") is True and "capacity" in r and r["capacity"] == 2000,
                  f"(seq={log0.get('seq')}, buffered={log0.get('buffered')})")
            seq_before = log0.get("seq", 0)

            # ---------- ① 合规自检：故意造问题图纸 ----------
            await b.call("cad_new_document")
            env = await b.call("cad_setup_mechanical_environment")
            check("[V2.0-1] mechanical environment ready", env,
                  lambda r: r.get("ok") is True and r.get("auto_layer_by_kind") is True)

            # 两条重合直线 + 一条碎线 + 一个未闭合多段线
            l1 = await b.call("cad_add_line", start=[0, 0], end=[100, 0], layer="0")
            l2 = await b.call("cad_add_line", start=[0, 0], end=[100, 0], layer="0")
            l3 = await b.call("cad_add_line", start=[0, 10], end=[0.2, 10], layer="0")
            pl = await b.call("cad_add_polyline",
                              points=[[10, 10], [60, 10], [60, 60], [10, 60]],
                              closed=False, layer="0")
            # 一个标注（会自动落「标注」层，合规）与一个错层文字
            d1 = await b.call("cad_add_dim_linear", p1=[0, 0], p2=[100, 0],
                              dim_line_point=[50, -15], angle_deg=0.0)
            txt = await b.call("cad_add_text", text="错层文字", position=[0, 40], layer="0")

            comp = await b.call("cad_check_drawing_compliance", min_line_length=1.0)
            cats = set()
            for it in comp.get("issues", []):
                cats.add(it.get("category"))
            s = comp.get("summary", {})
            print("        compliance summary:", json.dumps(s, ensure_ascii=False))
            print("        categories hit:", sorted(cats))
            check("[V2.0-1] report structure", comp,
                  lambda r: r.get("ok") is True and "summary" in r and "compliant" in r
                  and set(r.get("checks", [])) >= {"图层错误", "标注不规范", "未闭合轮廓", "重复线", "碎线"})
            check("[V2.0-1] duplicate line detected", comp,
                  lambda r: "重复线" in [i["category"] for i in r.get("issues", [])])
            check("[V2.0-1] short line detected", comp,
                  lambda r: "碎线" in [i["category"] for i in r.get("issues", [])])
            check("[V2.0-1] open polyline detected", comp,
                  lambda r: "未闭合轮廓" in [i["category"] for i in r.get("issues", [])])
            check("[V2.0-1] wrong-layer text reported", comp,
                  lambda r: "图层错误" in [i["category"] for i in r.get("issues", [])])
            check("[V2.0-1] compliant is False on dirty drawing", comp,
                  lambda r: r.get("compliant") is False and r.get("summary", {}).get("errors", 0) > 0)

            # 干净图纸：删掉问题后应合规（新建一张，只画规范对象）
            await b.call("cad_new_document")
            await b.call("cad_setup_mechanical_environment")
            await b.call("cad_add_rectangle", corner1=[0, 0], corner2=[120, 80])
            comp2 = await b.call("cad_check_drawing_compliance")
            print("        clean compliance summary:", json.dumps(comp2.get("summary", {}), ensure_ascii=False))
            check("[V2.0-1] clean drawing is compliant", comp2,
                  lambda r: r.get("compliant") is True,
                  f"(issues={comp2.get('summary', {}).get('issues')})")

            # ---------- ③ 图元数据反解析 ----------
            rev = await b.call("cad_reverse_engineer_drawing", include_script=True)
            script = rev.get("script", "")
            check("[V2.0-3] reverse returns entity list", rev,
                  lambda r: r.get("ok") is True and r.get("returned", 0) >= 1
                  and isinstance(r.get("entities"), list) and r.get("histogram"))
            check("[V2.0-3] replay script generated", rev,
                  lambda r: r.get("script_chars", 0) > 100 and "CadBridge" in r.get("script", "")
                  and "b.add_" in r.get("script", ""))
            check("[V2.0-3] standard block present", rev,
                  lambda r: "layers" in r.get("standard", {}) and "sysvars" in r.get("standard", {}))
            # 图元参数真实：矩形是 4 点闭合多段线
            polys = [e for e in rev.get("entities", []) if "Polyline" in e.get("type", "")]
            check("[V2.0-3] polyline params readable", rev,
                  lambda r: len(polys) >= 1 and polys[0].get("closed") is True
                  and len(polys[0].get("coordinates") or []) == 8,
                  f"(polys={len(polys)})")

            script_path = os.path.join(OUT, "replay.py")
            exp = await b.call("cad_export_reverse_script", path=script_path)
            check("[V2.0-3] script exported to disk", exp,
                  lambda r: r.get("ok") is True and r.get("size", 0) > 100
                  and os.path.isfile(script_path) and os.path.getsize(script_path) == r.get("size"),
                  f"(path={exp.get('path')}, size={exp.get('size')})")
            # 生成的脚本必须是合法 Python
            try:
                import ast
                ast.parse(open(script_path, encoding="utf-8").read())
                script_valid = True
            except Exception as e:
                script_valid = False
                print("        script parse error:", e)
            check("[V2.0-3] replay script is valid Python", {"ok": script_valid},
                  lambda r: r["ok"] is True)

            # ---------- ② 企业 .dwt 模板 ----------
            lt = await b.call("cad_list_templates")
            check("[V2.0-2] templates listed", lt,
                  lambda r: r.get("ok") is True and r.get("count", 0) >= 1
                  and os.path.isdir(r.get("template_dir", "") or ""),
                  f"(count={lt.get('count')}, dir={lt.get('template_dir')})")

            std = await b.call("cad_read_drawing_standard")
            check("[V2.0-2] drawing standard readable", std,
                  lambda r: r.get("ok") is True and isinstance(r.get("layers"), list)
                  and "轮廓" in [l.get("name") for l in r.get("layers", [])]
                  and "sysvars" in r and r["sysvars"].get("INSUNITS") is not None,
                  f"(layers={len(std.get('layers', []))}, insunits={std.get('sysvars', {}).get('INSUNITS')})")

            dwt_path = os.path.join(OUT, "企业标准模板.dwt")
            sat = await b.call("cad_save_as_template", path=dwt_path)
            check("[V2.0-2] save_as_template creates .dwt", sat,
                  lambda r: r.get("ok") is True and r.get("size", 0) > 1000
                  and os.path.isfile(dwt_path),
                  f"(size={sat.get('size')})")

            loaded = await b.call("cad_load_template", path=dwt_path)
            check("[V2.0-2] load_template opens new doc", loaded,
                  lambda r: r.get("ok") is True and r.get("documents_after", 0) > r.get("documents_before", 0)
                  and r.get("template", "").lower().endswith(".dwt"),
                  f"(before={loaded.get('documents_before')} after={loaded.get('documents_after')})")
            check("[V2.0-2] custom spec survives template", loaded,
                  lambda r: "轮廓" in [l.get("name") for l in r.get("standard", {}).get("layers", [])]
                  and r.get("standard", {}).get("sysvars", {}).get("INSUNITS") is not None)

            # 坏路径必须显式报错，不能静默建普通新图
            before_docs = (await b.call("cad_list_documents"))
            bad = await b.call("cad_load_template", path="__no_such_dwt__\\nope.dwt")
            check("[V2.0-2] missing template is rejected", bad,
                  lambda r: r.get("ok") is False and "找不到" in str(r.get("error", "")),
                  f"(error={str(bad.get('error', ''))[:120]})")

            # ---------- ④ 活动日志 + 工作报告 ----------
            log1 = await b.call("cad_get_activity_log", limit=2000)
            check("[log] seq advanced with every _call", log1,
                  lambda r: r.get("seq", 0) > seq_before,
                  f"(before={seq_before} after={log1.get('seq')})")
            check("[log] records carry action/ok/duration", log1,
                  lambda r: len(r.get("records", [])) > 0
                  and all(("action" in x and "ok" in x and "duration_ms" in x) for x in r["records"]))
            flt = await b.call("cad_get_activity_log", limit=2000, action="合规", ok_only=True)
            check("[log] action filter works", flt,
                  lambda r: all("合规" in x["action"] and x["ok"] is True for x in r.get("records", [])))

            report_path = os.path.join(OUT, "工程工作报告.md")
            rep = await b.call("cad_export_activity_report", path=report_path, limit=2000)
            check("[log] activity report exported", rep,
                  lambda r: r.get("ok") is True and r.get("size", 0) > 100
                  and os.path.isfile(report_path),
                  f"(size={rep.get('size')}, records={rep.get('records')})")
            if os.path.isfile(report_path):
                body = open(report_path, encoding="utf-8").read()
                check("[log] report has title + activity table", {"body": body},
                      lambda r: "工程工作报告" in r["body"] and "活动明细" in r["body"]
                      and "失败明细" in r["body"])

    print(f"\n=== V20 LIVE: PASS={PASS} FAIL={FAIL} total={PASS + FAIL} ===")
    print(f"=== ELAPSED_SEC={round(time.time() - t0, 1)} OUT={OUT} ===")
    return 0 if FAIL == 0 else 1


if __name__ == "__main__":
    sys.exit(asyncio.run(main()))
