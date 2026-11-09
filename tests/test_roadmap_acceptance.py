# -*- coding: utf-8 -*-
"""ROADMAP acceptance test — one repeatable line per version's 交付标准.

The ROADMAP (docs/ROADMAP.md) states a one-sentence delivery standard for each
version. This test turns those three sentences into LIVE, REPEATABLE checks
against the running bridge + AutoCAD 2024:

  V1.1  生成机械零件三视图 + 自动标注 + 分层规范
        -> one drawing: 3 orthographic views, the FULL dimension system
           (linear/aligned/radius/diameter/angular/ordinate), GB layers,
           and Chinese MText technical requirements.
  V1.2  批量整理整套设备图纸、自动出 BOM、自动转 PDF 交付
        -> build source DWGs, auto BOM from blocks, single-doc DXF/DWG/PDF,
           then batch-convert the whole set and verify every file on disk.
  V1.3  输入零件尺寸 → 自动 2D 三视图 + 自动 3D 模型
        -> 2D three-view rectangles + dimensions_to_3d(L,W,H) + view
           switches verified via VIEWDIR + profile_to_3d.

Design rules (same as test_backward_compat.py):
  * Status-independent: every phase opens its OWN new document, and no
    assertion depends on `auto_layer_by_kind` being OFF (it is a process-global
    switch that stays ON once setup_mechanical_environment ran). Explicit
    layers are used where determinism matters.
  * Self-contained: all artifacts go into a fixed, best-effort-cleaned system
    temp dir; real files are verified with os.path.getsize, not just "ok".
  * Harness notes (mcp 2.3.0): streamable_http_client (NOT streamablehttp),
    CallToolResult uses snake_case is_error / structured_content, do NOT pass
    a custom httpx client (mcp vendors httpx2; default read timeout is 300s).

Run (bridge must be live on 127.0.0.1:8770):
    <venv-python> tests/test_roadmap_acceptance.py
"""
import asyncio
import os
import shutil
import sys
import tempfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
sys.stderr.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)

from mcp import ClientSession
from mcp.client.streamable_http import streamable_http_client

URL = os.environ.get("AUTOCAD_BRIDGE_URL", "http://127.0.0.1:8770/mcp")
OUT = os.path.join(tempfile.gettempdir(), "dsh_roadmap_acceptance")
RESULTS = []  # (label, passed)


def unwrap(res):
    err = getattr(res, "is_error", None)
    if err is None:
        err = getattr(res, "isError", False)
    sc = getattr(res, "structured_content", None)
    if sc is None:
        sc = getattr(res, "structuredContent", None)
    if isinstance(sc, dict) and "result" in sc:
        sc = sc["result"]
    if sc is not None:
        return bool(err), sc
    import json
    return bool(err), json.loads(res.content[0].text) if res.content else {}


def chk(label, cond, detail=""):
    RESULTS.append((label, bool(cond)))
    print(("PASS " if cond else "FAIL ") + label + (" :: " + str(detail) if detail else ""), flush=True)


def size_of(path):
    try:
        return os.path.getsize(path)
    except OSError:
        return -1


class Bridge:
    def __init__(self, session):
        self.s = session

    async def call(self, tool_name, **args):
        r = await self.s.call_tool(tool_name, arguments=args)
        err, data = unwrap(r)
        return {"_error": True, "raw": data} if err else data

    async def entities(self):
        return await self.call("cad_list_entities", limit=500)


def ent_types(ents):
    return [e.get("type") for e in ents.get("entities", [])]


# ---------------------------------------------------------------------------
# V1.1 交付标准：生成机械零件三视图 + 自动标注 + 分层规范
# ---------------------------------------------------------------------------
async def accept_v11(b):
    V = "[V1.1] "
    d = await b.call("cad_new_document")
    chk(V + "new_document", d.get("ok"), str(d)[:120])

    env = await b.call("cad_setup_mechanical_environment")
    chk(V + "setup_mechanical_environment (分层规范)",
        env.get("ok") and env.get("auto_layer_by_kind") is True, str(env)[:160])

    # 三视图：主视 / 俯视 / 左视，三个矩形，省略 layer -> 自动落「轮廓」
    views = [([0, 0], [120, 80]), ([0, -100], [120, -72]), ([150, 0], [178, 80])]
    rects = []
    for c1, c2 in views:
        r = await b.call("cad_add_rectangle", corner1=c1, corner2=c2)
        rects.append(r)
        chk(V + f"三视图 rect {c1}->{c2}", r.get("ok"), str(r)[:120])

    layers = await b.call("cad_list_layers")
    names = {l.get("name") for l in layers.get("layers", [])}
    want_layers = {"轮廓", "中心线", "虚线", "标注", "文字", "剖面线"}
    chk(V + "6 个机械标准图层齐全", want_layers.issubset(names),
        str(sorted(names)))

    ents = await b.entities()
    by_h = {e.get("handle"): e for e in ents.get("entities", [])}
    auto = [by_h.get(r.get("handle", ""), {}).get("layer") for r in rects]
    chk(V + "绘图自动归属「轮廓」层", all(x == "轮廓" for x in auto), str(auto))

    # 完整尺寸标注体系：6 种
    dim_calls = [
        ("linear-H", "cad_add_dim_linear", dict(p1=[0, 0], p2=[120, 0], dim_line_point=[60, -15], angle_deg=0.0)),
        ("linear-V", "cad_add_dim_linear", dict(p1=[0, 0], p2=[0, 80], dim_line_point=[-15, 40], angle_deg=90.0)),
        ("aligned",  "cad_add_dim_aligned", dict(p1=[0, 0], p2=[60, 40], dim_line_point=[40, -20])),
        ("radial",   "cad_add_dim_radial", dict(center=[60, 40], chord_point=[85, 40], leader_length=10.0)),
        ("diametric","cad_add_dim_diametric", dict(chord1=[35, 40], chord2=[85, 40], leader_length=10.0)),
        ("angular",  "cad_add_dim_angular", dict(vertex=[0, 0], p1=[100, 0], p2=[60, 60], text_point=[120, 40])),
        ("ordinate", "cad_add_dim_ordinate", dict(definition_point=[300, 300], leader_end_point=[340, 330], use_x_axis=True)),
    ]
    for tag, tool, args in dim_calls:
        r = await b.call(tool, **args)
        chk(V + f"自动标注 {tag}", r.get("ok") and r.get("measurement") is not None
            and float(r.get("measurement") or 0) > 0 and bool(r.get("handle")),
            f"measurement={r.get('measurement')} {r.get('error','')}")

    # 两条线性标注精确值（ROADMAP「自动标注」的可量化证据）
    lh = await b.call("cad_add_dim_linear", p1=[0, 0], p2=[120, 0], dim_line_point=[60, -30], angle_deg=0.0)
    chk(V + "水平标注 measurement=120", lh.get("measurement") == 120.0, str(lh.get("measurement")))
    lv = await b.call("cad_add_dim_linear", p1=[0, 0], p2=[0, 80], dim_line_point=[-30, 40], angle_deg=90.0)
    chk(V + "竖直标注 measurement=80", lv.get("measurement") == 80.0, str(lv.get("measurement")))

    # 中文技术要求（MText）
    mt = await b.call("cad_add_mtext", text="技术要求\\P未注公差 GB/T 1804-m", position=[0, -140], width=160.0)
    chk(V + "中文 MText 技术要求", mt.get("ok"), str(mt)[:120])

    e2 = await b.entities()
    t2 = ent_types(e2)
    chk(V + "6 种标注对象都在图纸里",
        all(any(k in (x or "") for x in t2) for k in
            ["RotatedDimension", "AlignedDimension", "RadialDimension",
             "DiametricDimension", "2LineAngularDimension", "OrdinateDimension"]),
        str(t2))
    chk(V + "MText 对象在图纸里", "AcDbMText" in t2, str(t2))

    z = await b.call("cad_zoom_extents")
    chk(V + "zoom_extents", z.get("ok"), str(z)[:120])


# ---------------------------------------------------------------------------
# V1.2 交付标准：批量整理整套设备图纸、自动出 BOM、自动转 PDF 交付
# ---------------------------------------------------------------------------
async def accept_v12(b):
    V = "[V1.2] "
    src1 = os.path.join(OUT, "equip_a.dwg")
    src2 = os.path.join(OUT, "equip_b.dwg")
    outdir = os.path.join(OUT, "batch_out")

    d = await b.call("cad_new_document")
    chk(V + "new_document (设备图纸)", d.get("ok"), str(d)[:120])
    await b.call("cad_add_rectangle", corner1=[0, 0], corner2=[120, 80])

    # 属性块 -> BOM 数据源
    blk = await b.call("cad_create_block", name="ACC_PART",
                       entities=[{"type": "circle", "center": [0, 0], "radius": 10}],
                       attributes=[{"tag": "PARTNO", "value": "A-001", "prompt": "Part No",
                                    "height": 2.5, "position": [0, -15], "mode": 0}])
    chk(V + "create_block (属性块)", blk.get("ok") and blk.get("created") is True, str(blk)[:160])
    for pos, val in (([60, 40], "A-001"), ([90, 40], "A-002")):
        ins = await b.call("cad_insert_block", name="ACC_PART", position=pos,
                           attributes={"PARTNO": val})
        chk(V + f"insert_block @ {pos}", ins.get("ok"), str(ins)[:120])

    cnt = await b.call("cad_count_blocks")
    chk(V + "count_blocks >= 2", cnt.get("ok") and cnt.get("total_refs", 0) >= 2, str(cnt)[:160])
    bom = await b.call("cad_generate_bom")
    chk(V + "自动出 BOM 表 (右下角对齐)",
        bom.get("ok") and bom.get("columns") == 4 and bom.get("rows", 0) >= 2
        and isinstance(bom.get("position"), list), str(bom)[:200])

    # 两套源 DWG（Wblock 写出独立文件，不作为文档打开 -> batch 会真的 Open+Close）
    e1 = await b.call("cad_export_dwg", path=src1)
    chk(V + "源 DWG A 落盘", e1.get("ok") and size_of(e1.get("path", "")) > 0,
        f"size={size_of(e1.get('path',''))} path={e1.get('path')}")
    e2 = await b.call("cad_export_dwg", path=src2)
    chk(V + "源 DWG B 落盘", e2.get("ok") and size_of(e2.get("path", "")) > 0,
        f"size={size_of(e2.get('path',''))} path={e2.get('path')}")

    # 单文档 DXF / PDF 交付
    dxf = await b.call("cad_export_dxf", path=os.path.join(OUT, "single.dxf"))
    chk(V + "单文档转 DXF", dxf.get("ok") and size_of(dxf.get("path", "")) > 0,
        f"size={size_of(dxf.get('path',''))}")
    pdf = await b.call("cad_export_pdf", path=os.path.join(OUT, "single.pdf"))
    chk(V + "单文档转 PDF", pdf.get("ok") and size_of(pdf.get("path", "")) > 0,
        f"size={size_of(pdf.get('path',''))}")

    # 批量整理整套图纸 -> 每种格式 -> 关闭
    batch = await b.call("cad_batch_convert", paths=[src1, src2], out_dir=outdir,
                         formats=["dxf", "dwg", "pdf"], close_after=True)
    chk(V + "batch_convert succeeded=2 failed=0",
        batch.get("ok") and batch.get("succeeded") == 2 and batch.get("failed") == 0,
        str(batch)[:240])
    for r in batch.get("results", []):
        tag = os.path.basename(str(r.get("input", "?")))
        outs = r.get("outputs", {})
        for fmt in ("dxf", "dwg", "pdf"):
            o = outs.get(fmt, {})
            p = o.get("path", "")
            chk(V + f"batch {tag} {fmt} 文件非空",
                bool(r.get("ok")) and int(o.get("size", -1)) > 0 and size_of(p) > 0,
                f"size={o.get('size')} disk={size_of(p)}")


# ---------------------------------------------------------------------------
# V1.3 交付标准：输入零件尺寸 → 自动 2D 三视图 + 自动 3D 模型
# ---------------------------------------------------------------------------
async def accept_v13(b):
    V = "[V1.3] "
    d = await b.call("cad_new_document")
    chk(V + "new_document (零件)", d.get("ok"), str(d)[:120])

    # 2D 三视图（沿用 V1.1 的绘图 + 分层）
    rects = []
    for c1, c2 in (([0, 0], [44, 40]), ([0, -50], [44, -22]), ([60, 0], [88, 40])):
        r = await b.call("cad_add_rectangle", corner1=c1, corner2=c2)
        rects.append(r)
    chk(V + "自动 2D 三视图 (3 矩形)", all(r.get("ok") for r in rects),
        str([r.get("ok") for r in rects]))

    # 一句话：输入零件尺寸 -> 3D box + 等轴测 + 着色
    d3 = await b.call("cad_dimensions_to_3d", length=44, width=28, height=40)
    chk(V + "dimensions_to_3d bbox=[44,28,40]",
        d3.get("ok") and d3.get("bbox", {}).get("max") == [44.0, 28.0, 40.0],
        str(d3.get("bbox")) + " " + str(d3.get("error", "")))
    chk(V + "自动等轴测视图 verified",
        d3.get("viewdir") == [1.0, -1.0, 1.0] and d3.get("verified") is True,
        f"viewdir={d3.get('viewdir')} verified={d3.get('verified')}")
    chk(V + "自动着色已发送", d3.get("shade_sent") is True, str(d3.get("shade")))

    h = d3.get("handle", "")
    si = await b.call("cad_solid_info", handle=h)
    chk(V + "solid_info 实体存活", si.get("ok") and si.get("alive"), str(si)[:120])

    # 3D 视图自动切换（VIEWDIR 可读验证）
    for view, want in (("front", [0.0, -1.0, 0.0]), ("top", [0.0, 0.0, 1.0]),
                       ("right", [1.0, 0.0, 0.0])):
        r = await b.call("cad_set_3d_view", view=view)
        chk(V + f"set_3d_view {view} verified",
            r.get("ok") and r.get("viewdir") == want and r.get("verified") is True,
            f"viewdir={r.get('viewdir')} verified={r.get('verified')}")

    # 2D 草图 + 厚度 -> 拉伸实体
    p3 = await b.call("cad_profile_to_3d",
                      points=[[0, 0], [20, 0], [20, 10], [0, 10]], thickness=12.0)
    chk(V + "profile_to_3d 厚度=12",
        p3.get("ok") and p3.get("bbox", {}).get("max", [0, 0, 0])[2] == 12.0
        and "solid" in str(p3.get("object_name", "")).lower(),
        str(p3.get("bbox")) + " " + str(p3.get("object_name")) + " " + str(p3.get("error", "")))

    ents = await b.entities()
    t = ent_types(ents)
    chk(V + "图纸里 >=2 个 3D 实体",
        sum(1 for x in t if "3dsolid" in (x or "").lower()) >= 2,
        str([x for x in t if "3d" in (x or "").lower()]))
    chk(V + "2D 三视图与 3D 模型共存",
        sum(1 for x in t if (x or "") == "AcDbPolyline") >= 3
        and any("3dsolid" in (x or "").lower() for x in t), str(t))


async def main():
    # 固定输出目录：每轮开始尽力清空，重复跑不会累积；跑完保留产物供人工核对。
    try:
        shutil.rmtree(OUT, ignore_errors=True)
    except OSError:
        pass
    os.makedirs(OUT, exist_ok=True)

    async with streamable_http_client(URL) as (rs, ws):
        async with ClientSession(rs, ws) as session:
            await session.initialize()
            b = Bridge(session)

            names = [t.name for t in (await session.list_tools()).tools]
            chk("[meta] tool table = 70 (V1.0+V1.1+V1.2+V1.3)", len(names) == 70, f"got {len(names)}")

            await accept_v11(b)
            await accept_v12(b)
            await accept_v13(b)

    ok = sum(1 for _, c in RESULTS if c)
    print(f"\n=== ROADMAP ACCEPTANCE: PASS={ok} FAIL={len(RESULTS)-ok} total={len(RESULTS)} ===", flush=True)
    for v in ("[meta]", "[V1.1]", "[V1.2]", "[V1.3]"):
        sub = [c for n, c in RESULTS if n.startswith(v)]
        print(f"  {v}: {sum(sub)}/{len(sub)}", flush=True)
    for n, c in RESULTS:
        if not c:
            print("  FAILED:", n, flush=True)
    print(f"ARTIFACTS: {OUT}", flush=True)
    return 0 if ok == len(RESULTS) else 1


if __name__ == "__main__":
    sys.exit(asyncio.run(main()))
