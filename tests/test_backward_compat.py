# -*- coding: utf-8 -*-
"""Cross-version backward-compatibility regression (V1.0 / V1.1 / V1.2 / V1.3).

This is a DURABLE, REPEATABLE test — it can be run any number of times against
the SAME warm bridge process, unlike a one-shot probe. It proves the ROADMAP's
"fully backward compatible / production-ready" rule by calling tools from all
four generations on ONE drawing, interleaved, and re-checking that earlier
generations still work AFTER later ones ran.

Status independence (important)
--------------------------------
`auto_layer_by_kind` is a PROCESS-GLOBAL switch on the CadBridge singleton,
turned on by `cad_setup_mechanical_environment` and never reset (including by
`cad_reconnect` / `cad_new_document`). Therefore V1.0's classic
"omit layer -> use CAD current layer" default cannot be asserted on a warm
process without a fresh restart.

Rather than depend on a cold start, Phase A passes an EXPLICIT ``layer="0"``.
`_layer_for()` checks ``if layer: return layer`` FIRST, so an explicit layer
always wins regardless of the switch state. That makes the whole test
deterministic; the switch's own opt-in behaviour is asserted in Phase B.

Run (bridge must be live on 127.0.0.1:8770):
    <venv-python> tests/test_backward_compat.py

Harness notes (mcp 2.3.0): streamable_http_client (not streamablehttp_client),
CallToolResult uses snake_case is_error / structured_content, do NOT pass a
custom httpx client (mcp vendors httpx2; default read timeout is 300s).
"""
import asyncio
import os
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
sys.stderr.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)

from mcp import ClientSession
from mcp.client.streamable_http import streamable_http_client

URL = os.environ.get("AUTOCAD_BRIDGE_URL", "http://127.0.0.1:8770/mcp")
RESULTS = []


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


def chk(name, cond, detail=""):
    RESULTS.append((name, bool(cond)))
    print(("PASS " if cond else "FAIL ") + name + (" :: " + str(detail) if detail else ""), flush=True)


class Bridge:
    def __init__(self, session):
        self.s = session

    async def call(self, tool_name, **args):
        r = await self.s.call_tool(tool_name, arguments=args)
        err, data = unwrap(r)
        return {"_error": True, "raw": data} if err else data

    async def entities(self):
        return await self.call("cad_list_entities", limit=500)


def by_handle(ents, handle):
    for e in ents.get("entities", []):
        if e.get("handle") == handle:
            return e
    return {}


def types(ents):
    return [e.get("type") for e in ents.get("entities", [])]


async def main():
    async with streamable_http_client(URL) as (rs, ws):
        async with ClientSession(rs, ws) as session:
            await session.initialize()
            b = Bridge(session)

            names = [t.name for t in (await session.list_tools()).tools]
            # V1.0+V1.1+V1.2+V1.3 = 70 tools. 后续版本（V2.0 等）会继续往上加，
            # 所以这里断言「至少 70 且四代代表工具都在」，而不是锁死一个数字。
            chk("all four generations registered (>=70 tools)",
                len(names) >= 70 and all(x in names for x in (
                    "cad_add_line", "cad_setup_mechanical_environment",
                    "cad_batch_convert", "cad_dimensions_to_3d")),
                f"got {len(names)}")

            # ------------------------------------------------------------------
            # Phase A: V1.0 baseline. Explicit layer="0" -> deterministic on a
            # warm process too. These two entities must survive every later
            # generation UNCHANGED (same handle, still on layer "0").
            # ------------------------------------------------------------------
            d = await b.call("cad_new_document")
            chk("V1.0 cad_new_document", d.get("ok"), str(d)[:120])

            l0 = await b.call("cad_add_line", start=[0, 0], end=[100, 0], layer="0")
            chk("V1.0 add_line", l0.get("ok") and l0.get("length") == 100.0, str(l0)[:120])
            h_line0 = l0.get("handle", "")
            c0 = await b.call("cad_add_circle", center=[50, 0], radius=10, layer="0")
            chk("V1.0 add_circle", c0.get("ok"), str(c0)[:120])
            h_circle0 = c0.get("handle", "")

            eA = await b.entities()
            chk("V1.0 list_entities total=2", eA.get("ok") and eA.get("total") == 2,
                f"total={eA.get('total')}")
            chk("V1.0 explicit layer=0 honoured",
                by_handle(eA, h_line0).get("layer") == "0" and by_handle(eA, h_circle0).get("layer") == "0",
                str([(e.get("handle"), e.get("layer")) for e in eA.get("entities", [])]))

            # ------------------------------------------------------------------
            # Phase B: V1.1 environment + dims + text. Flips the opt-in switch;
            # V1.0 tools must still work and the two Phase-A entities must NOT
            # be deleted or moved off layer "0".
            # ------------------------------------------------------------------
            env = await b.call("cad_setup_mechanical_environment")
            chk("V1.1 setup_mechanical_environment (switch on)",
                env.get("ok") and env.get("auto_layer_by_kind") is True, str(env)[:160])

            l1 = await b.call("cad_add_line", start=[0, -30], end=[100, -30])
            chk("V1.0 add_line still works after V1.1 env", l1.get("ok"), str(l1)[:120])
            e_after = await b.entities()
            chk("V1.1 auto-layer by kind -> 轮廓",
                by_handle(e_after, l1.get("handle", "")).get("layer") == "轮廓",
                str([(e.get("handle"), e.get("layer")) for e in e_after.get("entities", [])]))
            chk("V1.1 did NOT rewrite Phase-A entities",
                by_handle(e_after, h_line0).get("layer") == "0",
                str(by_handle(e_after, h_line0)))

            dim = await b.call("cad_add_dim_linear", p1=[0, 0], p2=[100, 0],
                               dim_line_point=[50, -15], angle_deg=0.0)
            chk("V1.1 add_dim_linear measurement=100",
                dim.get("ok") and dim.get("measurement") == 100.0,
                str(dim.get("measurement")) + " " + str(dim.get("error", "")))

            mt = await b.call("cad_add_mtext", text="Regress\\Pfour generations", position=[0, -60], width=120.0)
            chk("V1.1 add_mtext", mt.get("ok"), str(mt)[:120])

            eB = await b.entities()
            tB = types(eB)
            chk("V1.0 entities survived V1.1",
                h_line0 in [e.get("handle") for e in eB.get("entities", [])]
                and h_circle0 in [e.get("handle") for e in eB.get("entities", [])],
                f"total={eB.get('total')}")
            chk("V1.1 dimension object present", "AcDbRotatedDimension" in tB, str(tB))
            chk("V1.1 MText object present", "AcDbMText" in tB, str(tB))

            # ------------------------------------------------------------------
            # Phase C: V1.2 mechanical aid + read-only inventories + style unify.
            # unify_styles uses style-only (layer="") so no layer is disturbed.
            # ------------------------------------------------------------------
            cl = await b.call("cad_add_centerlines", corner1=[0, 0], corner2=[100, 60])
            chk("V1.2 add_centerlines", cl.get("ok"), str(cl)[:120])
            cnt = await b.call("cad_count_blocks")
            chk("V1.2 count_blocks (read-only)", cnt.get("ok"), str(cnt)[:120])
            xr = await b.call("cad_list_xrefs")
            chk("V1.2 list_xrefs (read-only)", xr.get("ok"), str(xr)[:120])
            uni = await b.call("cad_unify_styles", style="DSH-CN", layer="")
            chk("V1.2 unify_styles (style only)",
                uni.get("ok") and uni.get("layer_changed") == 0 and uni.get("style_changed", 0) >= 1,
                str(uni)[:160])

            eC = await b.entities()
            chk("V1.2 centerlines added entities", eC.get("ok") and eC.get("total") > eB.get("total"),
                f"before={eB.get('total')} after={eC.get('total')}")
            chk("V1.1 dimension still present after V1.2", "AcDbRotatedDimension" in types(eC), str(types(eC)))

            # ------------------------------------------------------------------
            # Phase D: V1.3 3D modelling + view switch, then go BACK to
            # V1.0/V1.1. The "对象已打开进行通知" bug lived exactly at this seam.
            # ------------------------------------------------------------------
            prim = await b.call("cad_add_3d_primitive", kind="box", center=[300, 0, 0],
                                params={"length": 10, "width": 20, "height": 30})
            chk("V1.3 add_3d_primitive",
                prim.get("ok") and prim.get("bbox", {}).get("max") == [305.0, 10.0, 15.0],
                str(prim.get("bbox")) + " " + str(prim.get("error", "")))
            h_box = prim.get("handle", "")

            si = await b.call("cad_solid_info", handle=h_box)
            chk("V1.3 solid_info alive", si.get("ok") and si.get("alive"), str(si)[:120])

            d3 = await b.call("cad_dimensions_to_3d", length=44, width=28, height=40)
            chk("V1.3 dimensions_to_3d",
                d3.get("ok") and d3.get("bbox", {}).get("max") == [44.0, 28.0, 40.0],
                str(d3.get("bbox")) + " " + str(d3.get("error", "")))

            vv = await b.call("cad_set_3d_view", view="front")
            chk("V1.3 set_3d_view front",
                vv.get("ok") and vv.get("viewdir") == [0.0, -1.0, 0.0],
                f"viewdir={vv.get('viewdir')} verified={vv.get('verified')}")

            # THE regression guard: old tools after a 3D build + VSCURRENT
            # command + viewport reassignment (the historic failure point).
            l2 = await b.call("cad_add_line", start=[0, -90], end=[100, -90], layer="0")
            chk("V1.0 add_line works AFTER 3D+view", l2.get("ok"), str(l2)[:120])
            dim2 = await b.call("cad_add_dim_linear", p1=[0, -90], p2=[100, -90],
                                dim_line_point=[50, -105], angle_deg=0.0)
            chk("V1.1 add_dim_linear works AFTER 3D+view",
                dim2.get("ok") and dim2.get("measurement") == 100.0,
                str(dim2.get("measurement")) + " " + str(dim2.get("error", "")))

            # ------------------------------------------------------------------
            # Phase E: final consistency.
            # ------------------------------------------------------------------
            eD = await b.entities()
            tD = types(eD)
            hs = [e.get("handle") for e in eD.get("entities", [])]
            chk("V1.0 entities survived ALL later generations",
                h_line0 in hs and h_circle0 in hs, f"total={eD.get('total')}")
            chk("V1.0 entities still on layer 0 at the end",
                by_handle(eD, h_line0).get("layer") == "0" and by_handle(eD, h_circle0).get("layer") == "0",
                str([(by_handle(eD, h).get("handle"), by_handle(eD, h).get("layer")) for h in (h_line0, h_circle0)]))
            chk("3D solids present at end",
                sum(1 for t in tD if "Solid" in (t or "")) >= 2,
                str([t for t in tD if "Solid" in (t or "")]))
            chk("entity count grew A->B->C->D",
                eA.get("total") < eB.get("total") < eC.get("total") < eD.get("total"),
                f"A={eA.get('total')} B={eB.get('total')} C={eC.get('total')} D={eD.get('total')}")

            z = await b.call("cad_zoom_extents")
            chk("V1.0 zoom_extents still works", z.get("ok"), str(z)[:120])

    ok = sum(1 for _, c in RESULTS if c)
    print(f"\n=== BACKWARD-COMPAT REGRESSION: PASS={ok} FAIL={len(RESULTS)-ok} total={len(RESULTS)} ===", flush=True)
    for n, c in RESULTS:
        if not c:
            print("  FAILED:", n, flush=True)
    return 0 if ok == len(RESULTS) else 1


if __name__ == "__main__":
    sys.exit(asyncio.run(main()))
