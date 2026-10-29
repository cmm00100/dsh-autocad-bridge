# -*- coding: utf-8 -*-
"""DN300 工业蝶式阀门 3D 模型（多色装配，非单一实体）。

设计参数（PN16 对夹式软密封蝶阀，DN300）：
  公称通径      DN300      流道 Ø300
  结构长度      78 mm      对夹式
  法兰外径      Ø460       PN16 DN300
  螺栓孔        12-Ø26     分布圆 Ø410
  阀体外径      Ø380
  阀板          Ø296 x 42
  阀杆          Ø50
  颈部          Ø110，法兰面高 300
  顶法兰        Ø200 x 20，ISO 5211 F16，8-Ø18 分布圆 Ø165
  手柄          260 x 34 x 14

坐标系：X 水平 / Y 为流道轴线 / Z 为阀杆轴线（向上）。

实测 API 语义（不要改）：
  AddCylinder(center, r, h)  center = 几何中心，轴向 Z
  AddBox(center, l, w, h)    center = 几何中心
  Rotate3D(p1, p2, 角)       角 = 弧度
  Boolean(op, other)         op: 0=并集 1=交集 2=差集
"""
import math
import os
import time

import pythoncom
from win32com.client import VARIANT, gencache

BUSY = {-2147418111, -2147417846}

# 产物写到脚本旁边的 out/ 目录 —— 换台机器、换个目录都能跑
OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "out")
os.makedirs(OUT_DIR, exist_ok=True)
OUT = os.path.join(OUT_DIR, "valve-dn300-3d.dwg")

# ---------------- 设计尺寸（真实 mm） ----------------
DN = 300.0
FACE_TO_FACE = 78.0
FLANGE_OD = 460.0
FLANGE_T = 22.0
BOLT_N, BOLT_D, BOLT_BCD = 12, 26.0, 410.0
BODY_OD = 380.0
DISC_D, DISC_T = 296.0, 42.0
STEM_D = 50.0
NECK_D, NECK_TOP = 110.0, 280.0
TOP_OD, TOP_T = 200.0, 20.0
TOP_BOLT_N, TOP_BOLT_D, TOP_BCD = 8, 18.0, 165.0
LEVER_L, LEVER_W, LEVER_T = 260.0, 34.0, 14.0

TOP_FACE = NECK_TOP + TOP_T          # 300
BODY_R = BODY_OD / 2.0


def retry(call, label="", attempts=120, delay=0.3):
    for _ in range(attempts):
        try:
            return call()
        except pythoncom.com_error as exc:
            hr = exc.hresult if exc.hresult < 0 else exc.hresult - (1 << 32)
            if hr not in BUSY:
                raise
            time.sleep(delay)
    raise RuntimeError("AutoCAD 一直忙，放弃：" + label)


def P(x, y, z=0.0):
    return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [float(x), float(y), float(z)])


def new_document(app):
    """Documents.Add() 不幂等：忙时抛错但文档已建好，盲重试会堆出一堆空文档。"""
    before = int(retry(lambda: app.Documents.Count))
    for _ in range(200):
        try:
            return app.Documents.Add()
        except pythoncom.com_error as exc:
            hr = exc.hresult if exc.hresult < 0 else exc.hresult - (1 << 32)
            if hr not in BUSY:
                raise
            now = int(retry(lambda: app.Documents.Count))
            if now > before:
                return retry(lambda: app.Documents.Item(now - 1))
            time.sleep(0.3)
    raise RuntimeError("新建文档失败")


def axis_y(solid, cx, cy, cz):
    """把轴向 Z 的圆柱转成轴向 Y。"""
    retry(lambda: solid.Rotate3D(P(cx, cy, cz), P(cx + 1.0, cy, cz), math.pi / 2))


def axis_x(solid, cx, cy, cz):
    """把轴向 Z 的圆柱转成轴向 X。"""
    retry(lambda: solid.Rotate3D(P(cx, cy, cz), P(cx, cy + 1.0, cz), math.pi / 2))


def main():
    pythoncom.CoInitialize()
    app = retry(lambda: gencache.EnsureDispatch("AutoCAD.Application.24", bForDemand=False))
    retry(lambda: setattr(app, "Visible", True), "visible")
    doc = new_document(app)
    sp = retry(lambda: doc.ModelSpace)
    try:
        retry(lambda: doc.SetVariable("INSUNITS", 4))  # 毫米
    except Exception:
        pass

    def layer(name, color):
        def ensure():
            try:
                doc.Layers.Item(name)
            except pythoncom.com_error:
                obj = doc.Layers.Add(name)
                obj.color = color

        retry(ensure, "layer " + name)

    L_BODY, L_DISC, L_STEM, L_NECK, L_TOP, L_LEVER = "阀体", "阀板", "阀杆", "支架", "顶法兰", "手柄"
    layer(L_BODY, 8)    # 深灰
    layer(L_DISC, 5)    # 蓝
    layer(L_STEM, 2)    # 黄
    layer(L_NECK, 4)    # 青
    layer(L_TOP, 6)     # 品红
    layer(L_LEVER, 1)   # 红

    # ---------- 1. 阀体：外筒 + 两端法兰，然后挖流道与螺栓孔 ----------
    half_face = FACE_TO_FACE / 2.0
    body = retry(lambda: sp.AddCylinder(P(0, 0, 0), BODY_R, FACE_TO_FACE))
    axis_y(body, 0, 0, 0)                       # Y -39..39

    fy = half_face - FLANGE_T / 2.0             # 法兰中心离中面
    flanges = []
    for sign in (-1.0, 1.0):
        f = retry(lambda: sp.AddCylinder(P(0, sign * fy, 0), FLANGE_OD / 2.0, FLANGE_T))
        axis_y(f, 0, sign * fy, 0)
        flanges.append(f)

    for f in flanges:
        retry(lambda: body.Boolean(0, f), "union flange")

    # 流道 Ø300
    bore = retry(lambda: sp.AddCylinder(P(0, 0, 0), DN / 2.0, FACE_TO_FACE + 40.0))
    axis_y(bore, 0, 0, 0)
    retry(lambda: body.Boolean(2, bore), "subtract bore")

    # 12-Ø26 螺栓孔（错开 15°，避开颈部）
    for i in range(BOLT_N):
        ang = math.radians(15.0 + i * (360.0 / BOLT_N))
        hx = (BOLT_BCD / 2.0) * math.cos(ang)
        hz = (BOLT_BCD / 2.0) * math.sin(ang)
        hole = retry(lambda: sp.AddCylinder(P(hx, 0, hz), BOLT_D / 2.0, FACE_TO_FACE + 40.0))
        axis_y(hole, hx, 0, hz)
        retry(lambda: body.Boolean(2, hole), "subtract bolt hole")

    retry(lambda: setattr(body, "Layer", L_BODY))

    # ---------- 2. 阀板：Ø296 x 42，轴线 Y，中间挖 Ø50 阀杆孔 ----------
    disc = retry(lambda: sp.AddCylinder(P(0, 0, 0), DISC_D / 2.0, DISC_T))
    axis_y(disc, 0, 0, 0)
    stem_bore = retry(lambda: sp.AddCylinder(P(0, 0, 0), STEM_D / 2.0, DISC_T + 40.0))
    axis_y(stem_bore, 0, 0, 0)
    retry(lambda: disc.Boolean(2, stem_bore), "disc stem hole")
    retry(lambda: setattr(disc, "Layer", L_DISC))

    # ---------- 3. 阀杆：Ø50，Z 从 -150 到 300 ----------
    stem_lo, stem_hi = -150.0, TOP_FACE
    stem = retry(lambda: sp.AddCylinder(P(0, 0, (stem_lo + stem_hi) / 2.0), STEM_D / 2.0, stem_hi - stem_lo))
    retry(lambda: setattr(stem, "Layer", L_STEM))

    # ---------- 4. 颈部 Ø110：Z 190..280（穿过阀杆，做成环） ----------
    neck = retry(lambda: sp.AddCylinder(P(0, 0, (BODY_R + NECK_TOP) / 2.0), NECK_D / 2.0, NECK_TOP - BODY_R))
    neck_hole = retry(lambda: sp.AddCylinder(P(0, 0, (BODY_R + NECK_TOP) / 2.0), STEM_D / 2.0, NECK_TOP - BODY_R + 40.0))
    retry(lambda: neck.Boolean(2, neck_hole), "neck bore")
    retry(lambda: setattr(neck, "Layer", L_NECK))

    # ---------- 5. 顶法兰 Ø200 x 20 + 8-Ø18 ----------
    top = retry(lambda: sp.AddCylinder(P(0, 0, NECK_TOP + TOP_T / 2.0), TOP_OD / 2.0, TOP_T))
    for i in range(TOP_BOLT_N):
        ang = math.radians(i * (360.0 / TOP_BOLT_N) + 22.5)
        hx = (TOP_BCD / 2.0) * math.cos(ang)
        hy = (TOP_BCD / 2.0) * math.sin(ang)
        hole = retry(lambda: sp.AddCylinder(P(hx, hy, NECK_TOP + TOP_T / 2.0), TOP_BOLT_D / 2.0, TOP_T + 20.0))
        retry(lambda: top.Boolean(2, hole), "subtract top hole")
    retry(lambda: setattr(top, "Layer", L_TOP))

    # ---------- 6. 手柄：沿 +X 伸出 ----------
    lever = retry(lambda: sp.AddBox(
        P(LEVER_L / 2.0, 0, TOP_FACE + LEVER_T / 2.0), LEVER_L, LEVER_W, LEVER_T))
    retry(lambda: setattr(lever, "Layer", L_LEVER))

    # ---------- 收尾 ----------
    retry(lambda: doc.Regen(1))
    print("=== DN300 蝶阀 3D 部件 ===")
    for name, obj in (("阀体", body), ("阀板", disc), ("阀杆", stem), ("支架", neck), ("顶法兰", top), ("手柄", lever)):
        lo, hi = retry(lambda: obj.GetBoundingBox())
        print("  %-4s bbox X[%8.2f,%8.2f] Y[%8.2f,%8.2f] Z[%8.2f,%8.2f]"
              % (name, lo[0], hi[0], lo[1], hi[1], lo[2], hi[2]))

    print("图元总数:", int(retry(lambda: sp.Count)))
    ctr = {}
    for i in range(int(retry(lambda: sp.Count))):
        ent = retry(lambda: sp.Item(i))
        lay = str(retry(lambda: ent.Layer))
        ctr[lay] = ctr.get(lay, 0) + 1
    print("按图层:", ctr)
    for lay_name in (L_BODY, L_DISC, L_STEM, L_NECK, L_TOP, L_LEVER):
        col = int(retry(lambda: doc.Layers.Item(lay_name).color))
        print("  图层 %-6s 颜色索引 %d" % (lay_name, col))

    # 着色 + 等轴测
    for var, val in (("VSCURRENT", "概念"),):
        try:
            retry(lambda: doc.SetVariable(var, val))
        except Exception:
            try:
                retry(lambda: doc.SetVariable("SHADEMODE", "Conceptual"))
            except Exception:
                pass
    try:
        retry(lambda: doc.SetVariable("VPOINT", VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [1.0, -1.2, 0.9])))
    except Exception:
        pass
    retry(lambda: doc.SendCommand("_.ZOOM\n_E\n"))
    retry(lambda: doc.Regen(1))

    retry(lambda: doc.SaveAs(OUT))
    print("saved ->", OUT)


if __name__ == "__main__":
    main()
