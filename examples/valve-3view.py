# -*- coding: utf-8 -*-
"""DN300 工业蝶式阀门 —— A0 三视图 + 尺寸标注（v3）

v1 的致命问题（用户已指出）：中文全显示成 ???、内容越出图框、标注线杂乱穿插、
表格与视图糊在一起、颜色花花绿绿。v3 的对策：

  1. 字体：统一 DSH-CN 样式 = gbenor.shx + gbcbig.shx 大字体（实测可显示中文/Ø/℃）
  2. 布局：先算死三个视图 + 表格的包围盒，各占一块互不重叠的区域，最后自检越界
  3. 标注：每个视图只标必要尺寸，全部引到视图外侧空白区，不跨视图
  4. 线宽：粗实线 0.50 / 细线 0.18 / 中心线红色 CENTER 线型，靠线宽分粗细
  5. 尺寸样式在画标注之前就设好（否则已有标注不跟随新样式）

投影（第一角法）：
  主视图 = X-Z，正对法兰面，见 Ø460 圆
  俯视图 = X-Y，与主视图长对正
  左视图 = Y-Z，与主视图高平齐
"""
import math
import os
import time

import pythoncom
from win32com.client import VARIANT, dynamic, gencache

BUSY = {-2147418111, -2147417846}

# 产物写到脚本旁边的 out/ 目录 —— 换台机器、换个目录都能跑
OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "out")
os.makedirs(OUT_DIR, exist_ok=True)
OUT = os.path.join(OUT_DIR, "valve-dn300-drawing.dwg")

# ==================== 零件尺寸（与 valve3d.py 同源） ====================
DN = 300.0
FACE_TO_FACE = 78.0
FLANGE_OD = 460.0
FLANGE_T = 22.0
BOLT_N, BOLT_D, BOLT_BCD = 12, 26.0, 410.0
BODY_OD = 380.0
STEM_D = 50.0
NECK_D, NECK_TOP = 110.0, 280.0
TOP_OD, TOP_T = 200.0, 20.0
TOP_BOLT_N, TOP_BOLT_D, TOP_BCD = 8, 18.0, 165.0
LEVER_L, LEVER_W, LEVER_T = 260.0, 34.0, 14.0
TOP_FACE = NECK_TOP + TOP_T          # 300 = 流道中心 → 顶法兰面

RAD = math.radians
R_FLANGE = FLANGE_OD / 2.0           # 230
R_BODY = BODY_OD / 2.0               # 190
R_BORE = DN / 2.0                    # 150
R_BCD = BOLT_BCD / 2.0               # 205
R_NECK = NECK_D / 2.0                # 55
R_TOP = TOP_OD / 2.0                 # 100
R_TOPBCD = TOP_BCD / 2.0             # 82.5
HALF_H = FACE_TO_FACE / 2.0          # 39

# ==================== 图框与分区布局（A0 = 1189 x 841，1:1） ====================
FW, FH = 1189.0, 841.0
M_LEFT, M_OTHER = 25.0, 5.0
IN_L, IN_R = M_LEFT, FW - M_OTHER        # 25 .. 1184
IN_B, IN_T = M_OTHER, FH - M_OTHER       # 5 .. 836

MVX, MVY = 300.0, 470.0                  # 主视图中心
TVX, TVY = 300.0, 150.0                  # 俯视图中心（X 与主视图对齐）
LVX, LVY = 790.0, 480.0                  # 左视图中心

CN_STYLE = "DSH-CN"


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


def clear_stale_files(path):
    """删掉目标文件与 AutoCAD 残留的 .dwl/.dwl2 锁文件，否则 SaveAs 报「保存文档时出错」。"""
    for candidate in (path, path + ".dwl", path + ".dwl2", path + "l", path + "l2"):
        try:
            if os.path.exists(candidate):
                os.remove(candidate)
                print("删除旧文件:", candidate)
        except Exception as exc:
            print("  删除失败", candidate, type(exc).__name__, str(exc)[:60])


def close_docs_named(app, wanted):
    """SaveAs 到已被别的文档占用的路径会失败，先关掉同名文档。"""
    wanted = os.path.normcase(os.path.abspath(wanted))
    for _ in range(20):
        hit = None
        for i in range(int(retry(lambda: app.Documents.Count))):
            doc = retry(lambda: app.Documents.Item(i))
            full = ""
            try:
                full = str(retry(lambda: doc.FullName))
            except Exception:
                full = ""
            name = str(retry(lambda: doc.Name))
            if full and full != "None" and os.path.normcase(os.path.abspath(full)) == wanted:
                hit = doc
                break
            if not full or full == "None":
                # 未保存文档：名字直接比对
                if os.path.normcase(name) == os.path.normcase(os.path.basename(wanted)):
                    hit = doc
                    break
        if hit is None:
            return
        print("关闭同名文档:", str(retry(lambda: hit.Name)))
        retry(lambda: hit.Close(False))


def new_document(app):
    """Documents.Add() 非幂等：忙时报错但文档其实已建好，必须核对 Count。"""
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


class Drawing:
    def __init__(self, doc):
        self.doc = doc
        self.sp = retry(lambda: doc.ModelSpace)

    # ---------- 样式 / 图层 ----------
    def ensure_style(self, name, font="gbenor.shx", big="gbcbig.shx"):
        """建/改文字样式。

        坑：ts.Add() 返回的对象直接用 dynamic 设属性会报
        `Property '<unknown>.fontFile' can not be set.`，
        必须用 ts.Item(name) 重新取一次再设。
        """
        ts = retry(lambda: self.doc.TextStyles)
        try:
            retry(lambda: ts.Item(name))
        except pythoncom.com_error:
            retry(lambda: ts.Add(name))
        style = retry(lambda: ts.Item(name))
        late = dynamic.Dispatch(style)
        late.fontFile = font
        late.bigFontFile = big
        late.Height = 0.0
        late.Width = 1.0
        return style

    def layer(self, name, color=7, lineweight=25, linetype=None):
        """lineweight 单位 0.01mm：50 = 0.50mm，25 = 0.25，18 = 0.18。"""
        def impl():
            try:
                lay = self.doc.Layers.Item(name)
            except pythoncom.com_error:
                lay = self.doc.Layers.Add(name)
            late = dynamic.Dispatch(lay)
            try:
                late.color = color
            except Exception:
                pass
            try:
                late.Lineweight = lineweight
            except Exception:
                pass
            if linetype:
                try:
                    late.Linetype = linetype
                except Exception:
                    pass

        retry(impl, "layer " + name)

    # ---------- 图元 ----------
    def line(self, x1, y1, x2, y2, layer):
        e = retry(lambda: self.sp.AddLine(P(x1, y1), P(x2, y2)))
        retry(lambda: setattr(e, "Layer", layer))
        return e

    def rect(self, x0, y0, x1, y1, layer):
        self.line(x0, y0, x1, y0, layer)
        self.line(x1, y0, x1, y1, layer)
        self.line(x1, y1, x0, y1, layer)
        self.line(x0, y1, x0, y0, layer)

    def circle(self, cx, cy, r, layer):
        e = retry(lambda: self.sp.AddCircle(P(cx, cy), r))
        retry(lambda: setattr(e, "Layer", layer))
        return e

    def text(self, s, x, y, h, layer="文字", align="left"):
        e = retry(lambda: self.sp.AddText(s, P(x, y), h))
        retry(lambda: setattr(e, "Layer", layer))
        retry(lambda: setattr(e, "StyleName", CN_STYLE))
        if align != "left":
            try:
                lo, hi = retry(lambda: e.GetBoundingBox())
                shift = -(hi[0] - lo[0]) / 2.0 if align == "center" else -(hi[0] - lo[0])
                retry(lambda: e.Move(P(0.0, 0.0), P(shift, 0.0)))
            except Exception:
                pass
        return e


def main():
    pythoncom.CoInitialize()
    app = retry(lambda: gencache.EnsureDispatch("AutoCAD.Application.24", bForDemand=False))
    retry(lambda: setattr(app, "Visible", True))

    close_docs_named(app, OUT)
    clear_stale_files(OUT)
    doc = new_document(app)

    try:
        doc.Name  # noqa: B018  —— 触发一次 COM 往返，确认文档任可用
    except Exception:
        pass

    d = Drawing(doc)
    d.ensure_style(CN_STYLE)

    for var, val in (("INSUNITS", 4), ("LWDISPLAY", 1), ("LUNITS", 2), ("DIMASSOC", 2),
                     ("MEASUREMENT", 1), ("ORTHOMODE", 0)):
        try:
            retry(lambda: doc.SetVariable(var, val))
        except Exception:
            pass

    # 尺寸样式必须在画任何标注之前设置，否则既有标注不会跟随
    for var, val in (("DIMTXT", 7.0), ("DIMASZ", 3.5), ("DIMEXE", 2.5),
                     ("DIMEXO", 1.5), ("DIMGAP", 1.2), ("DIMDEC", 0),
                     ("DIMZIN", 0), ("DIMSCALE", 1.0), ("DIMTAD", 1),
                     ("DIMTIX", 0), ("DIMSOXD", 0), ("DIMCLRD", 7),
                     ("DIMCLRE", 7), ("DIMCLRT", 7), ("DIMTXSTY", CN_STYLE),
                     ("DIMBLK", "."), ("DIMLDRBLK", ".")):
        try:
            retry(lambda: doc.SetVariable(var, val))
        except Exception as exc:
            print("  变量设置失败", var, type(exc).__name__)

    # ---------- 图层 ----------
    d.layer("图框", 7, 50)
    d.layer("粗实线", 7, 50)
    d.layer("细实线", 7, 18)
    d.layer("中心线", 1, 18, "CENTER")
    d.layer("虚线", 7, 25, "DASHED")
    d.layer("标注", 7, 18)
    d.layer("文字", 7, 25)
    d.layer("表格", 7, 35)

    # ==================== 图框 + 标题栏 ====================
    d.rect(0.0, 0.0, FW, FH, "图框")
    d.rect(IN_L, IN_B, IN_R, IN_T, "图框")

    TB_W, TB_H = 180.0, 56.0
    tbx, tby = IN_R - TB_W, IN_B
    d.rect(tbx, tby, IN_R, tby + TB_H, "表格")
    for i in range(1, 4):
        d.line(tbx, tby + i * 14.0, IN_R, tby + i * 14.0, "表格")
    d.line(tbx + 40.0, tby, tbx + 40.0, tby + TB_H, "表格")
    d.line(tbx + 110.0, tby, tbx + 110.0, tby + TB_H, "表格")
    d.text("DN300 蝶式阀门", tbx + 45.0, tby + 40.0, 8.5, "文字")
    d.text("三视图 + 尺寸标注", tbx + 45.0, tby + 26.0, 6.0, "文字")
    d.text("比例 1:1", tbx + 45.0, tby + 13.0, 6.0, "文字")
    d.text("单位 mm", tbx + 114.0, tby + 40.0, 6.0, "文字")
    d.text("材料 WCB", tbx + 114.0, tby + 26.0, 6.0, "文字")
    d.text("共 1 张", tbx + 114.0, tby + 13.0, 6.0, "文字")

    # 图名（明细表下方的空白区）
    d.text("DN300 工业蝶式阀门", 1030.0, 780.0, 11.0, "文字", align="center")
    d.text("对夹式  PN16  软密封", 1030.0, 760.0, 7.0, "文字", align="center")

    # ==================== 主视图（X-Z） ====================
    d.circle(MVX, MVY, R_FLANGE, "粗实线")        # Ø460
    d.circle(MVX, MVY, R_BODY, "粗实线")          # Ø380
    d.circle(MVX, MVY, R_BORE, "虚线")            # Ø300 流道
    d.circle(MVX, MVY, R_BCD, "中心线")           # Ø410 螺栓圆
    for i in range(BOLT_N):
        a = RAD(15.0 + i * (360.0 / BOLT_N))
        d.circle(MVX + R_BCD * math.cos(a), MVY + R_BCD * math.sin(a),
                 BOLT_D / 2.0, "粗实线")

    neck_bot = MVY + R_FLANGE                     # 700
    neck_top_y = MVY + NECK_TOP                   # 750
    top_top_y = MVY + TOP_FACE                    # 770
    # 颈部 Ø110
    d.line(MVX - R_NECK, neck_bot, MVX - R_NECK, neck_top_y, "粗实线")
    d.line(MVX + R_NECK, neck_bot, MVX + R_NECK, neck_top_y, "粗实线")
    # 顶法兰 Ø200 x 20
    d.rect(MVX - R_TOP, neck_top_y, MVX + R_TOP, top_top_y, "粗实线")
    # 顶法兰上的 8-Ø18（前后被遮，画虚线）
    d.line(MVX - R_TOPBCD, neck_top_y, MVX - R_TOPBCD, top_top_y, "虚线")
    d.line(MVX + R_TOPBCD, neck_top_y, MVX + R_TOPBCD, top_top_y, "虚线")
    # 阀杆 Ø50
    d.line(MVX - STEM_D / 2.0, MVY + R_BODY, MVX - STEM_D / 2.0, top_top_y, "虚线")
    d.line(MVX + STEM_D / 2.0, MVY + R_BODY, MVX + STEM_D / 2.0, top_top_y, "虚线")
    # 手柄端面 34 x 14
    d.rect(MVX - LEVER_W / 2.0, top_top_y, MVX + LEVER_W / 2.0,
           top_top_y + LEVER_T, "粗实线")

    d.line(MVX - R_FLANGE - 30.0, MVY, MVX + R_FLANGE + 30.0, MVY, "中心线")
    d.line(MVX, MVY - R_FLANGE - 30.0, MVX, top_top_y + LEVER_T + 25.0, "中心线")
    d.text("主视图", MVX - 120.0, MVY + R_FLANGE + 16.0, 8.0, "文字", align="center")

    # ==================== 俯视图（X-Y） ====================
    d.rect(TVX - R_FLANGE, TVY - HALF_H, TVX + R_FLANGE, TVY + HALF_H, "粗实线")
    for xx in (R_BODY, R_BORE):
        d.line(TVX - xx, TVY - HALF_H, TVX - xx, TVY + HALF_H, "虚线")
        d.line(TVX + xx, TVY - HALF_H, TVX + xx, TVY + HALF_H, "虚线")
    d.circle(TVX, TVY, R_NECK, "粗实线")
    d.circle(TVX, TVY, R_TOP, "粗实线")
    d.circle(TVX, TVY, R_TOPBCD, "中心线")
    for i in range(TOP_BOLT_N):
        a = RAD(i * (360.0 / TOP_BOLT_N) + 22.5)
        d.circle(TVX + R_TOPBCD * math.cos(a), TVY + R_TOPBCD * math.sin(a),
                 TOP_BOLT_D / 2.0, "粗实线")
    d.rect(TVX, TVY - LEVER_W / 2.0, TVX + LEVER_L, TVY + LEVER_W / 2.0, "粗实线")

    d.line(TVX - R_FLANGE - 30.0, TVY, TVX + LEVER_L + 25.0, TVY, "中心线")
    d.line(TVX, TVY - HALF_H - 26.0, TVX, TVY + HALF_H + 26.0, "中心线")
    d.text("俯视图", TVX, TVY - HALF_H - 78.0, 8.0, "文字", align="center")

    # ==================== 左视图（Y-Z） ====================
    lx, ly = LVX, LVY
    d.rect(lx - HALF_H, ly - R_FLANGE, lx + HALF_H, ly + R_FLANGE, "粗实线")
    # 两侧法兰台阶：厚 22，中间收成 Ø380
    body_x0 = lx - HALF_H + FLANGE_T
    body_x1 = lx + HALF_H - FLANGE_T
    for s in (-1.0, 1.0):
        y_edge = ly + s * R_FLANGE
        y_body = ly + s * R_BODY
        x_out = lx + s * HALF_H
        x_in = lx + s * (HALF_H - FLANGE_T)
        d.line(x_out, y_edge, x_in, y_edge, "粗实线")
        d.line(x_in, y_edge, x_in, y_body, "粗实线")
    d.line(body_x0, ly + R_BODY, body_x1, ly + R_BODY, "粗实线")
    d.line(body_x0, ly - R_BODY, body_x1, ly - R_BODY, "粗实线")
    # 流道 Ø300
    d.line(lx - HALF_H, ly - R_BORE, lx + HALF_H, ly - R_BORE, "虚线")
    d.line(lx - HALF_H, ly + R_BORE, lx + HALF_H, ly + R_BORE, "虚线")
    # 颈部 Ø110 与顶法兰 Ø200
    d.line(lx - R_NECK, ly + R_BODY, lx - R_NECK, ly + NECK_TOP, "粗实线")
    d.line(lx + R_NECK, ly + R_BODY, lx + R_NECK, ly + NECK_TOP, "粗实线")
    d.rect(lx - R_TOP, ly + NECK_TOP, lx + R_TOP, ly + TOP_FACE, "粗实线")
    # 手柄端面 34 x 14
    d.rect(lx - LEVER_W / 2.0, ly + TOP_FACE,
           lx + LEVER_W / 2.0, ly + TOP_FACE + LEVER_T, "粗实线")

    d.line(lx, ly - R_FLANGE - 30.0, lx, ly + TOP_FACE + LEVER_T + 25.0, "中心线")
    d.line(lx - HALF_H - 30.0, ly, lx + HALF_H + 30.0, ly, "中心线")
    d.text("左视图", lx + 130.0, ly - R_FLANGE - 20.0, 8.0, "文字", align="center")

    # ==================== 尺寸标注 ====================
    def dim_rotated(p1, p2, loc, angle_deg, txt=None):
        def impl():
            dim = d.sp.AddDimRotated(P(p1[0], p1[1]), P(p2[0], p2[1]),
                                     P(loc[0], loc[1]), RAD(angle_deg))
            dim.Layer = "标注"
            if txt:
                dim.TextOverride = txt
            return dim
        return retry(impl, "dim rotated")

    def dim_dia(center, r, angle_deg, leader):
        a = RAD(angle_deg)
        p1 = (center[0] + r * math.cos(a), center[1] + r * math.sin(a))
        p2 = (center[0] - r * math.cos(a), center[1] - r * math.sin(a))

        def impl():
            dim = d.sp.AddDimDiametric(P(p1[0], p1[1]), P(p2[0], p2[1]), leader)
            dim.Layer = "标注"
            return dim
        return retry(impl, "dim dia")

    # ---- 主视图 ----
    dim_dia((MVX, MVY), R_FLANGE, 42.0, 50.0)                 # Ø460
    dim_dia((MVX, (neck_bot + neck_top_y) / 2.0), R_NECK, 180.0, 22.0)   # Ø110 向左
    dim_dia((MVX, (neck_top_y + top_top_y) / 2.0), R_TOP, 180.0, 20.0)   # Ø200 向左
    # 中心高 300：放在主视图右侧空白区（左侧离内框太近，文字会压线）
    dim_rotated((MVX + R_FLANGE + 30.0, MVY), (MVX + R_FLANGE + 30.0, MVY + TOP_FACE),
                (MVX + R_FLANGE + 55.0, MVY + TOP_FACE / 2.0), 90.0)

    # 引注放到主视图右下方的空白处，避开中心区多圈线与中心高标注
    a_bolt = RAD(-42.0)
    bx = MVX + R_BCD * math.cos(a_bolt)
    by = MVY + R_BCD * math.sin(a_bolt)
    tx, ty = MVX + R_FLANGE + 26.0, MVY - R_FLANGE + 62.0
    d.text("12-Ø26 均布于", tx, ty, 6.0, "文字")
    d.text("Ø410 分布圆", tx, ty - 9.0, 6.0, "文字")
    d.line(tx - 4.0, ty + 2.0, bx + 22.0, by + 6.0, "细实线")
    d.line(bx + 22.0, by + 6.0, bx + 4.0, by + 1.0, "细实线")

    # ---- 俯视图 ----
    dim_rotated((TVX - R_FLANGE, TVY), (TVX + R_FLANGE, TVY),
                (TVX, TVY - HALF_H - 58.0), 0.0, "460")
    dim_rotated((TVX - R_FLANGE, TVY - HALF_H), (TVX - R_FLANGE, TVY + HALF_H),
                (TVX - R_FLANGE - 38.0, TVY), 90.0, "78")
    dim_rotated((TVX, TVY + LEVER_W / 2.0), (TVX + LEVER_L, TVY + LEVER_W / 2.0),
                (TVX + LEVER_L / 2.0, TVY + LEVER_W / 2.0 + 42.0), 0.0, "260")
    dim_rotated((TVX + LEVER_L, TVY - LEVER_W / 2.0), (TVX + LEVER_L, TVY + LEVER_W / 2.0),
                (TVX + LEVER_L + 32.0, TVY), 90.0, "34")
    d.text("8-Ø18 均布于 Ø165", TVX + 150.0, TVY + 96.0, 6.0, "文字")
    d.line(TVX + 148.0, TVY + 93.0, TVX + 78.0, TVY + 60.0, "细实线")
    d.line(TVX + 78.0, TVY + 60.0, TVX + R_TOPBCD * math.cos(RAD(45.0)),
           TVY + R_TOPBCD * math.sin(RAD(45.0)), "细实线")

    # ---- 左视图 ----
    dim_rotated((lx - HALF_H, ly - R_FLANGE), (lx + HALF_H, ly - R_FLANGE),
                (lx, ly - R_FLANGE - 56.0), 0.0, "78")
    dim_rotated((lx - HALF_H, ly - R_FLANGE), (lx - HALF_H, ly + R_FLANGE),
                (lx - HALF_H - 38.0, ly), 90.0, "460")
    dim_rotated((lx - HALF_H, ly - R_FLANGE), (lx - HALF_H + FLANGE_T, ly - R_FLANGE),
                (lx - HALF_H + FLANGE_T / 2.0, ly - R_FLANGE - 26.0), 0.0, "22")
    # 中心高 300 在左视图右侧也标一处，方便对照
    dim_rotated((lx + HALF_H + 18.0, ly), (lx + HALF_H + 18.0, ly + TOP_FACE),
                (lx + HALF_H + 42.0, ly + TOP_FACE / 2.0), 90.0)

    # ==================== 技术要求（右侧下部） ====================
    notes = [
        "技术要求",
        "1. 公称通径 DN300，公称压力 PN16，对夹式连接。",
        "2. 结构长度 78，法兰 Ø460，12-Ø26 均布于 Ø410。",
        "3. 阀体 WCB，阀板 304，阀杆 2Cr13，密封 EPDM。",
        "4. 阀杆 Ø50，顶部 ISO 5211 F16 法兰 Ø200，8-Ø18。",
        "5. 介质温度 -10℃ ~ 120℃，适用水、油、气体。",
        "6. 未注铸造圆角 R3，未注公差按 GB/T 1804-m。",
        "7. 装配后按 1.5 倍公称压力做强度试验，",
        "   密封试验按 1.1 倍公称压力。",
    ]
    nx, ny = 880.0, 545.0
    for i, s in enumerate(notes):
        d.text(s, nx, ny - i * 20.0, 8.0 if i == 0 else 6.0, "文字")
    d.line(nx, ny + 13.0, nx + 290.0, ny + 13.0, "细实线")

    # ==================== 明细表（右侧上部） ====================
    rows = [
        ("1", "阀体", "WCB", "1", "铸件"),
        ("2", "阀板", "304", "1", "精铸"),
        ("3", "阀杆", "2Cr13", "1", "磨削"),
        ("4", "密封圈", "EPDM", "1", "模压"),
        ("5", "支架", "HT200", "1", "铸件"),
        ("6", "手柄", "Q235", "1", "焊接"),
    ]
    headers = ["序号", "名称", "材料", "数量", "备注"]
    col_x = [880.0, 920.0, 1010.0, 1090.0, 1130.0, 1180.0]
    top_y = 700.0
    hdr_h, row_h = 16.0, 15.0
    tbl_h = hdr_h + row_h * len(rows)

    d.text("零件明细表", 880.0, top_y + 8.0, 7.5, "文字")
    d.rect(col_x[0], top_y - tbl_h, col_x[-1], top_y, "表格")
    d.line(col_x[0], top_y - hdr_h, col_x[-1], top_y - hdr_h, "表格")
    for x in col_x[1:-1]:
        d.line(x, top_y, x, top_y - tbl_h, "表格")
    for i, h in enumerate(headers):
        d.text(h, (col_x[i] + col_x[i + 1]) / 2.0, top_y - hdr_h + 5.0, 6.0,
               "文字", align="center")
    for r, row in enumerate(rows):
        ry = top_y - hdr_h - (r + 1) * row_h
        d.line(col_x[0], ry, col_x[-1], ry, "表格")
        for c, val in enumerate(row):
            d.text(val, (col_x[c] + col_x[c + 1]) / 2.0, ry + 5.0, 6.0,
                   "文字", align="center")

    # ==================== 收尾 + 自检 ====================
    retry(lambda: doc.Regen(1))
    retry(lambda: doc.SendCommand("_.ZOOM\n_E\n"))
    retry(lambda: doc.Regen(1))

    total = int(retry(lambda: d.sp.Count))
    counts, layers = {}, {}
    out_of_frame = []
    for i in range(total):
        ent = retry(lambda: d.sp.Item(i))
        oname = str(retry(lambda: ent.ObjectName))
        counts[oname] = counts.get(oname, 0) + 1
        layers[str(retry(lambda: ent.Layer))] = layers.get(str(retry(lambda: ent.Layer)), 0) + 1
        try:
            lo, hi = retry(lambda: ent.GetBoundingBox())
        except Exception:
            continue
        tol = 1.0
        if (lo[0] < IN_L - tol or lo[1] < IN_B - tol or
                hi[0] > IN_R + tol or hi[1] > IN_T + tol):
            out_of_frame.append((i, oname, str(retry(lambda: ent.Layer)),
                                 round(lo[0], 1), round(lo[1], 1),
                                 round(hi[0], 1), round(hi[1], 1)))

    print("图元总数:", total)
    print("按类型:", counts)
    print("按图层:", layers)
    print("越过内框的图元数:", len(out_of_frame))
    for row in out_of_frame[:30]:
        print("  idx=%d %s layer=%s bbox=(%.1f,%.1f)-(%.1f,%.1f)" % row)

    retry(lambda: doc.SaveAs(OUT))
    print("saved ->", OUT)


if __name__ == "__main__":
    main()
