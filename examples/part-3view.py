# -*- coding: utf-8 -*-
"""参考零件图（轴测图）→ 三视图（第一角法）+ 尺寸标注

零件判读（用户已确认）：
  总体 44 长 × 28 宽 × 40 高
  底板    X 0..44, Y 0..28, Z 0..12          高 12
  左立块  X 0..24, Y 0..28, Z 12..40         宽 24
  右凸台  X 32..44, Y 0..28, Z 12..40        宽 12
  中间空档 X 24..32（宽 8），从 Z=12 通到顶
  顶面槽  X 12..22, Z 32..40                 宽 10 深 8
  左端斜面 (0,28) → (12,40)，水平投影 12
  底部缺口 X 24..44, Z 0..6                  宽 20 高 6

图纸：A3 420x297，比例 2:1（几何按 2 倍画，DIMLFAC=0.5 让标注读出真实尺寸）
字体：DSH-CN = gbenor.shx + gbcbig.shx（AutoCAD 2024 Fonts 目录只有 shx，无 ttf）
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
OUT = os.path.join(OUT_DIR, "part-3view.dwg")

CN_STYLE = "DSH-CN"

# ==================== 零件尺寸（1x 单位，mm） ====================
L = 44.0        # 总长 X
W = 28.0        # 总宽 Y
H = 40.0        # 总高 Z
BASE_H = 12.0   # 底板高
TOWER_X = 24.0  # 左立块右端
BOSS_X = 32.0   # 右凸台左端
GROOVE_X0, GROOVE_X1 = 12.0, 22.0   # 顶面槽
GROOVE_Z0 = 32.0                    # 槽底
SLANT_Z = 28.0                      # 左端面高
GAP_X0, GAP_X1 = TOWER_X, BOSS_X    # 24..32
NOTCH_Z = 6.0                       # 底部缺口高

SC = 2.0        # 比例 2:1
DIMLFAC = 1.0 / SC

# ==================== 图幅（A3 420x297） ====================
FW, FH = 420.0, 297.0
M_LEFT, M_OTHER = 25.0, 10.0
IN_L, IN_R = M_LEFT, FW - M_OTHER        # 25 .. 410
IN_B, IN_T = M_OTHER, FH - M_OTHER       # 10 .. 287

FX, FZ = 56.0, 152.0     # 主视图原点（零件 X=0 / Z=0）
TX, TY = 56.0, 62.0      # 俯视图原点（零件 X=0 / Y=0）
LX, LZ = 212.0, 152.0    # 左视图原点（零件 Y=0 / Z=0）


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


# 三套坐标映射：零件坐标 → 图纸坐标
def fp(x, z):
    """主视图 X-Z"""
    return (FX + x * SC, FZ + z * SC)


def tp(x, y):
    """俯视图 X-Y"""
    return (TX + x * SC, TY + y * SC)


def lp(y, z):
    """左视图 Y-Z"""
    return (LX + y * SC, LZ + z * SC)


def clear_stale_files(path):
    for candidate in (path, path + ".dwl", path + ".dwl2", path + "l", path + "l2"):
        try:
            if os.path.exists(candidate):
                os.remove(candidate)
                print("删除旧文件:", candidate)
        except Exception as exc:
            print("  删除失败", candidate, type(exc).__name__, str(exc)[:60])


def close_docs_named(app, wanted):
    wanted = os.path.normcase(os.path.abspath(wanted))
    for _ in range(20):
        hit = None
        for i in range(int(retry(lambda: app.Documents.Count))):
            doc = retry(lambda: app.Documents.Item(i))
            try:
                full = str(retry(lambda: doc.FullName))
            except Exception:
                full = ""
            if full and full != "None" and os.path.normcase(os.path.abspath(full)) == wanted:
                hit = doc
                break
            if not full or full == "None":
                name = str(retry(lambda: doc.Name))
                if os.path.normcase(name) == os.path.normcase(os.path.basename(wanted)):
                    hit = doc
                    break
        if hit is None:
            return
        print("关闭同名文档:", str(retry(lambda: hit.Name)))
        retry(lambda: hit.Close(False))


def new_document(app):
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

    def ensure_style(self, name, font="gbenor.shx", big="gbcbig.shx"):
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

    def line(self, p1, p2, layer):
        e = retry(lambda: self.sp.AddLine(P(p1[0], p1[1]), P(p2[0], p2[1])))
        retry(lambda: setattr(e, "Layer", layer))
        return e

    def poly(self, pts, layer, closed=True):
        seq = list(pts) + ([pts[0]] if closed else [])
        for a, b in zip(seq, seq[1:]):
            self.line(a, b, layer)

    def rect(self, x0, y0, x1, y1, layer):
        self.poly([(x0, y0), (x1, y0), (x1, y1), (x0, y1)], layer)

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

    def dim(self, p1, p2, loc, angle, txt=None):
        def impl():
            d = self.sp.AddDimRotated(P(p1[0], p1[1]), P(p2[0], p2[1]),
                                      P(loc[0], loc[1]), float(angle))
            d.Layer = "标注"
            if txt:
                d.TextOverride = txt
            return d

        return retry(impl, "dim")


def main():
    pythoncom.CoInitialize()
    app = retry(lambda: gencache.EnsureDispatch("AutoCAD.Application.24", bForDemand=False))
    retry(lambda: setattr(app, "Visible", True))

    close_docs_named(app, OUT)
    clear_stale_files(OUT)
    doc = new_document(app)

    d = Drawing(doc)
    d.ensure_style(CN_STYLE)

    for var, val in (("INSUNITS", 4), ("LWDISPLAY", 1), ("LUNITS", 2), ("DIMASSOC", 2),
                     ("MEASUREMENT", 1), ("ORTHOMODE", 0), ("LTSCALE", 2.0),
                     ("FILLMODE", 1), ("MIRRTEXT", 0)):
        try:
            retry(lambda: doc.SetVariable(var, val))
        except Exception:
            pass

    # 尺寸样式必须在画标注之前设好
    for var, val in (("DIMTXT", 3.5), ("DIMASZ", 2.5), ("DIMEXE", 1.5), ("DIMEXO", 1.0),
                     ("DIMGAP", 1.0), ("DIMDEC", 0), ("DIMZIN", 0), ("DIMSCALE", 1.0),
                     ("DIMTAD", 1), ("DIMTIX", 0), ("DIMSOXD", 0), ("DIMLFAC", DIMLFAC),
                     ("DIMCLRD", 7), ("DIMCLRE", 7), ("DIMCLRT", 7),
                     ("DIMTXSTY", CN_STYLE), ("DIMBLK", "."), ("DIMLDRBLK", ".")):
        try:
            retry(lambda: doc.SetVariable(var, val))
        except Exception as exc:
            print("  变量设置失败", var, type(exc).__name__)

    d.layer("图框", 7, 50)
    d.layer("粗实线", 7, 50)
    d.layer("细实线", 7, 18)
    d.layer("虚线", 7, 25, "DASHED")
    d.layer("标注", 7, 18)
    d.layer("文字", 7, 25)
    d.layer("表格", 7, 35)

    # ==================== 图框 ====================
    d.rect(0.0, 0.0, FW, FH, "图框")
    d.rect(IN_L, IN_B, IN_R, IN_T, "图框")

    # ==================== 标题栏（右下 140x35） ====================
    TB_W, TB_H = 140.0, 35.0
    tbx, tby = IN_R - TB_W, IN_B
    d.rect(tbx, tby, IN_R, tby + TB_H, "表格")
    d.line((tbx, tby + 12.0), (IN_R, tby + 12.0), "表格")
    d.line((tbx, tby + 23.0), (IN_R, tby + 23.0), "表格")
    d.line((tbx + 34.0, tby), (tbx + 34.0, tby + TB_H), "表格")
    d.line((tbx + 92.0, tby), (tbx + 92.0, tby + TB_H), "表格")
    d.text("支座", tbx + 40.0, tby + 24.0, 6.0, "文字")
    d.text("三视图", tbx + 40.0, tby + 13.0, 4.5, "文字")
    d.text("比例 2:1", tbx + 40.0, tby + 3.0, 4.5, "文字")
    d.text("第一角法", tbx + 96.0, tby + 24.0, 4.5, "文字")
    d.text("件号 1", tbx + 96.0, tby + 13.0, 4.5, "文字")
    d.text("单位 mm", tbx + 96.0, tby + 3.0, 4.5, "文字")

    # ==================== 主视图（X-Z） ====================
    profile = [(0, 0), (24, 0), (24, 6), (44, 6), (44, 40), (32, 40), (32, 12),
               (24, 12), (24, 40), (22, 40), (22, 32), (12, 32), (12, 40), (0, 28)]
    d.poly([fp(x, z) for x, z in profile], "粗实线", closed=True)

    # ==================== 俯视图（X-Y） ====================
    d.poly([tp(0.0, 0.0), tp(L, 0.0), tp(L, W), tp(0.0, W)], "粗实线", closed=True)
    for xx in (GROOVE_X0, GROOVE_X1, TOWER_X, BOSS_X):
        d.line(tp(xx, 0.0), tp(xx, W), "粗实线")

    # ==================== 左视图（Y-Z） ====================
    d.poly([lp(0.0, 0.0), lp(W, 0.0), lp(W, H), lp(0.0, H)], "粗实线", closed=True)
    d.line(lp(0.0, SLANT_Z), lp(W, SLANT_Z), "粗实线")     # 左端面顶边（可见）
    d.line(lp(0.0, BASE_H), lp(W, BASE_H), "虚线")         # 中间空档底面（不可见）
    d.line(lp(0.0, GROOVE_Z0), lp(W, GROOVE_Z0), "虚线")   # 顶面槽底（不可见）
    d.line(lp(0.0, NOTCH_Z), lp(W, NOTCH_Z), "虚线")       # 底部缺口顶面（不可见）

    # ==================== 尺寸标注 ====================
    # --- 主视图：水平（X 方向）尺寸，全部放在视图上方两排 ---
    y_row1 = FZ + H * SC + 12.0
    y_row2 = FZ + H * SC + 30.0
    d.dim(fp(0.0, SLANT_Z), fp(GROOVE_X0, H), (fp(0.0, 0)[0], y_row1), 0.0)   # 12 斜面投影
    d.dim(fp(GROOVE_X0, H), fp(GROOVE_X1, H), (fp(0.0, 0)[0], y_row1), 0.0)   # 10 槽宽
    d.dim(fp(GAP_X0, H), fp(GAP_X1, H), (fp(0.0, 0)[0], y_row1), 0.0)         # 8 空档宽
    d.dim(fp(BOSS_X, H), fp(L, H), (fp(0.0, 0)[0], y_row1), 0.0)              # 12 凸台宽
    d.dim(fp(0.0, SLANT_Z), fp(TOWER_X, H), (fp(0.0, 0)[0], y_row2), 0.0)     # 24 立块宽

    # --- 主视图：竖直（Z 方向）尺寸，左侧 28 / 右侧 40 与 6 ---
    d.dim(fp(0.0, 0.0), fp(0.0, SLANT_Z), (FX - 8.0, fp(0.0, 0.0)[1]), math.pi / 2.0)  # 28 左端面高
    d.dim(fp(L, 0.0), fp(L, H), (fp(L, 0.0)[0] + 16.0, fp(0.0, 0.0)[1]), math.pi / 2.0)  # 40 总高
    d.dim(fp(L, 0.0), fp(L, NOTCH_Z), (fp(L, 0.0)[0] + 6.0, fp(0.0, 0.0)[1]), math.pi / 2.0)  # 6 缺口高

    # --- 俯视图：总长 44、缺口宽 20、总宽 28 ---
    d.dim(tp(0.0, 0.0), tp(L, 0.0), (tp(0.0, 0.0)[0], TY - 26.0), 0.0)   # 44 总长
    d.dim(tp(GAP_X0, 0.0), tp(L, 0.0), (tp(0.0, 0.0)[0], TY - 14.0), 0.0)  # 20 缺口宽
    d.dim(tp(L, 0.0), tp(L, W), (tp(L, 0.0)[0] + 14.0, TY), math.pi / 2.0)          # 28 总宽

    # --- 左视图：槽深 8、底板高 12 ---
    d.dim(lp(0.0, GROOVE_Z0), lp(0.0, H), (LX - 14.0, LZ), math.pi / 2.0)   # 8 槽深
    d.dim(lp(W, 0.0), lp(W, BASE_H), (lp(W, 0.0)[0] + 14.0, LZ), math.pi / 2.0)  # 12 底板高

    # ==================== 技术要求 ====================
    note_x, note_y = 292.0, 250.0
    d.text("技术要求", note_x, note_y, 5.0, "文字")
    notes = [
        "1. 未注圆角 R1~R2。",
        "2. 未注公差按 GB/T 1804-m。",
        "3. 去除毛刺、飞边。",
        "4. 材料：HT200。",
        "5. 未注尺寸见三维模型。",
        "6. 三视图按第一角法布置，",
        "   比例 2:1，单位 mm。",
    ]
    for i, line_text in enumerate(notes):
        d.text(line_text, note_x, note_y - 11.0 * (i + 1), 3.5, "文字")

    # ==================== 收尾 ====================
    retry(lambda: doc.Regen(1))
    try:
        retry(lambda: doc.SendCommand("_.ZOOM\n_E\n"))
    except Exception as exc:
        print("  ZOOM 失败", type(exc).__name__)

    # ---------- 自检 ----------
    print("\n===== 自检 =====")
    kinds, layers = {}, {}
    bad = []
    for i in range(int(retry(lambda: d.sp.Count))):
        ent = retry(lambda: d.sp.Item(i))
        try:
            kind = str(retry(lambda: ent.EntityName))
        except Exception:
            kind = "?"
        kinds[kind] = kinds.get(kind, 0) + 1
        lay = str(retry(lambda: ent.Layer))
        layers[lay] = layers.get(lay, 0) + 1
        try:
            lo, hi = retry(lambda: ent.GetBoundingBox())
            x0, y0, x1, y1 = min(lo[0], hi[0]), min(lo[1], hi[1]), max(lo[0], hi[0]), max(lo[1], hi[1])
            tol = 1.0
            if (x0 < IN_L - tol or x1 > IN_R + tol or
                    y0 < IN_B - tol or y1 > IN_T + tol):
                bad.append((kind, lay, round(x0, 1), round(y0, 1), round(x1, 1), round(y1, 1)))
        except Exception:
            pass
    print("图元总数:", sum(kinds.values()))
    print("类型:", kinds)
    print("图层:", layers)
    print("越过内框的图元数:", len(bad))
    for row in bad[:12]:
        print("  越界:", row)

    # 文字互相重叠检查
    texts = []
    for i in range(int(retry(lambda: d.sp.Count))):
        ent = retry(lambda: d.sp.Item(i))
        try:
            if not str(retry(lambda: ent.EntityName)).endswith("Text"):
                continue
            s = str(dynamic.Dispatch(ent).TextString)
            lo, hi = retry(lambda: ent.GetBoundingBox())
            texts.append((s, min(lo[0], hi[0]), min(lo[1], hi[1]),
                          max(lo[0], hi[0]), max(lo[1], hi[1])))
        except Exception:
            pass
    over = 0
    for i in range(len(texts)):
        for j in range(i + 1, len(texts)):
            a, b = texts[i], texts[j]
            w = min(a[3], b[3]) - max(a[1], b[1])
            h = min(a[4], b[4]) - max(a[2], b[2])
            if w > 0 and h > 0:
                area = w * h
                amin = (a[3] - a[1]) * (a[4] - a[2])
                bmin = (b[3] - b[1]) * (b[4] - b[2])
                if area > 0.4 * min(amin, bmin):
                    over += 1
                    print("  文字重叠:", a[0][:16], "|", b[0][:16])
    print("文字重叠对数:", over)

    # 标注图元 bbox 与文字 bbox 的碰撞检查
    ann, txt = [], []
    for i in range(int(retry(lambda: d.sp.Count))):
        ent = retry(lambda: d.sp.Item(i))
        lay = str(retry(lambda: ent.Layer))
        if lay not in ("标注", "文字"):
            continue
        try:
            lo, hi = retry(lambda: ent.GetBoundingBox())
            box = (min(lo[0], hi[0]), min(lo[1], hi[1]), max(lo[0], hi[0]), max(lo[1], hi[1]))
        except Exception:
            continue
        if lay == "标注":
            ann.append(box)
        else:
            try:
                label = str(dynamic.Dispatch(ent).TextString)[:14]
            except Exception:
                label = "?"
            txt.append((label, box))
    clashes = 0
    for box in ann:
        for label, tb in txt:
            w = min(box[2], tb[2]) - max(box[0], tb[0])
            h = min(box[3], tb[3]) - max(box[1], tb[1])
            if w > 0 and h > 0:
                area = w * h
                tarea = (tb[2] - tb[0]) * (tb[3] - tb[1])
                if area > 0.35 * tarea:
                    clashes += 1
                    print("  标注压文字:", label, round(area / tarea, 2))
    print("标注压文字对数:", clashes)

    retry(lambda: doc.SaveAs(OUT))
    print("\nsaved ->", OUT)


if __name__ == "__main__":
    main()
