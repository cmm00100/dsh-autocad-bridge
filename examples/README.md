# 示例脚本

三个能直接跑出图纸的脚本。它们不走 MCP 工具，而是**直接调 COM**——这样你能看清每一步到底做了什么，也方便改。

跑之前请确认：

1. AutoCAD 2024 已经开着，并且**没有未保存的改动**（脚本会往当前活动文档里画）
2. 装了 `pywin32`：`<你的 python> -m pip install pywin32`
3. 当前没有别的命令在 CAD 里执行（键盘上按一下 ESC）

产出的 `.dwg` 会写到脚本旁边的 `out/` 目录，目录不存在会自动建。

---

## `part-3view.py` — 零件三视图

一个 44 × 28 × 40 的方块类零件，带斜面、顶面槽、底部缺口。A3 图幅，比例 2:1。

```powershell
python examples\part-3view.py
```

产出 `out/part-3view.dwg`。

覆盖的技术点：

- 第一角法三视图对位（主视图 X-Z、俯视图 X-Y 长对正、左视图 Y-Z 高平齐）
- `DIMLFAC = 0.5` 让 2 倍画出的几何标注出真实尺寸
- 真标注对象（`AddDimRotated`），不是画线冒充
- 中文文字用 `DSH-CN` 样式

---

## `valve-3view.py` — DN300 蝶阀三视图

同样的零件，A0 图幅，1:1。含明细表、技术要求、标题栏。

```powershell
python examples\valve-3view.py
```

产出 `out/valve-dn300-drawing.dwg`。

覆盖的技术点：

- 三视图分区布局（先算死各视图包围盒，避免重叠）
- 直径标注（`AddDimDiametric`）
- 引注文字 + 引线（`12-Ø26 均布于 Ø410` 这类）
- 图层与线宽分层：粗实线 0.50 / 细线 0.18 / 中心线红色 CENTER
- 越界自检（遍历图元 `GetBoundingBox()` 找跑出内框的）

---

## `valve-3d.py` — DN300 蝶阀 3D 装配

把同一个阀门造成 6 个独立实体、6 个图层 6 种颜色。

```powershell
python examples\valve-3d.py
```

产出 `out/valve-dn300-3d.dwg`。

覆盖的技术点：

- `AddCylinder` / `AddBox` 的 **center 是几何中心**（不是底面圆心）
- `Rotate3D` 的角度单位是**弧度**（`math.pi/2`，不是 `90.0`）
- `Boolean` 的编码：**0=并集 / 1=交集 / 2=差集**
- 差集挖孔：流道、螺栓孔、阀杆孔
- 6 个部件分图层分色，便于单独选取

---

## 这几个脚本共用的一套「生存技巧」

看脚本时留意这几处，能省很多事：

```python
# 1. 坐标参数必须包成 SAFEARRAY，否则 E_INVALIDARG
def P(x, y, z=0.0):
    return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [float(x), float(y), float(z)])

# 2. CAD 忙的时候重试，而不是直接失败
BUSY = {-2147418111, -2147417846}

# 3. 中文文字必须用 shx 大字体组合
style = doc.TextStyles.Item("DSH-CN")
style.fontFile = "gbenor.shx"
style.bigFontFile = "gbcbig.shx"

# 4. 尺寸变量要在画标注之前设好，否则已有标注不跟随
doc.SetVariable("DIMTXT", 7.0)
```

更多同类问题见 [../docs/api-semantics.md](../docs/api-semantics.md)。
