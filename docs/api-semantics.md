# AutoCAD COM 语义：和官方文档不一致的地方

这份文档里的每一条都是**实测出来的**，不是从文档抄的。有些和 Autodesk 文档明确矛盾，按文档写会得到静默的错误结果。

写这些脚本时如果只信文档，会浪费大量时间在「为什么对象跑到了奇怪的位置」「为什么实体消失了」上面。

---

## 参数必须显式包成 SAFEARRAY

**症状**：晚绑定 `win32com.client.Dispatch` 下，所有绘图方法全失败：

```
(-2147352567, '发生意外。', (0, None, None, None, 0, -2147024809), None)
```

内层 hresult `-2147024809` = `E_INVALIDARG 0x80070057`。

**最有误导性的地方**：同一时刻 `Layers.Add` 和 `ZoomExtents` 却**成功**。这极易让人误判成「COM 权限问题」或「版本不匹配」。

**根因**：这些方法的坐标参数是 `SAFEARRAY(double)`。晚绑定下把 Python 的 `list` / `tuple` / `array.array('d')` 直接传进去，会触发 `E_INVALIDARG`。

```python
def dbl_array(values):
    return VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [float(v) for v in values])
```

同时改用**早绑定** `gencache.EnsureDispatch(prog_id, bForDemand=False)`（首次会生成 `gen_py` 缓存，1–3 秒；之后很快）。

实测五种传参方式，只有 `VARIANT(VT_ARRAY|VT_R8)` 和「gencache + VARIANT」能成功。

---

## `AddCylinder` / `AddBox` 的 center 是几何中心

不是底面圆心，不是角点。

```python
AddCylinder(P(0,0,0), 22.5, 68.0)
# → bbox [-22.5,-22.5,-34.0] .. [22.5,22.5,34.0]   沿 Z 对称居中

AddBox(P(0,0,0), 10, 20, 30)
# → bbox [-5,-10,-15] .. [5,10,15]
```

把 center 当底心写，所有实体都会平移半个尺寸。这个错误不会报错，只会让你看着模型发愣。

---

## `Rotate3D` 的角度单位是**弧度**

```python
Rotate3D(P(0,0,0), P(1,0,0), 90.0)      # 错！转了 90 弧度
# → bbox [-22.5,-40.478,-35.349] .. [22.5,40.478,35.349]   躺倒的怪姿势

Rotate3D(P(0,0,0), P(1,0,0), math.pi/2) # 对
# → bbox [-22.5,-34.0,-22.5] .. [22.5,34.0,22.5]           正确的 Y 轴圆柱
```

同一个坑还有一处：`AddDimRotated(p1, p2, linePoint, rotationAngle)` 的角度**也是弧度**。传 `90.0` 会让竖排尺寸全错，`Measurement` 读出 50.064 / 71.52 这类垃圾值。

> 记住一句话：这个 API 里凡是「角度」，先怀疑是弧度。

---

## `Boolean` 的 operation 编码与文档相反

| 值 | 实际含义 |
|---|---|
| `0` | **并集** |
| `1` | **交集** |
| `2` | **差集** |
| `3` / `4` | 报错「参数 Operation 无效」 |

实测方式：圆柱（Z −10..10）与盒子（Z 5..25）逐个试：

- `op 0 → bbox [-10,-10,-10]..[10,10,25]` 并集
- `op 1 → bbox [-10,-10,5]..[10,10,10]` 交集
- `op 2 → bbox [-10,-10,-10]..[10,10,10]` 差集

按文档用 `1=并集` 会让整个模型被挖空。

---

## 求交集为空时，双方一起被删

```python
cp = body.Copy()
sph = space.AddSphere(center, r)
cp.Boolean(1, sph)
# 若交集为空 → cp 和 sph 都被 AutoCAD 删除，Count 从 3 变 2
```

**推论**：探针绝不能直接打主体，必须 `body.Copy()` 一份副本去挨刀。副本死了不影响主体。

**另一个推论**：判断对象是否被删**不能读 `Handle`**——被删掉的 COM 对象读 `Handle` 有时仍不报错，会让探测结果全是「还活着」。可靠做法是立刻读 `GetBoundingBox()`，删掉的对象必抛 `com_error`。

我们用这个原理做了 17 点球探针校验（`verify3d.py`）：在材料处放小球求交 → 副本应存活；在孔洞处放小球求交 → 副本应消失。

---

## `MassProperties` 取不到

```python
solid.MassProperties          # AttributeError: 'IAcad3DSolid' object has no attribute
```

早期绑定和 `dynamic.Dispatch` 都失败。所以**体积校验走不通**，这不是我的用法问题，是接口本身在这个版本上不可用。用球探针替代。

---

## `Documents.Add()` 不是幂等的

CAD 忙的时候它会**报错，但文档已经建好了**。外面套一个「失败就重试」的循环，就会堆出十几个空文档（我们曾经误建 17 个同名文档）。

正确做法：调用前记下 `Documents.Count`，失败后先核对 Count 是否已经增加，再决定要不要重试。

顺带：清理这类残留要逐个 `Documents.Item(0).Close(False)`，`Close(False)` 的 `False` 表示不保存。

---

## `SaveAs` 会因锁文件失败

```
com_error ... '保存文档时出错' ... -2145320861
```

两个原因：

1. 目标路径已被另一个 AutoCAD 文档实例占用
2. 残留 `.dwl` / `.dwl2` 锁文件

标准处理：

```python
close_docs_named(app, wanted)   # 先关掉同名文档
clear_stale_files(path)         # 再删 .dwl / .dwl2
```

---

## 文字样式：两步才能设属性

```python
st = doc.TextStyles.Add("DSH-CN")
st.fontFile = "gbenor.shx"      # Property '<unknown>.fontFile' can not be set.
```

`Add()` 的返回值设不了属性。必须重新取一次：

```python
doc.TextStyles.Add("DSH-CN")
st = doc.TextStyles.Item("DSH-CN")
st.fontFile = "gbenor.shx"      # 可以了
```

另外 `bigFontFile` **不在**早期绑定的 `IAcadTextStyle` 属性表里，只能走 `dynamic.Dispatch(st).bigFontFile`。

---

## 中文只能用 SHX 大字体

这是「三视图满屏 `???`」的根因。

- 默认 `Standard` 样式的 `fontFile=txt`（SHX 单线字体，**不含中文**）
- AutoCAD 2024 的 `Fonts\` 目录里**只有 .shx，没有任何 .ttf / .ttc**
- 所以 `fontFile = "SimSun"` / `"宋体"` / `"simsun.ttc"` 全部报「文件处理器错误」

唯一可行的组合：

```python
st = doc.TextStyles.Item("DSH-CN")
st.fontFile = "gbenor.shx"       # 西文
st.bigFontFile = "gbcbig.shx"    # 中文大字体
doc.ActiveTextStyle = st
```

实测中文、`Ø`、`℃` 全部正常显示。

---

## 尺寸变量必须在画标注之前设

已有的标注对象**不会**跟随后来改的 `DIM*` 变量。所以：

```python
doc.SetVariable("DIMTXT", 7.0)
doc.SetVariable("DIMTXSTY", "DSH-CN")
doc.SetVariable("DIMLFAC", 1.0 / SCALE)   # 2:1 画图时让标注读出真实尺寸
# ... 全部设完，再开始 AddDimRotated
```

另外 `DIMBLK = "."` / `DIMLDRBLK = "."` 表示「用默认箭头」，不设会画出奇怪的自定义块。

---

## `SetVariable` 对某些变量不可用

```python
doc.SetVariable("VSCURRENT", "概念")   # com_error ... 设置系统变量时出错 ... -2145320858
doc.GetVariable("VSCURRENT")           # com_error ... 获取系统变量时出错 ... -2145320859
```

视觉样式这类变量只能走命令行：

```python
doc.SendCommand("_.VSCURRENT\n_Conceptual\n")
doc.SendCommand("_.GRIDMODE\n0\n")
doc.SendCommand("_.UCSICON\n_OFF\n")
doc.SendCommand("-_.VPOINT\n1,-1.6,0.9\n")
doc.SendCommand("_.ZOOM\n_E\n")
```

`VPOINT` 走 `SetVariable` 时参数必须包成 `VARIANT(VT_ARRAY|VT_R8)`。

---

## 读文字要用晚绑定

```python
ent.TextString   # AttributeError: '<...IAcadEntity...>' object has no attribute 'TextString'
```

`IAcadEntity` 基类没有 `TextString`。读之前先降级：

```python
dynamic.Dispatch(ent).TextString
```

---

## 控制台编码

Windows 控制台是 GBK，Python 打印 `▶`、`Ø` 这类字符会 `UnicodeEncodeError`。脚本开头：

```python
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")
```

并设 `PYTHONUTF8=1` / `PYTHONIOENCODING=utf-8`（插件已经替子进程设好了）。

---

## 补充：`PostCommand` 是异步的

`cad_send_command` 返回 `sent: true` 只代表**已发出**，不代表**已执行完**。发完命令后立刻读图元列表，可能读不到刚创建的东西。需要确定性时，用具体的 `Add*` 方法而不是 `SendCommand`。
