# Excel 渲染技术笔记（已探明/已验证，勿重新推导）

目标环境 **Excel 2016**。不要使用更新版本才有的 API（例如 DataLabel 没有任意角度 Rotation）。

## 为什么是散点 + 面积复合图

- 原生 histogram 图表类型不能与其他系列复合。
- area/column 系列绑死 category 轴，无法按连续 X 值定位。
- 原生 `xlLine` 系列不能携带任意 XValues（`Series.XValues` 仅 scatter/bubble 可用）。
- 所以：散点画轮廓/曲线/竖线，面积画填充，×10000 归一化映射叠加。方案是被 Excel 逼出来的，不要试图"简化"掉。

## 已探明的轴拓扑（2026-11-10 经批准的 COM 探针，Excel 服务端状态）

1. **加入 xlArea 系列会把主 category 轴降格为 ordinal**：`CategoryType` 报 `xlAutomaticScale = -4105`，此后写 `MinimumScale/MaximumScale/MajorUnit/MinimumScaleIsAuto` 抛 E_FAIL `80004005`。
   **解药 = smooth 开关**：把第一个散点系列 `ChartType → xlXYScatterSmoothNoMarkers → 改回原类型`，即 `ChartCanvas.RepairPrimaryAxis`。**load-bearing，绝不能删**（早期"轴看着正常就删掉"的建议已撤回）。之后隐藏次轴不会再次降格。
2. **面积系列是第一个加入的系列时，`Series.AxisGroup = xlSecondary` 被静默忽略**——主组系列存在之前次组不存在。先设 AxisGroup 再设 ChartType 也没用。Builder 的 z-order 恰好把 histoShade 放第一，所以真实管线必中。
   **解法 = `Painters.PrepareAreaAxis`：所有系列加完后，对 area 系列重断言 `AxisGroup = xlSecondary`**，随后 RepairPrimaryAxis 才能让每次轴写入成功。
3. **复合图的 `Chart.ChartType` 返回 -4111**（XlChartType 里无名字的值）——**永远不要对它分支**；`Series.ChartType` 保持正确（75/1）。
4. **次 category 轴必须 `CategoryType = xlTimeScale`**：默认 category 轴按序号标签定位，time-scale 轴按标签数值定位——这是 ×10000 映射能落对的机制（源自用户旧 Macros.bas）。两处加固：显式钉死次轴 bounds 1..10001（不信任日期轴自动缩放）；**次 value 轴强制同步主轴范围**（主轴还要容纳正态曲线峰，次轴独立自动缩放会让填充比轮廓矮）。

## ×10000 映射

- `ToScaledX = AREA_BASE + Int((X - XMin) / span * AREA_SCALE + 0.5)`，AREA_BASE=1，AREA_SCALE=10000。
- **Spec 层永远存原始数据坐标**；缩放只发生在 Painters 渲染时。
- `ChartCanvas.FromChartObject` 要求两轴显式 Min/Max（拒绝自动缩放，ERR_AUTO_SCALE）——映射需要确定边界。

## ApplyAxes 顺序（顺序敏感，勿调换）

```
PrepareAreaAxis → RepairPrimaryAxis → ApplyPrimaryCategory → ApplyPrimaryValue → ConfigureAreaAxis
```

## 轴标签接管（todo4 机制）

- 原生刻度标签用 `";;;"` 自定义数字格式隐藏（**仍占位**，标签带高度保留）。
- axisLabels = 隐藏散点系列：点在 (x, 0)，`MarkerStyle = xlMarkerStyleNone`，`Format.Line.Visible = msoFalse`；逐点 `HasDataLabel = True` + `DataLabel.Position = xlLabelPositionBelow` + `.Text` + 逐点 `Font.Color`（与对应参考线同色）。
- **未验证风险**：线+标记都隐藏时 DataLabel 是否可见——需用户侧渲染确认（回归清单第一项）。
- 2026-11-13 起：**每条 bin 边界都打标签**（stride/MAX_AXIS_LABELS 已删）。边界 AddLabel 传 `eps = 0#`：BinCount > 50 时 bin 宽 < 2% span，否则相邻边界会被去重互相吞掉。辅助标签（LSL/USL/Target/mean/sigma）保留 eps = 2% span 去重；**边界在前、辅助在后** → 辅助值落在边界 eps 内时彩色标签被丢弃（数值几乎相同，可接受，属设计）。
- 重叠缓解（按性价比）：减 Decimals、加宽图表、减 bin 数。**不要旋转标签**：2016 DataLabel 无任意 Rotation（Orientation 只有 4 个固定方向），标签带只预留一行横排高度，竖排会裁切。

## 刻度对齐 bin 网格（todo1）

- xLo/xHi 向外吸附到 `EdgeLow + k*BinWidth`：floor 用 `Int`，ceiling 用 `-Int(-x)`。
- `MajorUnit = BinWidth` → 刻度 Min + k*unit 恰好落在 bin 边界。
- `X_PAD_RATIO = 0.02` 现在**只用于统计框内缩**，不再外扩轴范围。

## 纵轴固定贴左（2026-11-13）

- `ApplyPrimaryCategory` 内 `.Crosses = xlAxisCrossesMinimum` → 纵轴永远交于横轴最小值处（图形左缘），负 X 数据不会把纵轴挤到中间。
- `Axis.Crosses` 语义：**另一条轴**穿过**本轴**的位置。对横轴设 Minimum 等价于 UI"纵轴交叉于最小分类"。
- 用枚举而非 `CrossesAt`（category 轴上 CrossesAt 按分类序号解释，有歧义）。
- 纵轴 min 恒为 0，横轴天然在底部，无需对值轴做对称设置。

## 注释（统计框）定位

- `DataToPixelX/Y`：`PlotArea.InsideLeft/Top/Width/Height` 线性映射。
- `AddAnnotation`：AddTextbox + TextFrame2（WordWrap=False、AutoSize=ShapeToFitText、四边 Margin=0）；**DoEvents 强制布局后再读 Width**（否则右缘偏移算错）；(X,Y) 固定为**右上角锚点**：`shp.Left = px - shp.Width`，`shp.Top = py`。
- 多行文本：Spec 用 vbCrLf，写入 Shape 前替换为 vbLf（Excel 文本框只认 vbLf）。
- 2026-11-13 已删除全部水平对齐管线（TextAlignEnum / ToMsoAnchor / ToMsoParagraphAlign / .HorizontalAnchor / .ParagraphFormat.Alignment）。统计框段落现为默认左对齐（外观微变，Minitab 统计框本就左对齐），框位置不变。`msoAnchorLeft/msoAnchorRight` **不存在**，不要再引用。

## 其他已验证事实 / 教训

- `SetElement` 的 msoElement* 用命名常量（如 msoElementPrimaryValueGridLinesNone）；历史教训：不要猜枚举数值。
- 面积系列填充透明度、次轴隐藏等均在 Painters.ConfigureAreaAxis / PaintAreaFill 内，勿散落。
- Trust Center 阻止程序化访问 VBProject；不要尝试注入代码验证。
- 探针脚本（axes_probe*.ps1）与 GBK 编辑脚本在旧机器 `Temp/alispc_gen/`，**不随项目迁移**；结论已全部固化在本目录。