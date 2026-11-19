# 架构与编码风格

## 分层契约（@Folder 标注）

| 层 | 模块 | 契约 |
|---|---|---|
| src.Model | SampleData, Histogram, CapabilityResult, Arr(工具) | 纯计算，**零 Excel 类型**。唯一例外：SampleData.Initialize 接受 Range 是 late-bound（IDispatch 调 Areas/Value2），不需要 Excel 引用。 |
| src.Spec | SeriesSpec, AxisSpec, AnnotationSpec, CapabilitySpec, PlotOptions, CapabilityPlotSpec | 纯数据容器，无 Excel、无统计学。 |
| src.Build | CapabilityPlotBuilder | BuildSpec 装配过程 Excel-free；Plot / Render 是面向 Excel 的 overload。 |
| src.Render | ChartCanvas, Painters, SeriesDump | 只认识 Spec，不认识统计学。.bas 用 `Option Private Module`。 |
| src.Test | CapabilityDemo | 冒烟测试（硬编码数据 + 真实工作表数据两条路径）。 |
| （工作簿侧） | ThisWorkbook.doccls, Sheet2.doccls | 用户维护，agent 不碰。 |

数据流：
```
SampleData ─→ Histogram(Build/StepPolygon/NormalCurve)
                 │
CapabilitySpec + PlotOptions ─→ Builder.BuildSpec ─→ CapabilityPlotSpec
                                                        │
                                          ChartCanvas + Painters ─→ Excel Chart
                                                        │（可选）
                                          SeriesDump.WriteData(ps, ws) ─→ 工作表数值
```

## 语义系列与 z-order

7 个语义系列 → 9~11 个真实 Excel Series。`CapabilityPlotSpec.AllSeries()` 按绘制顺序展平（Excel 先加的在下层）：

1. HistoShade（areaFill，次轴，直方图填充）
2. HistoProfile（xyLine，直方图轮廓）
3. DistCurve（xyLine，平滑，**仅 overall** 正态曲线）
4. MeanLine（2 点竖线）
5. TolLines（Collection：LSL / USL / 可选 Target）
6. SigmaLines（Collection：Mean ± k/2·sigma，0~2 条）
7. AxisLabels（隐藏散点 DataLabel，接管原生刻度标签）

Nothing = 未装配 = 跳过。渲染层不猜测"应该有什么"，装配层决定一切。

## 模块速查（src/ 17 文件）

| 文件 | 职责 |
|---|---|
| SampleData.cls | 数据摄入（Range late-bound / 一维数组 / 单行或单列二维数组；表头 hasHeader/header），N/Mean/StdevOverall/StdevWithin/min/max。非数值/空/错误单元格跳过；有效值 <2 报错。 |
| Histogram.cls | 等宽分箱（Sturges/Sqrt/FixedCount/FixedWidth，MAX_BINS 上限），Edges(0-based)/Counts/StepPolygon(每柱从 0 起)/NormalCurve(频次标定)。末边界钉死 e(k)=hi。 |
| CapabilityResult.cls | Cp/Cpk/Pp/Ppk/Cpm，Variant 返回（无规格时 Empty→显示 N/A）。Cpm 用 within-sigma/Taguchi（用户已对照 Minitab 确认关闭争议）。 |
| Arr.bas | 数组工具（Arr.Max 等），纯计算。 |
| CapabilitySpec.cls | LSL/USL/Target 全可选（Create 参数 Empty=未设），SigmaMultiplier 默认 6；双规格时 lower<upper 校验。 |
| PlotOptions.cls | 分箱策略/曲线点数/各线颜色与开关/YAxisMode(频次|密度)/Decimals(默认3, clamp 0..10)/ShowAxisLabels(默认True)/Title。零 Excel 类型。 |
| SeriesSpec.cls | 单系列：xs/ys + RenderKindEnum(renderXYLine/renderAreaFill/renderAxisLabels) + AxisGroup + 线型/填充/平滑 + Labels/LabelColors(Variant, 校验一维且与 xs 等长)。 |
| AxisSpec.cls | Min/Max/MajorUnit/Title/NumberFormat/Visible；HasMin/HasMax/HasMajorUnit；Create 全 Optional。 |
| AnnotationSpec.cls | 文本框数据（text/X/Y/fontName/fontSize/FontColor/Visible）。**锚点约定：(X,Y) = 文本框右上角**（2026-11-13 删除 align 管线后固定）。 |
| CapabilityPlotSpec.cls | 7 语义系列 + xAxis/yAxis + StatBox 聚合；AllSeries() 展平；集合属性 Set 传 Nothing 自动 New。 |
| CapabilityPlotBuilder.cls | BuildSpec（纯装配）/ Plot（一步到位，可选 dumpSheet）/ Render（重绘已装配规格）。 |
| ChartCanvas.cls | ChartObject 包装：FromChartObject(要求两轴显式 Min/Max)、ClearCanvas、AddSeries(按 RenderAs 分发)、ApplyAxes、RepairPrimaryAxis、FormatFrame、AddAnnotation、DataToPixelX/Y。 |
| Painters.bas | PaintXYLine / PaintAreaFill(ToScaledXs) / PaintAxisLabels / PrepareAreaAxis(重断言次轴) / ConfigureAreaAxis(xlTimeScale 1..10001, 隐藏次轴) / ToScaledX / ToMsoDash。 |
| SeriesDump.bas | WriteData(ps, ws)：ClearContents 后按系列导出 X/Y 块，列 stride 3。注意 `Write` 是保留字，方法名必须 WriteData。 |
| CapabilityDemo.bas | DrawCapabilityDemo（Sheet3 真实数据）/ CheckCapabilityDemo（60 硬编码点，EXPECTED_SERIES=10，MsgBox PASS/FAIL）。 |
| ThisWorkbook.doccls / Sheet2.doccls | 工作簿侧，用户维护。 |

## VBA 编码约定

- `Option Explicit` 全模块。
- 状态封装：`Private Type T{Name}` + `Private this As T{Name}`；`Class_Initialize` 设默认值。
- 工厂模式：`Public Function Create(...) As T` + `Public Sub Initialize(...)`；`@PredeclaredId`（属性行由用户经 Rubberduck 维护）。
- 属性：只读用 Get-only；**对象属性用 `Property Set`（不是 Let）**；集合属性 Set Nothing 时自动 New。
- 错误：每模块 `vbObjectError + NNNN` 分段（已核实：SampleData 1101~1103、SeriesSpec 1500 段(ERR_LABELS=1506)、AnnotationSpec 1701~1702、ChartCanvas 1801~1805、Builder 1901/1903）。`Err.Raise number, "Module.Proc", 中文消息`。
- Rubberduck 注解：`@Folder("src.Layer")`、`@ModuleDescription`、`@Description`、`@PredeclaredId`。
- 注释：中文（GBK）；`'=== 节 ===` 横幅；写"为什么"和 Excel 怪癖，不翻译代码。
- 命名：PascalCase 过程/属性，camelCase 字段/局部，`this.X` 访问结构体字段。
- 校验在 Spec 边界做：SeriesSpec.ValidatePair/ValidateLabelArray；CapabilitySpec 规格顺序；PlotOptions.Decimals clamp。
- **显示精度规则**：数据永远全精度；圆整只发生在 `Format$` / `TickLabels.NumberFormat`。绝不圆整数据本身（会破坏分箱、刻度对齐、统计）。

## 枚举作用域规则（用户纠正，权威）

- 类模块里的 `Public Enum` 提升到**项目作用域**。
- 类型位置写**裸枚举名**：`ByVal dash As LineDashEnum` ✓；`As SeriesSpec.LineDashEnum` ✗（编译错"用户定义类型未定义"）。
- 表达式里 `EnumName.member` 自限定合法（`Case BinPolicyEnum.binSturges`）；裸成员也合法（`binFixedCount`）。
- `ClassName.Procedure` 合法（`SeriesSpec.Create`、`Histogram.FromSample`）；库限定类型合法（`Excel.Chart`）。
- **不要一刀切去点号**。历史上生成代码曾有 32 处非法 `ClassName.EnumName` 引用。

## 保留字与命名陷阱（编译级教训）

- **语句关键字做过程名**：`Write`（Write #n）、`Load`、`Reset` —— 声明行可能通过，**调用点**报 461"方法和数据成员未找到"或编译失败。先例：SeriesDump.Write→WriteData（用户改名）、Load→Initialize、Reset→ClearCanvas。`Name` 做属性合法（Excel 自家 Series.Name），规则不统一。
- **参数与函数同名**（VBA 大小写不敏感）= 重复声明：`Function Fmt(ByVal fmt As String)` 事故 → 合并为 `FormatNumber(value As Variant, fmt As String)`，Empty→"N/A"，Double 自动提升 Variant，无需拆两个函数。
- 461 指向调用点而声明看似正常 → 先怀疑保留字。
- 跨模块成员解析错误（如复制粘贴导致的错误类型签名）会编译失败整个项目 → 静态检查必须跨模块，不能只看单文件。

## 统计口径

- StdevOverall：样本标准差（分母 N-1）。
- StdevWithin：MRbar / d2，d2=1.128（移动极差长度 2）；假设 subgroup=1。子组化不在范围，需要时另开接口。
- 密度模式：Y = 频次 / (N × BinWidth)，只在 Builder 缩放 py/cy/curveTop/maxBar，Histogram/Painters 不知情。