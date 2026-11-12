VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CapabilityPlotBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@ModuleDescription "把 SampleData + CapabilitySpec + PlotOptions 装配成 CapabilityPlotSpec，再交给 ChartCanvas 落地。"
'@PredeclaredId
'@Folder("src.Build")
Option Explicit

Private Const ERR_NO_SAMPLE As Long = vbObjectError + 1901
Private Const ERR_NO_OPTIONS As Long = vbObjectError + 1903

Private Const X_PAD_RATIO As Double = 0.02      ' 统计框右边缘内缩比例（不再用于轴范围）
Private Const Y_HEADROOM As Double = 1.25       ' 纵轴顶部留白，给统计框腾位置
Private Const MAX_AXIS_LABELS As Long = 10      ' bin 边界标签抽稀上限
Private Const LABEL_EPS_RATIO As Double = 0.02  ' 辅助线标签去重半径（占轴跨度比例）

'=== 装配（纯函数，不碰 Excel）===============================================
' 本函数及其全部 Private 辅助过程只依赖 Model / Spec / Arr，
' 可以在没有图表、没有工作簿的情况下直接调用并断言结果。

'@Description "由数据、规格线、选项生成完整的绘图规格。 series 的添加顺序即 z-order。"
Public Function BuildSpec(ByVal source As SampleData, _
                          ByVal spec As CapabilitySpec, _
                          ByVal opt As PlotOptions) As CapabilityPlotSpec
    If source Is Nothing Then
        Err.Raise ERR_NO_SAMPLE, "CapabilityPlotBuilder.BuildSpec", "source 不能为 Nothing。"
    End If
    If spec Is Nothing Then Set spec = CapabilitySpec.Create()
    If opt Is Nothing Then
        Err.Raise ERR_NO_OPTIONS, "CapabilityPlotBuilder.BuildSpec", "opt 不能为 Nothing。"
    End If

    ' 显示精度：只影响格式字符串，数据系列保持全精度
    Dim fmt As String
    fmt = "0"
    If opt.Decimals > 0 Then fmt = "0." & String$(opt.Decimals, "0")

    Dim hist As Histogram
    Set hist = Histogram.FromSample(source, opt.BinPolicy, opt.BinParam)

    Dim px As Variant
    Dim py As Variant
    hist.StepPolygon px, py

    Dim cx As Variant
    Dim cy As Variant
    Dim curveTop As Double
    If opt.ShowDistCurve Then
        ' 只画 overall 一条曲线，这是已定的范围。
        hist.NormalCurve source.StdevOverall, opt.CurvePoints, cx, cy
        curveTop = Arr.Max(cy)
    End If

    '--- 纵轴模式：概率密度 = 频数 / (N * BinWidth)。只缩放 Y，X 不动 ---
    Dim maxBar As Double
    maxBar = CDbl(hist.MaxCount)
    If opt.YAxisMode = yAxisDensity Then
        Dim yScale As Double
        yScale = 1# / (source.Count * hist.BinWidth)
        ScaleYs py, yScale
        If opt.ShowDistCurve Then
            ScaleYs cy, yScale
            curveTop = curveTop * yScale
        End If
        maxBar = maxBar * yScale
    End If

    '--- 横轴范围：直方图两侧各多留一个箱宽（与曲线跨度一致），再纳入所有竖线 ---
    Dim xLo As Double
    Dim xHi As Double
    xLo = hist.EdgeLow - hist.BinWidth
    xHi = hist.EdgeHigh + hist.BinWidth

    ExtendRange xLo, xHi, source.Mean
    If spec.HasLowerSpec Then ExtendRange xLo, xHi, spec.lowerSpec
    If spec.HasUpperSpec Then ExtendRange xLo, xHi, spec.upperSpec
    If spec.HasTarget Then
        If opt.ShowTarget Then ExtendRange xLo, xHi, spec.target
    End If
    If opt.ShowSigma Then
        ExtendRange xLo, xHi, source.Mean - HalfSpan(spec, source)
        ExtendRange xLo, xHi, source.Mean + HalfSpan(spec, source)
    End If

    '--- todo1：轴边界向外对齐到 bin 网格（锚点 EdgeLow），刻度精确落在柱边界 ---
    Dim bw As Double
    bw = hist.BinWidth
    xLo = hist.EdgeLow + bw * Int((xLo - hist.EdgeLow) / bw)
    xHi = hist.EdgeLow - bw * Int(-(xHi - hist.EdgeLow) / bw)

    Dim pad As Double
    pad = (xHi - xLo) * X_PAD_RATIO      ' 仅用于统计框内缩
    If pad <= 0# Then pad = 1#

    '--- 纵轴范围：必须同时容纳最高箱和曲线峰值，否则填色与轮廓会错位 ---
    Dim yTop As Double
    yTop = maxBar
    If curveTop > yTop Then yTop = curveTop
    yTop = yTop * Y_HEADROOM
    If yTop <= 0# Then yTop = 1#

    Dim result As CapabilityPlotSpec
    Set result = New CapabilityPlotSpec
    result.Title = opt.Title

    '--- 系列。添加顺序就是绘制顺序，填色必须最先，否则盖住轮廓线 ---
    Dim s As SeriesSpec

    Set s = SeriesSpec.Create("histoShade", px, py, _
                              RenderKindEnum.renderAreaFill, AxisGroupEnum.axisSecondary)
    s.FillColor = opt.ShadeColor
    s.FillTransparency = opt.ShadeTransparency
    Set result.HistoShade = s

    Set s = SeriesSpec.Create("histoProfile", px, py, _
                              RenderKindEnum.renderXYLine, AxisGroupEnum.axisPrimary)
    s.LineColor = opt.ProfileColor
    s.LineWeight = 1.25
    Set result.HistoProfile = s

    If opt.ShowDistCurve Then
        Set s = SeriesSpec.Create("distCurve", cx, cy, _
                                  RenderKindEnum.renderXYLine, AxisGroupEnum.axisPrimary)
        s.Smooth = True                 ' 渲染层据此选 xlXYScatterSmoothNoMarkers，绝不是 xlLine
        s.LineColor = opt.CurveColor
        s.LineWeight = 1.75
        Set result.DistCurve = s
    End If

    Set s = VerticalLine("mean", source.Mean, yTop, opt.MeanColor, LineDashEnum.dashSolid)
    s.LineWeight = 1.25
    Set result.MeanLine = s

    If opt.ShowTol Then
        If spec.HasLowerSpec Then
            result.TolLines.Add VerticalLine("LSL", spec.lowerSpec, yTop, _
                                             opt.TolColor, LineDashEnum.dashSolid)
        End If
        If spec.HasUpperSpec Then
            result.TolLines.Add VerticalLine("USL", spec.upperSpec, yTop, _
                                             opt.TolColor, LineDashEnum.dashSolid)
        End If
        If spec.HasTarget Then
            If opt.ShowTarget Then
                result.TolLines.Add VerticalLine("Target", spec.target, yTop, _
                                                 opt.TargetColor, LineDashEnum.dashDash)
            End If
        End If
    End If

    If opt.ShowSigma Then
        ' ±k/2 倍样本标准差，与 distCurve 同源（overall），默认 k=6 即 ±3σ。
        Dim half As Double
        half = HalfSpan(spec, source)
        result.SigmaLines.Add VerticalLine("Sigma3xLow", source.Mean - half, yTop, _
                                           opt.SigmaColor, LineDashEnum.dashDash)
        result.SigmaLines.Add VerticalLine("Sigma3xHigh", source.Mean + half, yTop, _
                                           opt.SigmaColor, LineDashEnum.dashDash)
    End If

    '--- 轴。范围必须显式给定，ChartCanvas 会拒绝自动缩放的轴 ---
    '--- todo4：横轴标签系列（bin 边界抽稀 + 辅助线精确值，epsilon 去重）---
    If opt.ShowAxisLabels Then
        Set result.AxisLabels = BuildAxisLabels(hist, source, spec, opt, xLo, xHi, fmt)
    End If

    Dim xFmt As String
    xFmt = fmt
    If opt.ShowAxisLabels Then xFmt = ";;;"    ' 隐藏原生刻度标签，由 axisLabels 接管
    Set result.xAxis = AxisSpec.Create(xLo, xHi, bw, "", xFmt, True)

    Dim yTitle As String
    Dim yFmt As String
    yTitle = ""
    yFmt = "0"                              ' 频数是整数
    If opt.YAxisMode = yAxisDensity Then
        yTitle = "Density"
        yFmt = fmt                          ' 密度按 Decimals 显示
    End If
    Set result.yAxis = AxisSpec.Create(0#, yTop, , yTitle, yFmt, True)

    '--- 统计框，贴在绘图区右上角 ---
    If opt.ShowStatBox Then
        Set result.StatBox = AnnotationSpec.Create(StatText(source, spec, fmt), _
                                                   xHi - pad, yTop * 0.98, _
                                                   TextAlignEnum.alignRight, 9, "Arial")
    End If

    Set BuildSpec = result
End Function

'=== 落地（碰 Excel）=========================================================

'@Description "一步到位：装配 + 渲染。spec / opt 可省略；dumpSheet 非 Nothing 时同步输出系列值。"
Public Sub Plot(ByVal target As ChartObject, _
                ByVal source As SampleData, _
                Optional ByVal spec As CapabilitySpec = Nothing, _
                Optional ByVal opt As PlotOptions = Nothing, _
                Optional ByVal dumpSheet As Worksheet = Nothing)
    If opt Is Nothing Then Set opt = PlotOptions.Create()
    Dim ps As CapabilityPlotSpec
    Set ps = Me.BuildSpec(source, spec, opt)
    Me.Render target, ps
    If Not dumpSheet Is Nothing Then SeriesDump.WriteData ps, dumpSheet
End Sub

'@Description "渲染一份已经装配好的规格。方便先断言 BuildSpec 再画，或手绘微调后重画。"
Public Sub Render(ByVal target As ChartObject, ByVal ps As CapabilityPlotSpec)
    Dim canvas As ChartCanvas
    Set canvas = ChartCanvas.FromChartObject(target, ps.xAxis, ps.yAxis)

    canvas.ClearCanvas

    Dim s As SeriesSpec
    For Each s In ps.AllSeries()
        canvas.AddSeries s
    Next s

    canvas.ApplyAxes
    canvas.FormatFrame ps.Title
    canvas.AddAnnotation ps.StatBox
End Sub

'=== 内部辅助（全部无 Excel 依赖）===========================================

' 原地扩展区间，是 mutator 而不是把返回值塞进出参。
Private Sub ExtendRange(ByRef lo As Double, ByRef hi As Double, ByVal value As Double)
    If value < lo Then lo = value
    If value > hi Then hi = value
End Sub

' 规格线到均值的半跨度：SigmaMultiplier / 2 * 样本标准差。
Private Function HalfSpan(ByVal spec As CapabilitySpec, ByVal source As SampleData) As Double
    HalfSpan = spec.SigmaMultiplier / 2# * source.StdevOverall
End Function

' 原地缩放 Y 数组（密度模式用）。X 数组不经过这里。
Private Sub ScaleYs(ByRef ys As Variant, ByVal factor As Double)
    Dim i As Long
    For i = LBound(ys) To UBound(ys)
        ys(i) = ys(i) * factor
    Next i
End Sub

Private Function VerticalLine(ByVal Name As String, ByVal xValue As Double, _
                              ByVal yTop As Double, ByVal color As Long, _
                              ByVal dash As LineDashEnum) As SeriesSpec
    Dim xs(1 To 2) As Double
    Dim ys(1 To 2) As Double
    xs(1) = xValue
    xs(2) = xValue
    ys(1) = 0#
    ys(2) = yTop

    Dim result As SeriesSpec
    Set result = SeriesSpec.Create(Name, xs, ys, _
                                   RenderKindEnum.renderXYLine, AxisGroupEnum.axisPrimary)
    result.LineColor = color
    result.DashStyle = dash
    result.LineWeight = 1#
    Set VerticalLine = result
End Function

' todo4：收集横轴标签。bin 边界按 MAX_AXIS_LABELS 抽稀；辅助线精确值按
' 规格 -> 均值 -> Sigma 的优先级加入，与已有标签距离 < eps 时跳过。
Private Function BuildAxisLabels(ByVal hist As Histogram, _
                                 ByVal source As SampleData, _
                                 ByVal spec As CapabilitySpec, _
                                 ByVal opt As PlotOptions, _
                                 ByVal xLo As Double, ByVal xHi As Double, _
                                 ByVal fmt As String) As SeriesSpec
    Dim cap As Long
    cap = hist.BinCount + 1 + 6          ' 上限：bin 边界 + 最多 6 条辅助线
    Dim xs() As Double
    Dim ys() As Double
    Dim lb() As String
    Dim lc() As Long
    ReDim xs(1 To cap)
    ReDim ys(1 To cap)
    ReDim lb(1 To cap)
    ReDim lc(1 To cap)
    Dim cnt As Long
    cnt = 0
    Dim eps As Double
    eps = (xHi - xLo) * LABEL_EPS_RATIO

    ' 1) bin 边界（用 Edges 的钉住值，与柱形严格一致）
    Dim edges As Variant
    edges = hist.edges
    Dim k As Long
    Dim stride As Long
    stride = Int(hist.BinCount / MAX_AXIS_LABELS) + 1
    For k = 0 To hist.BinCount
        If k Mod stride = 0 Then
            AddLabel xs, ys, lb, lc, cnt, CDbl(edges(k)), fmt, eps, RGB(60, 60, 60)
        End If
    Next k

    ' 2) 辅助线精确值，颜色与对应线一致
    If opt.ShowTol Then
        If spec.HasLowerSpec Then AddLabel xs, ys, lb, lc, cnt, spec.lowerSpec, fmt, eps, opt.TolColor
        If spec.HasUpperSpec Then AddLabel xs, ys, lb, lc, cnt, spec.upperSpec, fmt, eps, opt.TolColor
        If spec.HasTarget Then
            If opt.ShowTarget Then AddLabel xs, ys, lb, lc, cnt, spec.target, fmt, eps, opt.TargetColor
        End If
    End If
    AddLabel xs, ys, lb, lc, cnt, source.Mean, fmt, eps, opt.MeanColor
    If opt.ShowSigma Then
        Dim half As Double
        half = HalfSpan(spec, source)
        AddLabel xs, ys, lb, lc, cnt, source.Mean - half, fmt, eps, opt.SigmaColor
        AddLabel xs, ys, lb, lc, cnt, source.Mean + half, fmt, eps, opt.SigmaColor
    End If

    ' 截断到实际数量
    Dim ax() As Double
    Dim ay() As Double
    Dim al() As String
    Dim ac() As Long
    ReDim ax(1 To cnt)
    ReDim ay(1 To cnt)
    ReDim al(1 To cnt)
    ReDim ac(1 To cnt)
    For k = 1 To cnt
        ax(k) = xs(k)
        ay(k) = ys(k)
        al(k) = lb(k)
        ac(k) = lc(k)
    Next k

    Dim result As SeriesSpec
    Set result = SeriesSpec.Create("axisLabels", ax, ay, _
                                   RenderKindEnum.renderAxisLabels, AxisGroupEnum.axisPrimary)
    result.Labels = al
    result.LabelColors = ac
    Set BuildAxisLabels = result
End Function

' 追加一个标签；与已有标签距离 < eps 时跳过（只去重，不防视觉重叠）。
Private Sub AddLabel(ByRef xs() As Double, ByRef ys() As Double, _
                     ByRef lb() As String, ByRef lc() As Long, _
                     ByRef cnt As Long, ByVal X As Double, _
                     ByVal fmt As String, ByVal eps As Double, ByVal color As Long)
    Dim i As Long
    For i = 1 To cnt
        If Abs(xs(i) - X) < eps Then Exit Sub
    Next i
    cnt = cnt + 1
    xs(cnt) = X
    ys(cnt) = 0#
    lb(cnt) = Format$(X, fmt)
    lc(cnt) = color
End Sub

Private Function StatText(ByVal source As SampleData, ByVal spec As CapabilitySpec, ByVal fmt As String) As String
    Dim cap As CapabilityResult
    Set cap = CapabilityResult.FromSample(source, spec)

    Dim t As String
    t = "N = " & CStr(source.Count) & vbCrLf
    t = t & "Mean = " & FormatNumber(source.Mean, fmt) & vbCrLf
    t = t & "StDev(Overall) = " & FormatNumber(source.StdevOverall, fmt) & vbCrLf
    t = t & "StDev(Within) = " & FormatNumber(source.StdevWithin, fmt) & vbCrLf
    t = t & "Cp = " & FormatNumber(cap.Cp, fmt) & "    Cpk = " & FormatNumber(cap.Cpk, fmt) & vbCrLf
    t = t & "Pp = " & FormatNumber(cap.Pp, fmt) & "    Ppk = " & FormatNumber(cap.Ppk, fmt) & vbCrLf
    t = t & "Cpm = " & FormatNumber(cap.Cpm, fmt)
    If spec.HasLowerSpec Then
        If spec.HasUpperSpec Then
            t = t & vbCrLf & "LSL = " & FormatNumber(spec.lowerSpec, fmt) & "   USL = " & FormatNumber(spec.upperSpec, fmt)
        Else
            t = t & vbCrLf & "LSL = " & FormatNumber(spec.lowerSpec, fmt)
        End If
    ElseIf spec.HasUpperSpec Then
        t = t & vbCrLf & "USL = " & FormatNumber(spec.upperSpec, fmt)
    End If
    If spec.HasTarget Then t = t & vbCrLf & "Target = " & FormatNumber(spec.target, fmt)
    StatText = t
End Function

' 数值 -> 显示文本。Empty（能力指数不适用）显示 "N/A"，其余按 fmt 格式化。
' 入参用 Variant：Double（均值/标准差/规格值）与可能为 Empty 的指数统一走这里。
Private Function FormatNumber(ByVal value As Variant, ByVal fmt As String) As String
    If IsEmpty(value) Then
        FormatNumber = "N/A"
    Else
        FormatNumber = Format$(CDbl(value), fmt)
    End If
End Function


