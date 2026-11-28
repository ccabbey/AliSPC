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
' Modified: 2026-11-29 01:41:19

Private Const ERR_NO_SAMPLE As Long = vbObjectError + 1901
Private Const ERR_NO_OPTIONS As Long = vbObjectError + 1903
Private Const ERR_NO_GROUPED As Long = vbObjectError + 1904      ' 分组数据为 Nothing 或没有分组
Private Const ERR_NO_TARGETS As Long = vbObjectError + 1905       ' 独立图表目标集合为 Nothing
Private Const ERR_TARGET_COUNT As Long = vbObjectError + 1906     ' 图表数量与分组数量不一致

Private Const X_PAD_RATIO As Double = 0.02      ' 统计框右边缘内缩比例（不再用于轴范围）
Private Const Y_HEADROOM As Double = 1.25       ' 纵轴顶部留白，给统计框腾位置
Private Const LABEL_EPS_RATIO As Double = 0.02  ' 辅助线标签去重半径（占轴跨度比例）
Private Const MAX_OVERLAY_BINS As Long = 1000    ' 多组叠加直方图的公共分箱数量上限

'=== 装配（纯函数，不碰 Excel）===============================================
' 本函数及其全部 Private 辅助过程只依赖 Model / Spec / Arr，
' 可以在没有图表、没有工作簿的情况下直接调用并断言结果。

'@Description "由数据、规格线、选项生成完整的绘图规格。 series 的添加顺序即 z-order。"
Public Function BuildSpec(ByVal source As SampleData, _
                          ByVal spec As CapabilitySpec, _
                          ByVal opt As PlotOptions, _
                          Optional ByVal titleOverride As String = "") As CapabilityPlotSpec
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
    maxBar = CDbl(hist.maxCount)
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
    ' 标题优先级：显式 opt.Title > 数据源表头（SampleData.Header）。
    result.title = titleOverride
    If Len(result.title) = 0 Then result.title = opt.title
    If Len(result.title) = 0 Then result.title = source.header

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

    Dim outlierXs() As Double
    Dim outlierYs() As Double
    Dim outlierLabels() As String
    Dim outlierColors() As Long
    Dim outlierCount As Long
    Dim outlierLimit As Double
    outlierLimit = source.Count * opt.OutlierLabelRatio
    ReDim outlierXs(1 To hist.binCount)
    ReDim outlierYs(1 To hist.binCount)
    ReDim outlierLabels(1 To hist.binCount)
    ReDim outlierColors(1 To hist.binCount)

    Dim i As Long
    For i = 0 To hist.binCount - 1
        If hist.counts(i) > 0 Then
            If CDbl(hist.counts(i)) <= outlierLimit Then
                outlierCount = outlierCount + 1
                outlierXs(outlierCount) = (CDbl(hist.edges(i)) + CDbl(hist.edges(i + 1))) / 2#
                outlierYs(outlierCount) = CDbl(hist.counts(i))
                If opt.YAxisMode = yAxisDensity Then
                    outlierYs(outlierCount) = outlierYs(outlierCount) / (source.Count * hist.BinWidth)
                End If
                outlierLabels(outlierCount) = "n=" & CStr(hist.counts(i))
                outlierColors(outlierCount) = opt.ProfileColor
            End If
        End If
    Next i

    If outlierCount > 0 Then
        ReDim Preserve outlierXs(1 To outlierCount)
        ReDim Preserve outlierYs(1 To outlierCount)
        ReDim Preserve outlierLabels(1 To outlierCount)
        ReDim Preserve outlierColors(1 To outlierCount)
        Set s = SeriesSpec.Create("outlierLabels", outlierXs, outlierYs, _
                                  RenderKindEnum.renderDataLabels, AxisGroupEnum.axisPrimary)
        s.labels = outlierLabels
        s.LabelColors = outlierColors
        Set result.outlierLabels = s
    End If

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
    '--- 轴标签系列（bin 边界全显示 + 辅助线精确值，epsilon 去重）---
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
                                                   9, "Arial")
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

'@Description "按分组构建独立直方图 Spec；每个分组占一张图。"
Public Function BuildGroupedSeparate(ByVal grouped As GroupedSampleData, _
                                     Optional ByVal spec As CapabilitySpec = Nothing, _
                                     Optional ByVal opt As PlotOptions = Nothing) As Collection
    If grouped Is Nothing Then
        Err.Raise ERR_NO_GROUPED, "CapabilityPlotBuilder.BuildGroupedSeparate", _
                  "grouped 不能为 Nothing。"
    End If
    If opt Is Nothing Then Set opt = PlotOptions.Create()

    Dim result As Collection
    Set result = New Collection
    Dim i As Long
    Dim groupOpt As PlotOptions
    Dim ps As CapabilityPlotSpec
    Dim title As String

    For i = 1 To grouped.GroupCount
        Set groupOpt = opt.Clone
        title = grouped.groupName(i)
        If Len(opt.title) > 0 Then title = opt.title & " - " & title
        Set ps = Me.BuildSpec(grouped.GroupSample(i), spec, groupOpt, title)
        result.Add ps
    Next i

    Set BuildGroupedSeparate = result
End Function

'@Description "构建多组同画布叠加直方图；分组 shade 透明度由 PlotOptions.GroupedShadeTransparency 控制"
Public Function BuildGroupedOverlay(ByVal grouped As GroupedSampleData, _
                                    Optional ByVal spec As CapabilitySpec = Nothing, _
                                    Optional ByVal opt As PlotOptions = Nothing) As CapabilityPlotSpec
    If grouped Is Nothing Then
        Err.Raise ERR_NO_GROUPED, "CapabilityPlotBuilder.BuildGroupedOverlay", _
                  "grouped 不能为 Nothing。"
    End If
    If opt Is Nothing Then Set opt = PlotOptions.Create()

    Dim result As CapabilityPlotSpec
    Set result = New CapabilityPlotSpec
    result.title = opt.title
    If Len(result.title) = 0 Then result.title = "Grouped Histogram"
    result.LegendVisible = True

    Dim groupSpecs As Collection
    Dim histograms As Collection
    Dim shades As Collection
    Dim profiles As Collection
    Dim curves As Collection
    Dim outliers As Collection
    Set groupSpecs = New Collection
    Set histograms = New Collection
    Set shades = New Collection
    Set profiles = New Collection
    Set curves = New Collection
    Set outliers = New Collection

    Dim xLo As Double
    Dim xHi As Double
    Dim commonLow As Double
    Dim commonHigh As Double
    Dim commonBinCount As Long
    Dim commonWidth As Double
    Dim yTop As Double
    Dim i As Long
    Dim groupOpt As PlotOptions
    Dim ps As CapabilityPlotSpec
    Dim hist As Histogram

    For i = 1 To grouped.GroupCount
        Set groupOpt = opt.Clone
        groupOpt.title = ""
        groupOpt.ShowTol = False
        groupOpt.ShowTarget = False
        groupOpt.ShowSigma = False
        groupOpt.ShowStatBox = False
        groupOpt.ShowAxisLabels = False
        groupOpt.ShadeTransparency = opt.GroupedShadeTransparency
        groupOpt.ProfileColor = GroupColor(i)
        groupOpt.ShadeColor = GroupColor(i)
        groupOpt.CurveColor = GroupColor(i)

        Set hist = Histogram.FromSample(grouped.GroupSample(i), _
                                        opt.BinPolicy, opt.BinParam)
        histograms.Add hist

        Set ps = Me.BuildSpec(grouped.GroupSample(i), spec, groupOpt, _
                              grouped.groupName(i))
        groupSpecs.Add ps

        If i = 1 Then
            xLo = CDbl(ps.xAxis.Min)
            xHi = CDbl(ps.xAxis.Max)
            commonLow = hist.EdgeLow
            commonHigh = hist.EdgeHigh
            commonBinCount = hist.binCount
        Else
            If CDbl(ps.xAxis.Min) < xLo Then xLo = CDbl(ps.xAxis.Min)
            If CDbl(ps.xAxis.Max) > xHi Then xHi = CDbl(ps.xAxis.Max)
            If hist.EdgeLow < commonLow Then commonLow = hist.EdgeLow
            If hist.EdgeHigh > commonHigh Then commonHigh = hist.EdgeHigh
            If hist.binCount > commonBinCount Then commonBinCount = hist.binCount
        End If
    Next i

    If opt.BinPolicy = binFixedWidth And opt.BinParam > 0# Then
        commonBinCount = Int((commonHigh - commonLow) / opt.BinParam + 0.5)
    End If
    If commonBinCount < 1 Then commonBinCount = 1
    If commonBinCount > MAX_OVERLAY_BINS Then commonBinCount = MAX_OVERLAY_BINS
    commonWidth = (commonHigh - commonLow) / CDbl(commonBinCount)

    Dim px As Variant
    Dim py As Variant
    Dim maxCount As Long
    Dim maxBar As Double
    Dim yScale As Double
    Dim curveYs As Variant
    Dim curveScale As Double
    Dim curveTop As Double
    Dim s As SeriesSpec

    For i = 1 To histograms.Count
        Set hist = histograms(i)
        Set ps = groupSpecs(i)

        hist.StepPolygonUniform commonLow, commonHigh, commonBinCount, _
                                px, py, maxCount

        maxBar = CDbl(maxCount)
        If opt.YAxisMode = yAxisDensity Then
            yScale = 1# / (grouped.GroupSample(i).Count * commonWidth)
            ScaleYs py, yScale
            maxBar = maxBar * yScale
        End If
        If maxBar > yTop Then yTop = maxBar

        Set s = SeriesSpec.Create("histoShade", px, py, _
                                  RenderKindEnum.renderAreaFill, _
                                  AxisGroupEnum.axisSecondary)
        s.FillColor = GroupColor(i)
        s.FillTransparency = opt.GroupedShadeTransparency
        s.Name = grouped.groupName(i) & " shade"
        shades.Add s

        Set s = SeriesSpec.Create("histoProfile", px, py, _
                                  RenderKindEnum.renderXYLine, _
                                  AxisGroupEnum.axisPrimary)
        s.LineColor = GroupColor(i)
        s.LineWeight = 1.25
        s.Name = grouped.groupName(i)
        profiles.Add s

        If Not ps.DistCurve Is Nothing Then
            Set s = ps.DistCurve
            s.Name = grouped.groupName(i) & " normal"

            If opt.YAxisMode <> yAxisDensity Then
                curveYs = s.ys
                curveScale = commonWidth / hist.BinWidth
                ScaleYs curveYs, curveScale
                s.ys = curveYs
            End If

            curveTop = Arr.Max(s.ys)
            If curveTop > yTop Then yTop = curveTop
            curves.Add s
        End If

        If Not ps.outlierLabels Is Nothing Then
            Set s = ps.outlierLabels
            s.Name = grouped.groupName(i) & " outlier"
            outliers.Add s
        End If
    Next i

    yTop = yTop * Y_HEADROOM
    If yTop <= 0# Then yTop = 1#

    Dim fmt As String
    fmt = "0"
    If opt.Decimals > 0 Then fmt = "0." & String$(opt.Decimals, "0")

    Dim xFmt As String
    Dim yTitle As String
    Dim yFmt As String
    xFmt = fmt
    If opt.ShowAxisLabels Then
        Set result.AxisLabels = BuildUniformAxisLabels( _
                                    commonLow, commonHigh, commonBinCount, fmt)
        xFmt = ";;;"
    End If
    yTitle = ""
    yFmt = "0"
    If opt.YAxisMode = yAxisDensity Then
        yTitle = "Density"
        yFmt = fmt
    End If

    Set result.xAxis = AxisSpec.Create(xLo, xHi, , "", xFmt, True)
    Set result.yAxis = AxisSpec.Create(0#, yTop, , yTitle, yFmt, True)

    For Each s In shades
        result.GroupSeries.Add s
    Next s
    For Each s In profiles
        result.GroupSeries.Add s
    Next s
    For Each s In curves
        result.GroupSeries.Add s
    Next s
    For Each s In outliers
        result.GroupSeries.Add s
    Next s

    Set BuildGroupedOverlay = result
End Function

'@Description "把分组叠加直方图渲染到指定 ChartObject。"
Public Sub PlotGroupedOverlay(ByVal target As ChartObject, _
                              ByVal grouped As GroupedSampleData, _
                              Optional ByVal spec As CapabilitySpec = Nothing, _
                              Optional ByVal opt As PlotOptions = Nothing)
    Dim ps As CapabilityPlotSpec
    Set ps = Me.BuildGroupedOverlay(grouped, spec, opt)
    Me.Render target, ps
End Sub

'@Description "把分组独立直方图依次渲染到 ChartObject 集合。"
Public Sub PlotGroupedSeparate(ByVal targets As Collection, _
                               ByVal grouped As GroupedSampleData, _
                               Optional ByVal spec As CapabilitySpec = Nothing, _
                               Optional ByVal opt As PlotOptions = Nothing)
    If targets Is Nothing Then
        Err.Raise ERR_NO_TARGETS, "CapabilityPlotBuilder.PlotGroupedSeparate", _
                  "targets 不能为 Nothing。"
    End If

    Dim specs As Collection
    Set specs = Me.BuildGroupedSeparate(grouped, spec, opt)
    If targets.Count <> specs.Count Then
        Err.Raise ERR_TARGET_COUNT, "CapabilityPlotBuilder.PlotGroupedSeparate", _
                  "图表数量 " & targets.Count & " 与分组数量 " & specs.Count & " 不一致。"
    End If

    Dim i As Long
    For i = 1 To specs.Count
        Me.Render targets(i), specs(i)
    Next i
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
    canvas.FormatFrame ps.title, ps.LegendVisible
    canvas.AddAnnotation ps.StatBox
End Sub

'=== 内部辅助（全部无 Excel 依赖）===========================================

' 原地扩展区间，是 mutator 而不是把返回值塞进出参。
Private Function GroupColor(ByVal index As Long) As Long
    Select Case (index - 1) Mod 8
        Case 0: GroupColor = RGB(76, 120, 168)    ' 雾蓝 #4C78A8
        Case 1: GroupColor = RGB(245, 133, 24)    ' 暖橙 #F58518
        Case 2: GroupColor = RGB(84, 162, 75)     ' 鼠尾草绿 #54A24B
        Case 3: GroupColor = RGB(114, 183, 178)   ' 灰青 #72B7B2
        Case 4: GroupColor = RGB(180, 121, 162)   ' 灰紫 #B279A2
        Case 5: GroupColor = RGB(157, 117, 93)    ' 灰棕 #9D755D
        Case 6: GroupColor = RGB(228, 87, 86)     ' 珊瑚红 #E45756
        Case 7: GroupColor = RGB(186, 176, 172)   ' 暖灰 #BAB0AC
    End Select
End Function

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

Private Function BuildUniformAxisLabels(ByVal low As Double, _
                                        ByVal high As Double, _
                                        ByVal binCount As Long, _
                                        ByVal fmt As String) As SeriesSpec
    Dim width As Double
    Dim xs() As Double
    Dim ys() As Double
    Dim labels() As String
    Dim colors() As Long
    Dim i As Long

    width = (high - low) / CDbl(binCount)
    ReDim xs(1 To binCount + 1)
    ReDim ys(1 To binCount + 1)
    ReDim labels(1 To binCount + 1)
    ReDim colors(1 To binCount + 1)

    For i = 0 To binCount
        xs(i + 1) = low + CDbl(i) * width
        If i = binCount Then xs(i + 1) = high
        ys(i + 1) = 0#
        labels(i + 1) = Format$(xs(i + 1), fmt)
        colors(i + 1) = RGB(60, 60, 60)
    Next i

    Dim result As SeriesSpec
    Set result = SeriesSpec.Create("axisLabels", xs, ys, _
                                   RenderKindEnum.renderAxisLabels, _
                                   AxisGroupEnum.axisPrimary)
    result.labels = labels
    result.LabelColors = colors
    Set BuildUniformAxisLabels = result
End Function

' 收集轴标签：bin 边界全部显示（eps=0 恒不去重）；辅助线精确值按 eps 去重。
' 辅助标签（规格/均值/Sigma）与已有标签距离 < eps 时跳过，避免近似重复。
Private Function BuildAxisLabels(ByVal hist As Histogram, _
                                 ByVal source As SampleData, _
                                 ByVal spec As CapabilitySpec, _
                                 ByVal opt As PlotOptions, _
                                 ByVal xLo As Double, ByVal xHi As Double, _
                                 ByVal fmt As String) As SeriesSpec
    Dim cap As Long
    cap = hist.binCount + 1 + 6          ' 上限：bin 边界 + 最多 6 条辅助线
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

    ' 1) bin 边界：每条都显示（含末边界），eps=0 保证不被去重吞掉。
    Dim edges As Variant
    edges = hist.edges
    Dim k As Long
    For k = 0 To hist.binCount
        AddLabel xs, ys, lb, lc, cnt, CDbl(edges(k)), fmt, 0#, RGB(60, 60, 60)
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
    result.labels = al
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


