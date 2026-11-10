VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CapabilityPlotBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "Rubberduck" ,"Predeclared Class Module"
'@ModuleDescription "把 SampleData + CapabilitySpec + PlotOptions 装配成 CapabilityPlotSpec，再交给 ChartCanvas 落地。"
'@PredeclaredId
'@Folder("src.Build")
Option Explicit

Private Const ERR_NO_SAMPLE As Long = vbObjectError + 1901
Private Const ERR_NO_SPEC As Long = vbObjectError + 1902
Private Const ERR_NO_OPTIONS As Long = vbObjectError + 1903

Private Const X_PAD_RATIO As Double = 0.02      ' 横轴两端留白比例
Private Const Y_HEADROOM As Double = 1.25       ' 纵轴顶部留白，给统计框腾位置

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
    If spec Is Nothing Then
        Err.Raise ERR_NO_SPEC, "CapabilityPlotBuilder.BuildSpec", "spec 不能为 Nothing。"
    End If
    If opt Is Nothing Then
        Err.Raise ERR_NO_OPTIONS, "CapabilityPlotBuilder.BuildSpec", "opt 不能为 Nothing。"
    End If

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

    Dim pad As Double
    pad = (xHi - xLo) * X_PAD_RATIO
    If pad <= 0# Then pad = 1#
    xLo = xLo - pad
    xHi = xHi + pad

    '--- 纵轴范围：必须同时容纳最高箱和曲线峰值，否则填色与轮廓会错位 ---
    Dim yTop As Double
    yTop = CDbl(hist.MaxCount)
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
    Set result.xAxis = AxisSpec.Create(xLo, xHi, , "", "General", True)
    Set result.yAxis = AxisSpec.Create(0#, yTop, , "", "General", True)

    '--- 统计框，贴在绘图区右上角 ---
    If opt.ShowStatBox Then
        Set result.StatBox = AnnotationSpec.Create(StatText(source, spec), _
                                                   xHi - pad, yTop * 0.98, _
                                                   TextAlignEnum.alignRight, 9, "Arial")
    End If

    Set BuildSpec = result
End Function

'=== 落地（碰 Excel）=========================================================

'@Description "一步到位：装配 + 渲染到已有 ChartObject。"
Public Sub Plot(ByVal target As ChartObject, _
                ByVal source As SampleData, _
                ByVal spec As CapabilitySpec, _
                ByVal opt As PlotOptions)
    Dim ps As CapabilityPlotSpec
    Set ps = Me.BuildSpec(source, spec, opt)
    Me.Render target, ps
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

Private Function StatText(ByVal source As SampleData, ByVal spec As CapabilitySpec) As String
    Dim cap As CapabilityResult
    Set cap = CapabilityResult.FromSample(source, spec)

    Dim t As String
    t = "N = " & CStr(source.Count) & vbCrLf
    t = t & "Mean = " & Fmt(source.Mean) & vbCrLf
    t = t & "StDev(Overall) = " & Fmt(source.StdevOverall) & vbCrLf
    t = t & "StDev(Within) = " & Fmt(source.StdevWithin) & vbCrLf
    t = t & "Cp = " & FmtOpt(cap.Cp) & "    Cpk = " & FmtOpt(cap.Cpk) & vbCrLf
    t = t & "Pp = " & FmtOpt(cap.Pp) & "    Ppk = " & FmtOpt(cap.Ppk) & vbCrLf
    t = t & "Cpm = " & FmtOpt(cap.Cpm)
    If spec.HasLowerSpec Then
        If spec.HasUpperSpec Then
            t = t & vbCrLf & "LSL = " & Fmt(spec.lowerSpec) & "   USL = " & Fmt(spec.upperSpec)
        Else
            t = t & vbCrLf & "LSL = " & Fmt(spec.lowerSpec)
        End If
    ElseIf spec.HasUpperSpec Then
        t = t & vbCrLf & "USL = " & Fmt(spec.upperSpec)
    End If
    If spec.HasTarget Then t = t & vbCrLf & "Target = " & Fmt(spec.target)
    StatText = t
End Function

Private Function Fmt(ByVal value As Double) As String
    Fmt = Format$(value, "0.000")
End Function

' 能力指数用 Variant + Empty 表示「不适用」（单侧规格时 Cp / Cpm 无定义）。
Private Function FmtOpt(ByVal value As Variant) As String
    If IsEmpty(value) Then
        FmtOpt = "N/A"
    Else
        FmtOpt = Format$(CDbl(value), "0.000")
    End If
End Function



