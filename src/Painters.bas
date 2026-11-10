Attribute VB_Name = "Painters"
'@ModuleDescription "两种绘制原语。用面积图伪装直方图的全部脏活（×10000 缩放、时间刻度副轴）都封在本模块内。"
'@Folder("src.Render")
Option Explicit
Option Private Module

' 面积图横轴放大倍数。经验值，等宽分箱下 10000 足够把量化误差压到绘图区宽度的 1/20000。
' 这是全项目唯一出现该数字的地方；副轴上下限也由这里推导，外部不需要知道。
Private Const AREA_SCALE As Double = 10000
Private Const AREA_BASE As Double = 1

'=== 原语一：散点连线 =========================================================
' 承担 histoProfile / distCurve / mean / Tol / Sigma3x 五种语义系列。
' 必须是 XY 散点而不是 xlLine：只有散点系列的 XValues 才按数值定位，
' xlLine 的 X 来自 Categories（序号），叠加到面积图上会错位。

Public Sub PaintXYLine(ByVal canvas As ChartCanvas, ByVal spec As SeriesSpec)
    Dim s As Excel.Series
    Set s = canvas.Chart.SeriesCollection.NewSeries

    s.Name = spec.Name
    If spec.Smooth Then
        s.ChartType = xlXYScatterSmoothNoMarkers
    Else
        s.ChartType = xlXYScatterLinesNoMarkers
    End If

    If spec.AxisGroup = AxisGroupEnum.axisSecondary Then
        s.AxisGroup = xlSecondary
    Else
        s.AxisGroup = xlPrimary
    End If

    s.XValues = spec.xs
    s.Values = spec.ys
    s.MarkerStyle = xlMarkerStyleNone

    With s.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = spec.LineColor
        .Weight = spec.LineWeight
        .DashStyle = ToMsoDash(spec.DashStyle)
    End With
End Sub

'=== 原语二：面积图填色 =======================================================
' 只承担 histoShade。X 被归一化后放大到 1..10001，落在一条隐藏的时间刻度副横轴上；
' 纵轴与主轴强制同范围，保证填色和散点轮廓逐点重合。

Public Sub PaintAreaFill(ByVal canvas As ChartCanvas, ByVal spec As SeriesSpec)
    Dim s As Excel.Series
    Set s = canvas.Chart.SeriesCollection.NewSeries

    s.Name = spec.Name
    s.ChartType = xlArea
    s.AxisGroup = xlSecondary

    ' 先写 Values 再写 XValues 也可以，但副轴此刻还是普通分类轴，
    ' 真正的按值定位要等 ConfigureAreaAxis 把 CategoryType 改成 xlTimeScale。
    s.XValues = ToScaledXs(spec.xs, canvas.XMin, canvas.XMax)
    s.Values = spec.ys

    With s.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = spec.FillColor
        .Transparency = spec.FillTransparency
    End With

    ' 轮廓线交给 histoProfile，这里不描边，避免同一位置画两条线。
    s.Format.Line.Visible = msoFalse
End Sub

'=== 副横轴：分两步 ===========================================================
' 拆开是因为旧项目验证过的顺序是「先让副轴显示出来 -> 修复主轴 -> 再配置副轴 -> 最后隐藏」。
' 一次性配置完再隐藏，主横轴刻度有不刷新的历史问题。

'@Description "把副横轴/副纵轴显示出来，为后续配置做准备。必须在任何散点系列修复之前调用。"
Public Sub PrepareAreaAxis(ByVal cht As Excel.Chart)
    ' 面积系列若是全图第一个被加的，AxisGroup=xlSecondary 会被静默忽略。
    ' 此刻所有系列已存在，副轴组建得起来，重申一次即可归位。
    Dim i As Long
    Dim s As Excel.Series
    For i = 1 To cht.SeriesCollection.Count
        Set s = cht.SeriesCollection(i)
        If s.ChartType = xlArea Then s.AxisGroup = xlSecondary
    Next i

    On Error Resume Next
    cht.SetElement msoElementSecondaryCategoryAxisShow
    cht.SetElement msoElementSecondaryValueAxisShow
    On Error GoTo 0
End Sub

'@Description "配置面积图专用副轴：横轴按数值比例定位，纵轴与主轴同范围，然后隐藏两轴。"
Public Sub ConfigureAreaAxis(ByVal cht As Excel.Chart, _
                             ByVal YMin As Double, ByVal YMax As Double)
    With cht.Axes(xlCategory, xlSecondary)
        ' 关键一步。默认的分类轴按「第几个标签」等距排位，
        ' 时间刻度轴才按标签的数值大小比例排位，×10000 的映射因此成立。
        .CategoryType = xlTimeScale
        .MinimumScaleIsAuto = False
        .MinimumScale = AREA_BASE
        .MaximumScaleIsAuto = False
        .MaximumScale = AREA_BASE + AREA_SCALE
    End With

    ' 副纵轴必须显式跟主轴同步：主纵轴还要容纳正态曲线峰值，
    ' 放任它自动缩放会让填色比轮廓矮一截。
    With cht.Axes(xlValue, xlSecondary)
        .MinimumScaleIsAuto = False
        .MinimumScale = YMin
        .MaximumScaleIsAuto = False
        .MaximumScale = YMax
    End With

    On Error Resume Next
    cht.SetElement msoElementSecondaryValueAxisNone
    cht.SetElement msoElementSecondaryCategoryAxisNone
    On Error GoTo 0
End Sub

'=== 内部实现 ================================================================

Private Function ToScaledXs(ByVal xs As Variant, _
                            ByVal XMin As Double, ByVal XMax As Double) As Variant
    Dim lb As Long
    Dim ub As Long
    lb = LBound(xs)
    ub = UBound(xs)

    Dim result() As Double
    ReDim result(1 To ub - lb + 1)

    Dim span As Double
    span = XMax - XMin

    Dim i As Long
    Dim j As Long
    j = 1
    For i = lb To ub
        result(j) = ToScaledX(CDbl(xs(i)), XMin, span)
        j = j + 1
    Next i
    ToScaledXs = result
End Function

' 四舍五入到整数刻度。用 Int(x + 0.5) 而不是 Round，避免 VBA 的银行家舍入
' 在 .5 处产生方向不确定的偏差（虽然量级只有万分之一，没必要引入不确定性）。
Private Function ToScaledX(ByVal X As Double, _
                           ByVal XMin As Double, ByVal span As Double) As Double
    If span <= 0# Then
        ToScaledX = AREA_BASE
    Else
        ToScaledX = AREA_BASE + Int((X - XMin) / span * AREA_SCALE + 0.5)
    End If
End Function

Private Function ToMsoDash(ByVal style As LineDashEnum) As MsoLineDashStyle
    Select Case style
        Case LineDashEnum.dashDash
            ToMsoDash = msoLineDash
        Case LineDashEnum.dashDot
            ToMsoDash = msoLineRoundDot
        Case LineDashEnum.dashLongDash
            ToMsoDash = msoLineLongDash
        Case Else
            ToMsoDash = msoLineSolid
    End Select
End Function

