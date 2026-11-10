VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ChartCanvas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "Rubberduck" ,"Predeclared Class Module"
'@ModuleDescription "把一个 ChartObject 包装成画布：清场、加系列、配轴、加标注。只认识 Spec，不认识统计。"
'@PredeclaredId
'@Folder("src.Render")
Option Explicit

Private Const ERR_NO_TARGET As Long = vbObjectError + 1801
Private Const ERR_NO_AXISPEC As Long = vbObjectError + 1802
Private Const ERR_AUTO_SCALE As Long = vbObjectError + 1803
Private Const ERR_AXIS_ORDER As Long = vbObjectError + 1804
Private Const ERR_BAD_RENDER As Long = vbObjectError + 1805

Private Const ANNO_PREFIX As String = "AliSPC_Anno_"

Private Type TCanvas
    target As ChartObject
    xAxis As AxisSpec
    yAxis As AxisSpec
    XMin As Double
    XMax As Double
    YMin As Double
    YMax As Double
    HasArea As Boolean
End Type

Private this As TCanvas

'=== 工厂 ====================================================================
' 轴范围必须在构造时就给定：面积图的 ×10000 映射需要知道主轴的确定上下限，
' 自动缩放的主轴无法参与映射，因此这里直接拒绝 HasMin/HasMax 为假的 AxisSpec。

Public Function FromChartObject(ByVal target As ChartObject, _
                                ByVal xAxis As AxisSpec, _
                                ByVal yAxis As AxisSpec) As ChartCanvas
    Dim result As ChartCanvas
    Set result = New ChartCanvas
    result.Initialize target, xAxis, yAxis
    Set FromChartObject = result
End Function

Public Sub Initialize(ByVal target As ChartObject, _
                       ByVal xAxis As AxisSpec, ByVal yAxis As AxisSpec)
    If target Is Nothing Then
        Err.Raise ERR_NO_TARGET, "ChartCanvas.Initialize", "target 不能为 Nothing。"
    End If
    If xAxis Is Nothing Or yAxis Is Nothing Then
        Err.Raise ERR_NO_AXISPEC, "ChartCanvas.Initialize", "XAxis / YAxis 不能为 Nothing。"
    End If
    If Not xAxis.HasMin Or Not xAxis.HasMax Then
        Err.Raise ERR_AUTO_SCALE, "ChartCanvas.Initialize", _
                  "XAxis 必须显式指定 Min 和 Max：面积图映射不支持自动缩放的主横轴。"
    End If
    If Not yAxis.HasMin Or Not yAxis.HasMax Then
        Err.Raise ERR_AUTO_SCALE, "ChartCanvas.Initialize", _
                  "YAxis 必须显式指定 Min 和 Max：副纵轴要与主纵轴强制同范围。"
    End If

    Set this.target = target
    Set this.xAxis = xAxis
    Set this.yAxis = yAxis
    this.XMin = CDbl(xAxis.Min)
    this.XMax = CDbl(xAxis.Max)
    this.YMin = CDbl(yAxis.Min)
    this.YMax = CDbl(yAxis.Max)
    this.HasArea = False

    If this.XMax <= this.XMin Then
        Err.Raise ERR_AXIS_ORDER, "ChartCanvas.Initialize", "XAxis.Max 必须大于 Min。"
    End If
    If this.YMax <= this.YMin Then
        Err.Raise ERR_AXIS_ORDER, "ChartCanvas.Initialize", "YAxis.Max 必须大于 Min。"
    End If
End Sub

'=== 只读视图 ================================================================

Public Property Get Chart() As Excel.Chart
    Set Chart = this.target.Chart
End Property

Public Property Get XMin() As Double
    XMin = this.XMin
End Property

Public Property Get XMax() As Double
    XMax = this.XMax
End Property

Public Property Get YMin() As Double
    YMin = this.YMin
End Property

Public Property Get YMax() As Double
    YMax = this.YMax
End Property

'=== 清场 ====================================================================

'@Description "删掉图里所有系列和本组件加过的标注，保留 ChartObject 本身以便反复重绘。"
Public Sub ClearCanvas()
    Dim cht As Excel.Chart
    Set cht = Me.Chart

    Do While cht.SeriesCollection.Count > 0
        cht.SeriesCollection(1).Delete
    Loop

    Dim i As Long
    For i = cht.Shapes.Count To 1 Step -1
        If Left$(cht.Shapes(i).Name, Len(ANNO_PREFIX)) = ANNO_PREFIX Then
            cht.Shapes(i).Delete
        End If
    Next i

    this.HasArea = False
End Sub

'=== 加系列 ==================================================================

'@Description "按 RenderAs 分派到绘制原语，返回系列序号；被跳过时返回 0。"
Public Function AddSeries(ByVal spec As SeriesSpec) As Long
    AddSeries = 0
    If spec Is Nothing Then Exit Function

    ' Excel 的 Series 没有可见性开关，隐藏就是「不添加」。
    ' 这样也不会污染图例和自动缩放。
    If Not spec.Visible Then Exit Function

    Select Case spec.RenderAs
        Case RenderKindEnum.renderXYLine
            Painters.PaintXYLine Me, spec
        Case RenderKindEnum.renderAreaFill
            Painters.PaintAreaFill Me, spec
            this.HasArea = True
        Case Else
            Err.Raise ERR_BAD_RENDER, "ChartCanvas.AddSeries", _
                      "未知的 RenderAs：" & CStr(spec.RenderAs) & "。"
    End Select

    AddSeries = Me.Chart.SeriesCollection.Count
End Function

'=== 配轴 ====================================================================

'@Description "写入主横轴/主纵轴范围与外观，并按需配置面积图副轴。必须在所有系列加完之后调用。"
Public Sub ApplyAxes()
    Dim cht As Excel.Chart
    Set cht = Me.Chart

    If this.HasArea Then Painters.PrepareAreaAxis cht
    Me.RepairPrimaryAxis
    ApplyPrimaryCategory cht
    ApplyPrimaryValue cht
    If this.HasArea Then Painters.ConfigureAreaAxis cht, this.YMin, this.YMax
End Sub

Private Sub ApplyPrimaryCategory(ByVal cht As Excel.Chart)
    With cht.Axes(xlCategory, xlPrimary)
        .MinimumScaleIsAuto = False
        .MinimumScale = this.XMin
        .MaximumScaleIsAuto = False
        .MaximumScale = this.XMax

        If this.xAxis.HasMajorUnit Then
            .MajorUnitIsAuto = False
            .MajorUnit = CDbl(this.xAxis.MajorUnit)
        End If

        .HasTitle = this.xAxis.TitleVisible
        If this.xAxis.TitleVisible Then .AxisTitle.text = this.xAxis.Title

        .TickLabels.numberFormat = this.xAxis.numberFormat
        .TickLabels.Font.Name = "Arial"
        .TickLabels.Font.Size = 9

        If this.xAxis.Visible Then
            .Format.Line.Visible = msoTrue
            .Format.Line.ForeColor.RGB = RGB(0, 0, 0)
            .TickLabelPosition = xlTickLabelPositionNextToAxis
        Else
            .Format.Line.Visible = msoFalse
            .TickLabelPosition = xlTickLabelPositionNone
        End If
    End With
End Sub

Private Sub ApplyPrimaryValue(ByVal cht As Excel.Chart)
    With cht.Axes(xlValue, xlPrimary)
        .MinimumScaleIsAuto = False
        .MinimumScale = this.YMin
        .MaximumScaleIsAuto = False
        .MaximumScale = this.YMax

        If this.yAxis.HasMajorUnit Then
            .MajorUnitIsAuto = False
            .MajorUnit = CDbl(this.yAxis.MajorUnit)
        End If

        .HasTitle = this.yAxis.TitleVisible
        If this.yAxis.TitleVisible Then .AxisTitle.text = this.yAxis.Title

        .TickLabels.numberFormat = this.yAxis.numberFormat
        .TickLabels.Font.Name = "Arial"
        .TickLabels.Font.Size = 9

        If this.yAxis.Visible Then
            .Format.Line.Visible = msoTrue
            .Format.Line.ForeColor.RGB = RGB(0, 0, 0)
            .TickLabelPosition = xlTickLabelPositionNextToAxis
        Else
            .Format.Line.Visible = msoFalse
            .TickLabelPosition = xlTickLabelPositionNone
        End If
    End With
End Sub

'@Description "经验性修复：散点 + 面积组合图刚建好时主横轴刻度可能不刷新，把第一个散点系列改成平滑再改回原类型即可强制重算。沿用旧项目验证过的做法。"
Public Sub RepairPrimaryAxis()
    Dim cht As Excel.Chart
    Set cht = Me.Chart

    Dim i As Long
    Dim original As XlChartType
    For i = 1 To cht.SeriesCollection.Count
        original = cht.SeriesCollection(i).ChartType
        If IsScatterType(original) Then
            cht.SeriesCollection(i).ChartType = xlXYScatterSmoothNoMarkers
            cht.SeriesCollection(i).ChartType = original
            Exit For
        End If
    Next i
End Sub

Private Function IsScatterType(ByVal t As XlChartType) As Boolean
    Select Case t
        Case xlXYScatter, xlXYScatterLines, xlXYScatterLinesNoMarkers, _
             xlXYScatterSmooth, xlXYScatterSmoothNoMarkers
            IsScatterType = True
    End Select
End Function

'=== 图表外观 ================================================================

'@Description "标题、图例、网格线、绘图区边框。与轴刻度无关的那部分外观。"
Public Sub FormatFrame(ByVal Title As String)
    Dim cht As Excel.Chart
    Set cht = Me.Chart

    If Len(Title) > 0 Then
        cht.HasTitle = True
        cht.ChartTitle.text = Title
        cht.ChartTitle.Font.Name = "Arial"
        cht.ChartTitle.Font.Size = 12
    Else
        cht.HasTitle = False
    End If

    cht.HasLegend = False

    On Error Resume Next
    cht.SetElement msoElementPrimaryValueGridLinesNone
    cht.SetElement msoElementPrimaryCategoryGridLinesNone
    On Error GoTo 0

    With cht.PlotArea.Format
        .Fill.Visible = msoTrue
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = RGB(0, 0, 0)
        .Line.Weight = 0.75
    End With

    With cht.ChartArea.Format
        .Fill.Visible = msoTrue
        .Line.Visible = msoFalse
    End With
End Sub

'=== 标注 ====================================================================

'@Description "按数据坐标放一个无边框无填充的文本框。调用前轴范围必须已经写入图表。"
Public Sub AddAnnotation(ByVal spec As AnnotationSpec)
    If spec Is Nothing Then Exit Sub
    If Not spec.Visible Then Exit Sub
    If Len(spec.text) = 0 Then Exit Sub

    Dim cht As Excel.Chart
    Set cht = Me.Chart

    ' 强制 Excel 完成一次布局，否则 PlotArea.Inside* 可能还是 0。
    DoEvents

    Dim shp As Shape
    Set shp = cht.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 10, 10)
    shp.Name = ANNO_PREFIX & CStr(cht.Shapes.Count)

    With shp.TextFrame2
        .WordWrap = msoFalse
        .AutoSize = msoAutoSizeShapeToFitText
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        .MarginBottom = 0
        .VerticalAnchor = msoAnchorTop
        .HorizontalAnchor = ToMsoAnchor(spec.align)

        With .TextRange
            ' Spec 里用 vbCrLf 计行，Excel 文本框只认 vbLf。
            .text = Replace(spec.text, vbCrLf, vbLf)
            .Font.Name = spec.fontName
            .Font.Size = spec.fontSize
            .Font.Fill.ForeColor.RGB = spec.FontColor
            .ParagraphFormat.Alignment = ToMsoParagraphAlign(spec.align)
        End With
    End With

    shp.Fill.Visible = msoFalse
    shp.Line.Visible = msoFalse

    DoEvents      ' 等 AutoSize 生效后再读 Width，否则右对齐会偏

    Dim px As Single
    Dim py As Single
    px = Me.DataToPixelX(spec.X)
    py = Me.DataToPixelY(spec.Y)

    Select Case spec.align
        Case TextAlignEnum.alignCenter
            shp.Left = px - shp.Width / 2
        Case TextAlignEnum.alignRight
            shp.Left = px - shp.Width
        Case Else
            shp.Left = px
    End Select
    shp.Top = py
End Sub

'@Description "数据横坐标 -> 图表内像素（磅）。预留给以后改用 Shape 填充的方案。"
Public Function DataToPixelX(ByVal X As Double) As Single
    Dim pa As PlotArea
    Set pa = Me.Chart.PlotArea
    DataToPixelX = CSng(pa.InsideLeft + (X - this.XMin) / (this.XMax - this.XMin) * pa.InsideWidth)
End Function

'@Description "数据纵坐标 -> 图表内像素（磅），向下为正。"
Public Function DataToPixelY(ByVal Y As Double) As Single
    Dim pa As PlotArea
    Set pa = Me.Chart.PlotArea
    DataToPixelY = CSng(pa.InsideTop + (this.YMax - Y) / (this.YMax - this.YMin) * pa.InsideHeight)
End Function

Private Function ToMsoAnchor(ByVal a As TextAlignEnum) As MsoHorizontalAnchor
    Select Case a
        Case TextAlignEnum.alignCenter
            ToMsoAnchor = msoAnchorCenter
        Case TextAlignEnum.alignRight
            ToMsoAnchor = msoAnchorCenter
        Case Else
            ToMsoAnchor = msoAnchorCenter
    End Select
End Function

Private Function ToMsoParagraphAlign(ByVal a As TextAlignEnum) As MsoParagraphAlignment
    Select Case a
        Case TextAlignEnum.alignCenter
            ToMsoParagraphAlign = msoAlignCenter
        Case TextAlignEnum.alignRight
            ToMsoParagraphAlign = msoAlignRight
        Case Else
            ToMsoParagraphAlign = msoAlignLeft
    End Select
End Function



