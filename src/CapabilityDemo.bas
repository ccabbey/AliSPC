Attribute VB_Name = "CapabilityDemo"
'@ModuleDescription "硬编码数据的冒烟测试入口。不读工作表，直接在 Excel 2016 里画出能力图。"
'@Folder("src.Test")
Option Explicit

Private Const DEMO_SHEET As String = "AliSPC_Demo"
Private Const CHART_NAME As String = "CapabilityDemoChart"
Private Const EXPECTED_SERIES As Long = 9      ' Shade+Profile+Curve+Mean+LSL+USL+Target+2 条 Sigma

'@Description "主入口：在 AliSPC_Demo 表上画一张完整能力图。可反复运行，会清场重绘。"
Public Sub DrawCapabilityDemo()
    Dim src As SampleData
    Set src = SampleData.FromArray(Sheets("Sheet3").Range("A1:A196"))

    Dim spec As CapabilitySpec
    Set spec = CapabilitySpec.Create(38.5, 39.4, 38.95)

    Dim opt As PlotOptions
    Set opt = PlotOptions.Create
    opt.BinPolicy = binFixedWidth
    opt.BinParam = 0.065
    opt.Title = "Process Capability - Demo"

    Dim ws As Worksheet
    Set ws = DemoSheet()

    Dim co As ChartObject
    Set co = DemoChart(ws)

    CapabilityPlotBuilder.Plot co, src, spec, opt
    co.Activate
End Sub

'@Description "只跑 Model + Spec + BuildSpec，不碰图表。出图有问题时先用这个定位是数据侧还是渲染侧。"
Public Sub CheckCapabilityDemo()
    Dim src As SampleData
    Set src = SampleData.FromArray(DemoValues())

    Dim spec As CapabilitySpec
    Set spec = CapabilitySpec.Create(9#, 11#, 10#)

    Dim opt As PlotOptions
    Set opt = PlotOptions.Create
    opt.BinPolicy = binFixedCount
    opt.BinParam = 10

    Dim ps As CapabilityPlotSpec
    Set ps = CapabilityPlotBuilder.BuildSpec(src, spec, opt)

    Dim cap As CapabilityResult
    Set cap = CapabilityResult.FromSample(src, spec)

    Dim msg As String
    msg = "N = " & CStr(src.Count) & vbCrLf
    msg = msg & "Mean = " & Format$(src.Mean, "0.0000") & vbCrLf
    msg = msg & "StdevOverall = " & Format$(src.StdevOverall, "0.0000") & vbCrLf
    msg = msg & "StdevWithin = " & Format$(src.StdevWithin, "0.0000") & vbCrLf
    msg = msg & "Cp = " & ShowOpt(cap.Cp) & "   Cpk = " & ShowOpt(cap.Cpk) & vbCrLf
    msg = msg & "Pp = " & ShowOpt(cap.Pp) & "   Ppk = " & ShowOpt(cap.Ppk) & vbCrLf
    msg = msg & "Cpm = " & ShowOpt(cap.Cpm) & vbCrLf
    msg = msg & "XAxis = [" & Format$(ps.xAxis.Min, "0.000") & ", " & Format$(ps.xAxis.Max, "0.000") & "]" & vbCrLf
    msg = msg & "YAxis = [" & Format$(ps.yAxis.Min, "0.000") & ", " & Format$(ps.yAxis.Max, "0.000") & "]" & vbCrLf
    msg = msg & vbCrLf & "z-order:" & vbCrLf

    Dim s As SeriesSpec
    Dim i As Long
    For Each s In ps.AllSeries()
        i = i + 1
        msg = msg & "  " & i & ". " & s.Name & "  (" & s.PointCount & " pts)" & vbCrLf
    Next s

    If ps.SeriesCount = EXPECTED_SERIES Then
        msg = msg & vbCrLf & "PASS: SeriesCount = " & ps.SeriesCount
    Else
        msg = msg & vbCrLf & "FAIL: SeriesCount = " & ps.SeriesCount & _
              "，期望 " & EXPECTED_SERIES
    End If

    Debug.Print msg
    MsgBox msg, vbInformation, "AliSPC BuildSpec 冒烟测试"
End Sub

'=== 内部实现 ================================================================

Private Function DemoSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(DEMO_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
                 After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = DEMO_SHEET
    End If
    Set DemoSheet = ws
End Function

Private Function DemoChart(ByVal ws As Worksheet) As ChartObject
    Dim co As ChartObject
    For Each co In ws.ChartObjects
        If co.Name = CHART_NAME Then
            Set DemoChart = co
            Exit Function
        End If
    Next co

    ' 用 AddChart2 直接建成散点图，沿用旧项目验证过的做法；
    ' 新建的 ChartObject 一定是集合里的最后一个。
    ws.Shapes.AddChart2 -1, xlXYScatterLinesNoMarkers, _
                        ws.Range("B2").Left, ws.Range("B2").Top, 640, 420
    Set co = ws.ChartObjects(ws.ChartObjects.Count)
    co.Name = CHART_NAME
    Set DemoChart = co
End Function

Private Function ShowOpt(ByVal value As Variant) As String
    If IsEmpty(value) Then
        ShowOpt = "N/A"
    Else
        ShowOpt = Format$(CDbl(value), "0.000")
    End If
End Function

' 硬编码样本：60 个测量值，均值约 10.00，样本标准差约 0.56，
' 规格 LSL=9 / USL=11 / Target=10，会有少量超差点，便于观察图形。
Private Function DemoValues() As Variant
    Dim v(1 To 60) As Double
    v(1) = 9.62: v(2) = 10.53: v(3) = 9.615: v(4) = 9.643
    v(5) = 10.417: v(6) = 10.921: v(7) = 11.14: v(8) = 8.887
    v(9) = 10.097: v(10) = 10.333: v(11) = 10.293: v(12) = 9.955
    v(13) = 9.443: v(14) = 9.218: v(15) = 9.636: v(16) = 11.062
    v(17) = 9.653: v(18) = 10.367: v(19) = 9.542: v(20) = 9.752
    v(21) = 9.593: v(22) = 8.732: v(23) = 10.295: v(24) = 10.714
    v(25) = 10.523: v(26) = 10.812: v(27) = 10.292: v(28) = 10.324
    v(29) = 9.429: v(30) = 9.367: v(31) = 9.385: v(32) = 10.154
    v(33) = 10.475: v(34) = 11.158: v(35) = 10.245: v(36) = 10.135
    v(37) = 9.421: v(38) = 9.505: v(39) = 10.359: v(40) = 10.14
    v(41) = 10.438: v(42) = 9.084: v(43) = 10.663: v(44) = 9.804
    v(45) = 9.577: v(46) = 9.371: v(47) = 9.71: v(48) = 9.242
    v(49) = 10.029: v(50) = 9.419: v(51) = 10.459: v(52) = 9.703
    v(53) = 10.152: v(54) = 10.346: v(55) = 10.08: v(56) = 9.871
    v(57) = 10.242: v(58) = 9.938: v(59) = 9.846: v(60) = 10.94
    DemoValues = v
End Function

