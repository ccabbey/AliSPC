Attribute VB_Name = "CapabilityDemo"
'@ModuleDescription "能力图演示：单组、多列宽表、逆透视长表"
'@Folder("src.Test")
Option Explicit
' Modified: 2026-11-28 21:48:34

Private Const DEMO_SHEET As String = "AliSPC_Demo"
Private Const CHART_NAME As String = "CapabilityDemoChart"

'@Description "演示单组数据的能力图"
Public Sub Draw_A_Single_Capa_Chart_Demo()
    Dim src As SampleData
    Set src = SampleData.FromRange( _
        ThisWorkbook.Worksheets("DemoData").Range("A1:A1000"), True)

    Dim spec As CapabilitySpec
    Set spec = DemoSpec()
    Dim opt As PlotOptions
    Set opt = DemoOptions("Single Capability Demo", 20)

    Dim ws As Worksheet
    Set ws = DemoSheet()
    Dim co As ChartObject
    Set co = DemoChart(ws)

    CapabilityPlotBuilder.Plot co, src, spec, opt, ws
    co.Activate
End Sub

'@Description "演示多列宽表数据的分组能力图"
Public Sub Draw_Capa_Chart_With_Pivoted_Data_Demo()
    Dim grouped As GroupedSampleData
    Set grouped = GroupedSampleData.FromColumns( _
        ThisWorkbook.Worksheets("DemoData").Range("C1").CurrentRegion, True)

    Dim spec As CapabilitySpec
    Set spec = DemoSpec()
    Dim opt As PlotOptions
    Set opt = DemoOptions("Pivoted Capability Demo")

    Dim ws As Worksheet
    Set ws = DemoSheet()
    Dim co As ChartObject
    Set co = DemoChart(ws)

    CapabilityPlotBuilder.PlotGroupedOverlay co, grouped, spec, opt
    co.Activate
End Sub

'@Description "演示逆透视长表数据的分组能力图"
Public Sub Draw_Capa_Chart_With_Unpivoted_Data_Demo()
    Dim grouped As GroupedSampleData
    Set grouped = GroupedSampleData.FromLongTable( _
        ThisWorkbook.Worksheets("DemoData").Range("F1").CurrentRegion, 1, 2, True)

    Dim spec As CapabilitySpec
    Set spec = DemoSpec()
    Dim opt As PlotOptions
    Set opt = DemoOptions("Unpivoted Capability Demo")

    Dim ws As Worksheet
    Set ws = DemoSheet()
    Dim co As ChartObject
    Set co = DemoChart(ws)

    CapabilityPlotBuilder.PlotGroupedOverlay co, grouped, spec, opt
    co.Activate
End Sub

Private Function DemoSpec() As CapabilitySpec
    Set DemoSpec = CapabilitySpec.Create(-3, 3, 0)
End Function

Private Function DemoOptions(ByVal title As String, _
                             Optional ByVal binCount As Long = 20) As PlotOptions
    Dim opt As PlotOptions
    Set opt = PlotOptions.Create
    opt.title = title
    opt.Decimals = 2
    opt.BinPolicy = binFixedCount
    opt.BinParam = binCount
    opt.YAxisMode = yAxisDensity
    Set DemoOptions = opt
End Function

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

    ws.Shapes.AddChart2 -1, xlXYScatterLinesNoMarkers, _
                        ws.Range("B2").Left, ws.Range("B2").Top, 640, 420
    Set co = ws.ChartObjects(ws.ChartObjects.Count)
    co.Name = CHART_NAME
    Set DemoChart = co
End Function

