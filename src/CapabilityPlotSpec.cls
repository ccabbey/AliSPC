VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CapabilityPlotSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'@Folder("src.Spec")
'==============================================================================
' CapabilityPlotSpec - Minitab 风格过程能力图的完整绘制意图
'                      （解耦边界，零 Excel 引用，全部数据坐标）
'
' 这是整个架构的契约面：
'   左侧（src.Model）只产出本对象，不知道 Excel 存在；
'   右侧（src.Render）只消费本对象，不知道统计学存在。
'
' 六个语义系列 -> 两个绘图原语 -> 8~10 个真实 Excel Series：
'   HistoShade    areaFill     副轴   1 个（已定：不做超差着色）
'   HistoProfile  xyLine       主轴   1 个
'   DistCurve     xyLine       主轴   1 个（已定：只画 overall）
'   MeanLine      xyLine(2点)  主轴   1 个
'   TolLines      xyLine(2点)  主轴   2~3 个（LSL / USL / 可选 Target）
'   SigmaLines    xyLine(2点)  主轴   0 或 2 个（Mean ± 3*sigma）
'
' 语义层用具名属性而不是泛型集合 + 字符串键：编译期可查、IntelliSense 可见、
' 拼错系列名会直接编译失败而不是运行时静默画错。
' TolLines / SigmaLines 用 Collection，因为它们的条数本来就是可变的。
'
' AllSeries() 的返回顺序就是 z-order（Excel 按系列索引顺序绘制，先加的在下）：
'   HistoShade 必须第一个，否则填色会盖住轮廓线和曲线。
'
' 后续若要加超差红色填充：把 HistoShade 从 SeriesSpec 改成 Collection，
' AllSeries 里多展平两项即可，渲染层零改动。这是本设计预留的唯一扩展点。
'==============================================================================
Option Explicit

Private Type TCapabilityPlotSpec
    Title As String
    HistoShade As SeriesSpec
    HistoProfile As SeriesSpec
    DistCurve As SeriesSpec
    MeanLine As SeriesSpec
    TolLines As Collection
    SigmaLines As Collection
    xAxis As AxisSpec
    yAxis As AxisSpec
    StatBox As AnnotationSpec
End Type

Private this As TCapabilityPlotSpec

Private Sub Class_Initialize()
    this.Title = ""
    Set this.TolLines = New Collection
    Set this.SigmaLines = New Collection
    Set this.xAxis = AxisSpec.Create
    Set this.yAxis = AxisSpec.Create(0)      ' 频数轴下限钉在 0
End Sub

'=== 图表级 ===================================================================

Public Property Get Title() As String
    Title = this.Title
End Property

Public Property Let Title(ByVal value As String)
    this.Title = value
End Property

Public Property Get xAxis() As AxisSpec
    Set xAxis = this.xAxis
End Property

Public Property Set xAxis(ByVal value As AxisSpec)
    Set this.xAxis = value
End Property

Public Property Get yAxis() As AxisSpec
    Set yAxis = this.yAxis
End Property

Public Property Set yAxis(ByVal value As AxisSpec)
    Set this.yAxis = value
End Property

Public Property Get StatBox() As AnnotationSpec
    Set StatBox = this.StatBox
End Property

Public Property Set StatBox(ByVal value As AnnotationSpec)
    Set this.StatBox = value
End Property

'=== 六个语义系列 =============================================================

Public Property Get HistoShade() As SeriesSpec
    Set HistoShade = this.HistoShade
End Property

Public Property Set HistoShade(ByVal value As SeriesSpec)
    Set this.HistoShade = value
End Property

Public Property Get HistoProfile() As SeriesSpec
    Set HistoProfile = this.HistoProfile
End Property

Public Property Set HistoProfile(ByVal value As SeriesSpec)
    Set this.HistoProfile = value
End Property

Public Property Get DistCurve() As SeriesSpec
    Set DistCurve = this.DistCurve
End Property

Public Property Set DistCurve(ByVal value As SeriesSpec)
    Set this.DistCurve = value
End Property

Public Property Get MeanLine() As SeriesSpec
    Set MeanLine = this.MeanLine
End Property

Public Property Set MeanLine(ByVal value As SeriesSpec)
    Set this.MeanLine = value
End Property

' 元素类型必须是 SeriesSpec。Builder 直接 .TolLines.Add spec 即可。
Public Property Get TolLines() As Collection
    Set TolLines = this.TolLines
End Property

Public Property Set TolLines(ByVal value As Collection)
    If value Is Nothing Then
        Set this.TolLines = New Collection
    Else
        Set this.TolLines = value
    End If
End Property

Public Property Get SigmaLines() As Collection
    Set SigmaLines = this.SigmaLines
End Property

Public Property Set SigmaLines(ByVal value As Collection)
    If value Is Nothing Then
        Set this.SigmaLines = New Collection
    Else
        Set this.SigmaLines = value
    End If
End Property

'=== 展平 =====================================================================

' 按 z-order 展平成渲染层唯一需要认识的集合。
' 跳过 Nothing（未装配的可选系列），但保留 Visible = False 的项，
' 由渲染层决定跳过 —— 这样"装配了什么"和"画不画"是两个独立的决定。
Public Function AllSeries() As Collection
    Dim result As Collection
    Dim item As SeriesSpec

    Set result = New Collection

    AppendIfPresent result, this.HistoShade
    AppendIfPresent result, this.HistoProfile
    AppendIfPresent result, this.DistCurve
    AppendIfPresent result, this.MeanLine

    For Each item In this.TolLines
        AppendIfPresent result, item
    Next item

    For Each item In this.SigmaLines
        AppendIfPresent result, item
    Next item

    Set AllSeries = result
End Function

Public Property Get SeriesCount() As Long
    SeriesCount = AllSeries().Count
End Property

'=== 内部实现 =================================================================

Private Sub AppendIfPresent(ByVal target As Collection, ByVal spec As SeriesSpec)
    If Not spec Is Nothing Then target.Add spec
End Sub
