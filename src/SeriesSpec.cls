VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SeriesSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("src.Spec")
'@PredeclaredId
'==============================================================================
' SeriesSpec - 单个数据系列的绘制意图（哑数据对象，零 Excel 引用）
'
' 解耦契约：Xs / Ys 一律是【数据坐标】。本类不知道百分位、不知道 x10000、
'           不知道副轴分类轴的量化行为。那些全部封装在渲染层 PaintAreaFill。
'
' 为什么用自己的 Enum 而不是 XlAxisGroup / MsoLineDashStyle：
'   那两个是 Excel 类型库的枚举，一旦引用，Spec 层就被绑死在 Excel 上，
'   脱离 Excel 单测和跨项目复用都会破功。这里定义语义枚举，由渲染层
'   用 Select Case 显式映射到 Excel 常量。映射代码只出现在 Painters 里。
'
' RenderAs 只有两个值，对应两个绘图原语：
'   renderXYLine   -> xlXYScatterLines / SmoothNoMarkers，主数值轴
'                     histoProfile、distCurve、mean、Tol、Sigma3x 全走这个
'   renderAreaFill -> xlArea，副分类轴（x10000），仅 histoShade 走这个
'   竖直线不是独立原语，它就是只有 2 个点的 renderXYLine。
'
' 注意：distCurve 必须是 renderXYLine，绝不能用原生折线图 xlLine ——
'       xlLine 的 X 只能来自 Categories（分类槽位序号），Series.XValues
'       对它无效，正态曲线的点会被均匀铺在槽位上而不是落在真实 X 值上。
'==============================================================================
Option Explicit

Private Const ERR_NOT_ARRAY As Long = vbObjectError + 1501
Private Const ERR_NOT_1D As Long = vbObjectError + 1502
Private Const ERR_LENGTH As Long = vbObjectError + 1503
Private Const ERR_TRANSPARENCY As Long = vbObjectError + 1504
Private Const ERR_WEIGHT As Long = vbObjectError + 1505

Public Enum RenderKindEnum
    renderXYLine = 0
    renderAreaFill = 1
End Enum

Public Enum AxisGroupEnum
    axisPrimary = 0
    axisSecondary = 1
End Enum

Public Enum LineDashEnum
    dashSolid = 0
    dashDash = 1
    dashDot = 2
    dashLongDash = 3
End Enum

Private Type TSeriesSpec
    Name As String
    xs As Variant
    ys As Variant
    RenderAs As RenderKindEnum
    AxisGroup As AxisGroupEnum
    Smooth As Boolean
    Visible As Boolean
    LineColor As Long
    LineWeight As Single
    DashStyle As LineDashEnum
    FillColor As Long
    FillTransparency As Single
End Type

Private this As TSeriesSpec

'=== 工厂 =====================================================================

Public Function Create(ByVal Name As String, _
                       ByVal xs As Variant, _
                       ByVal ys As Variant, _
                       ByVal RenderAs As RenderKindEnum, _
                       Optional ByVal AxisGroup As AxisGroupEnum = 0) As SeriesSpec
    Dim result As SeriesSpec
    Set result = New SeriesSpec
    result.Initialize Name, xs, ys, RenderAs, AxisGroup
    Set Create = result
End Function

' 0 = axisPrimary。用字面量而非枚举成员作 Optional 默认值，避免类模块内
' 枚举常量作默认值时的编译歧义。
Public Sub Initialize(ByVal Name As String, _
                       ByVal xs As Variant, _
                       ByVal ys As Variant, _
                       ByVal RenderAs As RenderKindEnum, _
                       ByVal AxisGroup As AxisGroupEnum)
    ValidatePair xs, ys
    this.Name = Name
    this.xs = xs
    this.ys = ys
    this.RenderAs = RenderAs
    this.AxisGroup = AxisGroup
End Sub

Private Sub Class_Initialize()
    this.RenderAs = RenderKindEnum.renderXYLine
    this.AxisGroup = AxisGroupEnum.axisPrimary
    this.Smooth = False
    this.Visible = True
    this.LineColor = RGB(80, 80, 80)
    this.LineWeight = 1.5
    this.DashStyle = LineDashEnum.dashSolid
    this.FillColor = RGB(150, 150, 150)
    this.FillTransparency = 0.35
End Sub

'=== 属性 =====================================================================

Public Property Get Name() As String
    Name = this.Name
End Property

Public Property Let Name(ByVal value As String)
    this.Name = value
End Property

Public Property Get xs() As Variant
    xs = this.xs
End Property

Public Property Let xs(ByVal value As Variant)
    ValidatePair value, this.ys
    this.xs = value
End Property

Public Property Get ys() As Variant
    ys = this.ys
End Property

Public Property Let ys(ByVal value As Variant)
    ValidatePair this.xs, value
    this.ys = value
End Property

Public Property Get RenderAs() As RenderKindEnum
    RenderAs = this.RenderAs
End Property

Public Property Let RenderAs(ByVal value As RenderKindEnum)
    this.RenderAs = value
End Property

Public Property Get AxisGroup() As AxisGroupEnum
    AxisGroup = this.AxisGroup
End Property

Public Property Let AxisGroup(ByVal value As AxisGroupEnum)
    this.AxisGroup = value
End Property

' 仅对 renderXYLine 有意义：True -> xlXYScatterSmoothNoMarkers（曲线用）
'                          False -> xlXYScatterLinesNoMarkers（阶梯/直线用）
Public Property Get Smooth() As Boolean
    Smooth = this.Smooth
End Property

Public Property Let Smooth(ByVal value As Boolean)
    this.Smooth = value
End Property

' False 时渲染层跳过该系列。Sigma3x 关闭时置 False 即可，不必删对象。
Public Property Get Visible() As Boolean
    Visible = this.Visible
End Property

Public Property Let Visible(ByVal value As Boolean)
    this.Visible = value
End Property

Public Property Get LineColor() As Long
    LineColor = this.LineColor
End Property

Public Property Let LineColor(ByVal value As Long)
    this.LineColor = value
End Property

' 单位：磅（point），与 Excel Series.Format.Line.Weight 一致。
Public Property Get LineWeight() As Single
    LineWeight = this.LineWeight
End Property

Public Property Let LineWeight(ByVal value As Single)
    If value <= 0 Then
        Err.Raise ERR_WEIGHT, "SeriesSpec.LineWeight", _
                  "线宽必须为正，当前为 " & value & "。"
    End If
    this.LineWeight = value
End Property

Public Property Get DashStyle() As LineDashEnum
    DashStyle = this.DashStyle
End Property

Public Property Let DashStyle(ByVal value As LineDashEnum)
    this.DashStyle = value
End Property

' 仅对 renderAreaFill 有意义。
Public Property Get FillColor() As Long
    FillColor = this.FillColor
End Property

Public Property Let FillColor(ByVal value As Long)
    this.FillColor = value
End Property

' 0 = 不透明，1 = 全透明。
Public Property Get FillTransparency() As Single
    FillTransparency = this.FillTransparency
End Property

Public Property Let FillTransparency(ByVal value As Single)
    If value < 0 Or value > 1 Then
        Err.Raise ERR_TRANSPARENCY, "SeriesSpec.FillTransparency", _
                  "透明度必须在 0~1 之间，当前为 " & value & "。"
    End If
    this.FillTransparency = value
End Property

Public Property Get PointCount() As Long
    PointCount = ElementCount(this.xs)
End Property

'=== 内部实现 =================================================================

Private Sub ValidatePair(ByVal xs As Variant, ByVal ys As Variant)
    Dim nx As Long
    Dim ny As Long

    ' 允许一侧暂时为空（Property Let 逐个赋值时另一侧可能还没设）
    If IsEmpty(xs) And IsEmpty(ys) Then Exit Sub

    If Not IsEmpty(xs) Then
        If Not IsArray(xs) Then
            Err.Raise ERR_NOT_ARRAY, "SeriesSpec.ValidatePair", "Xs 必须是一维数组。"
        End If
        If HasSecondDimension(xs) Then
            Err.Raise ERR_NOT_1D, "SeriesSpec.ValidatePair", "Xs 必须是一维数组。"
        End If
    End If

    If Not IsEmpty(ys) Then
        If Not IsArray(ys) Then
            Err.Raise ERR_NOT_ARRAY, "SeriesSpec.ValidatePair", "Ys 必须是一维数组。"
        End If
        If HasSecondDimension(ys) Then
            Err.Raise ERR_NOT_1D, "SeriesSpec.ValidatePair", "Ys 必须是一维数组。"
        End If
    End If

    nx = ElementCount(xs)
    ny = ElementCount(ys)

    If Not IsEmpty(xs) And Not IsEmpty(ys) Then
        If nx <> ny Then
            Err.Raise ERR_LENGTH, "SeriesSpec.ValidatePair", _
                      "Xs 与 Ys 长度不一致：" & nx & " vs " & ny & "。"
        End If
        If nx < 1 Then
            Err.Raise ERR_LENGTH, "SeriesSpec.ValidatePair", "系列至少需要 1 个点。"
        End If
    End If
End Sub

Private Function ElementCount(ByVal Arr As Variant) As Long
    Dim lb As Long
    Dim ub As Long

    If IsEmpty(Arr) Then Exit Function
    If Not IsArray(Arr) Then Exit Function

    On Error Resume Next
    lb = LBound(Arr, 1)
    ub = UBound(Arr, 1)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    ElementCount = ub - lb + 1
End Function

Private Function HasSecondDimension(ByVal Arr As Variant) As Boolean
    Dim Probe As Long
    On Error Resume Next
    Probe = UBound(Arr, 2)
    HasSecondDimension = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function
