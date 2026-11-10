VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Histogram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("src.Model")
'@PredeclaredId
'==============================================================================
' Histogram - 等宽分箱、阶梯多边形、正态拟合曲线（纯计算，零 Excel 引用）
'
' 解耦契约：所有输出都是【数据坐标】，不是百分位、不是像素。
'           面积图所需的 0~1 归一化与 x10000 放大只存在于渲染层的
'           PaintAreaFill 内部，本类完全不知道那件事发生过。
'
' 等宽保证（渲染层的载重假设）：
'   跨度 span 被均分为 k 份，width = span / k，edges(i) = low + i * width。
'   即使调用方用 binFixedWidth 指定宽度，也会反算 k 后重新均分，
'   使 edges 精确落在 [low, high] 上。因此 bin 永远等间距 —— 这是
'   渲染层 x10000 分类轴映射成立的前提，不要改成不等宽分箱。
'
' 逐柱归零多边形（StepPolygon）：
'   每柱的左右边都从 0 轴起画，相邻柱共享零角点。
'   输出为 3k+1 个点，X 非递减、bin 边界处重复。
'   histoProfile 与 histoShade 共用同一组点。
'
' 正态曲线的频数缩放（最容易踩的坑）：
'   直方图纵轴是频数，pdf 积分为 1，直接叠加会贴地。
'   必须乘 N * BinWidth，使曲线面积等于直方图面积 —— 与 Minitab 一致。
'==============================================================================
Option Explicit

Private Const PI As Double = 3.14159265358979
Private Const MAX_BINS As Long = 1000
Private Const DEFAULT_CURVE_POINTS As Long = 200

Private Const ERR_NO_SAMPLE As Long = vbObjectError + 1201
Private Const ERR_POLICY As Long = vbObjectError + 1202
Private Const ERR_PARAM As Long = vbObjectError + 1203
Private Const ERR_SIGMA As Long = vbObjectError + 1204
Private Const ERR_POINTS As Long = vbObjectError + 1205

Public Enum BinPolicyEnum
    binSturges = 0                  ' k = ceil(log2 N) + 1
    binSqrt = 1                     ' k = ceil(sqrt N)
    binFixedCount = 2               ' k = binParam
    binFixedWidth = 3               ' width ~= binParam，反算 k 后重新均分
End Enum

Private Type THistogram
    source As SampleData
    policy As BinPolicyEnum
    BinCount As Long
    BinWidth As Double
    EdgeLow As Double
    EdgeHigh As Double
    Edges As Variant                ' 0..k，共 k+1 个边界
    Counts As Variant               ' 0..k-1，共 k 个频数
    MaxCount As Long
End Type

Private this As THistogram

'=== 工厂 =====================================================================

Public Function FromSample(ByVal source As SampleData, _
                           ByVal policy As BinPolicyEnum, _
                           Optional ByVal BinParam As Double = 0#) As Histogram
    Dim result As Histogram
    Set result = New Histogram
    result.Build source, policy, BinParam
    Set FromSample = result
End Function

'=== 只读属性 =================================================================

Public Property Get BinCount() As Long
    BinCount = this.BinCount
End Property

Public Property Get BinWidth() As Double
    BinWidth = this.BinWidth
End Property

Public Property Get EdgeLow() As Double
    EdgeLow = this.EdgeLow
End Property

Public Property Get EdgeHigh() As Double
    EdgeHigh = this.EdgeHigh
End Property

Public Property Get MaxCount() As Long
    MaxCount = this.MaxCount
End Property

Public Property Get policy() As BinPolicyEnum
    policy = this.policy
End Property

Public Property Get source() As SampleData
    Set source = this.source
End Property

Public Property Get Edges() As Variant
    Edges = this.Edges
End Property

Public Property Get Counts() As Variant
    Counts = this.Counts
End Property

'=== 绘图数据产出 =============================================================

' 逐柱归零多边形：histoProfile 与 histoShade 共用同一组点。
' 输出为 1-based 一维 Double 数组，长度 3k+1：
'   (e0,0) 之后每柱三点 (e_i,c_i) (e_i+1,c_i) (e_i+1,0)。
' 相邻柱共享零角点，故每柱左右边都从 0 轴起画，且不沿 0 轴画底边。
Public Sub StepPolygon(ByRef xs As Variant, ByRef ys As Variant)
    Dim ax() As Double
    Dim ay() As Double
    Dim n As Long
    Dim i As Long
    Dim p As Long
    Dim k As Long

    k = this.BinCount
    n = 3 * k + 1
    ReDim ax(1 To n)
    ReDim ay(1 To n)

    p = 1
    ax(p) = this.Edges(0)
    ay(p) = 0#

    For i = 0 To k - 1
        p = p + 1
        ax(p) = this.Edges(i)
        ay(p) = this.Counts(i)

        p = p + 1
        ax(p) = this.Edges(i + 1)
        ay(p) = this.Counts(i)

        p = p + 1
        ax(p) = this.Edges(i + 1)
        ay(p) = 0#
    Next i

    xs = ax
    ys = ay
End Sub

' 正态拟合曲线，纵轴已按 N * BinWidth 缩放到频数尺度。
' X 跨度为 [EdgeLow - BinWidth, EdgeHigh + BinWidth]，两端各留一个 bin 宽，
' 与 Minitab 的曲线延伸一致；末点精确落在 x1，不受浮点累加漂移影响。
' sigma 由调用方传入：当前只画 overall，传 SampleData.StdevOverall。
Public Sub NormalCurve(ByVal sigma As Double, _
                       ByVal nPoints As Long, _
                       ByRef xs As Variant, _
                       ByRef ys As Variant)
    Dim ax() As Double
    Dim ay() As Double
    Dim x0 As Double
    Dim x1 As Double
    Dim dx As Double
    Dim scale1 As Double
    Dim mu As Double
    Dim i As Long

    If sigma <= 0# Then
        Err.Raise ERR_SIGMA, "Histogram.NormalCurve", _
                  "sigma 必须为正，当前为 " & sigma & "。样本可能全部相同。"
    End If
    If nPoints < 2 Then
        Err.Raise ERR_POINTS, "Histogram.NormalCurve", _
                  "曲线至少需要 2 个点，当前为 " & nPoints & "。"
    End If

    x0 = this.EdgeLow - this.BinWidth
    x1 = this.EdgeHigh + this.BinWidth
    dx = (x1 - x0) / (nPoints - 1)
    mu = this.source.Mean
    scale1 = this.source.Count * this.BinWidth

    ReDim ax(1 To nPoints)
    ReDim ay(1 To nPoints)

    For i = 1 To nPoints
        ax(i) = x0 + (i - 1) * dx
        ay(i) = scale1 * NormPdf(ax(i), mu, sigma)
    Next i
    ax(nPoints) = x1
    ay(nPoints) = scale1 * NormPdf(x1, mu, sigma)

    xs = ax
    ys = ay
End Sub

' 曲线默认点数，供 Builder 取用，避免各处散落魔数。
Public Property Get DefaultCurvePoints() As Long
    DefaultCurvePoints = DEFAULT_CURVE_POINTS
End Property

'=== 内部实现 =================================================================

Public Sub Build(ByVal source As SampleData, _
                  ByVal policy As BinPolicyEnum, _
                  ByVal BinParam As Double)
    Dim lo As Double
    Dim hi As Double
    Dim span As Double
    Dim k As Long
    Dim i As Long
    Dim idx As Long
    Dim v As Double
    Dim e() As Double
    Dim c() As Long

    If source Is Nothing Then
        Err.Raise ERR_NO_SAMPLE, "Histogram.Build", "source 不能为 Nothing。"
    End If

    Set this.source = source
    this.policy = policy

    lo = source.minValue
    hi = source.maxValue

    ' 退化情形：所有值相同，跨度为 0 无法分箱。对称撑开 1 个单位，
    ' 保证图形可画且结果确定（不依赖 mean 的量级）。
    If hi - lo <= 0# Then
        lo = lo - 0.5
        hi = hi + 0.5
    End If
    span = hi - lo

    k = ResolveBinCount(source.Count, span, policy, BinParam)
    If k < 1 Then k = 1
    If k > MAX_BINS Then k = MAX_BINS

    this.BinCount = k
    this.BinWidth = span / k          ' 重新均分：保证严格等宽
    this.EdgeLow = lo
    this.EdgeHigh = hi

    ReDim e(0 To k)
    For i = 0 To k
        e(i) = lo + i * this.BinWidth
    Next i
    e(k) = hi                         ' 末边界钉死，消除浮点漂移
    this.Edges = e

    ReDim c(0 To k - 1)
    For i = 1 To source.Count
        v = source.Values(i)
        idx = Int((v - lo) / this.BinWidth)
        ' 区间约定 [left, right)，末箱闭合；越界值夹到首/末箱
        If idx < 0 Then idx = 0
        If idx >= k Then idx = k - 1
        c(idx) = c(idx) + 1
    Next i
    this.Counts = c

    this.MaxCount = 0
    For i = 0 To k - 1
        If c(i) > this.MaxCount Then this.MaxCount = c(i)
    Next i
End Sub

Private Function ResolveBinCount(ByVal n As Long, _
                                 ByVal span As Double, _
                                 ByVal policy As BinPolicyEnum, _
                                 ByVal BinParam As Double) As Long
    Dim k As Long

    Select Case policy
        Case BinPolicyEnum.binSturges
            k = CeilLng(Log(n) / Log(2#)) + 1

        Case BinPolicyEnum.binSqrt
            k = CeilLng(Sqr(n))

        Case BinPolicyEnum.binFixedCount
            If BinParam < 1# Then
                Err.Raise ERR_PARAM, "Histogram.ResolveBinCount", _
                          "binFixedCount 需要 binParam >= 1，当前为 " & BinParam & "。"
            End If
            k = CLng(BinParam)

        Case BinPolicyEnum.binFixedWidth
            If BinParam <= 0# Then
                Err.Raise ERR_PARAM, "Histogram.ResolveBinCount", _
                          "binFixedWidth 需要 binParam > 0，当前为 " & BinParam & "。"
            End If
            k = CeilLng(span / BinParam)

        Case Else
            Err.Raise ERR_POLICY, "Histogram.ResolveBinCount", _
                      "未知的分箱策略：" & policy & "。"
    End Select

    ResolveBinCount = k
End Function

Private Function NormPdf(ByVal X As Double, ByVal mu As Double, ByVal sd As Double) As Double
    Dim z As Double
    z = (X - mu) / sd
    NormPdf = Exp(-0.5 * z * z) / (sd * Sqr(2# * PI))
End Function

' VBA 没有 Ceiling，用 -Int(-x) 实现；入参为正。
Private Function CeilLng(ByVal X As Double) As Long
    CeilLng = -Int(-X)
End Function
