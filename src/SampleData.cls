VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SampleData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@Folder("src.Model")
'@PredeclaredId
'==============================================================================
' SampleData - 样本数据与描述统计（纯计算，零 Excel 引用）
'
' 解耦契约：本类不声明任何 Excel 对象模型类型（Range / Worksheet / Chart /
'           Application / Workbook 一律不允许），脱离 Excel 仍可编译与单测。
'           数据入口接受 Variant 数组或 Excel Range 对象：后者以 late-bound
'           方式读 Areas / Value2，只依赖 IDispatch，不需要 Excel 引用。
'
' 输入约定：
'   - 一维数组（任意下标基），或
'   - 二维数组但必须是向量（某一维长度为 1），即单行或单列，或
'   - 单区域 Excel Range（单单元格也接受，内部包装成一元数组）。
'   多行多列的块、多区域 Range 会报错，避免静默展平或只取第一块掩盖取数错误。
'   非数值 / 空 / Null / 单元格错误值一律跳过，不计入 N。
'
' sigma 估计：
'   StdevOverall = 样本标准差，分母 N-1
'   StdevWithin  = MRbar / d2，d2 = 1.128（移动极差长度 2）
'   StdevWithin 假定子组大小为 1（单值-移动极差）。子组化数据的 Rbar/d2
'   估计不在本类范围内；需要时新增独立入口，不要在这里加分支。
'
' 用法：
'   Dim s As SampleData
'   Set s = SampleData.FromArray(Sheets("Data").Range("A2:A500"))
'   Set s = SampleData.FromArray(Sheets("Data").Range("A2:A500").Value2)  ' 等价
'   Debug.Print s.Count, s.Mean, s.StdevOverall, s.StdevWithin
'==============================================================================
Option Explicit

Private Const D2_MR As Double = 1.128

Private Const ERR_EMPTY As Long = vbObjectError + 1101
Private Const ERR_SHAPE As Long = vbObjectError + 1102
Private Const ERR_TOO_FEW As Long = vbObjectError + 1103

Private Type TSample
    Values As Variant               ' 1-based 一维 Double 数组
    Count As Long
    Mean As Double
    StdevOverall As Double
    StdevWithin As Double
    minValue As Double
    maxValue As Double
End Type

Private this As TSample

'=== 工厂 =====================================================================

Public Function FromArray(ByVal data As Variant) As SampleData
    Dim result As SampleData
    Set result = New SampleData
    result.Initialize data
    Set FromArray = result
End Function

'=== 只读属性 =================================================================

Public Property Get Count() As Long
    Count = this.Count
End Property

Public Property Get Mean() As Double
    Mean = this.Mean
End Property

Public Property Get StdevOverall() As Double
    StdevOverall = this.StdevOverall
End Property

Public Property Get StdevWithin() As Double
    StdevWithin = this.StdevWithin
End Property

Public Property Get minValue() As Double
    minValue = this.minValue
End Property

Public Property Get maxValue() As Double
    maxValue = this.maxValue
End Property

' 返回内部数组的副本（VBA 数组赋值即拷贝），调用方改动不影响本对象。
Public Property Get Values() As Variant
    Values = this.Values
End Property

'=== 内部实现 =================================================================

Public Sub Initialize(ByVal data As Variant)
    Dim flat() As Double
    Dim dims As Long
    Dim total As Long
    Dim n As Long
    Dim i As Long
    Dim j As Long
    Dim v As Variant
    Dim rows As Long
    Dim cols As Long
    Dim src As Object
    Dim areaCount As Long

    ' Range 入口：late-bound 取数，不引入 Excel 类型。
    If IsObject(data) Then
        Set src = data
        If src Is Nothing Then
            Err.Raise ERR_EMPTY, "SampleData.Initialize", "Range 对象为 Nothing。"
        End If
        On Error Resume Next
        areaCount = src.Areas.Count
        If Err.Number <> 0 Then
            On Error GoTo 0
            Err.Raise ERR_SHAPE, "SampleData.Initialize", _
                      "对象不是 Excel Range（读 Areas 失败）。"
        End If
        If areaCount > 1 Then
            On Error GoTo 0
            Err.Raise ERR_SHAPE, "SampleData.Initialize", _
                      "不支持多区域 Range（" & areaCount & " 块），请逐块传入。"
        End If
        data = src.Value2
        If Err.Number <> 0 Then
            On Error GoTo 0
            Err.Raise ERR_SHAPE, "SampleData.Initialize", "读取 Range.Value2 失败。"
        End If
        On Error GoTo 0
        ' 单单元格 Range 的 Value2 是标量，包装成一元数组走统一路径
        If Not IsArray(data) Then data = Array(data)
    End If

    If IsEmpty(data) Then
        Err.Raise ERR_EMPTY, "SampleData.Initialize", "数据为 Empty。"
    End If
    If IsNull(data) Then
        Err.Raise ERR_EMPTY, "SampleData.Initialize", "数据为 Null。"
    End If
    If Not IsArray(data) Then
        Err.Raise ERR_SHAPE, "SampleData.Initialize", _
                  "需要数组或 Range。单个数值请用 Array(x) 包装。"
    End If

    dims = ArrayDimensions(data)
    If dims = 0 Then
        Err.Raise ERR_EMPTY, "SampleData.Initialize", "数组未分配（无元素）。"
    End If

    If dims = 1 Then
        total = UBound(data, 1) - LBound(data, 1) + 1
    ElseIf dims = 2 Then
        rows = UBound(data, 1) - LBound(data, 1) + 1
        cols = UBound(data, 2) - LBound(data, 2) + 1
        If rows <> 1 And cols <> 1 Then
            Err.Raise ERR_SHAPE, "SampleData.Initialize", _
                      "二维数组必须是单行或单列向量，当前为 " & _
                      rows & " 行 x " & cols & " 列。"
        End If
        total = rows * cols
    Else
        Err.Raise ERR_SHAPE, "SampleData.Initialize", _
                  "只支持一维或二维数组，当前维数为 " & dims & "。"
    End If

    If total < 1 Then
        Err.Raise ERR_EMPTY, "SampleData.Initialize", "数组没有元素。"
    End If

    ReDim flat(1 To total)
    n = 0

    If dims = 1 Then
        For i = LBound(data, 1) To UBound(data, 1)
            v = data(i)
            If IsAcceptable(v) Then
                n = n + 1
                flat(n) = CDbl(v)
            End If
        Next i
    Else
        For i = LBound(data, 1) To UBound(data, 1)
            For j = LBound(data, 2) To UBound(data, 2)
                v = data(i, j)
                If IsAcceptable(v) Then
                    n = n + 1
                    flat(n) = CDbl(v)
                End If
            Next j
        Next i
    End If

    If n < 2 Then
        Err.Raise ERR_TOO_FEW, "SampleData.Initialize", _
                  "有效数值只有 " & n & " 个，至少需要 2 个才能估计标准差。"
    End If

    If n < total Then ReDim Preserve flat(1 To n)

    this.Values = flat
    this.Count = n
    ComputeStats
End Sub

Private Sub ComputeStats()
    Dim i As Long
    Dim sum As Double
    Dim dev As Double
    Dim sumSq As Double
    Dim mrSum As Double
    Dim lo As Double
    Dim hi As Double

    sum = 0#
    lo = this.Values(1)
    hi = this.Values(1)

    For i = 1 To this.Count
        sum = sum + this.Values(i)
        If this.Values(i) < lo Then lo = this.Values(i)
        If this.Values(i) > hi Then hi = this.Values(i)
    Next i

    this.Mean = sum / this.Count
    this.minValue = lo
    this.maxValue = hi

    ' 两遍法算离差平方和，比 E(x^2)-E(x)^2 数值稳定得多
    sumSq = 0#
    For i = 1 To this.Count
        dev = this.Values(i) - this.Mean
        sumSq = sumSq + dev * dev
    Next i
    this.StdevOverall = Sqr(sumSq / (this.Count - 1))

    ' 移动极差：MR_i = |x_i - x_(i-1)|，共 N-1 个
    mrSum = 0#
    For i = 2 To this.Count
        mrSum = mrSum + Abs(this.Values(i) - this.Values(i - 1))
    Next i
    this.StdevWithin = (mrSum / (this.Count - 1)) / D2_MR
End Sub

Private Function IsAcceptable(ByVal v As Variant) As Boolean
    ' 顺序要紧：IsNumeric(Null) 返回 Null，直接放进 If 会触发错误 94
    If IsError(v) Then Exit Function
    If IsNull(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If IsObject(v) Then Exit Function
    IsAcceptable = IsNumeric(v)
End Function

Private Function ArrayDimensions(ByVal Arr As Variant) As Long
    Dim d As Long
    Dim Probe As Long

    On Error Resume Next
    For d = 1 To 60
        Probe = UBound(Arr, d)
        If Err.Number <> 0 Then
            Err.Clear
            Exit For
        End If
    Next d
    On Error GoTo 0

    ArrayDimensions = d - 1
End Function




