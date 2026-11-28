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
' Modified: 2026-11-28 22:41:35
Option Explicit

Private Const D2_MR As Double = 1.128

Private Const ERR_EMPTY As Long = vbObjectError + 1101
Private Const ERR_SHAPE As Long = vbObjectError + 1102
Private Const ERR_TOO_FEW As Long = vbObjectError + 1103

Private Type TSample
    values As Variant
    Count As Long
    Mean As Double
    StdevOverall As Double
    StdevWithin As Double
    minValue As Double
    maxValue As Double
    header As String
End Type

Private this As TSample

Public Function FromRange(ByVal source As Object, _
                           Optional ByVal hasHeader As Boolean = False, _
                           Optional ByVal header As String = "") As SampleData
    Dim data As Variant
    data = ReadRange(source)
    Set FromRange = FromArray(data, hasHeader, header)
End Function

Public Function FromArray(ByVal data As Variant, _
                          Optional ByVal hasHeader As Boolean = False, _
                          Optional ByVal header As String = "") As SampleData
    Dim result As SampleData
    Set result = New SampleData
    result.Initialize data, hasHeader, header
    Set FromArray = result
End Function

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

Public Property Get values() As Variant
    values = this.values
End Property

Public Property Get header() As String
    header = this.header
End Property

Public Sub Initialize(ByVal data As Variant, _
                       Optional ByVal hasHeader As Boolean = False, _
                       Optional ByVal header As String = "")
    Dim flat() As Double
    Dim dims As Long
    Dim total As Long
    Dim n As Long
    Dim i As Long
    Dim v As Variant
    Dim rows As Long
    Dim cols As Long
    Dim firstRow As Long
    Dim firstCol As Long
    Dim skipFirst As Boolean

    If IsObject(data) Then
        Err.Raise ERR_SHAPE, "SampleData.Initialize", _
                  "FromArray 只接受数组；Range 请使用 FromRange。"
    End If
    If IsEmpty(data) Or IsNull(data) Then
        Err.Raise ERR_EMPTY, "SampleData.Initialize", "数据为 Empty 或 Null。"
    End If
    If Not IsArray(data) Then
        Err.Raise ERR_SHAPE, "SampleData.Initialize", _
                  "FromArray 只接受一维数组或单行/单列二维数组。"
    End If

    dims = ArrayDimensions(data)
    If dims = 0 Then
        Err.Raise ERR_EMPTY, "SampleData.Initialize", "数据为未分配数组。"
    End If

    firstRow = LBound(data, 1)
    If dims = 1 Then
        total = UBound(data, 1) - firstRow + 1
    ElseIf dims = 2 Then
        firstCol = LBound(data, 2)
        rows = UBound(data, 1) - firstRow + 1
        cols = UBound(data, 2) - firstCol + 1
        If rows <> 1 And cols <> 1 Then
            Err.Raise ERR_SHAPE, "SampleData.Initialize", _
                      "单组数据只能是单行或单列，不能是多行多列矩阵。"
        End If
        total = rows * cols
    Else
        Err.Raise ERR_SHAPE, "SampleData.Initialize", _
                  "只支持一维数组或单行/单列二维数组。"
    End If

    If total < 1 Then
        Err.Raise ERR_EMPTY, "SampleData.Initialize", "数据没有元素。"
    End If

    skipFirst = hasHeader
    If hasHeader Then
        If dims = 1 Then
            this.header = HeaderText(data(firstRow))
        Else
            this.header = HeaderText(data(firstRow, firstCol))
        End If
    ElseIf Len(header) > 0 Then
        this.header = header
    End If

    ReDim flat(1 To total)
    n = 0
    For i = 0 To total - 1
        If Not (skipFirst And i = 0) Then
            If dims = 1 Then
                v = data(firstRow + i)
            ElseIf rows = 1 Then
                v = data(firstRow, firstCol + i)
            Else
                v = data(firstRow + i, firstCol)
            End If

            If IsAcceptable(v) Then
                n = n + 1
                flat(n) = CDbl(v)
            End If
        End If
    Next i

    If n < 2 Then
        Err.Raise ERR_TOO_FEW, "SampleData.Initialize", _
                  "有效值只有 " & n & " 个，至少需要 2 个。"
    End If
    If n < total Then ReDim Preserve flat(1 To n)

    this.values = flat
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
    lo = this.values(1)
    hi = this.values(1)

    For i = 1 To this.Count
        sum = sum + this.values(i)
        If this.values(i) < lo Then lo = this.values(i)
        If this.values(i) > hi Then hi = this.values(i)
    Next i

    this.Mean = sum / this.Count
    this.minValue = lo
    this.maxValue = hi

    sumSq = 0#
    For i = 1 To this.Count
        dev = this.values(i) - this.Mean
        sumSq = sumSq + dev * dev
    Next i
    this.StdevOverall = Sqr(sumSq / (this.Count - 1))

    mrSum = 0#
    For i = 2 To this.Count
        mrSum = mrSum + Abs(this.values(i) - this.values(i - 1))
    Next i
    this.StdevWithin = (mrSum / (this.Count - 1)) / D2_MR
End Sub

Private Function ReadRange(ByVal source As Object) As Variant
    Dim areaCount As Long
    Dim data As Variant

    If source Is Nothing Then
        Err.Raise ERR_EMPTY, "SampleData.ReadRange", "Range 不能为 Nothing。"
    End If

    On Error Resume Next
    areaCount = source.Areas.Count
    If Err.Number <> 0 Then
        On Error GoTo 0
        Err.Raise ERR_SHAPE, "SampleData.ReadRange", "数据源必须是 Excel Range。"
    End If
    If areaCount > 1 Then
        On Error GoTo 0
        Err.Raise ERR_SHAPE, "SampleData.ReadRange", _
                  "不支持多区域 Range，请先合并为连续区域。"
    End If

    data = source.Value2
    If Err.Number <> 0 Then
        On Error GoTo 0
        Err.Raise ERR_SHAPE, "SampleData.ReadRange", "读取 Range.Value2 失败。"
    End If
    On Error GoTo 0

    If Not IsArray(data) Then data = Array(data)
    ReadRange = data
End Function

Private Function HeaderText(ByVal v As Variant) As String
    If Not arr.IsScalarValue(v) Then Exit Function
    HeaderText = CStr(v)
End Function

Private Function IsAcceptable(ByVal v As Variant) As Boolean
    If Not arr.IsScalarValue(v) Then Exit Function
    IsAcceptable = IsNumeric(v)
End Function

Private Function ArrayDimensions(ByVal arr As Variant) As Long
    Dim d As Long
    Dim probe As Long

    On Error Resume Next
    For d = 1 To 60
        probe = UBound(arr, d)
        If Err.Number <> 0 Then
            Err.Clear
            Exit For
        End If
    Next d
    On Error GoTo 0

    ArrayDimensions = d - 1
End Function

