Attribute VB_Name = "Arr"
'@ModuleDescription "极简一维数组工具。只放本项目真正用到的几个函数，不复用 CPArrayLib。"
'@Folder("lib")
Option Explicit
Option Private Module
' Modified: 2026-11-28 22:41:35

Private Const ERR_NOT_ALLOCATED As Long = vbObjectError + 1701
Private Const ERR_BAD_COUNT As Long = vbObjectError + 1702

'@Description "判断动态数组是否已分配（LBound <= UBound）。未分配时返回 False，不抛错。"
Public Function IsAllocated(ByVal arr As Variant) As Boolean
    If Not IsArray(arr) Then Exit Function
    On Error GoTo notAllocated
    IsAllocated = (UBound(arr) >= LBound(arr))
    Exit Function
notAllocated:
    IsAllocated = False
End Function

'@Description "判断 Variant 是否为可处理的单元格值；排除 Error、Null、Empty、Object 和数组"
Public Function IsScalarValue(ByVal value As Variant) As Boolean
    If IsError(value) Then Exit Function
    If IsNull(value) Then Exit Function
    If IsEmpty(value) Then Exit Function
    If IsObject(value) Then Exit Function
    If IsArray(value) Then Exit Function
    IsScalarValue = True
End Function

'@Description "生成 1 基等差数列，含两端点。nPoints < 2 时抛错。"
Public Function Linspace(ByVal x0 As Double, ByVal x1 As Double, _
                         ByVal nPoints As Long) As Variant
    If nPoints < 2 Then
        Err.Raise ERR_BAD_COUNT, "Arr.Linspace", "nPoints 必须 >= 2。"
    End If
    Dim result() As Double
    ReDim result(1 To nPoints)
    Dim step As Double
    step = (x1 - x0) / (nPoints - 1)
    Dim i As Long
    For i = 1 To nPoints - 1
        result(i) = x0 + step * (i - 1)
    Next i
    result(nPoints) = x1          ' 末点钉死，避免浮点累加漂移
    Linspace = result
End Function

'@Description "返回一维数组最大值。空数组抛错。"
Public Function Max(ByVal arr As Variant) As Double
    RequireAllocated arr, "Arr.Max"
    Dim result As Double
    result = CDbl(arr(LBound(arr)))
    Dim i As Long
    For i = LBound(arr) + 1 To UBound(arr)
        If CDbl(arr(i)) > result Then result = CDbl(arr(i))
    Next i
    Max = result
End Function

'@Description "返回一维数组最小值。空数组抛错。"
Public Function Min(ByVal arr As Variant) As Double
    RequireAllocated arr, "Arr.Min"
    Dim result As Double
    result = CDbl(arr(LBound(arr)))
    Dim i As Long
    For i = LBound(arr) + 1 To UBound(arr)
        If CDbl(arr(i)) < result Then result = CDbl(arr(i))
    Next i
    Min = result
End Function

'@Description "返回最大值所在下标（沿用原数组的基）。空数组抛错。"
Public Function ArgMax(ByVal arr As Variant) As Long
    RequireAllocated arr, "Arr.ArgMax"
    Dim best As Long
    best = LBound(arr)
    Dim i As Long
    For i = LBound(arr) + 1 To UBound(arr)
        If CDbl(arr(i)) > CDbl(arr(best)) Then best = i
    Next i
    ArgMax = best
End Function

Private Sub RequireAllocated(ByVal arr As Variant, ByVal source As String)
    If Not IsAllocated(arr) Then
        Err.Raise ERR_NOT_ALLOCATED, source, "数组未分配或为空。"
    End If
End Sub

