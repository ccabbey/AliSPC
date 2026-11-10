Attribute VB_Name = "Arr"
'@ModuleDescription "极简一维数组工具。只放本项目真正用到的几个函数，不复用 CPArrayLib。"
'@Folder("lib")
Option Explicit
Option Private Module

Private Const ERR_NOT_ALLOCATED As Long = vbObjectError + 1701
Private Const ERR_BAD_COUNT As Long = vbObjectError + 1702

'@Description "判断动态数组是否已分配（LBound <= UBound）。未分配时返回 False，不抛错。"
Public Function IsAllocated(ByVal Arr As Variant) As Boolean
    If Not IsArray(Arr) Then Exit Function
    On Error GoTo notAllocated
    IsAllocated = (UBound(Arr) >= LBound(Arr))
    Exit Function
notAllocated:
    IsAllocated = False
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
Public Function Max(ByVal Arr As Variant) As Double
    RequireAllocated Arr, "Arr.Max"
    Dim result As Double
    result = CDbl(Arr(LBound(Arr)))
    Dim i As Long
    For i = LBound(Arr) + 1 To UBound(Arr)
        If CDbl(Arr(i)) > result Then result = CDbl(Arr(i))
    Next i
    Max = result
End Function

'@Description "返回一维数组最小值。空数组抛错。"
Public Function Min(ByVal Arr As Variant) As Double
    RequireAllocated Arr, "Arr.Min"
    Dim result As Double
    result = CDbl(Arr(LBound(Arr)))
    Dim i As Long
    For i = LBound(Arr) + 1 To UBound(Arr)
        If CDbl(Arr(i)) < result Then result = CDbl(Arr(i))
    Next i
    Min = result
End Function

'@Description "返回最大值所在下标（沿用原数组的基）。空数组抛错。"
Public Function ArgMax(ByVal Arr As Variant) As Long
    RequireAllocated Arr, "Arr.ArgMax"
    Dim best As Long
    best = LBound(Arr)
    Dim i As Long
    For i = LBound(Arr) + 1 To UBound(Arr)
        If CDbl(Arr(i)) > CDbl(Arr(best)) Then best = i
    Next i
    ArgMax = best
End Function

Private Sub RequireAllocated(ByVal Arr As Variant, ByVal source As String)
    If Not IsAllocated(Arr) Then
        Err.Raise ERR_NOT_ALLOCATED, source, "数组未分配或为空。"
    End If
End Sub

