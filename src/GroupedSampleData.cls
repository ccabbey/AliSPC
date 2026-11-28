VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GroupedSampleData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@ModuleDescription "多组样本数据；支持多列宽表和逆透视长表"
'@PredeclaredId
'@Folder("src.Model")
' Modified: 2026-11-28 22:41:35
Option Explicit

Private Const ERR_NO_DATA As Long = vbObjectError + 1301          ' 数据源为 Nothing 或没有有效数据
Private Const ERR_SHAPE As Long = vbObjectError + 1302            ' 数据源形状或列参数不合法
Private Const ERR_NO_GROUP As Long = vbObjectError + 1303         ' 没有可用的分组
Private Const ERR_GROUP_TOO_SMALL As Long = vbObjectError + 1304  ' 某个分组有效样本不足 2 个
Private Const ERR_INDEX As Long = vbObjectError + 1305            ' 分组索引越界

Private Type TGroupedSample
    Samples As Collection
End Type

Private this As TGroupedSample

Private Sub Class_Initialize()
    Set this.Samples = New Collection
End Sub

Public Function FromColumns(ByVal source As Object, _
                            Optional ByVal hasHeader As Boolean = True) As GroupedSampleData
    Dim result As GroupedSampleData
    Set result = New GroupedSampleData
    result.InitializeFromColumns source, hasHeader
    Set FromColumns = result
End Function

Public Function FromLongTable(ByVal source As Object, _
                              ByVal valueColumn As Long, _
                              ByVal groupColumn As Long, _
                              Optional ByVal hasHeader As Boolean = False) As GroupedSampleData
    Dim result As GroupedSampleData
    Set result = New GroupedSampleData
    result.InitializeFromLongTable source, valueColumn, groupColumn, hasHeader
    Set FromLongTable = result
End Function

Public Property Get GroupCount() As Long
    GroupCount = this.Samples.Count
End Property

Public Function groupName(ByVal index As Long) As String
    groupName = Me.GroupSample(index).header
End Function

Public Function GroupSample(ByVal index As Long) As SampleData
    ValidateIndex index
    Set GroupSample = this.Samples(index)
End Function

Public Sub InitializeFromColumns(ByVal source As Object, _
                                  ByVal hasHeader As Boolean)
    Dim rowCount As Long
    Dim colCount As Long
    Dim firstDataRow As Long
    Dim c As Long
    Dim n As Long
    Dim i As Long
    Dim groupName As String
    Dim colData As Variant
    Dim values() As Double
    Dim v As Variant
    Dim sample As SampleData
    Dim valuesRange As Object

    ValidateRangeSource source
    rowCount = source.rows.Count
    colCount = source.Columns.Count
    If hasHeader Then
        If rowCount < 2 Then
            Err.Raise ERR_NO_GROUP, "GroupedSampleData.FromColumns", _
                      "只有表头，没有样本数据。"
        End If
        firstDataRow = 2
    Else
        firstDataRow = 1
    End If

    For c = 1 To colCount
        groupName = ""
        If hasHeader Then groupName = HeaderText(source.Cells(1, c).Value2)

        Set valuesRange = source.rows(firstDataRow) _
                                 .Resize(rowCount - firstDataRow + 1).Columns(c)
        colData = valuesRange.Value2
        If Not IsArray(colData) Then colData = Array(colData)

        n = 0
        For Each v In colData
            If IsNumericValue(v) Then n = n + 1
        Next v
        If n = 0 Then GoTo NextColumn
        If n < 2 Then
            Err.Raise ERR_GROUP_TOO_SMALL, "GroupedSampleData.FromColumns", _
                      "列 " & c & " 只有 " & n & " 个有效样本。"
        End If

        ReDim values(1 To n)
        i = 0
        For Each v In colData
            If IsNumericValue(v) Then
                i = i + 1
                values(i) = CDbl(v)
            End If
        Next v

        If Len(groupName) = 0 Then groupName = "Column " & CStr(c)
        Set sample = SampleData.FromArray(values, False, groupName)
        this.Samples.Add sample
NextColumn:
    Next c

    If this.Samples.Count = 0 Then
        Err.Raise ERR_NO_GROUP, "GroupedSampleData.FromColumns", "没有可用分组。"
    End If
End Sub

Public Sub InitializeFromLongTable(ByVal source As Object, _
                                    ByVal valueColumn As Long, _
                                    ByVal groupColumn As Long, _
                                    ByVal hasHeader As Boolean)
    Dim data As Variant
    Dim firstRow As Long
    Dim lastRow As Long
    Dim firstCol As Long
    Dim lastCol As Long
    Dim firstDataRow As Long
    Dim r As Long
    Dim key As Variant
    Dim n As Long
    Dim i As Long
    Dim values() As Double
    Dim groups As Object
    Dim order As Collection
    Dim groupValues As Collection
    Dim sample As SampleData

    ValidateRangeSource source
    data = source.Value2
    If Not IsArray(data) Then
        Err.Raise ERR_NO_DATA, "GroupedSampleData.FromLongTable", "长表没有数据。"
    End If
    If ArrayDimensions(data) <> 2 Then
        Err.Raise ERR_SHAPE, "GroupedSampleData.FromLongTable", "长表必须是二维区域。"
    End If

    firstRow = LBound(data, 1)
    lastRow = UBound(data, 1)
    firstCol = LBound(data, 2)
    lastCol = UBound(data, 2)

    If valueColumn < firstCol Or valueColumn > lastCol Then
        Err.Raise ERR_SHAPE, "GroupedSampleData.FromLongTable", "valueColumn 超出数据范围。"
    End If
    If groupColumn < firstCol Or groupColumn > lastCol Then
        Err.Raise ERR_SHAPE, "GroupedSampleData.FromLongTable", "groupColumn 超出数据范围。"
    End If
    If valueColumn = groupColumn Then
        Err.Raise ERR_SHAPE, "GroupedSampleData.FromLongTable", "数据列和分组列不能相同。"
    End If

    firstDataRow = firstRow
    If hasHeader Then firstDataRow = firstRow + 1
    If firstDataRow > lastRow Then
        Err.Raise ERR_NO_GROUP, "GroupedSampleData.FromLongTable", "只有表头，没有样本数据。"
    End If

    Set groups = CreateObject("Scripting.Dictionary")
    groups.CompareMode = vbTextCompare
    Set order = New Collection

    For r = firstDataRow To lastRow
        If IsNumericValue(data(r, valueColumn)) Then
            key = GroupText(data(r, groupColumn))
            If Len(key) > 0 Then
                If Not groups.Exists(key) Then
                    Set groupValues = New Collection
                    groups.Add key, groupValues
                    order.Add key
                End If
                Set groupValues = groups(key)
                groupValues.Add CDbl(data(r, valueColumn))
            End If
        End If
    Next r

    If groups.Count = 0 Then
        Err.Raise ERR_NO_GROUP, "GroupedSampleData.FromLongTable", "没有可用分组。"
    End If

    For Each key In order
        Set groupValues = groups(key)
        n = groupValues.Count
        If n < 2 Then
            Err.Raise ERR_GROUP_TOO_SMALL, "GroupedSampleData.FromLongTable", _
                      "分组 " & key & " 只有 " & n & " 个有效样本。"
        End If

        ReDim values(1 To n)
        For i = 1 To n
            values(i) = CDbl(groupValues(i))
        Next i
        Set sample = SampleData.FromArray(values, False, CStr(key))
        this.Samples.Add sample
    Next key
End Sub

Private Sub ValidateRangeSource(ByVal source As Object)
    If source Is Nothing Then
        Err.Raise ERR_NO_DATA, "GroupedSampleData.ValidateRangeSource", "Range 不能为 Nothing。"
    End If
    If source.Areas.Count <> 1 Then
        Err.Raise ERR_SHAPE, "GroupedSampleData.ValidateRangeSource", _
                  "只支持连续的单区域 Range。"
    End If
End Sub

Private Sub ValidateIndex(ByVal index As Long)
    If index < 1 Or index > this.Samples.Count Then
        Err.Raise ERR_INDEX, "GroupedSampleData.ValidateIndex", _
                  "分组索引越界：" & index
    End If
End Sub

Private Function HeaderText(ByVal v As Variant) As String
    If Not arr.IsScalarValue(v) Then Exit Function
    HeaderText = CStr(v)
End Function

Private Function GroupText(ByVal v As Variant) As String
    If Not arr.IsScalarValue(v) Then Exit Function
    GroupText = CStr(v)
End Function

Private Function IsNumericValue(ByVal v As Variant) As Boolean
    If Not arr.IsScalarValue(v) Then Exit Function
    IsNumericValue = IsNumeric(v)
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


