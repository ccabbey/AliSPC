Attribute VB_Name = "SeriesDump"
'@Folder("lib")
'@ModuleDescription "把 CapabilityPlotSpec 的系列实际值写到工作表，便于校验计算。只读 Spec，不碰图表。"
'@Folder("src.Render")
Option Explicit
Option Private Module

Private Const ERR_NOTHING As Long = vbObjectError + 2001

'@Description "清空目标工作表后按系列分块输出 X/Y 两列（块间空一列）。输出的是原始数据坐标，未做 x10000 缩放。"
Public Sub WriteData(ByVal ps As CapabilityPlotSpec, ByVal ws As Worksheet)
    If ps Is Nothing Then
        Err.Raise ERR_NOTHING, "SeriesDump.Write", "ps 不能为 Nothing。"
    End If
    If ws Is Nothing Then
        Err.Raise ERR_NOTHING, "SeriesDump.Write", "ws 不能为 Nothing。"
    End If

    ws.Cells.ClearContents

    Dim s As SeriesSpec
    Dim xv As Variant
    Dim yv As Variant
    Dim col As Long
    Dim r As Long
    Dim i As Long

    col = 1
    For Each s In ps.AllSeries()
        xv = s.xs
        yv = s.ys
        ws.Cells(1, col).value = s.Name
        ws.Cells(2, col).value = "X"
        ws.Cells(2, col + 1).value = "Y"
        r = 3
        For i = LBound(xv) To UBound(xv)
            ws.Cells(r, col).value = xv(i)
            ws.Cells(r, col + 1).value = yv(i)
            r = r + 1
        Next i
        col = col + 3
    Next s
End Sub

