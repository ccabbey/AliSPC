VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AxisSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("src.Spec")
'@PredeclaredId
'==============================================================================
' AxisSpec - 坐标轴意图（哑数据对象，零 Excel 引用）
'
' Min / Max / MajorUnit 用 Variant，Empty 表示"交给 Excel 自动"。
' 不要用 0 表达自动 —— 0 是合法的轴下限（频数轴就必须从 0 起）。
'
' 本类只描述【主数值轴】：XAxis 是测量值轴，YAxis 是频数轴。
' 面积图用的副分类轴（0~10000、隐藏）不在这里描述 —— 那是渲染层
' PaintAreaFill 的实现细节，由 ChartCanvas.EnsureSecondaryScaledAxis 自行建立。
' 把它写进 Spec 就等于让 x10000 泄漏出渲染层，解耦契约就破了。
'==============================================================================
Option Explicit

Private Const ERR_ORDER As Long = vbObjectError + 1601
Private Const ERR_UNIT As Long = vbObjectError + 1602
Private Const ERR_NOT_NUMERIC As Long = vbObjectError + 1603

Private Type TAxisSpec
    Min As Variant
    Max As Variant
    MajorUnit As Variant
    Title As String
    numberFormat As String
    Visible As Boolean
End Type

Private this As TAxisSpec

'=== 工厂 =====================================================================

Public Function Create(Optional ByVal minValue As Variant = Empty, _
                       Optional ByVal maxValue As Variant = Empty, _
                       Optional ByVal unitValue As Variant = Empty, _
                       Optional ByVal Title As String = "", _
                       Optional ByVal numberFormat As String = "General", _
                       Optional ByVal Visible As Boolean = True) As AxisSpec
    Dim result As AxisSpec
    Set result = New AxisSpec
    result.Initialize minValue, maxValue, unitValue, Title, numberFormat, Visible
    Set Create = result
End Function

Private Sub Class_Initialize()
    this.Min = Empty
    this.Max = Empty
    this.MajorUnit = Empty
    this.Title = ""
    this.numberFormat = "General"
    this.Visible = True
End Sub

'=== 属性 =====================================================================

Public Property Get Min() As Variant
    Min = this.Min
End Property

Public Property Let Min(ByVal value As Variant)
    this.Min = NormalizeNumeric(value, "Min")
    ValidateOrder
End Property

Public Property Get Max() As Variant
    Max = this.Max
End Property

Public Property Let Max(ByVal value As Variant)
    this.Max = NormalizeNumeric(value, "Max")
    ValidateOrder
End Property

Public Property Get MajorUnit() As Variant
    MajorUnit = this.MajorUnit
End Property

Public Property Let MajorUnit(ByVal value As Variant)
    this.MajorUnit = NormalizeNumeric(value, "MajorUnit")
    If Not IsEmpty(this.MajorUnit) Then
        If CDbl(this.MajorUnit) <= 0# Then
            Err.Raise ERR_UNIT, "AxisSpec.MajorUnit", _
                      "MajorUnit 必须为正，当前为 " & this.MajorUnit & "。"
        End If
    End If
End Property

Public Property Get HasMin() As Boolean
    HasMin = Not IsEmpty(this.Min)
End Property

Public Property Get HasMax() As Boolean
    HasMax = Not IsEmpty(this.Max)
End Property

Public Property Get HasMajorUnit() As Boolean
    HasMajorUnit = Not IsEmpty(this.MajorUnit)
End Property

Public Property Get Title() As String
    Title = this.Title
End Property

Public Property Let Title(ByVal value As String)
    this.Title = value
End Property

Public Property Get TitleVisible() As Boolean
    TitleVisible = (Len(this.Title) > 0)
End Property

Public Property Get numberFormat() As String
    numberFormat = this.numberFormat
End Property

Public Property Let numberFormat(ByVal value As String)
    If Len(value) = 0 Then
        this.numberFormat = "General"
    Else
        this.numberFormat = value
    End If
End Property

Public Property Get Visible() As Boolean
    Visible = this.Visible
End Property

Public Property Let Visible(ByVal value As Boolean)
    this.Visible = value
End Property

'=== 内部实现 =================================================================

Public Sub Initialize(ByVal minValue As Variant, _
                       ByVal maxValue As Variant, _
                       ByVal unitValue As Variant, _
                       ByVal Title As String, _
                       ByVal numberFormat As String, _
                       ByVal Visible As Boolean)
    this.Min = NormalizeNumeric(minValue, "Min")
    this.Max = NormalizeNumeric(maxValue, "Max")
    this.MajorUnit = NormalizeNumeric(unitValue, "MajorUnit")
    this.Title = Title
    If Len(numberFormat) = 0 Then
        this.numberFormat = "General"
    Else
        this.numberFormat = numberFormat
    End If
    this.Visible = Visible

    ValidateOrder
    If Not IsEmpty(this.MajorUnit) Then
        If CDbl(this.MajorUnit) <= 0# Then
            Err.Raise ERR_UNIT, "AxisSpec.Initialize", _
                      "MajorUnit 必须为正，当前为 " & this.MajorUnit & "。"
        End If
    End If
End Sub

Private Sub ValidateOrder()
    If IsEmpty(this.Min) Then Exit Sub
    If IsEmpty(this.Max) Then Exit Sub
    If CDbl(this.Min) >= CDbl(this.Max) Then
        Err.Raise ERR_ORDER, "AxisSpec.ValidateOrder", _
                  "轴下限必须小于上限，当前 Min=" & this.Min & " Max=" & this.Max & "。"
    End If
End Sub

' Empty / Null 归一为 Empty（自动）；其余必须是标量数值。
Private Function NormalizeNumeric(ByVal value As Variant, _
                                  ByVal fieldName As String) As Variant
    If IsEmpty(value) Then
        NormalizeNumeric = Empty
        Exit Function
    End If
    If IsNull(value) Then
        NormalizeNumeric = Empty
        Exit Function
    End If
    If IsError(value) Or IsObject(value) Or IsArray(value) Then
        Err.Raise ERR_NOT_NUMERIC, "AxisSpec.NormalizeNumeric", _
                  fieldName & " 必须是标量数值或 Empty。"
    End If
    If Not IsNumeric(value) Then
        Err.Raise ERR_NOT_NUMERIC, "AxisSpec.NormalizeNumeric", _
                  fieldName & " 必须是数值，当前为 """ & CStr(value) & """。"
    End If
    NormalizeNumeric = CDbl(value)
End Function
