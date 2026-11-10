VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CapabilitySpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("src.Model")
'@PredeclaredId
'==============================================================================
' CapabilitySpec - 规格限与目标值（纯值对象，零 Excel 引用）
'
' 单侧规格用 Empty / Null 表达"未指定"，不引入哨兵数值（0 和 -1 都可能是
' 真实规格值，用哨兵会出错）。因此：
'   CapabilitySpec.Create(10, 20)          双侧
'   CapabilitySpec.Create(10, Empty)       只有下限
'   CapabilitySpec.Create(Empty, 20)       只有上限
'   CapabilitySpec.Create(10, 20, 15)      双侧 + 目标值（Cpm 需要）
'
' SigmaMultiplier 默认 6，即 Cp = Tolerance / (6 * sigma)、
' Cpk 的单侧分母为 3 * sigma。改成 3 就是 3-sigma 能力口径；
' 不要为了别的口径在 CapabilityResult 里加分支，改这个属性即可。
'==============================================================================
Option Explicit

Private Const DEFAULT_SIGMA_MULTIPLIER As Double = 6#

Private Const ERR_NOT_NUMERIC As Long = vbObjectError + 1301
Private Const ERR_ORDER As Long = vbObjectError + 1302
Private Const ERR_NO_SPEC As Long = vbObjectError + 1303
Private Const ERR_MULTIPLIER As Long = vbObjectError + 1304

Private Type TCapabilitySpec
    lowerSpec As Double
    upperSpec As Double
    target As Double
    HasLower As Boolean
    HasUpper As Boolean
    HasTarget As Boolean
    SigmaMultiplier As Double
End Type

Private this As TCapabilitySpec

'=== 工厂 =====================================================================

Public Function Create(ByVal lowerSpec As Variant, _
                       ByVal upperSpec As Variant, _
                       Optional ByVal target As Variant = Empty) As CapabilitySpec
    Dim result As CapabilitySpec
    Set result = New CapabilitySpec
    result.Initialize lowerSpec, upperSpec, target
    Set Create = result
End Function

Private Sub Class_Initialize()
    this.SigmaMultiplier = DEFAULT_SIGMA_MULTIPLIER
End Sub

'=== 属性 =====================================================================

Public Property Get HasLowerSpec() As Boolean
    HasLowerSpec = this.HasLower
End Property

Public Property Get HasUpperSpec() As Boolean
    HasUpperSpec = this.HasUpper
End Property

Public Property Get HasTarget() As Boolean
    HasTarget = this.HasTarget
End Property

' 以下三个 Double 属性只在对应 Has* 为 True 时有意义。
Public Property Get lowerSpec() As Double
    lowerSpec = this.lowerSpec
End Property

Public Property Get upperSpec() As Double
    upperSpec = this.upperSpec
End Property

Public Property Get target() As Double
    target = this.target
End Property

' 双侧规格才有公差带；单侧返回 Empty，调用方据此显示 "--"。
Public Property Get Tolerance() As Variant
    If this.HasLower And this.HasUpper Then
        Tolerance = this.upperSpec - this.lowerSpec
    Else
        Tolerance = Empty
    End If
End Property

Public Property Get SigmaMultiplier() As Double
    SigmaMultiplier = this.SigmaMultiplier
End Property

Public Property Let SigmaMultiplier(ByVal value As Double)
    If value <= 0# Then
        Err.Raise ERR_MULTIPLIER, "CapabilitySpec.SigmaMultiplier", _
                  "SigmaMultiplier 必须为正，当前为 " & value & "。"
    End If
    this.SigmaMultiplier = value
End Property

'=== 内部实现 =================================================================

Public Sub Initialize(ByVal lowerSpec As Variant, _
                       ByVal upperSpec As Variant, _
                       ByVal target As Variant)
    Dim v As Variant

    v = ReadOptional(lowerSpec, "lowerSpec")
    this.HasLower = Not IsEmpty(v)
    If this.HasLower Then this.lowerSpec = CDbl(v)

    v = ReadOptional(upperSpec, "upperSpec")
    this.HasUpper = Not IsEmpty(v)
    If this.HasUpper Then this.upperSpec = CDbl(v)

    v = ReadOptional(target, "target")
    this.HasTarget = Not IsEmpty(v)
    If this.HasTarget Then this.target = CDbl(v)

    If Not this.HasLower And Not this.HasUpper Then
        Err.Raise ERR_NO_SPEC, "CapabilitySpec.Initialize", _
                  "上下规格限至少要指定一个。"
    End If

    If this.HasLower And this.HasUpper Then
        If this.lowerSpec >= this.upperSpec Then
            Err.Raise ERR_ORDER, "CapabilitySpec.Initialize", _
                      "下限必须小于上限，当前 LSL=" & this.lowerSpec & _
                      " USL=" & this.upperSpec & "。"
        End If
    End If

    ' 目标值允许落在规格带之外（Minitab 也允许），不做钳制。
End Sub

' Optional Variant 归一化：Empty / Null -> Empty（未指定），其余必须是标量数值。
' 编码方式与 AxisSpec.NormalizeNumeric 一致：用返回值的 Empty 表达"不存在"，
' 不用 Boolean + 出参。全项目只有 StepPolygon / NormalCurve 这类天然双输出
' 才使用 ByRef 出参。
Private Function ReadOptional(ByVal value As Variant, _
                              ByVal fieldName As String) As Variant
    If IsEmpty(value) Then
        ReadOptional = Empty
        Exit Function
    End If
    If IsNull(value) Then
        ReadOptional = Empty
        Exit Function
    End If

    If IsError(value) Then
        Err.Raise ERR_NOT_NUMERIC, "CapabilitySpec.ReadOptional", _
                  fieldName & " 是单元格错误值。"
    End If
    If IsObject(value) Or IsArray(value) Then
        Err.Raise ERR_NOT_NUMERIC, "CapabilitySpec.ReadOptional", _
                  fieldName & " 必须是标量数值或 Empty，不能是对象或数组。"
    End If
    If Not IsNumeric(value) Then
        Err.Raise ERR_NOT_NUMERIC, "CapabilitySpec.ReadOptional", _
                  fieldName & " 必须是数值或 Empty，当前为 """ & CStr(value) & """。"
    End If

    ReadOptional = CDbl(value)
End Function


