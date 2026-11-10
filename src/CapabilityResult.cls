VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CapabilityResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@Folder("src.Model")
'@PredeclaredId
'==============================================================================
' CapabilityResult - 过程能力指数（纯计算，零 Excel 引用）
'
' 口径：
'   C 系列用 StdevWithin（MRbar/d2），P 系列用 StdevOverall（样本标准差 N-1）
'   Cp  = (USL-LSL) / (k * sw)          k = spec.SigmaMultiplier，默认 6
'   Cpl = (Mean-LSL) / ((k/2) * sw)     默认即 3*sw
'   Cpu = (USL-Mean) / ((k/2) * sw)
'   Cpk = 可用侧的最小值
'   Cpm = (USL-LSL) / (k * Sqr(sw^2 + (Mean-Target)^2))    Taguchi，需目标值
'
' 不适用一律返回 Empty（不是 0、不是 -1），调用方用 IsEmpty 判断后显示 "--"。
' 触发 Empty 的情形：单侧规格下的 Cp/Pp/Cpm；sigma 为 0；无目标值时的 Cpm。
'
' 待核对：Cpm 的 sigma 取 within 还是 overall，各版本 Minitab 口径不完全一致。
'         本类按 within 实现（Taguchi 原始定义）。首次出图后请用同一组数据
'         对照 Minitab 的 Cpm 数值确认；若不符，只改 ComputeCpm 一处。
'
' 刻意不做：观测/预期超差 PPM。预期超差需要正态 CDF，而 VBA 侧只能靠
'           Application.WorksheetFunction.Norm_S_Dist —— 那会引入 Excel 引用，
'           破坏本层可脱离 Excel 单测的契约。要做就自己实现 erf 近似，
'           属于独立增量，不塞进本次范围。
'==============================================================================
Option Explicit

Private Const ERR_NO_SAMPLE As Long = vbObjectError + 1401
Private Const ERR_NO_SPEC As Long = vbObjectError + 1402

Private Type TCapabilityResult
    Cp As Variant
    Cpl As Variant
    Cpu As Variant
    Cpk As Variant
    Pp As Variant
    Ppl As Variant
    Ppu As Variant
    Ppk As Variant
    Cpm As Variant
End Type

Private this As TCapabilityResult

'=== 工厂 =====================================================================

Public Function FromSample(ByVal source As SampleData, _
                           ByVal spec As CapabilitySpec) As CapabilityResult
    Dim result As CapabilityResult
    Set result = New CapabilityResult
    result.Compute source, spec
    Set FromSample = result
End Function

'=== 只读属性（Variant：Empty 表示不适用）=====================================

Public Property Get Cp() As Variant
    Cp = this.Cp
End Property

Public Property Get Cpl() As Variant
    Cpl = this.Cpl
End Property

Public Property Get Cpu() As Variant
    Cpu = this.Cpu
End Property

Public Property Get Cpk() As Variant
    Cpk = this.Cpk
End Property

Public Property Get Pp() As Variant
    Pp = this.Pp
End Property

Public Property Get Ppl() As Variant
    Ppl = this.Ppl
End Property

Public Property Get Ppu() As Variant
    Ppu = this.Ppu
End Property

Public Property Get Ppk() As Variant
    Ppk = this.Ppk
End Property

Public Property Get Cpm() As Variant
    Cpm = this.Cpm
End Property

'=== 内部实现 =================================================================

Public Sub Compute(ByVal source As SampleData, ByVal spec As CapabilitySpec)
    Dim mu As Double
    Dim sw As Double
    Dim so As Double
    Dim k As Double
    Dim halfK As Double
    Dim hasBoth As Boolean
    Dim tol As Double

    If source Is Nothing Then
        Err.Raise ERR_NO_SAMPLE, "CapabilityResult.Compute", "source 不能为 Nothing。"
    End If
    If spec Is Nothing Then
        Err.Raise ERR_NO_SPEC, "CapabilityResult.Compute", "spec 不能为 Nothing。"
    End If

    mu = source.Mean
    sw = source.StdevWithin
    so = source.StdevOverall
    k = spec.SigmaMultiplier
    halfK = k / 2#
    hasBoth = spec.HasLowerSpec And spec.HasUpperSpec

    ' 默认全部不适用，下面只覆盖算得出来的项
    this.Cp = Empty
    this.Cpl = Empty
    this.Cpu = Empty
    this.Cpk = Empty
    this.Pp = Empty
    this.Ppl = Empty
    this.Ppu = Empty
    this.Ppk = Empty
    this.Cpm = Empty

    If hasBoth Then tol = spec.upperSpec - spec.lowerSpec

    '--- C 系列（within）---
    If sw > 0# Then
        If hasBoth Then this.Cp = tol / (k * sw)

        If spec.HasLowerSpec Then
            this.Cpl = (mu - spec.lowerSpec) / (halfK * sw)
        End If
        If spec.HasUpperSpec Then
            this.Cpu = (spec.upperSpec - mu) / (halfK * sw)
        End If
        this.Cpk = MinAvailable(this.Cpl, this.Cpu)

        If hasBoth And spec.HasTarget Then
            this.Cpm = ComputeCpm(tol, k, sw, mu, spec.target)
        End If
    End If

    '--- P 系列（overall）---
    If so > 0# Then
        If hasBoth Then this.Pp = tol / (k * so)

        If spec.HasLowerSpec Then
            this.Ppl = (mu - spec.lowerSpec) / (halfK * so)
        End If
        If spec.HasUpperSpec Then
            this.Ppu = (spec.upperSpec - mu) / (halfK * so)
        End If
        this.Ppk = MinAvailable(this.Ppl, this.Ppu)
    End If
End Sub

Private Function ComputeCpm(ByVal tol As Double, _
                            ByVal k As Double, _
                            ByVal sigma As Double, _
                            ByVal mu As Double, _
                            ByVal target As Double) As Double
    Dim dev As Double
    dev = mu - target
    ComputeCpm = tol / (k * Sqr(sigma * sigma + dev * dev))
End Function

' 单侧规格时只有一项可用，直接取该项；两侧都不可用返回 Empty。
Private Function MinAvailable(ByVal a As Variant, ByVal b As Variant) As Variant
    Dim hasA As Boolean
    Dim hasB As Boolean
    hasA = Not IsEmpty(a)
    hasB = Not IsEmpty(b)

    If hasA And hasB Then
        If CDbl(a) <= CDbl(b) Then
            MinAvailable = a
        Else
            MinAvailable = b
        End If
    ElseIf hasA Then
        MinAvailable = a
    ElseIf hasB Then
        MinAvailable = b
    Else
        MinAvailable = Empty
    End If
End Function


