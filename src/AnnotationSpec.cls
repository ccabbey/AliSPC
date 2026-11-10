VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AnnotationSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("src.Spec")
'@PredeclaredId
'==============================================================================
' AnnotationSpec - 图内文字标注（哑数据对象，零 Excel 引用）
'
' 主要用途：Minitab 能力图右上角的统计值文本块（N / Mean / StDev /
'           Cp / Cpk / Pp / Ppk）。由 Builder 把 CapabilityResult 格式化成
'           多行字符串后塞进 Text，本类只负责"放在哪、长什么样"。
'
' X / Y 是【数据坐标】，不是像素、不是图表百分比。渲染层负责用
' ChartCanvas.DataToPixelX/Y 换算成 Shape 位置。这样标注在缩放图表、
' 改轴范围时仍然锚在正确的数据位置上。
'
' 多行文本用 vbCrLf 分隔。渲染层建 Shape 后要设 .TextFrame2.WordWrap = False
' 并按行撑开高度，否则 Excel 会按默认宽度折行，排版会乱。
'==============================================================================
Option Explicit

Private Const ERR_EMPTY_TEXT As Long = vbObjectError + 1701
Private Const ERR_FONT_SIZE As Long = vbObjectError + 1702

Public Enum TextAlignEnum
    alignLeft = 0
    alignCenter = 1
    alignRight = 2
End Enum

Private Type TAnnotationSpec
    text As String
    X As Double
    Y As Double
    align As TextAlignEnum
    fontName As String
    fontSize As Single
    FontColor As Long
    Visible As Boolean
End Type

Private this As TAnnotationSpec

'=== 工厂 =====================================================================

Public Function Create(ByVal text As String, _
                       ByVal X As Double, _
                       ByVal Y As Double, _
                       Optional ByVal align As TextAlignEnum = 0, _
                       Optional ByVal fontSize As Single = 9, _
                       Optional ByVal fontName As String = "Arial") As AnnotationSpec
    Dim result As AnnotationSpec
    Set result = New AnnotationSpec
    result.Initialize text, X, Y, align, fontSize, fontName
    Set Create = result
End Function

Private Sub Class_Initialize()
    this.text = ""
    this.X = 0#
    this.Y = 0#
    this.align = TextAlignEnum.alignLeft
    this.fontName = "Arial"
    this.fontSize = 9
    this.FontColor = RGB(0, 0, 0)
    this.Visible = True
End Sub

'=== 属性 =====================================================================

Public Property Get text() As String
    text = this.text
End Property

Public Property Let text(ByVal value As String)
    this.text = value
End Property

Public Property Get X() As Double
    X = this.X
End Property

Public Property Let X(ByVal value As Double)
    this.X = value
End Property

Public Property Get Y() As Double
    Y = this.Y
End Property

Public Property Let Y(ByVal value As Double)
    this.Y = value
End Property

Public Property Get align() As TextAlignEnum
    align = this.align
End Property

Public Property Let align(ByVal value As TextAlignEnum)
    this.align = value
End Property

Public Property Get fontName() As String
    fontName = this.fontName
End Property

Public Property Let fontName(ByVal value As String)
    If Len(value) = 0 Then
        this.fontName = "Arial"
    Else
        this.fontName = value
    End If
End Property

Public Property Get fontSize() As Single
    fontSize = this.fontSize
End Property

Public Property Let fontSize(ByVal value As Single)
    If value <= 0 Then
        Err.Raise ERR_FONT_SIZE, "AnnotationSpec.FontSize", _
                  "字号必须为正，当前为 " & value & "。"
    End If
    this.fontSize = value
End Property

Public Property Get FontColor() As Long
    FontColor = this.FontColor
End Property

Public Property Let FontColor(ByVal value As Long)
    this.FontColor = value
End Property

Public Property Get Visible() As Boolean
    Visible = this.Visible
End Property

Public Property Let Visible(ByVal value As Boolean)
    this.Visible = value
End Property

' 行数，供渲染层估算 Shape 高度。
Public Property Get LineCount() As Long
    If Len(this.text) = 0 Then
        LineCount = 0
    Else
        LineCount = UBound(Split(this.text, vbCrLf)) + 1
    End If
End Property

'=== 内部实现 =================================================================

Public Sub Initialize(ByVal text As String, _
                       ByVal X As Double, _
                       ByVal Y As Double, _
                       ByVal align As TextAlignEnum, _
                       ByVal fontSize As Single, _
                       ByVal fontName As String)
    If Len(text) = 0 Then
        Err.Raise ERR_EMPTY_TEXT, "AnnotationSpec.Initialize", "标注文本不能为空。"
    End If
    If fontSize <= 0 Then
        Err.Raise ERR_FONT_SIZE, "AnnotationSpec.Initialize", _
                  "字号必须为正，当前为 " & fontSize & "。"
    End If

    this.text = text
    this.X = X
    this.Y = Y
    this.align = align
    this.fontSize = fontSize
    If Len(fontName) = 0 Then
        this.fontName = "Arial"
    Else
        this.fontName = fontName
    End If
End Sub
