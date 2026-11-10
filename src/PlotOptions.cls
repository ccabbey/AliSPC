VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PlotOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "Rubberduck" ,"Predeclared Class Module"
'@ModuleDescription "能力图绘制选项。纯配置对象，不引用任何 Excel 对象模型。"
'@PredeclaredId
'@Folder("src.Build")
Option Explicit

Private Type TPlotOptions
    BinPolicy As BinPolicyEnum
    BinParam As Double
    CurvePoints As Long
    ShowDistCurve As Boolean
    ShowTol As Boolean
    ShowTarget As Boolean
    ShowSigma As Boolean
    ShowStatBox As Boolean
    Title As String
    ProfileColor As Long
    ShadeColor As Long
    ShadeTransparency As Single
    CurveColor As Long
    MeanColor As Long
    TolColor As Long
    TargetColor As Long
    SigmaColor As Long
End Type

Private this As TPlotOptions

'@Description "返回一份默认选项。调用方随后用 Property Let 逐项覆盖。"
Public Function Create() As PlotOptions
    Set Create = New PlotOptions      ' 默认值全部写在 Class_Initialize，工厂不改 this
End Function

Private Sub Class_Initialize()
    this.BinPolicy = binSturges
    this.BinParam = 0#
    this.CurvePoints = 200
    this.ShowDistCurve = True
    this.ShowTol = True
    this.ShowTarget = True
    this.ShowSigma = True
    this.ShowStatBox = True
    this.Title = ""
    this.ProfileColor = RGB(60, 60, 60)
    this.ShadeColor = RGB(217, 217, 217)
    this.ShadeTransparency = 0#
    this.CurveColor = RGB(0, 90, 180)
    this.MeanColor = RGB(0, 140, 0)
    this.TolColor = RGB(200, 0, 0)
    this.TargetColor = RGB(0, 150, 150)
    this.SigmaColor = RGB(130, 130, 130)
End Sub

'--- 分箱 -------------------------------------------------------------------

Public Property Get BinPolicy() As BinPolicyEnum
    BinPolicy = this.BinPolicy
End Property

Public Property Let BinPolicy(ByVal value As BinPolicyEnum)
    this.BinPolicy = value
End Property

'@Description "binFixedCount 时为箱数，binFixedWidth 时为箱宽，其余策略忽略。"
Public Property Get BinParam() As Double
    BinParam = this.BinParam
End Property

Public Property Let BinParam(ByVal value As Double)
    this.BinParam = value
End Property

'--- 曲线与可见性 -----------------------------------------------------------

'@Description "正态拟合曲线的采样点数。"
Public Property Get CurvePoints() As Long
    CurvePoints = this.CurvePoints
End Property

Public Property Let CurvePoints(ByVal value As Long)
    If value < 2 Then value = 2
    this.CurvePoints = value
End Property

Public Property Get ShowDistCurve() As Boolean
    ShowDistCurve = this.ShowDistCurve
End Property

Public Property Let ShowDistCurve(ByVal value As Boolean)
    this.ShowDistCurve = value
End Property

Public Property Get ShowTol() As Boolean
    ShowTol = this.ShowTol
End Property

Public Property Let ShowTol(ByVal value As Boolean)
    this.ShowTol = value
End Property

'@Description "仅当 CapabilitySpec.HasTarget 为真时才有意义。"
Public Property Get ShowTarget() As Boolean
    ShowTarget = this.ShowTarget
End Property

Public Property Let ShowTarget(ByVal value As Boolean)
    this.ShowTarget = value
End Property

Public Property Get ShowSigma() As Boolean
    ShowSigma = this.ShowSigma
End Property

Public Property Let ShowSigma(ByVal value As Boolean)
    this.ShowSigma = value
End Property

Public Property Get ShowStatBox() As Boolean
    ShowStatBox = this.ShowStatBox
End Property

Public Property Let ShowStatBox(ByVal value As Boolean)
    this.ShowStatBox = value
End Property

Public Property Get Title() As String
    Title = this.Title
End Property

Public Property Let Title(ByVal value As String)
    this.Title = value
End Property

'--- 颜色 -------------------------------------------------------------------

Public Property Get ProfileColor() As Long
    ProfileColor = this.ProfileColor
End Property

Public Property Let ProfileColor(ByVal value As Long)
    this.ProfileColor = value
End Property

Public Property Get ShadeColor() As Long
    ShadeColor = this.ShadeColor
End Property

Public Property Let ShadeColor(ByVal value As Long)
    this.ShadeColor = value
End Property

Public Property Get ShadeTransparency() As Single
    ShadeTransparency = this.ShadeTransparency
End Property

Public Property Let ShadeTransparency(ByVal value As Single)
    If value < 0# Then value = 0#
    If value > 1# Then value = 1#
    this.ShadeTransparency = value
End Property

Public Property Get CurveColor() As Long
    CurveColor = this.CurveColor
End Property

Public Property Let CurveColor(ByVal value As Long)
    this.CurveColor = value
End Property

Public Property Get MeanColor() As Long
    MeanColor = this.MeanColor
End Property

Public Property Let MeanColor(ByVal value As Long)
    this.MeanColor = value
End Property

Public Property Get TolColor() As Long
    TolColor = this.TolColor
End Property

Public Property Let TolColor(ByVal value As Long)
    this.TolColor = value
End Property

Public Property Get TargetColor() As Long
    TargetColor = this.TargetColor
End Property

Public Property Let TargetColor(ByVal value As Long)
    this.TargetColor = value
End Property

Public Property Get SigmaColor() As Long
    SigmaColor = this.SigmaColor
End Property

Public Property Let SigmaColor(ByVal value As Long)
    this.SigmaColor = value
End Property
