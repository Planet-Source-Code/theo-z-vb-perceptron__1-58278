VERSION 5.00
Begin VB.UserControl ctlProgressBar 
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   300
   ScaleWidth      =   4800
   Begin VB.PictureBox picProgressBar 
      AutoRedraw      =   -1  'True
      ClipControls    =   0   'False
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   4695
      TabIndex        =   0
      Top             =   0
      Width           =   4755
   End
End
Attribute VB_Name = "ctlProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************
'** File Name   : ctlProgressBar.ctl                                          **
'** Language    : Visual Basic 6.0                                            **
'** Author      : Theo Zacharias (theo_yz@yahoo.com)                          **
'** Description : An ActiveX control to display progress bar with caption     **
'** Dependencies: -                                                           **
'** Members     :                                                             **
'**   * Properties : Appearance (r/w), BackColor (r/w), BorderStyle (r/w),    **
'**                  Caption (r/w), CaptionInteger (r/w), CaptionStyle (r/w), **
'**                  FillColor (r/w), Font (r/w), ForeColor (r/w), Max (r/w), **
'**                  Min (r/w), Value (r/w)                                   **
'**   * Methods    : -                                                        **
'**   * Events     : -                                                        **
'** Last modified on December 16, 2004                                        **
'*******************************************************************************

Option Explicit

'--- Public Enum Declaration
Public Enum genmAppearance
  Flat = 0
  ThreeD
End Enum

Public Enum genmBorderStyle
  None = 0
  FixedSingle
End Enum

Public Enum genmCaptionStyle
  NoCaption = 0
  Custom
  Percent
End Enum

'--- Property Variables
Private mblnCaptionInteger As Boolean
Private mlngFillColor As OLE_COLOR
Private mlngMax As Long
Private mlngMin As Long
Private msngValue As Single
Private mstrCaption As String
Private mudeCaptionStyle As genmCaptionStyle

'--- PropBag Names
Private Const mconAppearance As String = "Appearance"
Private Const mconBackColor As String = "BackColor"
Private Const mconBorderStyle As String = "BorderStyle"
Private Const mconCaption As String = "Caption"
Private Const mconCaptionInteger As String = "CaptionInteger"
Private Const mconCaptionStyle As String = "CaptionStyle"
Private Const mconFillColor As String = "FillColor"
Private Const mconFont As String = "Font"
Private Const mconForeColor As String = "ForeColor"
Private Const mconMax As String = "Max"
Private Const mconMin As String = "Min"
Private Const mconValue As String = "Value"

'--- Property Default Values
Private Const mconDefaultCaption As String = ""
Private Const mconDefaultCaptionInteger As Boolean = True
Private Const mconDefaultCaptionStyle As Long = Percent
Private Const mconDefaultFillColor As Long = vbBlue
Private Const mconDefaultMax As Long = 100
Private Const mconDefaultMin As Long = 0
Private Const mconDefaultValue As Long = 0

'-----------------------------------
' ActiveX Control Properties Events
'-----------------------------------

Private Sub UserControl_InitProperties()
  Max = mconDefaultMax
  Min = mconDefaultMin
  BackColor = UserControl.BackColor
  FillColor = mconDefaultFillColor
  CaptionStyle = mconDefaultCaptionStyle
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  With PropBag
    picProgressBar.Appearance = _
      .ReadProperty(Name:=mconAppearance, _
                    DefaultValue:=picProgressBar.Appearance)
    picProgressBar.BackColor = _
      .ReadProperty(Name:=mconBackColor, _
                    DefaultValue:=picProgressBar.BackColor)
    picProgressBar.BorderStyle = _
      .ReadProperty(Name:=mconBorderStyle, _
                    DefaultValue:=picProgressBar.BorderStyle)
    picProgressBar.ForeColor = _
      .ReadProperty(Name:=mconForeColor, _
                    DefaultValue:=picProgressBar.ForeColor)
    mblnCaptionInteger = .ReadProperty(Name:=mconCaptionInteger, _
                                       DefaultValue:=mconDefaultCaptionInteger)
    mlngFillColor = .ReadProperty(Name:=mconFillColor, _
                                  DefaultValue:=mconDefaultFillColor)
    mlngMax = .ReadProperty(Name:=mconMax, DefaultValue:=mconDefaultMax)
    mlngMin = .ReadProperty(Name:=mconMin, DefaultValue:=mconDefaultMin)
    msngValue = .ReadProperty(Name:=mconValue, DefaultValue:=mconDefaultValue)
    mstrCaption = .ReadProperty(Name:=mconCaption, _
                                DefaultValue:=mconDefaultCaption)
    mudeCaptionStyle = .ReadProperty(Name:=mconCaptionStyle, _
                                     DefaultValue:=mconDefaultCaptionStyle)
  End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  With PropBag
    .WriteProperty Name:=mconAppearance, Value:=picProgressBar.Appearance, _
                   DefaultValue:=picProgressBar.Appearance
    .WriteProperty Name:=mconBackColor, Value:=picProgressBar.BackColor, _
                   DefaultValue:=picProgressBar.BackColor
    .WriteProperty Name:=mconBorderStyle, Value:=picProgressBar.BorderStyle, _
                   DefaultValue:=picProgressBar.BorderStyle
    .WriteProperty Name:=mconCaption, Value:=mstrCaption, _
                   DefaultValue:=mconDefaultCaption
    .WriteProperty Name:=mconCaptionInteger, Value:=mblnCaptionInteger, _
                   DefaultValue:=mconDefaultCaptionInteger
    .WriteProperty Name:=mconCaptionStyle, Value:=mudeCaptionStyle, _
                   DefaultValue:=mconDefaultCaptionStyle
    .WriteProperty Name:=mconFillColor, Value:=mlngFillColor, _
                   DefaultValue:=mconDefaultFillColor
    .WriteProperty Name:=mconFont, Value:=UserControl.Font, _
                   DefaultValue:=Ambient.Font
    .WriteProperty Name:=mconForeColor, Value:=picProgressBar.ForeColor, _
                   DefaultValue:=picProgressBar.ForeColor
    .WriteProperty Name:=mconMax, Value:=mlngMax, DefaultValue:=mconDefaultMax
    .WriteProperty Name:=mconMin, Value:=mlngMin, DefaultValue:=mconDefaultMin
    .WriteProperty Name:=mconValue, Value:=msngValue, _
                   DefaultValue:=mconDefaultValue
  End With
End Sub

'-------------------------------
' Others ActiveX Control Events
'-------------------------------

Private Sub UserControl_Resize()
  picProgressBar.Width = UserControl.Width
  picProgressBar.Height = UserControl.Height
End Sub

'----------------------------
' ActiveX Control Properties
'----------------------------

' Purpose    : Sets whether an object is painted at run time with 3-D effects
' Input      : udeValue (the new Appearance property value)
Public Property Let Appearance(udeValue As genmAppearance)
  picProgressBar.Appearance = udeValue
  PropertyChanged mconAppearance
End Property

' Purpose    : Returns a value that determines whether an object is painted at
'              run time with 3-D effects
Public Property Get Appearance() As genmAppearance
  Appearance = picProgressBar.Appearance
End Property

' Purpose    : Sets the background color of the progress bar
' Input      : lngBackColor (the new BackColor property value)
Public Property Let BackColor(lngBackColor As OLE_COLOR)
  picProgressBar.BackColor = lngBackColor
  PropertyChanged mconBackColor
End Property

' Purpose    : Returns the background color of the progress bar
Public Property Get BackColor() As OLE_COLOR
  BackColor = picProgressBar.BackColor
End Property

' Purpose    : Sets the border style of the progress bar
' Input      : udeBorderStyle (the new BorderStyle property value)
Public Property Let BorderStyle(udeBorderStyle As genmBorderStyle)
  picProgressBar.BorderStyle = udeBorderStyle
  PropertyChanged mconBorderStyle
End Property

' Purpose    : Returns the border style of the progress bar
Public Property Get BorderStyle() As genmBorderStyle
  BorderStyle = picProgressBar.BorderStyle
End Property

' Purpose    : Sets the caption to be displayed on the progress bar
' Input      : strCaption (the new Caption property value)
' Note       : The caption is displayed only if CaptionStyle is set to 'Custom'
Public Property Let Caption(strCaption As String)
  mstrCaption = strCaption
  PropertyChanged mconCaption
End Property

' Purpose    : Returns the caption to be displayed on the progress bar
Public Property Get Caption() As String
  Caption = mstrCaption
End Property

' Purpose    : Sets a value that determines whether the percentage value is
'              converted to integer value before displayed on the progress bar
' Input      : blnCaptionInteger (the new CaptionInteger property value)
Public Property Let CaptionInteger(blnCaptionInteger As Boolean)
  mblnCaptionInteger = blnCaptionInteger
  PropertyChanged mconCaptionInteger
End Property

' Purpose    : Returnss a value that determines whether the percentage value is
'              converted to integer value before displayed on the progress bar
Public Property Get CaptionInteger() As Boolean
  CaptionInteger = mblnCaptionInteger
End Property

' Purpose    : Sets the caption style
' Input      : udeCaptionStyle (the new CaptionStyle property value)
Public Property Let CaptionStyle(udeCaptionStyle As genmCaptionStyle)
  mudeCaptionStyle = udeCaptionStyle
  PropertyChanged mconCaptionStyle
End Property

' Purpose    : Returns the caption style
Public Property Get CaptionStyle() As genmCaptionStyle
  CaptionStyle = mudeCaptionStyle
End Property

' Purpose    : Sets the fill color of the progress bar
' Input      : lngFillColor (the new FillColor property value)
Public Property Let FillColor(lngFillColor As OLE_COLOR)
  mlngFillColor = lngFillColor
  PropertyChanged mconFillColor
End Property

' Purpose    : Returns the fill color of the progress bar
Public Property Get FillColor() As OLE_COLOR
  FillColor = mlngFillColor
End Property

' Purpose    : Sets the font object used to display texts on the progress bar
' Input      : fntFont (the new Font property value)
' Effects    : All controls' font object in the progress bar have been set to
'              fntFont
Public Property Set Font(fntFont As Font)
  Dim obj As Object
    
  Set UserControl.Font = fntFont
  For Each obj In Controls
    Set obj.Font = UserControl.Font
  Next
  PropertyChanged mconFont
End Property

' Purpose    : Returns the font object used to display texts on the progress bar
Public Property Get Font() As Font
  Set Font = UserControl.Font
End Property

' Purpose    : Sets the fore color of the progress bar
' Input      : lngForeColor (the new ForeColor property value)
Public Property Let ForeColor(lngForeColor As OLE_COLOR)
  picProgressBar.ForeColor = lngForeColor
  PropertyChanged mconForeColor
End Property

' Purpose    : Returns the fore color of the progress bar
Public Property Get ForeColor() As OLE_COLOR
  ForeColor = picProgressBar.ForeColor
End Property

' Purpose    : Sets the progress bar's maximum value
' Input      : lngMax (the new Max property value)
Public Property Let Max(lngMax As Long)
  mlngMax = lngMax
  PropertyChanged mconMax
End Property

' Purpose    : Returns the progress bar's maximum value
Public Property Get Max() As Long
  Max = mlngMax
End Property

' Purpose    : Sets the progress bar's minimum value
' Input      : lngMin (the new Min property value)
Public Property Let Min(lngMin As Long)
  mlngMin = lngMin
  PropertyChanged mconMin
End Property

' Purpose    : Returns the progress bar's minimum value
Public Property Get Min() As Long
  Min = mlngMin
End Property

' Purpose    : Sets the progress bar's value
' Input      : sngValue (the new Value property value)
' Effects    : * The progress bar has been updated
'            : * If CaptionStyle is set to 'Percent', the progress bar's caption
'                has been updated
Public Property Let Value(sngValue As Single)
  msngValue = sngValue
  UpdateProgressBar sngValue:=sngValue
  PropertyChanged mconValue
End Property

' Purpose    : Returns the progress bar's value
Public Property Get Value() As Single
  Value = msngValue
End Property

'----------------------------------
' Private Functions and Procedures
'----------------------------------

' Purpose    : Update the progress bar interface with the new Value property
' Input      : sngValue (the new Value property of the progres bar)
Private Sub UpdateProgressBar(sngValue As Single)
  Dim strCaption As String               'the text displayed on the progress bar

  If sngValue > mlngMax Then
    sngValue = mlngMax
  ElseIf sngValue < mlngMin Then
    sngValue = mlngMin
  End If
    
  Select Case mudeCaptionStyle
    Case NoCaption
      strCaption = ""
    Case Custom
      strCaption = mstrCaption
    Case Percent
      strCaption = _
        Format(((sngValue - mlngMin) / (mlngMax - mlngMin)) * 100, _
               IIf(mblnCaptionInteger, "0", "0.00;;0")) & "%"
  End Select
  
  With picProgressBar
    .Cls
    .ScaleWidth = mlngMax - mlngMin
    .CurrentX = (.ScaleWidth / 2) - (.TextWidth(strCaption) / 2)
    .CurrentY = (.ScaleHeight - .TextHeight(strCaption)) / 2
    .DrawMode = vbNotXorPen
    picProgressBar.Print strCaption
    picProgressBar.Line (0, 0)-((sngValue - mlngMin), .Width), mlngFillColor, BF
  End With
End Sub
