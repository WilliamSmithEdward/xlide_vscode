VERSION 5.00
Object = "{C5412DA5-2E2F-11D3-85BF-00105AC8B715}#1.0#0"; "isAnalogLibrary.ocx"
Object = "{BF3128D8-55B8-11D4-8ED4-00E07D815373}#1.0#0"; "MBPrgBar.ocx"
Begin VB.UserControl MixSlider 
   BackColor       =   &H00C0C0C0&
   BackStyle       =   0  'Transparent
   ClientHeight    =   1200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   645
   ScaleHeight     =   1200
   ScaleWidth      =   645
   Begin isAnalogLibrary.iSliderX iSliderX9 
      Height          =   975
      Left            =   130
      TabIndex        =   0
      Top             =   140
      Width           =   375
      EndsMargin      =   12
      PointerIndicatorInactiveColor=   0
      PointerIndicatorActiveColor=   255
      KeyArrowStepSize=   1
      KeyPageStepSize =   10
      Orientation     =   0
      OrientationTickMarks=   0
      PointerHeight   =   4
      PointerStyle    =   5
      PointerWidth    =   10
      ShowFocusRect   =   0   'False
      TrackColor      =   0
      TrackStyle      =   1
      TickMajorStyle  =   0
      TickMinorStyle  =   0
      BackGroundColor =   -16777201
      ShowTicksMajor  =   0   'False
      ShowTicksMinor  =   0   'False
      ShowTickLabels  =   0   'False
      TickMajorCount  =   5
      TickMajorColor  =   255
      TickMajorLength =   7
      TickMinorAlignment=   1
      TickMinorCount  =   4
      TickMinorColor  =   16777215
      TickMinorLength =   3
      TickMargin      =   5
      TickLabelMargin =   5
      BeginProperty TickLabelFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TickLabelPrecision=   0
      CurrentMax      =   0
      CurrentMin      =   0
      PositionPercent =   1.52590218966964E-03
      Position        =   100
      PositionMax     =   65535
      PositionMin     =   0
      BorderStyle     =   0
      Object.Visible         =   -1  'True
      Enabled         =   -1  'True
      TickLabelFontColor=   -16777208
      BackGroundPicture=   "MixSlider.ctx":0000
      MinMaxFixed     =   0   'False
      ReverseScale    =   0   'False
      Transparent     =   0   'False
      PrecisionStyle  =   1
      AutoScaleDesiredTicks=   5
      AutoScaleMaxTicks=   6
      AutoScaleEnabled=   0   'False
      AutoScaleStyle  =   0
      MouseControlStyle=   0
      UpdateFrameRate =   60
      OptionSaveAllProperties=   0   'False
      MouseWheelStepSize=   1
      AutoFrameRate   =   0   'False
      Object.Width           =   25
      Object.Height          =   65
      AutoCenter      =   0   'False
      OffsetX         =   0
      OffsetY         =   0
      PointerBitmap   =   "MixSlider.ctx":0DEA
      ShowDisabledState=   0   'False
      PointerFillEnabled=   0   'False
      PointerFillColor=   16711680
      TickLabelFontName=   "Tahoma"
      TickLabelFontSize=   8
      TickLabelFontBold=   0   'False
      TickLabelFontItalic=   0   'False
      TickLabelFontUnderline=   0   'False
      TickLabelFontStrikeOut=   0   'False
      OPCItemCount    =   0
   End
   Begin MBProgressBar.ProgressBar ProgressBar2 
      Height          =   1005
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   1773
      BorderStyle     =   2
      CaptionType     =   0
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackPicture     =   "MixSlider.ctx":1454
      BarPicture      =   "MixSlider.ctx":1470
   End
End
Attribute VB_Name = "MixSlider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Event Declarations:
Event OnPositionChange() 'MappingInfo=iSliderX9,iSliderX9,-1,OnPositionChange
Event OnPositionChangeFinished() 'MappingInfo=iSliderX9,iSliderX9,-1,OnPositionChangeFinished


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=iSliderX9,iSliderX9,-1,Position
Public Property Get Value() As Double
    Value = iSliderX9.Position
End Property

Public Property Let Value(ByVal New_Value As Double)
    iSliderX9.Position() = New_Value
    PropertyChanged "Value"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=iSliderX9,iSliderX9,-1,PositionMin
Public Property Get Min() As Double
    Min = iSliderX9.PositionMin
End Property

Public Property Let Min(ByVal New_Min As Double)
    iSliderX9.PositionMin() = New_Min
    PropertyChanged "Min"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=iSliderX9,iSliderX9,-1,PositionMax
Public Property Get Max() As Double
    Max = iSliderX9.PositionMax
End Property

Public Property Let Max(ByVal New_Max As Double)
    iSliderX9.PositionMax() = New_Max
    PropertyChanged "Max"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=iSliderX9,iSliderX9,-1,PositionPercent
Public Property Get Percentage() As Double
    Percentage = iSliderX9.PositionPercent
End Property

Public Property Let Percentage(ByVal New_Percentage As Double)
    iSliderX9.PositionPercent() = New_Percentage
    PropertyChanged "Percentage"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    iSliderX9.Position = PropBag.ReadProperty("Value", 100)
    iSliderX9.PositionMin = PropBag.ReadProperty("Min", 0)
    iSliderX9.PositionMax = PropBag.ReadProperty("Max", 65535)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Value", iSliderX9.Position, 100)
    Call PropBag.WriteProperty("Min", iSliderX9.PositionMin, 0)
    Call PropBag.WriteProperty("Max", iSliderX9.PositionMax, 65535)
End Sub

Private Sub iSliderX9_OnPositionChange()
    RaiseEvent OnPositionChange
End Sub

Private Sub iSliderX9_OnPositionChangeFinished()
    RaiseEvent OnPositionChangeFinished
End Sub

