VERSION 5.00
Begin VB.UserControl Hyperlink 
   ClientHeight    =   1440
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3150
   ScaleHeight     =   1440
   ScaleWidth      =   3150
   ToolboxBitmap   =   "Hyperlink.ctx":0000
   Begin VB.Timer tmrMouse 
      Left            =   1980
      Top             =   450
   End
   Begin VB.Label lblText 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hyperlink Text"
      Height          =   195
      Left            =   270
      MouseIcon       =   "Hyperlink.ctx":0312
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   360
      Width           =   1020
   End
End
Attribute VB_Name = "Hyperlink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'
''''''''''''''''''''''''''''''''''''''''''
'
' This control is created by Faraz Azhar
'
''''''''''''''''''''''''''''''''''''''''''
'
Private Const SW_SHOWNORMAL As Long = 1

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" _
  (lpPoint As POINTAPI) As Long

Private Declare Function ScreenToClient Lib "user32" _
  (ByVal hwnd As Long, _
   lpPoint As POINTAPI) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function ShellExecute Lib "shell32" _
   Alias "ShellExecuteA" _
  (ByVal hwnd As Long, _
   ByVal lpOperation As String, _
   ByVal lpFile As String, _
   ByVal lpParameters As String, _
   ByVal lpDirectory As String, _
   ByVal nShowCmd As Long) As Long

' Link property
Public Enum hplResponse
    hplOpenURL
    hplClickEvent
End Enum
#If False Then
    Dim hplOpenURL
    Dim hplClickEvent
#End If

' Button states
Private Enum ButtonState
    stNormal
    stHot
    stDown
End Enum
#If False Then
    Dim stNormal
    Dim stHot
    Dim stDown
#End If

' Colors of the hyperlink
Private CLR_NORMAL  As OLE_COLOR
Private CLR_HOT     As OLE_COLOR
Private CLR_DOWN    As OLE_COLOR

' Events
Public Event Click()
Attribute Click.VB_UserMemId = -600
Public Event OpeningURL(URL As String)

' Properties
Private m_URL       As String
Private m_Response  As hplResponse
Private m_Underline As Boolean

' Misc vars
Private State       As ButtonState

Private Sub lblText_Click()
    ' User clicked on the link, see what is to be done.
    If m_Response = hplClickEvent Then
        ' Inform the main window.
        RaiseEvent Click
    Else
        ' URL is specified so we open the URL by ourselves.
        If m_URL = "" Then m_URL = lblText.Caption
        RaiseEvent OpeningURL(m_URL)
        ShellExecute GetDesktopWindow(), "open", m_URL, 0&, 0&, SW_SHOWNORMAL
    End If
End Sub

Private Sub lblText_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' when the label is clicked, change
    ' the colour to indicate it is down
    ChangeState stDown
End Sub

Private Sub lblText_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' if not already highlighted, set the
    ' label colour and start the timer to
    ' poll for the mouse cursor position
    With lblText
        If State = stNormal Then
            ChangeState stHot
            tmrMouse.Interval = 50
            tmrMouse.Enabled = True
        End If
    End With
End Sub

Private Sub lblText_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' mouse released, so restore the label to normal
    ChangeState stNormal
End Sub

Private Sub UserControl_Initialize()
    '
    CLR_DOWN = vbBlue
    CLR_NORMAL = vbBlack
    CLR_HOT = vbRed
    m_Underline = True
    '
    With lblText
        .Caption = "Hyperlink"
        .Move 0, 0
        .ForeColor = CLR_NORMAL
        UserControl.Width = .Width
        UserControl.Height = .Height
   End With
End Sub

Private Sub tmrMouse_Timer()
    '
    Dim Pt As POINTAPI
    Dim x As Long
    Dim y As Long
    Dim lLeft As Long
    Dim lTop As Long
    '
    With UserControl
        '
        lLeft = .Extender.Left
        lTop = .Extender.Top
        '
        GetCursorPos Pt
        ScreenToClient .ContainerHwnd, Pt
        '
        x = Pt.x * Screen.TwipsPerPixelX
        y = Pt.y * Screen.TwipsPerPixelY
        '
        If (x < lLeft) Or (x > (.Width + lLeft)) Or _
           (y < lTop) Or (y > (.Height + lTop)) Then
            '
            'the cursor has moved outside, so
            'reset the label appearance
            If State = stHot Then ChangeState stNormal
            '
            ' and disable the timer
            tmrMouse.Enabled = False
            '
        End If
        '
    End With
    '
End Sub

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/Sets the caption to display for the Hyperlink."
Attribute Caption.VB_UserMemId = -518
    Caption = lblText.Caption
End Property

Public Property Let Caption(ByVal vNewValue As String)
    lblText.Caption = vNewValue
    ' resize control
     Call UserControl_Resize
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Caption = PropBag.ReadProperty("Caption", "Hyperlink")
    URL = PropBag.ReadProperty("URL", "")
    ClickResponse = PropBag.ReadProperty("ClickResponse", hplOpenURL)
    BackColor = PropBag.ReadProperty("BackColor", vbButtonFace)
    ToolTipText2 = PropBag.ReadProperty("ToolTipText2", "")
    
    ColorNormal = PropBag.ReadProperty("ColorNormal", vbBlack)
    ColorHot = PropBag.ReadProperty("ColorHot", vbRed)
    ColorDown = PropBag.ReadProperty("ColorDown", vbBlue)
    Set lblText.Font = PropBag.ReadProperty("Font", UserControl.Font)
    HoverUnderline = PropBag.ReadProperty("HoverUnderline", True)
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = lblText.Width
    UserControl.Height = lblText.Height
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Caption", Caption, "Hyperlink"
    PropBag.WriteProperty "URL", URL, ""
    PropBag.WriteProperty "ClickResponse", ClickResponse, hplOpenURL
    PropBag.WriteProperty "BackColor", BackColor, vbButtonFace
    PropBag.WriteProperty "ToolTipText2", ToolTipText2, ""
    
    PropBag.WriteProperty "ColorNormal", ColorNormal, vbBlack
    PropBag.WriteProperty "ColorHot", ColorHot, vbRed
    PropBag.WriteProperty "ColorDown", ColorDown, vbBlue
    PropBag.WriteProperty "Font", lblText.Font, UserControl.Font
    PropBag.WriteProperty "HoverUnderline", HoverUnderline, True
End Sub

Public Property Get URL() As String
Attribute URL.VB_Description = "Returns/Sets the URL to be opened when user clicks on the Hyperlink and ClickResponse is set to OpenURL."
    URL = m_URL
End Property

Public Property Let URL(ByVal vNewValue As String)
    m_URL = vNewValue
    PropertyChanged "URL"
End Property

Public Property Get ClickResponse() As hplResponse
Attribute ClickResponse.VB_Description = "Returns/Sets what action to perform when user clicks on the Hyperlink."
    ClickResponse = m_Response
End Property

Public Property Let ClickResponse(ByVal vNewValue As hplResponse)
    m_Response = vNewValue
    PropertyChanged "ClickResponse"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/Sets the backcolor of the Hyperlink."
Attribute BackColor.VB_UserMemId = -501
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal vNewValue As OLE_COLOR)
    UserControl.BackColor = vNewValue
    PropertyChanged "BackColor"
End Property

Public Property Get ColorNormal() As OLE_COLOR
Attribute ColorNormal.VB_Description = "Specifies the color of the Hyperlink."
Attribute ColorNormal.VB_UserMemId = -513
    ColorNormal = CLR_NORMAL
End Property

Public Property Let ColorNormal(ByVal vNewValue As OLE_COLOR)
    CLR_NORMAL = vNewValue
    PropertyChanged "ColorNormal"
    ChangeState State
End Property

Public Property Get ColorHot() As OLE_COLOR
Attribute ColorHot.VB_Description = "Specifies the color of the Hyperlink when the user hovers the mouse over it."
    ColorHot = CLR_HOT
End Property

Public Property Let ColorHot(ByVal vNewValue As OLE_COLOR)
    CLR_HOT = vNewValue
    PropertyChanged "ColorHot"
    ChangeState State
End Property

Public Property Get ColorDown() As OLE_COLOR
Attribute ColorDown.VB_Description = "Specifies the color of the Hyperlink when the user presses it."
    ColorDown = CLR_DOWN
End Property

Public Property Let ColorDown(ByVal vNewValue As OLE_COLOR)
    CLR_DOWN = vNewValue
    PropertyChanged "ColorDown"
    ChangeState State
End Property

Public Property Get ToolTipText2() As String
Attribute ToolTipText2.VB_Description = "Returns/Sets the tooltip text for the Hyperlink."
    ToolTipText2 = lblText.ToolTipText
End Property

Public Property Let ToolTipText2(ByVal vNewValue As String)
    lblText.ToolTipText = vNewValue
End Property

Public Property Get Font() As StdFont
Attribute Font.VB_Description = "Returns/Sets the font styles for the Hyperlink."
Attribute Font.VB_UserMemId = -512
    Set Font = lblText.Font
End Property

Public Property Set Font(ByVal vNewValue As StdFont)
    Set lblText.Font = vNewValue
    Me.Caption = lblText.Caption    ' to resize UserControl
    PropertyChanged "Font"
End Property

Private Sub ChangeState(lState As ButtonState)
'    If lState = State Then Exit Sub
    '
    Select Case lState
        '
        Case stNormal
            lblText.ForeColor = CLR_NORMAL
            lblText.FontUnderline = False
            '
        Case stHot
            lblText.ForeColor = CLR_HOT
            lblText.FontUnderline = m_Underline
            '
        Case stDown
            lblText.ForeColor = CLR_DOWN
            lblText.FontUnderline = m_Underline
            '
    End Select
    '
    lblText.Refresh
    State = lState
End Sub

Public Property Get HoverUnderline() As Boolean
    HoverUnderline = m_Underline
End Property

Public Property Let HoverUnderline(ByVal vNewValue As Boolean)
    m_Underline = vNewValue
    ChangeState State
    PropertyChanged "HoverUnderline"
End Property
