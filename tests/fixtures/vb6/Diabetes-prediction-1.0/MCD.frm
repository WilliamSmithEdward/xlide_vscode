VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Demo - diabetes prediction prototype"
   ClientHeight    =   9075
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16725
   BeginProperty Font 
      Name            =   "Courier"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00404040&
   LinkTopic       =   "Form1"
   Picture         =   "MCD.frx":0000
   ScaleHeight     =   605
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1115
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Center_patt 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   11520
      ScaleHeight     =   311
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   335
      TabIndex        =   31
      Top             =   2880
      Width           =   5055
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         X1              =   168
         X2              =   168
         Y1              =   0
         Y2              =   304
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   336
         Y1              =   152
         Y2              =   152
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Processes k steps"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   27
      Top             =   6840
      Width           =   5175
      Begin VB.CommandButton Solve_n 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Analyze"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox sntext 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4440
         TabIndex        =   28
         Text            =   "20"
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Number of days for prediction ="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2160
         TabIndex        =   30
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Step by step"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   22
      Top             =   7920
      Width           =   5175
      Begin VB.CheckBox Anim_Step 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Animate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   24
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox ASS 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   23
         Text            =   "200"
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "ms"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1845
         TabIndex        =   26
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Animation step"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Patient blood sugar for a period of time (days)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   360
      TabIndex        =   21
      Top             =   4080
      Width           =   5295
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Extract limits from the input data"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   2280
         Value           =   1  'Checked
         Width           =   4935
      End
      Begin VB.TextBox LdT 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3720
         TabIndex        =   36
         Text            =   "60"
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox LuT 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3720
         TabIndex        =   34
         Text            =   "200"
         ToolTipText     =   "Upper glycemic limit"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Gly 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1845
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Text            =   "MCD.frx":92E0
         ToolTipText     =   "Preferably one or several measurements taken each day at equal time intervals."
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Ld (mg/dL)"
         Height          =   255
         Left            =   3720
         TabIndex        =   35
         ToolTipText     =   "Down glycemic limit"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Lu (mg/dL)"
         Height          =   255
         Left            =   3720
         TabIndex        =   33
         ToolTipText     =   "Upper glycemic limit"
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   5760
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   20
      Top             =   2640
      Width           =   5415
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Probability values of the last vector:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   5760
      TabIndex        =   15
      Top             =   120
      Width           =   5415
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "L - low blood sugar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   19
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "H - high blood sugar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   18
         Top             =   360
         Width           =   2055
      End
      Begin VB.Shape top_graph 
         Height          =   1935
         Left            =   240
         Top             =   360
         Width           =   15
      End
      Begin VB.Shape Yp 
         BackColor       =   &H00C0C0FF&
         BackStyle       =   1  'Opaque
         Height          =   1575
         Left            =   2760
         Top             =   720
         Width           =   2055
      End
      Begin VB.Shape Xp 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   1  'Opaque
         FillColor       =   &H00FFFFFF&
         Height          =   1575
         Left            =   480
         Top             =   720
         Width           =   2055
      End
      Begin VB.Line Line8 
         X1              =   240
         X2              =   5280
         Y1              =   2280
         Y2              =   2280
      End
   End
   Begin VB.PictureBox graf_val 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   11520
      ScaleHeight     =   111
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   335
      TabIndex        =   9
      ToolTipText     =   "Predicting the behavior of glucose levels over the next period (days in this case but other system of reference can be taken)"
      Top             =   480
      Width           =   5055
   End
   Begin VB.TextBox v1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      TabIndex        =   6
      Text            =   "0"
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox v2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Text            =   "1"
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox P22 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3000
      TabIndex        =   3
      Text            =   "0.4"
      Top             =   3120
      Width           =   495
   End
   Begin VB.TextBox P21 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2400
      TabIndex        =   2
      Text            =   "0.6"
      Top             =   3120
      Width           =   495
   End
   Begin VB.TextBox P12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3000
      TabIndex        =   1
      Text            =   "0.2"
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox P11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2400
      TabIndex        =   0
      Text            =   "0.8"
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Prediction of behavior:"
      Height          =   255
      Left            =   11520
      TabIndex        =   38
      Top             =   240
      Width           =   4455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Vector components - probability plot:"
      Height          =   255
      Left            =   11520
      TabIndex        =   37
      Top             =   2640
      Width           =   4455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "High"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   17
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Low"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   16
      Top             =   1200
      Width           =   735
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00808080&
      Height          =   3735
      Left            =   360
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label sn 
      BackStyle       =   0  'Transparent
      Caption         =   "k"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   14
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label L12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   13
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label L21 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   12
      Top             =   360
      Width           =   615
   End
   Begin VB.Label L22 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label L11 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5160
      TabIndex        =   10
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label y 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4680
      TabIndex        =   8
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label x 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   4080
      TabIndex        =   7
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   2880
      Width           =   255
   End
   Begin VB.Shape Anim_S1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   5
      FillColor       =   &H0080C0FF&
      Height          =   855
      Left            =   3720
      Shape           =   3  'Circle
      Top             =   840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Shape Anim_S0 
      BackColor       =   &H80000003&
      BorderColor     =   &H000080FF&
      BorderWidth     =   5
      FillColor       =   &H0080C0FF&
      Height          =   855
      Left            =   1320
      Shape           =   3  'Circle
      Top             =   840
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   ________________________________                          ____________________
'  /  Markov chains                 \________________________/       v3.00        |
' |                                                                               |
' |            Name:  Markov Chains Diabetes                                      |
' |        Category:  open source software                                        |
' |          Author:  Paul A. Gagniuc                                             |
' |            Book:  Markov Chains: From Theory to                               |
' |                   Implementation and Experimentation                          |
' |                                                                               |
' |    Date Created:  November 2013                                               |
' |          Update:  September 2021                                              |
' |       Tested On:  WinXP, WinVista, Win7, Win8, Win10                          |
' |           Email:  paul_gagniuc@acad.ro                                        |
' |             Use:  Markov chains example                                       |
' |                                                                               |
' |                  _____________________________                                |
' |_________________/                             \_______________________________|
'
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim Gtr As Variant
Dim M(1 To 2, 1 To 2) As String


Private Sub Check1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If LuT.Enabled = False Then LuT.Enabled = True Else LuT.Enabled = False
    If LdT.Enabled = False Then LdT.Enabled = True Else LdT.Enabled = False
End Sub


Function gli_limits()

    Dim Inp() As String
    Dim fu, fd As Variant
    Dim i As Integer
    fu = 0
    fd = 0
    
    Inp = Split(Gly.Text, ",")
    
        For i = 0 To UBound(Inp)
            If Val(Inp(i)) > Val(fu) Then fu = Inp(i)
        Next i
        
     fd = fu
        For i = 0 To UBound(Inp)
            If Val(Inp(i)) < Val(fd) Then fd = Inp(i)
        Next i
        
    LuT.Text = Val(fu) + 1
    LdT.Text = Val(fd) - 1
    
End Function

Private Sub Form_Load()

    If ItIsWin7 Then
        ' Is Win7 or above
        Anim_Step.Value = 1
    Else
        Anim_Step.Value = 1 ' Is Win XP or lower, preserve colors whether we are on Win7,8 or XP
        Frame4.BackColor = &HF0F0F0
        Frame1.BackColor = &HF0F0F0
        Frame3.BackColor = &HF0F0F0
        Frame5.BackColor = &HF0F0F0
        Anim_Step.BackColor = &HF0F0F0
    End If
    
    L12.Caption = Round(P12.Text, 2)
    L11.Caption = Round(P11.Text, 2)
    L21.Caption = Round(P21.Text, 2)
    L22.Caption = Round(P22.Text, 2)
    sn.Caption = sntext.Text
    
    Call draw_scale(3)
    
End Sub


Private Sub sntext_Change()
    sn.Caption = sntext.Text
End Sub


Function check_1()
    If Val(P11.Text) + Val(P12.Text) <> 1 Then MsgBox "Row 1 does not add up to 1 ! Probabilities on each row must add up to 1."
    If Val(P21.Text) + Val(P22.Text) <> 1 Then MsgBox "Row 2 does not add up to 1 ! Probabilities on each row must add up to 1."
End Function


Private Sub Solve_n_Click()

    Dim oldxx, oldyy, xx, yy, xxc, yyc, oldn As Variant
    Dim cicle, i As Integer
    
    v1.Text = 0
    v2.Text = 1
    
    Solve_n.Enabled = False
    Text1.Text = Empty
    Center_patt.Cls
    graf_val.Cls
    Call check_1
    
    oldxx = 0
    oldyy = 0
    
    cicle = Val(sntext.Text)
    
    Call ExtractProb(Transform(Gly.Text))
    
    Call draw_scale(cicle)
    
    For i = 0 To cicle
    
        x.Caption = (Val(v1.Text) * Val(P11.Text)) + (Val(v2.Text) * Val(P21.Text))
        y.Caption = (Val(v1.Text) * Val(P12.Text)) + (Val(v2.Text) * Val(P22.Text))
    
        If (v1.Text = x.Caption And v2.Text = y.Caption) Then
            Text1.Text = Text1.Text & "At [" & i & "] is the steady state vector !" & vbCrLf & vbCrLf
            
            Text1.Text = Text1.Text & "Considering the upper and lower limits (Lu=" & LuT.Text & " mg/dL and Ld=" & LdT.Text & " mg/dL) the threshold glucose level for this patient was calculated at " & _
            Gtr & " mg/dL (representing the half between of the two limits)" & vbCrLf & vbCrLf
    
            Text1.Text = Text1.Text & "Elevated levels or low levels of blood sugar are considered according to this threshold (" & Gtr & " mg/dL)." & vbCrLf & vbCrLf
            
            Text1.Text = Text1.Text & "Therefore, in the future the patient will have a LOW blood sugar (under " & Gtr & " mg/dL) about " & _
            Round(100 * Val(x.Caption), 2) & "% of the time, and a HIGH blood sugar (above " & Gtr & " mg/dL) about " & Round(100 * Val(y.Caption), 2) & "% of the time." & vbCrLf & vbCrLf
            
            
            Text1.Text = Text1.Text & "Patient's glycemic events indicate the following observations: if the patient has a high " & _
            "blood sugar it returns to a high blood sugar " & Round(100 * Val(P22.Text), 2) & "% of the time, and if it has a low " & _
            "blood sugar it returns to a low blood sugar level " & Round(100 * Val(P11.Text), 2) & "% of the time." & vbCrLf & vbCrLf
    
            Text1.Text = Text1.Text & "If the patient has a HIGH " & _
            "blood sugar it moves to a LOW blood sugar level " & Round(100 * Val(P21.Text), 2) & "% of the time, and if it has a LOW " & _
            "blood sugar it moves to a HIGH blood sugar level " & Round(100 * Val(P12.Text), 2) & "% of the time." & vbCrLf
            
            i = cicle
        Else
            v1.Text = x.Caption
            v2.Text = y.Caption
            '------------------------------------- Animate
            If Anim_Step.Value = 1 Then
                
                If Val(x.Caption) > Val(y.Caption) Then
                    Anim_S0.Visible = True
                    Anim_S1.Visible = False
                Else
                    Anim_S0.Visible = False
                    Anim_S1.Visible = True
                End If
    
                Call bar_function(x.Caption, y.Caption)
                Sleep (CLng(ASS.Text))
            End If
            '-------------------------------------
            
            If Val(x.Caption) > Val(y.Caption) Then
                Text1.Text = Text1.Text & "L[" & i + 1 & "] = [" & x.Caption & " - " & y.Caption & "]" & vbCrLf
            Else
                Text1.Text = Text1.Text & "H[" & i + 1 & "] = [" & x.Caption & " - " & y.Caption & "]" & vbCrLf
            End If
        End If
    
        xx = (graf_val.ScaleHeight / 100) * (100 * Val(x.Caption))
        yy = (graf_val.ScaleHeight / 100) * (100 * Val(y.Caption))
    
        xxc = (Center_patt.ScaleWidth / 100) * (100 * Val(x.Caption))
        yyc = (Center_patt.ScaleHeight / 100) * (100 * Val(y.Caption))
        Center_patt.Circle (xxc, Center_patt.ScaleHeight - yyc), 3, vbRed
    
        If i > 1 Then
            graf_val.Line (oldn, oldyy)-((graf_val.ScaleWidth / cicle) * i, yy), vbBlue
            graf_val.Line (oldn, oldxx)-((graf_val.ScaleWidth / cicle) * i, xx), vbRed
        End If
    
        oldn = (graf_val.ScaleWidth / cicle) * i
    
        oldxx = xx
        oldyy = yy
    
        DoEvents
    
    Next i
    
    Solve_n.Enabled = True
    
End Sub

Function draw_scale(ByVal k_stat As Integer)

    Dim zx, qx, zy, qy As Variant
    Dim sp As Variant
    Dim i As Integer
    
    Form1.Cls
    
    'X axis on graf_val OBJ
    '-------------------------------------
    sp = graf_val.ScaleWidth / k_stat
    For i = 0 To k_stat
    
        zx = graf_val.Left + (sp * i)
        qx = zx
        zy = graf_val.Top + graf_val.ScaleHeight
        qy = graf_val.Top + graf_val.ScaleHeight + 6
    
        If k_stat < 10 Then
            Form1.CurrentX = zx - 6
            Form1.CurrentY = qy
            Form1.Print "S" & i
        End If
    
        Form1.Line (zx, zy)-(qx, qy), &H808080
    
    Next i
    '-------------------------------------
    
    'Y axis on graf_val OBJ
    '-------------------------------------
        zx = graf_val.Left - 6
        qx = graf_val.Left
        zy = graf_val.Top
        qy = zy
        Form1.Line (zx, zy)-(qx, qy), &H808080
        Form1.CurrentX = zx - 7
        Form1.CurrentY = qy - 6
        Form1.Print "1"
    
        zx = graf_val.Left - 6
        qx = graf_val.Left
        zy = graf_val.Top + graf_val.ScaleHeight
        qy = zy
        Form1.Line (zx, zy)-(qx, qy), &H808080
        Form1.CurrentX = zx - 7
        Form1.CurrentY = qy - 6
        Form1.Print "0"
    '-------------------------------------
    
    'X axis on Center_patt OBJ
    '-------------------------------------
    sp = Center_patt.ScaleWidth / 4
    For i = 0 To 4
    
        zx = Center_patt.Left + (sp * i)
        qx = zx
        zy = Center_patt.Top + Center_patt.ScaleHeight
        qy = Center_patt.Top + Center_patt.ScaleHeight + 6
        Form1.CurrentX = zx - 10
        Form1.CurrentY = qy
        
        If i = 0 Then Form1.Print 0
        
        If i = 1 Then
            Form1.Print ".25"
        End If
        
        If i = 2 Then
            Form1.Print ".5"
        End If
        
        If i = 3 Then
            Form1.Print ".75"
        End If
        
        If i = 4 Then Form1.Print 1
        
        Form1.Line (zx, zy)-(qx, qy), &H808080
    
    Next i
    '-------------------------------------
    
    'Y axis on Center_patt OBJ
    '-------------------------------------
    sp = Center_patt.ScaleHeight / 4
    For i = 0 To 4
    
        zx = Center_patt.Left - 6
        qx = Center_patt.Left
        zy = Center_patt.Top + (sp * i)
        qy = zy
        Form1.CurrentX = zx - 25
        Form1.CurrentY = qy - 6
        
        If i = 4 Then
            Form1.CurrentX = zx - 16
            Form1.Print 0
        End If
        
        If i = 3 Then
            Form1.Print ".25"
        End If
        
        If i = 2 Then
            Form1.CurrentX = zx - 16
            Form1.Print ".5"
        End If
        
        If i = 1 Then
            Form1.Print ".75"
        End If
        
        If i = 0 Then
            Form1.CurrentX = zx - 16
            Form1.Print 1
        End If
        
        Form1.Line (zx, zy)-(qx, qy), &H808080
    
    Next i
    '-------------------------------------
End Function


Function bar_function(ByVal x As String, ByVal y As String)
    Xp.Height = (top_graph.Height / 100) * (x * 100)
    Yp.Height = (top_graph.Height / 100) * (y * 100)
    Xp.Top = top_graph.Top + (top_graph.Height - Xp.Height)
    Yp.Top = top_graph.Top + (top_graph.Height - Yp.Height)
End Function



Private Sub P12_Change()
    L12.Caption = Round(P12.Text, 2)
End Sub

Private Sub P11_Change()
    L11.Caption = Round(P11.Text, 2)
End Sub

Private Sub P21_Change()
    L21.Caption = Round(P21.Text, 2)
End Sub

Private Sub P22_Change()
    L22.Caption = Round(P22.Text, 2)
End Sub


Function Transform(ByVal R As String) As String

    Dim Inp() As String
    Dim s, l, Obs, Reg As String
    Dim Lu, Ld, n, i As Integer
    Dim Pr As Variant
    
    Inp = Split(R, ",")
    
    If Check1.Value = 1 Then
    Call gli_limits
    End If
    
    Lu = Val(LuT.Text)
    Ld = Val(LdT.Text)
    
    n = 2
    
    Pr = (Lu - Ld) / n
    Gtr = Ld + Pr * (n - 1)
    
    For i = 0 To UBound(Inp)
    
        s = (Inp(i) - Ld) / Pr
        s = Split(s, ".")(0)
    
        If s = 0 Then l = "A"
        If s = 1 Then l = "B"
        If s = 2 Then l = "C"
        If s = 3 Then l = "D"
    
        Obs = Obs & l
    
    Next i
    
    Transform = Obs
    
End Function


Function ExtractProb(ByVal s As String)

    Dim Eb, Es, DI1, DI2 As String
    Dim i, j, R, c As Integer
    Dim TB, TS As Variant
    
    Eb = "A"
    Es = "B"
    
    For i = 1 To 2
        For j = 1 To 2
          M(i, j) = 0
        Next j
    Next i
    
    TB = 0
    TS = 0
    
    For i = 2 To Len(s) - 1
            DI1 = Mid(s, i, 1)
            DI2 = Mid(s, i + 1, 1)
    
            If DI1 = Eb Then R = 1
            If DI1 = Es Then R = 2
            
            If DI2 = Eb Then c = 1
            If DI2 = Es Then c = 2
    
            M(R, c) = Val(M(R, c)) + 1
    
            If DI1 = Eb Then TB = TB + 1
            If DI1 = Es Then TS = TS + 1
    Next i
    
    'Text1.Text = Text1.Text & DrowMatrix(2, 2, M, "(C)", "Count:")
    
    For i = 1 To 2
        For j = 1 To 2
           If i = 1 Then M(i, j) = Val(M(i, j)) / TB
           If i = 2 Then M(i, j) = Val(M(i, j)) / TS
        Next j
    Next i
    
    'Text1.Text = Text1.Text & DrowMatrix(2, 2, M, "(P)", "Transition matrix M:")
    
    P11.Text = M(1, 1)
    P12.Text = M(1, 2)
    P21.Text = M(2, 1)
    P22.Text = M(2, 2)
    
End Function


Function DrowMatrix(ib, jb, ByVal M As Variant, ByVal model As String, ByVal msg As String) As String

    Dim Eb, Es, ct, y, u, v, o As String
    Dim i, j  As Integer
    Dim TB, TS As Variant
    
    Eb = "A"
    Es = "B"
    
    y = "|___|___|___|"
    ct = ct & vbCrLf & "____________"
    ct = ct & vbCrLf & "| " & model & " |  " & Eb & "  |  " & Es & "  | "
    ct = ct & vbCrLf & y & vbCrLf
    
    For i = 1 To ib
        For j = 1 To jb
        
        v = Round(M(i, j), 2)
        
            If Len(v) = 0 Then u = "|     "
            If Len(v) = 1 Then u = "|    "
            If Len(v) = 2 Then u = "|   "
            If Len(v) = 3 Then u = "|  "
            If Len(v) = 4 Then u = "| "
            If Len(v) = 5 Then u = "|"
            
            If j = jb Then o = "|" Else o = ""
            If j = 1 And i = 1 Then ct = ct & "|  " & Eb & "  "
            If j = 1 And i = 2 Then ct = ct & "|  " & Es & "  "
            
            ct = ct & u & v & o
            
        Next j
    
    ct = ct & vbCrLf & y & vbCrLf
    
    Next i
    
    DrowMatrix = msg & " M[" & Val(jb) & "," & Val(ib) & "]" & vbCrLf & ct & vbCrLf

End Function

