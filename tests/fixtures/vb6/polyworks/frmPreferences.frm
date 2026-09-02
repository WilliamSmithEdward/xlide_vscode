VERSION 5.00
Begin VB.Form frmPreferences 
   BackColor       =   &H004A3C31&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8175
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   8775
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   545
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtResetZoom 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5280
      TabIndex        =   89
      Tag             =   "font1"
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox txtMinZoom 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5280
      TabIndex        =   82
      Tag             =   "font1"
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox txtMaxZoom 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5280
      TabIndex        =   81
      Tag             =   "font1"
      Top             =   1320
      Width           =   735
   End
   Begin VB.PictureBox picTopmost 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   6720
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   78
      TabStop         =   0   'False
      Tag             =   "4"
      Top             =   7515
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picScenery 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   6720
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   77
      TabStop         =   0   'False
      Tag             =   "4"
      Top             =   6960
      Width           =   240
   End
   Begin VB.ComboBox cboSkin 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Tag             =   "font1"
      Top             =   3600
      Width           =   4215
   End
   Begin VB.TextBox txtHotkey 
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   13
      Left            =   8160
      TabIndex        =   23
      Tag             =   "font1"
      Top             =   3720
      Width           =   255
   End
   Begin VB.TextBox txtHotkey 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   12
      Left            =   6720
      TabIndex        =   22
      Tag             =   "font1"
      Top             =   3720
      Width           =   255
   End
   Begin VB.PictureBox picMore 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   120
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   74
      TabStop         =   0   'False
      Tag             =   "3"
      Top             =   6120
      Width           =   960
   End
   Begin VB.TextBox txtWayptKey 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   4
      Left            =   8040
      TabIndex        =   28
      Tag             =   "font1"
      Text            =   "X"
      ToolTipText     =   "Fly"
      Top             =   5520
      Width           =   255
   End
   Begin VB.TextBox txtWayptKey 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   3
      Left            =   7320
      TabIndex        =   27
      Tag             =   "font1"
      Text            =   "Z"
      ToolTipText     =   "Down"
      Top             =   5520
      Width           =   255
   End
   Begin VB.TextBox txtWayptKey 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   7320
      TabIndex        =   24
      Tag             =   "font1"
      Text            =   "W"
      ToolTipText     =   "Up"
      Top             =   4800
      Width           =   255
   End
   Begin VB.TextBox txtWayptKey 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   7680
      TabIndex        =   26
      Tag             =   "font1"
      Text            =   "S"
      ToolTipText     =   "Right"
      Top             =   5160
      Width           =   255
   End
   Begin VB.TextBox txtWayptKey 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   6960
      TabIndex        =   25
      Tag             =   "font1"
      Text            =   "A"
      ToolTipText     =   "Left"
      Top             =   5160
      Width           =   255
   End
   Begin VB.PictureBox picPrefabs 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   5880
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   71
      TabStop         =   0   'False
      Tag             =   "10"
      Top             =   5520
      Width           =   240
   End
   Begin VB.TextBox txtPrefabs 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   360
      TabIndex        =   9
      Tag             =   "font1"
      Top             =   5520
      Width           =   5415
   End
   Begin VB.TextBox txtUncomp 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   360
      TabIndex        =   8
      Tag             =   "font1"
      Top             =   4920
      Width           =   5415
   End
   Begin VB.PictureBox picUncomp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   5880
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   68
      TabStop         =   0   'False
      Tag             =   "10"
      Top             =   4920
      Width           =   240
   End
   Begin VB.TextBox txtHotkey 
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   11
      Left            =   8160
      TabIndex        =   21
      Tag             =   "font1"
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox txtHotkey 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   10
      Left            =   6720
      TabIndex        =   20
      Tag             =   "font1"
      Top             =   3240
      Width           =   255
   End
   Begin VB.TextBox txtHotkey 
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   9
      Left            =   8160
      TabIndex        =   19
      Tag             =   "font1"
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox txtHotkey 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   8
      Left            =   6720
      TabIndex        =   18
      Tag             =   "font1"
      Top             =   2760
      Width           =   255
   End
   Begin VB.TextBox txtHotkey 
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   7
      Left            =   8160
      TabIndex        =   17
      Tag             =   "font1"
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox txtHotkey 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   6
      Left            =   6720
      TabIndex        =   16
      Tag             =   "font1"
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox txtHotkey 
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   5
      Left            =   8160
      TabIndex        =   15
      Tag             =   "font1"
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox txtHotkey 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   4
      Left            =   6720
      TabIndex        =   14
      Tag             =   "font1"
      Top             =   1800
      Width           =   255
   End
   Begin VB.TextBox txtHotkey 
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   3
      Left            =   8160
      TabIndex        =   13
      Tag             =   "font1"
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox txtHotkey 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   6720
      TabIndex        =   12
      Tag             =   "font1"
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox txtHotkey 
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   8160
      TabIndex        =   11
      Tag             =   "font1"
      Top             =   840
      Width           =   255
   End
   Begin VB.TextBox txtHotkey 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   6720
      TabIndex        =   10
      Tag             =   "font1"
      Top             =   840
      Width           =   255
   End
   Begin VB.PictureBox picHotkeys 
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3360
      Left            =   7080
      ScaleHeight     =   224
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   720
      Width           =   960
   End
   Begin VB.TextBox txtOpacity2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      TabIndex        =   5
      Tag             =   "font1"
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox txtOpacity1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2760
      TabIndex        =   4
      Tag             =   "font1"
      Top             =   2280
      Width           =   495
   End
   Begin VB.PictureBox picGridColor2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3840
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   2640
      Width           =   255
   End
   Begin VB.PictureBox picFolder 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   5880
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   33
      TabStop         =   0   'False
      Tag             =   "10"
      Top             =   4320
      Width           =   240
   End
   Begin VB.TextBox txtHeight 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Tag             =   "font1"
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox txtWidth 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Tag             =   "font1"
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox txtDir 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   360
      TabIndex        =   7
      Tag             =   "font1"
      Top             =   4320
      Width           =   5415
   End
   Begin VB.TextBox txtDivisions 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Tag             =   "font1"
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox txtSpacing 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Tag             =   "font1"
      Top             =   2280
      Width           =   495
   End
   Begin VB.PictureBox picGridColor1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3840
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   2280
      Width           =   255
   End
   Begin VB.PictureBox picApply 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   7680
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   49
      TabStop         =   0   'False
      Tag             =   "2"
      Top             =   6120
      Width           =   960
   End
   Begin VB.PictureBox picOK 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   5520
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   48
      TabStop         =   0   'False
      Tag             =   "0"
      Top             =   6120
      Width           =   960
   End
   Begin VB.PictureBox picCancel 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   6600
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   47
      TabStop         =   0   'False
      Tag             =   "1"
      Top             =   6120
      Width           =   960
   End
   Begin VB.PictureBox picBackColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3840
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   1680
      Width           =   255
   End
   Begin VB.PictureBox picPointColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3840
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   960
      Width           =   255
   End
   Begin VB.PictureBox picSelectionColor 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3840
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   1320
      Width           =   255
   End
   Begin VB.ComboBox cboWireSrc 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmPreferences.frx":0000
      Left            =   1200
      List            =   "frmPreferences.frx":001C
      Style           =   2  'Dropdown List
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   7230
      Width           =   2415
   End
   Begin VB.ComboBox cboWireDest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmPreferences.frx":0072
      Left            =   1200
      List            =   "frmPreferences.frx":008E
      Style           =   2  'Dropdown List
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   7590
      Width           =   2415
   End
   Begin VB.ComboBox cboPolyDest 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmPreferences.frx":00E4
      Left            =   3720
      List            =   "frmPreferences.frx":0100
      Style           =   2  'Dropdown List
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   7590
      Width           =   2535
   End
   Begin VB.ComboBox cboPolySrc 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmPreferences.frx":0156
      Left            =   3720
      List            =   "frmPreferences.frx":0172
      Style           =   2  'Dropdown List
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   7230
      Width           =   2535
   End
   Begin VB.PictureBox picTitle 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   585
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   0
      Width           =   8775
      Begin VB.PictureBox picHide 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H004A3C31&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   8535
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   35
         TabStop         =   0   'False
         Tag             =   "3"
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.Label lblPref 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   31
      Left            =   6030
      TabIndex        =   90
      Tag             =   "font2"
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label lblPref 
      BackColor       =   &H00614B3D&
      Caption         =   "Reset:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   28
      Left            =   4440
      TabIndex        =   88
      Tag             =   "font2"
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label lblPref 
      BackColor       =   &H00614B3D&
      Caption         =   "Max:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   27
      Left            =   4440
      TabIndex        =   87
      Tag             =   "font2"
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label lblPref 
      BackColor       =   &H00614B3D&
      Caption         =   "Min:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   26
      Left            =   4440
      TabIndex        =   86
      Tag             =   "font2"
      Top             =   960
      Width           =   735
   End
   Begin VB.Label lblPref 
      BackColor       =   &H004A3C31&
      BackStyle       =   0  'Transparent
      Caption         =   "Zoom"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   25
      Left            =   4440
      TabIndex        =   85
      Tag             =   "font2"
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblPref 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   29
      Left            =   6030
      TabIndex        =   84
      Tag             =   "font2"
      Top             =   960
      Width           =   255
   End
   Begin VB.Label lblPref 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   30
      Left            =   6030
      TabIndex        =   83
      Tag             =   "font2"
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label lblPref 
      BackStyle       =   0  'Transparent
      Caption         =   "Fullscreen always on top"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   570
      Index           =   24
      Left            =   6960
      TabIndex        =   80
      Tag             =   "font2"
      Top             =   7515
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblPref 
      BackStyle       =   0  'Transparent
      Caption         =   "Use 4 verts for scenery"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   570
      Index           =   23
      Left            =   6960
      TabIndex        =   79
      Tag             =   "font2"
      Top             =   6960
      Width           =   1575
   End
   Begin VB.Label lblOther 
      AutoSize        =   -1  'True
      BackColor       =   &H004A3C31&
      Caption         =   "Other"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   6720
      TabIndex        =   76
      Top             =   6600
      Width           =   480
   End
   Begin VB.Shape fraPref 
      BorderColor     =   &H000B3C0D&
      Height          =   1365
      Index           =   5
      Left            =   6480
      Top             =   6720
      Width           =   2175
   End
   Begin VB.Label lblPref 
      BackColor       =   &H00614B3D&
      Caption         =   "Skin:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   22
      Left            =   360
      TabIndex        =   75
      Tag             =   "font2"
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label lblWayKeys 
      AutoSize        =   -1  'True
      BackColor       =   &H004A3C31&
      Caption         =   "Waypoint Keys"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   6720
      TabIndex        =   73
      Tag             =   "font2"
      Top             =   4440
      Width           =   1365
   End
   Begin VB.Shape fraPref 
      BorderColor     =   &H000B3C0D&
      Height          =   1455
      Index           =   3
      Left            =   6480
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Label lblPref 
      BackColor       =   &H004A3C31&
      BackStyle       =   0  'Transparent
      Caption         =   "Prefabs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   17
      Left            =   360
      TabIndex        =   72
      Tag             =   "font2"
      Top             =   5280
      Width           =   5415
   End
   Begin VB.Label lblPref 
      BackColor       =   &H004A3C31&
      BackStyle       =   0  'Transparent
      Caption         =   "Uncompiled"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   16
      Left            =   360
      TabIndex        =   70
      Tag             =   "font2"
      Top             =   4680
      Width           =   5415
   End
   Begin VB.Label lblPref 
      BackColor       =   &H004A3C31&
      BackStyle       =   0  'Transparent
      Caption         =   "OpenSoldat"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   360
      TabIndex        =   69
      Tag             =   "font2"
      Top             =   4080
      Width           =   5415
   End
   Begin VB.Label lblHotkeys 
      AutoSize        =   -1  'True
      BackColor       =   &H004F3D31&
      Caption         =   "HotKeys"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   6720
      TabIndex        =   66
      Tag             =   "font2"
      Top             =   360
      Width           =   765
   End
   Begin VB.Shape fraPref 
      BorderColor     =   &H000B3C0D&
      Height          =   3855
      Index           =   1
      Left            =   6480
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label lblPref 
      BackStyle       =   0  'Transparent
      Caption         =   "px"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   12
      Left            =   2190
      TabIndex        =   65
      Tag             =   "font2"
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label lblPref 
      BackStyle       =   0  'Transparent
      Caption         =   "px"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   2190
      TabIndex        =   64
      Tag             =   "font2"
      Top             =   960
      Width           =   375
   End
   Begin VB.Label lblPref 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   15
      Left            =   3270
      TabIndex        =   63
      Tag             =   "font2"
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label lblPref 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   14
      Left            =   3270
      TabIndex        =   62
      Tag             =   "font2"
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label lblPref 
      BackColor       =   &H004A3C31&
      BackStyle       =   0  'Transparent
      Caption         =   "Colors"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   2760
      TabIndex        =   61
      Tag             =   "font2"
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label lblPref 
      BackColor       =   &H004A3C31&
      BackStyle       =   0  'Transparent
      Caption         =   "Window"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   360
      TabIndex        =   59
      Tag             =   "font2"
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label lblPref 
      BackColor       =   &H00614B3D&
      Caption         =   "Height:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   58
      Tag             =   "font2"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lblPref 
      BackColor       =   &H00614B3D&
      Caption         =   "Width:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   57
      Tag             =   "font2"
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblPref 
      BackStyle       =   0  'Transparent
      Caption         =   "px"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   13
      Left            =   2070
      TabIndex        =   56
      Tag             =   "font2"
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label lblPref 
      BackColor       =   &H00614B3D&
      Caption         =   "Divisions:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   55
      Tag             =   "font2"
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label lblPref 
      BackColor       =   &H00614B3D&
      Caption         =   "Spacing:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   54
      Tag             =   "font2"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label lblPref 
      BackColor       =   &H004F3D31&
      BackStyle       =   0  'Transparent
      Caption         =   "Grid"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   360
      TabIndex        =   52
      Tag             =   "font2"
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label lblPref 
      BackColor       =   &H00614B3D&
      BackStyle       =   0  'Transparent
      Caption         =   "Wireframe"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   18
      Left            =   1200
      TabIndex        =   51
      Top             =   6915
      Width           =   2415
   End
   Begin VB.Label lblPref 
      BackColor       =   &H00614B3D&
      BackStyle       =   0  'Transparent
      Caption         =   "Polygon"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   19
      Left            =   3720
      TabIndex        =   50
      Top             =   6915
      Width           =   2535
   End
   Begin VB.Label lblDirs 
      AutoSize        =   -1  'True
      BackColor       =   &H004A3C31&
      Caption         =   "Directories"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   360
      TabIndex        =   46
      Tag             =   "font2"
      Top             =   3240
      Width           =   945
   End
   Begin VB.Shape fraPref 
      BorderColor     =   &H000B3C0D&
      Height          =   2655
      Index           =   2
      Left            =   120
      Top             =   3360
      Width           =   6255
   End
   Begin VB.Label lblDisplay 
      AutoSize        =   -1  'True
      BackColor       =   &H004F3D31&
      Caption         =   "Display"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   360
      TabIndex        =   45
      Tag             =   "font2"
      Top             =   360
      Width           =   675
   End
   Begin VB.Label lblPref 
      BackColor       =   &H00614B3D&
      Caption         =   "Pattern:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   2760
      TabIndex        =   44
      Tag             =   "font2"
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblPref 
      BackColor       =   &H00614B3D&
      Caption         =   "SRC:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   20
      Left            =   360
      TabIndex        =   43
      Top             =   7230
      Width           =   735
   End
   Begin VB.Label lblPref 
      BackColor       =   &H00614B3D&
      Caption         =   "DEST:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   21
      Left            =   360
      TabIndex        =   42
      Top             =   7590
      Width           =   735
   End
   Begin VB.Label lblPref 
      BackColor       =   &H00614B3D&
      Caption         =   "Back:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   2760
      TabIndex        =   41
      Tag             =   "font2"
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lblPref 
      BackColor       =   &H00614B3D&
      Caption         =   "Point:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   2760
      TabIndex        =   40
      Tag             =   "font2"
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblBlending 
      AutoSize        =   -1  'True
      BackColor       =   &H004A3C31&
      Caption         =   "Blending"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   360
      TabIndex        =   39
      Top             =   6600
      Width           =   750
   End
   Begin VB.Shape fraPref 
      BorderColor     =   &H000B3C0D&
      Height          =   2655
      Index           =   0
      Left            =   120
      Top             =   480
      Width           =   6255
   End
   Begin VB.Shape fraPref 
      BorderColor     =   &H000B3C0D&
      Height          =   1365
      Index           =   4
      Left            =   120
      Top             =   6720
      Width           =   6255
   End
End
Attribute VB_Name = "frmPreferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' preferences dialog - change application prefrences


' fix vb6 ide casing changes
#If False Then
    Public FileName, color, token, A, R, G, B, commonDialog, value, Val, X, Y, Z, Left, hWnd, Mid, Right, BackColor, bottom
    'Public FileName, color, token, A, R, G, B, commonDialog, value, Val, X, Y, Z, Left, hWnd, Mid, Right, BackColor, bottom
#End If


' consts - public


' consts - private

Private Const MIN_HEIGHT = 440
Private Const MAX_HEIGHT = 547


' types - public


' types - private


' vars - public


' vars - private

Private blendModes(0 To 7) As Integer

Private backgroundColor As TColor
Private pointColor As TColor
Private selectionColor As TColor
Private gridColor1 As TColor
Private gridColor2 As TColor

Private spacing As Integer
Private divisions As Integer
Private formWidth As Integer
Private formHeight As Integer
Private opacity1 As Integer
Private opacity2 As Integer
Private sceneryVerts As Boolean
Private topmost As Boolean

Private formMinZoom As Single
Private formMaxZoom As Single
Private formResetZoom As Single


' external functions - public


' external functions - private


' functions - public

Public Sub SetColors()

    On Error Resume Next

    Dim i As Integer
    Dim c As Control


    picTitle.Picture = LoadPicture(appPath & "\skins\" & gfxDir & "\titlebar_preferences.bmp")
    picHotkeys.Picture = LoadPicture(appPath & "\skins\" & gfxDir & "\tools.bmp")

    MouseEvent2 picHide, 0, 0, BUTTON_SMALL, 0, BUTTON_UP
    MouseEvent2 picOK, 0, 0, BUTTON_LARGE, 0, BUTTON_UP
    MouseEvent2 picCancel, 0, 0, BUTTON_LARGE, 0, BUTTON_UP
    MouseEvent2 picMore, 0, 0, BUTTON_LARGE, 0, BUTTON_UP
    MouseEvent2 picApply, 0, 0, BUTTON_LARGE, 0, BUTTON_UP
    MouseEvent2 picFolder, 0, 0, BUTTON_SMALL, 0, BUTTON_UP
    MouseEvent2 picUncomp, 0, 0, BUTTON_SMALL, 0, BUTTON_UP
    MouseEvent2 picPrefabs, 0, 0, BUTTON_SMALL, 0, BUTTON_UP
    MouseEvent2 picScenery, 0, 0, BUTTON_SMALL, sceneryVerts, BUTTON_UP
    MouseEvent2 picTopmost, 0, 0, BUTTON_SMALL, topmost, BUTTON_UP


    Me.BackColor = bgColor

    For Each c In lblPref
        c.BackColor = lblBackColor
        c.ForeColor = lblTextColor
    Next

    lblDisplay.BackColor = bgColor
    lblDisplay.ForeColor = lblTextColor
    lblHotkeys.BackColor = bgColor
    lblHotkeys.ForeColor = lblTextColor
    lblDirs.BackColor = bgColor
    lblDirs.ForeColor = lblTextColor
    lblWayKeys.BackColor = bgColor
    lblWayKeys.ForeColor = lblTextColor
    lblBlending.BackColor = bgColor
    lblBlending.ForeColor = lblTextColor
    lblOther.BackColor = bgColor
    lblOther.ForeColor = lblTextColor

    For Each c In txtHotkey
        c.BackColor = bgColor
        c.ForeColor = lblTextColor
    Next

    For Each c In txtWayptKey
        c.BackColor = bgColor
        c.ForeColor = lblTextColor
    Next

    For Each c In fraPref
        c.BorderColor = frameColor
    Next

    txtWidth.BackColor = txtBackColor
    txtWidth.ForeColor = txtTextColor
    txtHeight.BackColor = txtBackColor
    txtHeight.ForeColor = txtTextColor

    txtSpacing.BackColor = txtBackColor
    txtSpacing.ForeColor = txtTextColor
    txtDivisions.BackColor = txtBackColor
    txtDivisions.ForeColor = txtTextColor
    txtOpacity1.BackColor = txtBackColor
    txtOpacity1.ForeColor = txtTextColor
    txtOpacity2.BackColor = txtBackColor
    txtOpacity2.ForeColor = txtTextColor

    txtMinZoom.BackColor = txtBackColor
    txtMinZoom.ForeColor = txtTextColor
    txtMaxZoom.BackColor = txtBackColor
    txtMaxZoom.ForeColor = txtTextColor
    txtResetZoom.BackColor = txtBackColor
    txtResetZoom.ForeColor = txtTextColor

    txtDir.BackColor = txtBackColor
    txtDir.ForeColor = txtTextColor
    txtUncomp.BackColor = txtBackColor
    txtUncomp.ForeColor = txtTextColor
    txtPrefabs.BackColor = txtBackColor
    txtPrefabs.ForeColor = txtTextColor

    cboWireSrc.BackColor = txtBackColor
    cboWireSrc.ForeColor = txtTextColor
    cboWireDest.BackColor = txtBackColor
    cboWireDest.ForeColor = txtTextColor
    cboPolySrc.BackColor = txtBackColor
    cboPolySrc.ForeColor = txtTextColor
    cboPolyDest.BackColor = txtBackColor
    cboPolyDest.ForeColor = txtTextColor

    cboSkin.BackColor = txtBackColor
    cboSkin.ForeColor = txtTextColor

    SetFormFonts Me

End Sub


' functions - private

Private Function applyPreferences() As Boolean

    Dim i As Integer
    Dim mInitialWindowWidth As Long
    Dim mInitialWindowHeight As Long
    Dim deltaLeft As Long
    Dim deltaTop As Long

    On Error GoTo ErrorHandler

    MouseEvent2 picHide, 0, 0, BUTTON_SMALL, 0, BUTTON_UP
    MouseEvent2 picOK, 0, 0, BUTTON_LARGE, 0, BUTTON_UP
    MouseEvent2 picCancel, 0, 0, BUTTON_LARGE, 0, BUTTON_UP
    MouseEvent2 picMore, 0, 0, BUTTON_LARGE, 0, BUTTON_UP
    MouseEvent2 picApply, 0, 0, BUTTON_LARGE, 0, BUTTON_UP
    MouseEvent2 picFolder, 0, 0, BUTTON_SMALL, 0, BUTTON_UP

    If Right(txtDir.Text, 1) <> "\" Then txtDir.Text = txtDir.Text + "\"

    If Len(Dir(txtDir.Text, vbDirectory)) <> 0 And frmOpenSoldatMapEditor.OpenSoldatDir <> txtDir.Text Then
        If Not Len(Dir(txtDir.Text & "Maps\", vbDirectory)) <> 0 Then
            MsgBox "'Maps' folder does not exist in OpenSoldat directory."
            Exit Function
        End If
        If Not Len(Dir(txtDir.Text & "Textures\", vbDirectory)) <> 0 Then
            MsgBox "'Textures' folder does not exist in OpenSoldat directory."
            Exit Function
        End If
        If Not Len(Dir(txtDir.Text & "Scenery-gfx\", vbDirectory)) <> 0 Then
            MsgBox "'Scenery-gfx' folder does not exist in OpenSoldat directory."
            Exit Function
        End If

        frmOpenSoldatMapEditor.OpenSoldatDir = txtDir.Text
    ElseIf Len(Dir(txtDir.Text, vbDirectory)) = 0 Then
        MsgBox "OpenSoldat directory does not exist."
        Exit Function
    End If

    If Right(txtUncomp.Text, 1) <> "\" Then txtUncomp.Text = txtUncomp.Text + "\"

    If Len(Dir(txtUncomp.Text, vbDirectory)) <> 0 Then
        frmOpenSoldatMapEditor.uncompDir = txtUncomp.Text
    Else
        MsgBox "Uncompiled Maps directory does not exist."
        Exit Function
    End If

    If Right(txtPrefabs.Text, 1) <> "\" Then txtPrefabs.Text = txtPrefabs.Text + "\"

    If Len(Dir(txtPrefabs.Text, vbDirectory)) <> 0 Then
        frmOpenSoldatMapEditor.prefabDir = txtPrefabs.Text
    Else
        MsgBox "Prefabs Maps directory does not exist."
        Exit Function
    End If

    frmOpenSoldatMapEditor.wireBlendSrc = blendModes(cboWireSrc.ListIndex)
    frmOpenSoldatMapEditor.wireBlendDest = blendModes(cboWireDest.ListIndex)
    frmOpenSoldatMapEditor.polyBlendSrc = blendModes(cboPolySrc.ListIndex)
    frmOpenSoldatMapEditor.polyBlendDest = blendModes(cboPolyDest.ListIndex)

    frmOpenSoldatMapEditor.backgroundColor = RGB(backgroundColor.blue, backgroundColor.green, backgroundColor.red)
    frmOpenSoldatMapEditor.pointColor = RGB(pointColor.blue, pointColor.green, pointColor.red)
    frmOpenSoldatMapEditor.selectionColor = RGB(selectionColor.blue, selectionColor.green, selectionColor.red)
    frmOpenSoldatMapEditor.gridColor1 = RGB(gridColor1.blue, gridColor1.green, gridColor1.red)
    frmOpenSoldatMapEditor.gridColor2 = RGB(gridColor2.blue, gridColor2.green, gridColor2.red)

    mInitialWindowWidth = frmOpenSoldatMapEditor.Width
    mInitialWindowHeight = frmOpenSoldatMapEditor.Height

    If frmOpenSoldatMapEditor.Tag = vbNormal Then
        frmOpenSoldatMapEditor.Width = formWidth * Screen.TwipsPerPixelX
        frmOpenSoldatMapEditor.Height = formHeight * Screen.TwipsPerPixelY

        ' TODO: move into functions
        If Len(frmDisplay.Tag) <> 0 Then
            deltaLeft = frmOpenSoldatMapEditor.getLeftSnapDelta(frmOpenSoldatMapEditor, frmDisplay, mInitialWindowWidth, formWidth)
            deltaTop = frmOpenSoldatMapEditor.getTopSnapDelta(frmOpenSoldatMapEditor, frmDisplay, mInitialWindowHeight, formHeight)
            frmDisplay.Move (frmDisplay.Left + deltaLeft + (frmOpenSoldatMapEditor.Left - (frmOpenSoldatMapEditor.formLeft * Screen.TwipsPerPixelX))), (frmDisplay.Top + deltaTop + (frmOpenSoldatMapEditor.Top - (frmOpenSoldatMapEditor.formTop * Screen.TwipsPerPixelY)))
        End If
        If Len(frmInfo.Tag) <> 0 Then
            deltaLeft = frmOpenSoldatMapEditor.getLeftSnapDelta(frmOpenSoldatMapEditor, frmInfo, mInitialWindowWidth, formWidth)
            deltaTop = frmOpenSoldatMapEditor.getTopSnapDelta(frmOpenSoldatMapEditor, frmInfo, mInitialWindowHeight, formHeight)
            frmInfo.Move (frmInfo.Left + deltaLeft + (frmOpenSoldatMapEditor.Left - (frmOpenSoldatMapEditor.formLeft * Screen.TwipsPerPixelX))), (frmInfo.Top + deltaTop + (frmOpenSoldatMapEditor.Top - (frmOpenSoldatMapEditor.formTop * Screen.TwipsPerPixelY)))
        End If
        If Len(frmPalette.Tag) <> 0 Then
            deltaLeft = frmOpenSoldatMapEditor.getLeftSnapDelta(frmOpenSoldatMapEditor, frmPalette, mInitialWindowWidth, formWidth)
            deltaTop = frmOpenSoldatMapEditor.getTopSnapDelta(frmOpenSoldatMapEditor, frmPalette, mInitialWindowHeight, formHeight)
            frmPalette.Move (frmPalette.Left + deltaLeft + (frmOpenSoldatMapEditor.Left - (frmOpenSoldatMapEditor.formLeft * Screen.TwipsPerPixelX))), (frmPalette.Top + deltaTop + (frmOpenSoldatMapEditor.Top - (frmOpenSoldatMapEditor.formTop * Screen.TwipsPerPixelY)))
        End If
        If Len(frmScenery.Tag) <> 0 Then
            deltaLeft = frmOpenSoldatMapEditor.getLeftSnapDelta(frmOpenSoldatMapEditor, frmScenery, mInitialWindowWidth, formWidth)
            deltaTop = frmOpenSoldatMapEditor.getTopSnapDelta(frmOpenSoldatMapEditor, frmScenery, mInitialWindowHeight, formHeight)
            frmScenery.Move (frmScenery.Left + deltaLeft + (frmOpenSoldatMapEditor.Left - (frmOpenSoldatMapEditor.formLeft * Screen.TwipsPerPixelX))), (frmScenery.Top + deltaTop + (frmOpenSoldatMapEditor.Top - (frmOpenSoldatMapEditor.formTop * Screen.TwipsPerPixelY)))
        End If
        If Len(frmTexture.Tag) <> 0 Then
            deltaLeft = frmOpenSoldatMapEditor.getLeftSnapDelta(frmOpenSoldatMapEditor, frmTexture, mInitialWindowWidth, formWidth)
            deltaTop = frmOpenSoldatMapEditor.getTopSnapDelta(frmOpenSoldatMapEditor, frmTexture, mInitialWindowHeight, formHeight)
            frmTexture.Move (frmTexture.Left + deltaLeft + (frmOpenSoldatMapEditor.Left - (frmOpenSoldatMapEditor.formLeft * Screen.TwipsPerPixelX))), (frmTexture.Top + deltaTop + (frmOpenSoldatMapEditor.Top - (frmOpenSoldatMapEditor.formTop * Screen.TwipsPerPixelY)))
        End If
        If Len(frmTools.Tag) <> 0 Then
            deltaLeft = frmOpenSoldatMapEditor.getLeftSnapDelta(frmOpenSoldatMapEditor, frmTools, mInitialWindowWidth, formWidth)
            deltaTop = frmOpenSoldatMapEditor.getTopSnapDelta(frmOpenSoldatMapEditor, frmTools, mInitialWindowHeight, formHeight)
            frmTools.Move (frmTools.Left + deltaLeft + (frmOpenSoldatMapEditor.Left - (frmOpenSoldatMapEditor.formLeft * Screen.TwipsPerPixelX))), (frmTools.Top + deltaTop + (frmOpenSoldatMapEditor.Top - (frmOpenSoldatMapEditor.formTop * Screen.TwipsPerPixelY)))
        End If
        If Len(frmWaypoints.Tag) <> 0 Then
            deltaLeft = frmOpenSoldatMapEditor.getLeftSnapDelta(frmOpenSoldatMapEditor, frmWaypoints, mInitialWindowWidth, formWidth)
            deltaTop = frmOpenSoldatMapEditor.getTopSnapDelta(frmOpenSoldatMapEditor, frmWaypoints, mInitialWindowHeight, formHeight)
            frmWaypoints.Move (frmWaypoints.Left + deltaLeft + (frmOpenSoldatMapEditor.Left - (frmOpenSoldatMapEditor.formLeft * Screen.TwipsPerPixelX))), (frmWaypoints.Top + deltaTop + (frmOpenSoldatMapEditor.Top - (frmOpenSoldatMapEditor.formTop * Screen.TwipsPerPixelY)))
        End If
    End If

    frmOpenSoldatMapEditor.formWidth = formWidth
    frmOpenSoldatMapEditor.formHeight = formHeight

    frmOpenSoldatMapEditor.picResize.Top = frmOpenSoldatMapEditor.Height / Screen.TwipsPerPixelY - frmOpenSoldatMapEditor.picResize.Height
    frmOpenSoldatMapEditor.picResize.Left = frmOpenSoldatMapEditor.Width / Screen.TwipsPerPixelX - frmOpenSoldatMapEditor.picResize.Width

    frmOpenSoldatMapEditor.gridSpacing = spacing
    frmOpenSoldatMapEditor.gridDivisions = divisions
    frmOpenSoldatMapEditor.gridOp1 = opacity1 / 100 * 255
    frmOpenSoldatMapEditor.gridOp2 = opacity2 / 100 * 255


    frmOpenSoldatMapEditor.gMinZoom = formMinZoom / 100
    frmOpenSoldatMapEditor.gMaxZoom = formMaxZoom / 100
    frmOpenSoldatMapEditor.gResetZoom = formResetZoom / 100

    frmOpenSoldatMapEditor.sceneryVerts = sceneryVerts
    frmOpenSoldatMapEditor.topmost = topmost

    Debug.Assert txtHotkey.LBound = frmTools.picTools.LBound
    Debug.Assert txtHotkey.UBound = frmTools.picTools.UBound

    For i = txtHotkey.LBound To txtHotkey.UBound
        frmTools.SetHotKey i, MapVirtualKey(CInt(txtHotkey(i).Tag), 0)
        frmTools.picTools(i).ToolTipText = frmTools.picTools(i).Tag & " (" & (txtHotkey(i).Text) & ")"
    Next

    Debug.Assert txtWayptKey.LBound = frmWaypoints.picType.LBound
    Debug.Assert txtWayptKey.UBound = frmWaypoints.picType.UBound
    Debug.Assert txtWayptKey.LBound = frmWaypoints.lblType.LBound
    Debug.Assert txtWayptKey.UBound = frmWaypoints.lblType.UBound

    For i = txtWayptKey.LBound To txtWayptKey.UBound
        frmWaypoints.SetWaypointKey i, MapVirtualKey(CInt(txtWayptKey(i).Tag), 0)
        frmWaypoints.picType(i).ToolTipText = " (" & (txtWayptKey(i).Text) & ")"
        frmWaypoints.lblType(i).ToolTipText = " (" & (txtWayptKey(i).Text) & ")"
    Next

    If cboSkin.List(cboSkin.ListIndex) <> gfxDir Then
        gfxDir = cboSkin.List(cboSkin.ListIndex)
        frmOpenSoldatMapEditor.LoadColors
        frmOpenSoldatMapEditor.SetColors
        frmOpenSoldatMapEditor.InitGfx
        frmColor.SetColors
        frmDisplay.SetColors
        frmInfo.SetColors
        frmMap.SetColors
        frmPalette.SetColors
        frmPreferences.SetColors
        frmScenery.SetColors
        frmOpenSoldatMapEditor.SetColors
        frmTexture.SetColors
        frmTools.SetColors
        frmWaypoints.SetColors
        frmDisplay.RefreshButtons
    End If

    frmOpenSoldatMapEditor.SetPreferences

    applyPreferences = True

    Exit Function

ErrorHandler:

    MsgBox "Error applying preferences" & vbNewLine & Error

End Function


' events - public


' events - private

Private Sub picHide_Click()

    Unload Me
    frmOpenSoldatMapEditor.RegainFocus

End Sub

Private Sub picMore_Click()

    If Me.ScaleHeight < MAX_HEIGHT - 20 Then
        Me.Height = MAX_HEIGHT * Screen.TwipsPerPixelY
    Else
        Me.Height = MIN_HEIGHT * Screen.TwipsPerPixelY
    End If

End Sub

Private Sub picTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ReleaseCapture
    SendMessage Me.hWnd, WM_NCLBUTTONDOWN, 2, 0&

End Sub

Private Sub picApply_Click()

    applyPreferences

End Sub

Private Sub picCancel_Click()

    Unload Me
    frmOpenSoldatMapEditor.RegainFocus

End Sub

Private Sub picOK_Click()

    If applyPreferences Then
        Unload Me
        frmOpenSoldatMapEditor.RegainFocus
    End If

End Sub

Private Sub Form_Load()

    Dim i As Integer

    On Error GoTo ErrorHandler

    Me.Height = MIN_HEIGHT * Screen.TwipsPerPixelY

    sceneryVerts = frmOpenSoldatMapEditor.sceneryVerts
    topmost = frmOpenSoldatMapEditor.topmost

    Me.SetColors

    blendModes(0) = 1
    blendModes(1) = 2
    blendModes(2) = 3
    blendModes(3) = 4
    blendModes(4) = 9
    blendModes(5) = 10
    blendModes(6) = 5
    blendModes(7) = 6

    backgroundColor = GetRGB(frmOpenSoldatMapEditor.backgroundColor)
    pointColor = GetRGB(frmOpenSoldatMapEditor.pointColor)
    selectionColor = GetRGB(frmOpenSoldatMapEditor.selectionColor)
    gridColor1 = GetRGB(frmOpenSoldatMapEditor.gridColor1)
    gridColor2 = GetRGB(frmOpenSoldatMapEditor.gridColor2)

    For i = LBound(blendModes) To UBound(blendModes)
        If frmOpenSoldatMapEditor.wireBlendSrc = blendModes(i) Then cboWireSrc.ListIndex = i
        If frmOpenSoldatMapEditor.wireBlendDest = blendModes(i) Then cboWireDest.ListIndex = i
        If frmOpenSoldatMapEditor.polyBlendSrc = blendModes(i) Then cboPolySrc.ListIndex = i
        If frmOpenSoldatMapEditor.polyBlendDest = blendModes(i) Then cboPolyDest.ListIndex = i
    Next

    Me.picBackColor.BackColor = RGB(backgroundColor.red, backgroundColor.green, backgroundColor.blue)
    Me.picPointColor.BackColor = RGB(pointColor.red, pointColor.green, pointColor.blue)
    Me.picSelectionColor.BackColor = RGB(selectionColor.red, selectionColor.green, selectionColor.blue)
    Me.picGridColor1.BackColor = RGB(gridColor1.red, gridColor1.green, gridColor1.blue)
    Me.picGridColor2.BackColor = RGB(gridColor2.red, gridColor2.green, gridColor2.blue)

    txtWidth.Text = frmOpenSoldatMapEditor.formWidth
    txtHeight.Text = frmOpenSoldatMapEditor.formHeight
    formWidth = txtWidth.Text
    formHeight = txtHeight.Text

    txtSpacing.Text = frmOpenSoldatMapEditor.gridSpacing
    txtDivisions.Text = frmOpenSoldatMapEditor.gridDivisions
    spacing = txtSpacing.Text
    divisions = txtDivisions.Text
    opacity1 = frmOpenSoldatMapEditor.gridOp1 / 255 * 100
    txtOpacity1.Text = opacity1
    opacity2 = frmOpenSoldatMapEditor.gridOp2 / 255 * 100
    txtOpacity2.Text = opacity2

    txtMinZoom.Text = frmOpenSoldatMapEditor.gMinZoom * 100
    txtMaxZoom.Text = frmOpenSoldatMapEditor.gMaxZoom * 100
    txtResetZoom.Text = frmOpenSoldatMapEditor.gResetZoom * 100
    formMinZoom = txtMinZoom.Text
    formMaxZoom = txtMaxZoom.Text
    formResetZoom = txtResetZoom.Text

    For i = txtHotkey.LBound To txtHotkey.UBound
        txtHotkey(i).Text = Chr(MapVirtualKey(frmTools.GetHotKey(i), 1))
        txtHotkey(i).Tag = AscDef(txtHotkey(i).Text, 0)
    Next

    For i = txtWayptKey.LBound To txtWayptKey.UBound
        txtWayptKey(i).Text = Chr(MapVirtualKey(frmWaypoints.GetWaypointKey(i), 1))
        txtWayptKey(i).Tag = Asc(txtWayptKey(i).Text)
    Next

    Dim file As Variant

    file = Dir(appPath & "\skins\*.*", vbDirectory)
    Do While Len(file)
        If FileExists(appPath & "\skins\" & file & "\colors.ini") Then
            cboSkin.AddItem file
            If file = gfxDir Then cboSkin.ListIndex = cboSkin.ListCount - 1
        End If
        file = Dir
    Loop

    txtDir.Text = frmOpenSoldatMapEditor.OpenSoldatDir
    txtUncomp.Text = frmOpenSoldatMapEditor.uncompDir
    txtPrefabs.Text = frmOpenSoldatMapEditor.prefabDir

    Exit Sub

ErrorHandler:

    MsgBox "Error loading Preferences form" & vbNewLine & Error

End Sub

Private Sub picPointColor_Click()

    frmColor.InitColor pointColor.red, pointColor.green, pointColor.blue
    frmColor.Show 1
    picPointColor.BackColor = RGB(frmColor.red, frmColor.green, frmColor.blue)
    pointColor.red = frmColor.red
    pointColor.green = frmColor.green
    pointColor.blue = frmColor.blue

End Sub

Private Sub picSelectionColor_Click()

    frmColor.InitColor selectionColor.red, selectionColor.green, selectionColor.blue
    frmColor.Show 1
    picSelectionColor.BackColor = RGB(frmColor.red, frmColor.green, frmColor.blue)
    selectionColor.red = frmColor.red
    selectionColor.green = frmColor.green
    selectionColor.blue = frmColor.blue

End Sub

Private Sub picBackColor_Click()

    frmColor.InitColor backgroundColor.red, backgroundColor.green, backgroundColor.blue
    frmColor.Show 1
    picBackColor.BackColor = RGB(frmColor.red, frmColor.green, frmColor.blue)
    backgroundColor.red = frmColor.red
    backgroundColor.green = frmColor.green
    backgroundColor.blue = frmColor.blue

End Sub

Private Sub picGridColor1_Click()

    frmColor.InitColor gridColor1.red, gridColor1.green, gridColor1.blue
    frmColor.Show 1
    picGridColor1.BackColor = RGB(frmColor.red, frmColor.green, frmColor.blue)
    gridColor1.red = frmColor.red
    gridColor1.green = frmColor.green
    gridColor1.blue = frmColor.blue

End Sub

Private Sub picGridColor2_Click()

    frmColor.InitColor gridColor2.red, gridColor2.green, gridColor2.blue
    frmColor.Show 1
    picGridColor2.BackColor = RGB(frmColor.red, frmColor.green, frmColor.blue)
    gridColor2.red = frmColor.red
    gridColor2.green = frmColor.green
    gridColor2.blue = frmColor.blue

End Sub

Private Sub picFolder_Click()

    Dim folder As String

    folder = SelectFolder(Me)

    If Len(folder) > 0 Then
        txtDir.Text = folder
    End If

    MouseEvent2 picFolder, 0, 0, BUTTON_SMALL, 0, BUTTON_UP

End Sub

Private Sub picUncomp_Click()

    Dim folder As String

    folder = SelectFolder(Me)

    If Len(folder) > 0 Then
        txtUncomp.Text = folder
    End If

    MouseEvent2 picUncomp, 0, 0, BUTTON_SMALL, 0, BUTTON_UP

End Sub

Private Sub picPrefabs_Click()

    Dim folder As String

    folder = SelectFolder(Me)

    If Len(folder) > 0 Then
        txtPrefabs.Text = folder
    End If

    MouseEvent2 picPrefabs, 0, 0, BUTTON_SMALL, 0, BUTTON_UP

End Sub

Private Sub txtHotkey_GotFocus(Index As Integer)

    txtHotkey(Index).SelStart = 0
    txtHotkey(Index).SelLength = Len(txtHotkey(Index).Text)

End Sub

Private Sub txtHotkey_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    txtHotkey(Index).Tag = KeyCode

End Sub

Private Sub txtHotkey_KeyPress(Index As Integer, KeyAscii As Integer)

    txtHotkey(Index).Text = UCase(Chr(KeyAscii))
    KeyAscii = 0

End Sub

Private Sub txtMaxZoom_Change()

    If IsNumeric(txtMaxZoom.Text) = False And txtMaxZoom.Text <> "" Then
        txtMaxZoom.Text = formMaxZoom
    End If

End Sub

Private Sub txtMaxZoom_GotFocus()

    txtMaxZoom.SelStart = 0
    txtMaxZoom.SelLength = Len(txtMaxZoom.Text)

End Sub

Private Sub txtMaxZoom_LostFocus()

    If IsNumeric(txtMaxZoom.Text) = False And txtMaxZoom.Text <> "" Then
        txtMaxZoom.Text = formMaxZoom
    ElseIf txtMaxZoom.Text = "" Then
        txtMaxZoom.Text = formMaxZoom
    ElseIf txtMaxZoom.Text > 0 And txtMaxZoom.Text < 10000000 Then
        formMaxZoom = CSng(txtMaxZoom.Text)
        txtMaxZoom.Text = formMaxZoom
    Else
        txtMaxZoom.Text = formMaxZoom
    End If

End Sub

Private Sub txtMinZoom_Change()

    If IsNumeric(txtMinZoom.Text) = False And txtMinZoom.Text <> "" Then
        txtMinZoom.Text = formMinZoom
    End If

End Sub

Private Sub txtMinZoom_GotFocus()

    txtMinZoom.SelStart = 0
    txtMinZoom.SelLength = Len(txtMinZoom.Text)

End Sub

Private Sub txtMinZoom_LostFocus()

    If IsNumeric(txtMinZoom.Text) = False And txtMinZoom.Text <> "" Then
        txtMinZoom.Text = formMinZoom
    ElseIf txtMinZoom.Text = "" Then
        txtMinZoom.Text = formMinZoom
    ElseIf txtMinZoom.Text > 0 And txtMinZoom.Text < 10000000 Then
        formMinZoom = CSng(txtMinZoom.Text)
        txtMinZoom.Text = formMinZoom
    Else
        txtMinZoom.Text = formMinZoom
    End If

End Sub

Private Sub txtResetZoom_Change()

    If IsNumeric(txtResetZoom.Text) = False And txtResetZoom.Text <> "" Then
        txtResetZoom.Text = formResetZoom
    End If

End Sub

Private Sub txtResetZoom_GotFocus()

    txtResetZoom.SelStart = 0
    txtResetZoom.SelLength = Len(txtResetZoom.Text)

End Sub

Private Sub txtResetZoom_LostFocus()

    If IsNumeric(txtResetZoom.Text) = False And txtResetZoom.Text <> "" Then
        txtResetZoom.Text = formResetZoom
    ElseIf txtResetZoom.Text = "" Then
        txtResetZoom.Text = formResetZoom
    ElseIf txtResetZoom.Text >= formMinZoom And txtResetZoom.Text <= formMaxZoom Then
        formResetZoom = CSng(txtResetZoom.Text)
        txtResetZoom.Text = formResetZoom
    Else
        txtResetZoom.Text = formResetZoom
    End If

End Sub

Private Sub txtWayptKey_GotFocus(Index As Integer)

    txtWayptKey(Index).SelStart = 0
    txtWayptKey(Index).SelLength = Len(txtWayptKey(Index).Text)

End Sub

Private Sub txtWayptKey_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    txtWayptKey(Index).Tag = KeyCode

End Sub

Private Sub txtWayptKey_KeyPress(Index As Integer, KeyAscii As Integer)

    txtWayptKey(Index).Text = UCase(Chr(KeyAscii))
    KeyAscii = 0

End Sub

Private Sub txtSpacing_Change()

    If IsNumeric(txtSpacing.Text) = False And txtSpacing.Text <> "" Then
        txtSpacing.Text = spacing
    End If

End Sub

Private Sub txtSpacing_GotFocus()

    txtSpacing.SelStart = 0
    txtSpacing.SelLength = Len(txtSpacing.Text)

End Sub

Private Sub txtSpacing_LostFocus()

    If IsNumeric(txtSpacing.Text) = False And txtSpacing.Text <> "" Then
        txtSpacing.Text = spacing
    ElseIf txtSpacing.Text = "" Then
        txtSpacing.Text = spacing
    ElseIf txtSpacing.Text >= 10 And txtSpacing.Text <= 100 Then
        spacing = Int(txtSpacing.Text)
        txtSpacing.Text = spacing
    Else
        If txtSpacing.Text < 10 Then spacing = 10
        If txtSpacing.Text > 100 Then spacing = 100
        txtSpacing.Text = spacing
    End If

End Sub

Private Sub txtDivisions_Change()

    If IsNumeric(txtDivisions.Text) = False And txtDivisions.Text <> "" Then
        txtDivisions.Text = divisions
    End If

End Sub

Private Sub txtDivisions_GotFocus()

    txtDivisions.SelStart = 0
    txtDivisions.SelLength = Len(txtDivisions.Text)

End Sub

Private Sub txtDivisions_LostFocus()

    If IsNumeric(txtDivisions.Text) = False And txtDivisions.Text <> "" Then
        txtDivisions.Text = divisions
    ElseIf txtDivisions.Text = "" Then
        txtDivisions.Text = divisions
    ElseIf txtDivisions.Text >= 1 And txtDivisions.Text <= 10 Then
        divisions = Int(txtDivisions.Text)
        txtDivisions = divisions
    Else
        If txtDivisions.Text < 1 Then divisions = 1
        If txtDivisions.Text > Int(spacing / 4) Then divisions = Int(spacing / 4)
        txtDivisions.Text = divisions
    End If

End Sub

Private Sub txtOpacity1_Change()

    If IsNumeric(txtOpacity1.Text) = False And txtOpacity1.Text <> "" Then
        txtOpacity1.Text = opacity1
    End If

End Sub

Private Sub txtOpacity1_GotFocus()

    txtOpacity1.SelStart = 0
    txtOpacity1.SelLength = Len(txtOpacity1.Text)

End Sub

Private Sub txtOpacity1_LostFocus()

    If IsNumeric(txtOpacity1.Text) = False And txtOpacity1.Text <> "" Then
        txtOpacity1.Text = opacity1
    ElseIf txtOpacity1.Text = "" Then
        txtOpacity1.Text = opacity1
    ElseIf txtOpacity1.Text >= 10 And txtOpacity1.Text <= 100 Then
        opacity1 = Int(txtOpacity1.Text)
        txtOpacity1.Text = opacity1
    Else
        txtOpacity1.Text = opacity1
    End If

End Sub

Private Sub txtOpacity2_Change()

    If IsNumeric(txtOpacity2.Text) = False And txtOpacity2.Text <> "" Then
        txtOpacity2.Text = opacity2
    End If

End Sub

Private Sub txtOpacity2_GotFocus()

    txtOpacity2.SelStart = 0
    txtOpacity2.SelLength = Len(txtOpacity2.Text)

End Sub

Private Sub txtOpacity2_LostFocus()

    If IsNumeric(txtOpacity2.Text) = False And txtOpacity2.Text <> "" Then
        txtOpacity2.Text = opacity2
    ElseIf txtOpacity2.Text = "" Then
        txtOpacity2.Text = opacity2
    ElseIf txtOpacity2.Text >= 10 And txtOpacity2.Text <= 100 Then
        opacity2 = Int(txtOpacity2.Text)
        txtOpacity2.Text = opacity2
    Else
        txtOpacity2.Text = opacity2
    End If

End Sub

Private Sub txtWidth_Change()

    If IsNumeric(txtWidth.Text) = False And txtWidth.Text <> "" Then
        txtWidth.Text = formWidth
    End If

End Sub

Private Sub txtWidth_GotFocus()

    txtWidth.SelStart = 0
    txtWidth.SelLength = Len(txtWidth.Text)

End Sub

Private Sub txtWidth_LostFocus()

    If IsNumeric(txtWidth.Text) = False And txtWidth.Text <> "" Then
        txtWidth.Text = formWidth
    ElseIf txtWidth.Text = "" Then
        txtWidth.Text = formWidth
    ElseIf txtWidth.Text >= MAINFORM_MIN_WIDTH And txtWidth.Text <= Screen.Width / Screen.TwipsPerPixelX Then
        formWidth = Int(txtWidth.Text)
        txtWidth.Text = formWidth
    Else
        If txtWidth.Text < MAINFORM_MIN_WIDTH Then formWidth = MAINFORM_MIN_WIDTH
        If txtWidth.Text > (Screen.Width / Screen.TwipsPerPixelX) Then formWidth = (Screen.Width / Screen.TwipsPerPixelX)
        txtWidth.Text = formWidth
    End If

End Sub

Private Sub txtHeight_Change()

    If IsNumeric(txtHeight.Text) = False And txtHeight.Text <> "" Then
        txtHeight.Text = formHeight
    End If

End Sub

Private Sub txtHeight_GotFocus()

    txtHeight.SelStart = 0
    txtHeight.SelLength = Len(txtHeight.Text)

End Sub

Private Sub txtHeight_LostFocus()

    If IsNumeric(txtHeight.Text) = False And txtHeight.Text <> "" Then
        txtHeight.Text = formHeight
    ElseIf txtHeight.Text = "" Then
        txtHeight.Text = formHeight
    ElseIf txtHeight.Text >= MAINFORM_MIN_HEIGHT And txtHeight.Text <= Screen.Height / Screen.TwipsPerPixelY Then
        formHeight = Int(txtHeight.Text)
        txtHeight.Text = formHeight
    Else
        If txtHeight.Text < MAINFORM_MIN_HEIGHT Then formHeight = MAINFORM_MIN_HEIGHT
        If txtHeight.Text > (Screen.Height / Screen.TwipsPerPixelY) Then formHeight = (Screen.Height / Screen.TwipsPerPixelY)
        txtHeight.Text = formHeight
    End If

End Sub

Private Sub picMore_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseEvent2 picMore, X, Y, BUTTON_LARGE, 0, BUTTON_DOWN

End Sub

Private Sub picMore_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseEvent2 picMore, X, Y, BUTTON_LARGE, 0, BUTTON_MOVE

End Sub

Private Sub picCancel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseEvent2 picCancel, X, Y, BUTTON_LARGE, 0, BUTTON_DOWN

End Sub

Private Sub picCancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseEvent2 picCancel, X, Y, BUTTON_LARGE, 0, BUTTON_MOVE

End Sub

Private Sub picOK_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseEvent2 picOK, X, Y, BUTTON_LARGE, 0, BUTTON_DOWN

End Sub

Private Sub picOK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseEvent2 picOK, X, Y, BUTTON_LARGE, 0, BUTTON_MOVE

End Sub

Private Sub picApply_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseEvent2 picApply, X, Y, BUTTON_LARGE, 0, BUTTON_DOWN

End Sub

Private Sub picApply_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseEvent2 picApply, X, Y, BUTTON_LARGE, 0, BUTTON_MOVE

End Sub

Private Sub picHide_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseEvent2 picHide, X, Y, BUTTON_SMALL, 0, BUTTON_DOWN

End Sub

Private Sub picHide_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseEvent2 picHide, X, Y, BUTTON_SMALL, 0, BUTTON_MOVE

End Sub

Private Sub picHide_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseEvent2 picHide, X, Y, BUTTON_SMALL, 0, BUTTON_UP

End Sub

Private Sub picfolder_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseEvent2 picFolder, X, Y, BUTTON_SMALL, 0, BUTTON_DOWN

End Sub

Private Sub picfolder_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseEvent2 picFolder, X, Y, BUTTON_SMALL, 0, BUTTON_MOVE

End Sub

Private Sub picUncomp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseEvent2 picUncomp, X, Y, BUTTON_SMALL, 0, BUTTON_DOWN

End Sub

Private Sub picUncomp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseEvent2 picUncomp, X, Y, BUTTON_SMALL, 0, BUTTON_MOVE

End Sub

Private Sub picPrefabs_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseEvent2 picPrefabs, X, Y, BUTTON_SMALL, 0, BUTTON_DOWN

End Sub

Private Sub picPrefabs_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseEvent2 picPrefabs, X, Y, BUTTON_SMALL, 0, BUTTON_MOVE

End Sub

Private Sub picScenery_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseEvent2 picScenery, X, Y, BUTTON_SMALL, sceneryVerts, BUTTON_DOWN

End Sub

Private Sub picScenery_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseEvent2 picScenery, X, Y, BUTTON_SMALL, sceneryVerts, BUTTON_MOVE

End Sub

Private Sub picScenery_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    sceneryVerts = Not sceneryVerts
    MouseEvent2 picScenery, X, Y, BUTTON_SMALL, sceneryVerts, BUTTON_UP

End Sub

Private Sub picTopmost_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseEvent2 picTopmost, X, Y, BUTTON_SMALL, topmost, BUTTON_DOWN

End Sub

Private Sub picTopmost_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseEvent2 picTopmost, X, Y, BUTTON_SMALL, topmost, BUTTON_MOVE

End Sub

Private Sub picTopmost_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    topmost = Not topmost
    MouseEvent2 picTopmost, X, Y, BUTTON_SMALL, topmost, BUTTON_UP

End Sub
