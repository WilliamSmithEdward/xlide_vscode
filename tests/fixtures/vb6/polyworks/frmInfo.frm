VERSION 5.00
Begin VB.Form frmInfo 
   BackColor       =   &H004A3C31&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3120
   ClientLeft      =   120
   ClientTop       =   120
   ClientWidth     =   3120
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   208
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   208
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      ScaleWidth      =   208
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   0
      Width           =   3120
      Begin VB.PictureBox picPropMenu 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H004A3C31&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   2640
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   38
         TabStop         =   0   'False
         Tag             =   "8"
         Top             =   0
         Width           =   240
      End
      Begin VB.PictureBox picHide 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H004A3C31&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   2880
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   17
         TabStop         =   0   'False
         Tag             =   "3"
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.CommandButton cmdDefault 
      Default         =   -1  'True
      Height          =   495
      Left            =   240
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   3240
      Width           =   495
   End
   Begin VB.PictureBox picProp 
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   4
      Left            =   120
      ScaleHeight     =   129
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   193
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   2895
      Begin VB.TextBox txtLightProp 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   74
         Tag             =   "font1"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtLightProp 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   72
         Tag             =   "font1"
         Top             =   120
         Width           =   615
      End
      Begin VB.TextBox txtLightProp 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   68
         Tag             =   "font1"
         Top             =   840
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.PictureBox picLight 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   2280
         ScaleHeight     =   465
         ScaleWidth      =   465
         TabIndex        =   67
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00614B3D&
         Caption         =   "Range:"
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
         Left            =   120
         TabIndex        =   73
         Tag             =   "font2"
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00614B3D&
         Caption         =   "Z-coord:"
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
         Left            =   120
         TabIndex        =   71
         Tag             =   "font2"
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00614B3D&
         Caption         =   "Intensity:"
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
         Left            =   120
         TabIndex        =   70
         Tag             =   "font2"
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
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
         Index           =   28
         Left            =   1920
         TabIndex        =   69
         Tag             =   "font2"
         Top             =   840
         Visible         =   0   'False
         Width           =   255
      End
   End
   Begin VB.PictureBox picProp 
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   5
      Left            =   120
      ScaleHeight     =   129
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   193
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   360
      Width           =   2895
      Begin VB.Label lblCount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0x0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   1680
         TabIndex        =   65
         Tag             =   "font1"
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Dimensions:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   27
         Left            =   120
         TabIndex        =   64
         Tag             =   "font2"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label lblCount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "500"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   1680
         TabIndex        =   61
         Tag             =   "font1"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lblCount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "500/500"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   1680
         TabIndex        =   60
         Tag             =   "font1"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblCount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "128/128"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   1680
         TabIndex        =   59
         Tag             =   "font1"
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lblCount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "128/128"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   1680
         TabIndex        =   58
         Tag             =   "font1"
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblCount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "500/500 (500)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   1380
         TabIndex        =   57
         Tag             =   "font1"
         Top             =   360
         Width           =   1395
      End
      Begin VB.Label lblCount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "5000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   1680
         TabIndex        =   56
         Tag             =   "font1"
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Connections:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   55
         Tag             =   "font2"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Waypoints:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   54
         Tag             =   "font2"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Spawns:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   25
         Left            =   120
         TabIndex        =   53
         Tag             =   "font2"
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Colliders:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   26
         Left            =   120
         TabIndex        =   52
         Tag             =   "font2"
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Polygons:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   23
         Left            =   120
         TabIndex        =   51
         Tag             =   "font2"
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Scenery:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   24
         Left            =   120
         TabIndex        =   50
         Tag             =   "font2"
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.PictureBox picProp 
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   1
      Left            =   120
      ScaleHeight     =   129
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   193
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   2895
      Begin VB.TextBox txtScenProp 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1680
         TabIndex        =   4
         Tag             =   "font1"
         Top             =   120
         Width           =   735
      End
      Begin VB.TextBox txtScenProp 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   5
         Tag             =   "font1"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtScenProp 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   7
         Tag             =   "font1"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtScenProp 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   1320
         TabIndex        =   8
         Tag             =   "font1"
         Top             =   1560
         Width           =   615
      End
      Begin VB.ComboBox cboLevel 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         ItemData        =   "frmInfo.frx":0000
         Left            =   1320
         List            =   "frmInfo.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Tag             =   "font1"
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
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
         Index           =   20
         Left            =   2400
         TabIndex        =   46
         Tag             =   "font2"
         Top             =   480
         Width           =   255
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
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
         Index           =   19
         Left            =   2400
         TabIndex        =   45
         Tag             =   "font2"
         Top             =   120
         Width           =   255
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
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
         Index           =   21
         Left            =   1920
         TabIndex        =   31
         Tag             =   "font2"
         Top             =   1200
         Width           =   255
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "°"
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
         Left            =   1920
         TabIndex        =   30
         Tag             =   "font2"
         Top             =   1560
         Width           =   255
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00614B3D&
         Caption         =   "Scaling:"
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
         Left            =   120
         TabIndex        =   29
         Tag             =   "font2"
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "X:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   17
         Left            =   1320
         TabIndex        =   28
         Tag             =   "font2"
         Top             =   120
         Width           =   255
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Y:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   18
         Left            =   1320
         TabIndex        =   27
         Tag             =   "font2"
         Top             =   480
         Width           =   255
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00614B3D&
         Caption         =   "Opacity:"
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
         Left            =   120
         TabIndex        =   26
         Tag             =   "font2"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00614B3D&
         Caption         =   "Rotation:"
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
         Left            =   120
         TabIndex        =   25
         Tag             =   "font2"
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00614B3D&
         Caption         =   "Level:"
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
         Left            =   120
         TabIndex        =   24
         Tag             =   "font2"
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.PictureBox picProp 
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   3
      Left            =   120
      ScaleHeight     =   129
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   193
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   2895
      Begin VB.TextBox txtQuadX 
         Appearance      =   0  'Flat
         BackColor       =   &H004A3C31&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   11
         Tag             =   "font1"
         Text            =   "128"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtQuadY 
         Appearance      =   0  'Flat
         BackColor       =   &H004A3C31&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   12
         Tag             =   "font1"
         Text            =   "0"
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox txtQuadY 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H004A3C31&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Tag             =   "font1"
         Text            =   "0"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtQuadX 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H004A3C31&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Tag             =   "font1"
         Text            =   "0"
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblHeight 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2400
         TabIndex        =   63
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblDimensions 
         BackStyle       =   0  'Transparent
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
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Tag             =   "font2"
         Top             =   0
         Width           =   2655
      End
      Begin VB.Line diagonal 
         BorderColor     =   &H00FFFFFF&
         X1              =   64
         X2              =   128
         Y1              =   40
         Y2              =   104
      End
      Begin VB.Shape square 
         BorderColor     =   &H00FFFFFF&
         Height          =   975
         Left            =   960
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.PictureBox picProp 
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   2
      Left            =   120
      ScaleHeight     =   129
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   193
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   2895
      Begin VB.TextBox txtRotate 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   15
         Tag             =   "font1"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txtScale 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   14
         Tag             =   "font1"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txtScale 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1680
         TabIndex        =   13
         Tag             =   "font1"
         Top             =   120
         Width           =   735
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
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
         Index           =   12
         Left            =   2400
         TabIndex        =   48
         Tag             =   "font2"
         Top             =   120
         Width           =   255
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
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
         Index           =   13
         Left            =   2400
         TabIndex        =   47
         Tag             =   "font2"
         Top             =   480
         Width           =   255
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00614B3D&
         Caption         =   "Rotation:"
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
         Left            =   120
         TabIndex        =   44
         Tag             =   "font2"
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Y:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   11
         Left            =   1320
         TabIndex        =   43
         Tag             =   "font2"
         Top             =   480
         Width           =   255
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "X:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   10
         Left            =   1320
         TabIndex        =   42
         Tag             =   "font2"
         Top             =   120
         Width           =   255
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00614B3D&
         Caption         =   "Scaling:"
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
         Left            =   120
         TabIndex        =   41
         Tag             =   "font2"
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "°"
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
         Left            =   2040
         TabIndex        =   40
         Tag             =   "font2"
         Top             =   840
         Width           =   255
      End
   End
   Begin VB.PictureBox picProp 
      Appearance      =   0  'Flat
      BackColor       =   &H004A3C31&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   0
      Left            =   120
      ScaleHeight     =   129
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   193
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   2895
      Begin VB.TextBox txtBounciness 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   78
         Tag             =   "font1"
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox txtVertexAlpha 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   3
         Tag             =   "font1"
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox txtTexture 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   2
         Tag             =   "font1"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtTexture 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   1
         Tag             =   "font1"
         Top             =   600
         Width           =   975
      End
      Begin VB.ComboBox cboPolyType 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         ItemData        =   "frmInfo.frx":0026
         Left            =   840
         List            =   "frmInfo.frx":0078
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Tag             =   "font1"
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
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
         Index           =   35
         Left            =   2160
         TabIndex        =   79
         Tag             =   "font2"
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00614B3D&
         Caption         =   "Bounciness:"
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
         Index           =   34
         Left            =   120
         TabIndex        =   77
         Tag             =   "font2"
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label lblInfo 
         Alignment       =   2  'Center
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
         Index           =   33
         Left            =   2160
         TabIndex        =   76
         Tag             =   "font2"
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00614B3D&
         Caption         =   "Opacity:"
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
         Index           =   32
         Left            =   120
         TabIndex        =   75
         Tag             =   "font2"
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Y:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   16
         Left            =   1200
         TabIndex        =   37
         Tag             =   "font2"
         Top             =   960
         Width           =   255
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "X:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   15
         Left            =   1200
         TabIndex        =   36
         Tag             =   "font2"
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00614B3D&
         Caption         =   "Texture:"
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
         Left            =   120
         TabIndex        =   35
         Tag             =   "font2"
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H00614B3D&
         Caption         =   "Type:"
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
         Left            =   120
         TabIndex        =   33
         Tag             =   "font2"
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.Label lblIndex 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2280
      TabIndex        =   21
      Tag             =   "font1"
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label lblSelScenery 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Tag             =   "font1"
      Top             =   2760
      Width           =   2895
   End
   Begin VB.Label lblSelPolys 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label lblCoords 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Tag             =   "font1"
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Menu mnuProperties 
      Caption         =   "Properties"
      Visible         =   0   'False
      Begin VB.Menu mnuProp 
         Caption         =   "Polygon Properties"
         Index           =   0
      End
      Begin VB.Menu mnuProp 
         Caption         =   "Scenery Properties"
         Index           =   1
      End
      Begin VB.Menu mnuProp 
         Caption         =   "Transform"
         Index           =   2
      End
      Begin VB.Menu mnuProp 
         Caption         =   "Texture Settings"
         Index           =   3
      End
      Begin VB.Menu mnuProp 
         Caption         =   "Light Properties"
         Index           =   4
      End
      Begin VB.Menu mnuProp 
         Caption         =   "Map Info"
         Checked         =   -1  'True
         Index           =   5
      End
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' info dialog - display current map information


' fix vb6 ide casing changes
#If False Then
    Public FileName, color, token, A, R, G, B, commonDialog, value, Val, X, Y, Z, Left, hWnd, Mid, Right, BackColor, bottom
    'Public FileName, color, token, A, R, G, B, commonDialog, value, Val, X, Y, Z, Left, hWnd, Mid, Right, BackColor, bottom
#End If


' consts - public


' consts - private


' types - public


' types - private


' vars - public

Public xPos As Integer
Public yPos  As Integer
Public collapsed As Boolean

Public noChange As Boolean


' vars - private

Private formHeight As Integer

Private tempVal As Single

Private applyChange As Boolean


' external functions - public


' external functions - private


' functions - public

Public Sub SetForm()

    Me.Left = xPos * Screen.TwipsPerPixelX
    Me.Top = yPos * Screen.TwipsPerPixelY
    If collapsed Then
        Me.Height = 19 * Screen.TwipsPerPixelY
    Else
        Me.Height = formHeight * Screen.TwipsPerPixelY
    End If

End Sub

Public Sub SetColors()

    On Error Resume Next

    Dim i As Integer
    Dim c As Control

    picTitle.Picture = LoadPicture(appPath & "\skins\" & gfxDir & "\titlebar_properties.bmp")
    MouseEvent2 picHide, 0, 0, BUTTON_SMALL, 0, BUTTON_UP
    MouseEvent2 picPropMenu, 0, 0, BUTTON_SMALL, 0, BUTTON_UP

    Me.BackColor = bgColor
    For Each c In lblInfo
        c.BackColor = lblBackColor
        c.ForeColor = lblTextColor
    Next
    For Each c In picProp
        c.BackColor = bgColor
    Next

    For Each c In txtScenProp
        c.BackColor = txtBackColor
        c.ForeColor = txtTextColor
    Next
    For Each c In txtQuadX
        c.BackColor = bgColor
        c.ForeColor = lblTextColor
    Next
    For Each c In txtQuadY
        c.BackColor = bgColor
        c.ForeColor = lblTextColor
    Next
    For Each c In txtScale
        c.BackColor = txtBackColor
        c.ForeColor = txtTextColor
    Next
    For Each c In txtTexture
        c.BackColor = txtBackColor
        c.ForeColor = txtTextColor
    Next
    For Each c In lblCount
        c.BackColor = lblBackColor
        c.ForeColor = lblTextColor
    Next

    txtVertexAlpha.BackColor = txtBackColor
    txtVertexAlpha.ForeColor = txtTextColor
    txtBounciness.BackColor = txtBackColor
    txtBounciness.ForeColor = txtTextColor

    lblDimensions.BackColor = lblBackColor
    lblDimensions.ForeColor = lblTextColor

    txtRotate.BackColor = txtBackColor
    txtRotate.ForeColor = txtTextColor
    cboLevel.BackColor = txtBackColor
    cboLevel.ForeColor = txtTextColor
    cboPolyType.BackColor = txtBackColor
    cboPolyType.ForeColor = txtTextColor

    For Each c In txtLightProp
        c.BackColor = txtBackColor
        c.ForeColor = txtTextColor
    Next

    square.BorderColor = lblTextColor
    diagonal.BorderColor = lblTextColor

    SetFormFonts Me

End Sub


' functions - private


' events - public

Public Sub mnuProp_Click(Index As Integer)

    Dim i As Integer

    For i = mnuProp.LBound To mnuProp.UBound
        mnuProp(i).Checked = False
    Next
    For i = picProp.LBound To picProp.UBound
        picProp(i).Visible = False
    Next

    mnuProp(Index).Checked = True

    picProp(Index).Visible = True

End Sub


' events - private

Private Sub Form_Load()

    On Error GoTo ErrorHandler

    Me.SetColors

    formHeight = Me.ScaleHeight

    SetForm

    cboPolyType.ListIndex = 0
    lblDimensions.Caption = "Dimensions: " & frmOpenSoldatMapEditor.xTexture & " x " & frmOpenSoldatMapEditor.yTexture
    txtQuadX(0).Text = 0
    txtQuadY(0).Text = 0
    txtQuadX(1).Text = frmOpenSoldatMapEditor.xTexture
    txtQuadY(1).Text = frmOpenSoldatMapEditor.yTexture

    Exit Sub

ErrorHandler:

    MsgBox "Error loading Properties form" & vbNewLine & Error

End Sub

Private Sub cboPolyType_Click()

    If Not noChange Then
        frmOpenSoldatMapEditor.ApplyPolyType cboPolyType.ListIndex
    End If

    If cboPolyType.ListIndex = 18 Then
        txtBounciness.Enabled = True
    Else
        txtBounciness.Enabled = False
    End If

End Sub

Private Sub txtLightProp_GotFocus(Index As Integer)

    If IsNumeric(txtLightProp(Index).Text) Then
        tempVal = txtLightProp(Index).Text
    End If
    SelectAllText txtLightProp(Index)

End Sub

Private Sub txtLightProp_LostFocus(Index As Integer)

    If IsNumeric(txtLightProp(Index).Text) And applyChange Then
        If Index = 0 Then
            frmOpenSoldatMapEditor.ApplyLightProp txtLightProp(Index).Text, Index
        ElseIf Index = 1 And txtLightProp(Index).Text >= 0 Then
            frmOpenSoldatMapEditor.ApplyLightProp txtLightProp(Index).Text, Index
        ElseIf Index = 1 And txtLightProp(Index).Text >= 0 And txtLightProp(Index).Text <= 100 Then
           ' no op
        Else
            txtLightProp(Index).Text = tempVal
        End If
    Else
        txtLightProp(Index).Text = tempVal
    End If
    tempVal = 0

    applyChange = False

End Sub

Private Sub picLight_Click()

    frmOpenSoldatMapEditor.SetLightColor

End Sub

Private Sub txtQuadX_GotFocus(Index As Integer)

    If IsNumeric(txtQuadX(Index).Text) Then
        tempVal = txtQuadX(Index).Text
    End If
    SelectAllText txtQuadX(Index)

End Sub

Private Sub txtQuadX_LostFocus(Index As Integer)

    If Not IsNumeric(txtQuadX(Index).Text) Then
        txtQuadX(Index).Text = tempVal
    ElseIf txtQuadX(Index).Text < 0 Or txtQuadX(Index).Text > frmOpenSoldatMapEditor.xTexture Then
        txtQuadX(Index).Text = tempVal
    Else
        frmTexture.SetTexCoords txtQuadX(Index).Text, Index
    End If
    tempVal = 0

End Sub

Private Sub txtQuadY_GotFocus(Index As Integer)

    If IsNumeric(txtQuadY(Index).Text) Then
        tempVal = txtQuadY(Index).Text
    End If
    SelectAllText txtQuadY(Index)

End Sub

Private Sub txtQuadY_LostFocus(Index As Integer)

    If Not IsNumeric(txtQuadY(Index).Text) Then
        txtQuadY(Index).Text = tempVal
    ElseIf txtQuadY(Index).Text < 0 Or txtQuadY(Index).Text > frmOpenSoldatMapEditor.yTexture Then
        txtQuadY(Index).Text = tempVal
    Else
        frmTexture.SetTexCoords txtQuadY(Index).Text, Index + 2
    End If
    tempVal = 0

End Sub

Private Sub txtRotate_GotFocus()

    If IsNumeric(txtRotate.Text) Then
        tempVal = txtRotate.Text
    End If

End Sub

Private Sub txtRotate_LostFocus()

    If IsNumeric(txtRotate.Text) And applyChange Then
        frmOpenSoldatMapEditor.ApplyRotate (txtRotate.Text / 180 * PI)
    Else
        txtRotate.Text = tempVal
    End If
    tempVal = 0

End Sub

Private Sub txtScale_GotFocus(Index As Integer)

    If IsNumeric(txtScale(Index).Text) Then
        tempVal = txtScale(Index).Text
    End If

End Sub

Private Sub txtScale_LostFocus(Index As Integer)

    If IsNumeric(txtScale(Index).Text) And applyChange Then
        If Index = 0 Then
            frmOpenSoldatMapEditor.ApplyScale (txtScale(Index).Text / 100), 1
        ElseIf Index = 1 Then
            frmOpenSoldatMapEditor.ApplyScale 1, (txtScale(Index).Text / 100)
        End If
    Else
        txtScale(Index).Text = tempVal
    End If

    tempVal = 0
    applyChange = False

End Sub

Private Sub cboLevel_Click()

    If Not noChange Then
        frmOpenSoldatMapEditor.ApplySceneryProp cboLevel.ListIndex, 4
    End If

End Sub

Private Sub txtScenProp_GotFocus(Index As Integer)

    If IsNumeric(txtScenProp(Index).Text) Then
        tempVal = txtScenProp(Index).Text
    End If
    SelectAllText txtScenProp(Index)

End Sub

Private Sub txtScenProp_LostFocus(Index As Integer)

    If IsNumeric(txtScenProp(Index).Text) And applyChange Then
        If Index = 0 Or Index = 1 Then
            frmOpenSoldatMapEditor.ApplySceneryProp txtScenProp(Index).Text / 100, Index
        ElseIf Index = 2 And txtScenProp(Index).Text >= 0 And txtScenProp(Index).Text <= 100 Then
            frmOpenSoldatMapEditor.ApplySceneryProp (txtScenProp(Index).Text / 100) * 255, Index
        ElseIf Index = 3 And txtScenProp(Index).Text >= -360 And txtScenProp(Index).Text <= 360 Then
            frmOpenSoldatMapEditor.ApplySceneryProp txtScenProp(Index).Text / 180 * PI, Index
        Else
            txtScenProp(Index).Text = tempVal
        End If
    Else
        txtScenProp(Index).Text = tempVal
    End If
    tempVal = 0

    applyChange = False

End Sub

Private Sub txtTexture_GotFocus(Index As Integer)

    If IsNumeric(txtTexture(Index).Text) Then
        tempVal = txtTexture(Index).Text
    End If

End Sub

Private Sub txtTexture_LostFocus(Index As Integer)

    If IsNumeric(txtTexture(Index).Text) And applyChange Then
        frmOpenSoldatMapEditor.ApplyTextureCoords txtTexture(Index).Text, Index
    Else
        txtTexture(Index).Text = tempVal
    End If
    tempVal = 0

    applyChange = False

End Sub

Private Sub txtVertexAlpha_GotFocus()

    If IsNumeric(txtVertexAlpha.Text) Then
        tempVal = txtVertexAlpha.Text
    End If

End Sub

Private Sub txtVertexAlpha_LostFocus()

    If Not IsNumeric(txtVertexAlpha.Text) Then
        txtVertexAlpha.Text = tempVal
    ElseIf txtVertexAlpha.Text < 0 Or txtVertexAlpha.Text > 100 Then
        txtVertexAlpha.Text = tempVal
    ElseIf applyChange Then
        frmOpenSoldatMapEditor.ApplyVertexAlpha txtVertexAlpha.Text / 100
    End If

    tempVal = 0
    applyChange = False

End Sub

Private Sub txtBounciness_GotFocus()

    If IsNumeric(txtBounciness.Text) Then
        tempVal = txtBounciness.Text
    End If

End Sub

Private Sub txtBounciness_LostFocus()

    If Not IsNumeric(txtBounciness.Text) Then
        txtBounciness.Text = tempVal
    ElseIf txtBounciness.Text < 0 Then
        txtBounciness.Text = tempVal
    ElseIf applyChange Then
        frmOpenSoldatMapEditor.ApplyBounciness 1 + (txtBounciness.Text / 100)
    End If

    tempVal = 0
    applyChange = False

End Sub

Private Sub cmdDefault_Click()

    applyChange = True
    cmdDefault.SetFocus
    frmOpenSoldatMapEditor.RegainFocus

End Sub

Private Sub picTitle_DblClick()

    collapsed = Not collapsed
    If collapsed Then
        Me.Height = 19 * Screen.TwipsPerPixelY
    Else
        Me.Height = formHeight * Screen.TwipsPerPixelY
    End If

End Sub

Private Sub picTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ReleaseCapture
    SendMessage Me.hWnd, WM_NCLBUTTONDOWN, 2, 0&

    SnapForm Me, frmPalette
    SnapForm Me, frmWaypoints
    SnapForm Me, frmTools
    SnapForm Me, frmScenery
    SnapForm Me, frmDisplay
    SnapForm Me, frmTexture
    Me.Tag = SnapForm(Me, frmOpenSoldatMapEditor)

    xPos = Me.Left / Screen.TwipsPerPixelX
    yPos = Me.Top / Screen.TwipsPerPixelY

End Sub

Private Sub picHide_Click()

    Me.Hide
    frmOpenSoldatMapEditor.mnuInfo.Checked = False

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

Private Sub picPropMenu_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseEvent2 picPropMenu, X, Y, BUTTON_SMALL, 0, BUTTON_DOWN

    PopupMenu mnuProperties, , picPropMenu.Left + 32, picPropMenu.Top + 16

    MouseEvent2 picPropMenu, X, Y, BUTTON_SMALL, 0, BUTTON_UP

End Sub

Private Sub picPropMenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseEvent2 picPropMenu, X, Y, BUTTON_SMALL, 0, BUTTON_MOVE

End Sub
