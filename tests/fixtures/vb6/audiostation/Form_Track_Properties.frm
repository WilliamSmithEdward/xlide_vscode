VERSION 5.00
Begin VB.Form Form_Track_Properties 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "-"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6000
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form_Track_Properties.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "1047"
   Begin Audiostation.ButtonBig Button_Close 
      Height          =   390
      Left            =   4800
      TabIndex        =   1
      Top             =   3120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   688
      Caption         =   "T(1025)"
      TextAlignment   =   0
   End
   Begin VB.TextBox Textbox_Properties 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3495
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form_Track_Properties.frx":000C
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "Form_Track_Properties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public File As String
Private Sub Button_Close_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim Builder As New clsStringBuilder

With Form_Main
    Builder.AppendNL GetTranslation(1048) & Space(1) & .AdioTagging.ReadTag(File, tArtist, v2)
    Builder.AppendNL GetTranslation(1049) & Space(1) & .AdioTagging.ReadTag(File, tTitle, v2)
    Builder.AppendNL GetTranslation(1050) & Space(1) & .AdioTagging.ReadTag(File, tAlbum, v2)
    Builder.AppendNL GetTranslation(1051) & Space(1) & .AdioTagging.ReadTag(File, tYear, v2)
    Builder.AppendNL GetTranslation(1052) & Space(1) & .AdioTagging.ReadTag(File, tGenre, v2)
    Builder.AppendNL vbNullString
    Builder.AppendNL "Biterate: " & .AdioTagging.ReadProperty(File, pBitrate) & "kbps"
    Builder.AppendNL "Channel Mode: " & .AdioTagging.ReadProperty(File, pChannels)
    Builder.AppendNL "Frequency: " & .AdioTagging.ReadProperty(File, pFrequency)
    Builder.AppendNL "Length: " & .AdioTagging.ReadProperty(File, pDurationInSeconds)
    Builder.AppendNL "Size: " & Extensions.SetBytes(FileLen(File))
End With

Textbox_Properties.Text = Builder.toString

Call TranslateFormAndControls(Me)
End Sub

