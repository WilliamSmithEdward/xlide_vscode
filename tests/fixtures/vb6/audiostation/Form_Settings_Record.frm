VERSION 5.00
Begin VB.Form Form_Settings 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "-"
   ClientHeight    =   6270
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   6660
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form_Settings_Record.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "1010"
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   " T(1018) "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5385
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.Frame Frame_MidiPlayer 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   1095
         Left            =   3360
         TabIndex        =   20
         Top             =   4080
         Width           =   2895
         Begin VB.CheckBox Checkbox_AutoStopMidi 
            BackColor       =   &H00C0C0C0&
            Caption         =   "T(1064)"
            Height          =   255
            Left            =   0
            TabIndex        =   23
            Top             =   825
            Value           =   1  'Checked
            Width           =   2175
         End
         Begin VB.OptionButton OptionButton_RemainingTimeMidi 
            BackColor       =   &H00C0C0C0&
            Caption         =   "T(1063)"
            Height          =   255
            Left            =   0
            TabIndex        =   22
            Top             =   465
            Width           =   2895
         End
         Begin VB.OptionButton OptionButton_ElapsedTimeMidi 
            BackColor       =   &H00C0C0C0&
            Caption         =   "T(1062)"
            Height          =   255
            Left            =   0
            TabIndex        =   21
            Top             =   240
            Value           =   -1  'True
            Width           =   2775
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "T(1067)"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   0
            TabIndex        =   24
            Top             =   0
            Width           =   780
         End
      End
      Begin VB.CheckBox Checkbox_AutoStop 
         BackColor       =   &H00C0C0C0&
         Caption         =   "T(1064)"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   4920
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.OptionButton OptionButton_RemainingTime 
         BackColor       =   &H00C0C0C0&
         Caption         =   "T(1063)"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   4560
         Width           =   3135
      End
      Begin VB.OptionButton OptionButton_ElapsedTime 
         BackColor       =   &H00C0C0C0&
         Caption         =   "T(1062)"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   4335
         Value           =   -1  'True
         Width           =   3015
      End
      Begin VB.ComboBox Combox_CDRomDrive 
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   420
         Width           =   3975
      End
      Begin VB.ComboBox Combox_Language 
         Height          =   315
         ItemData        =   "Form_Settings_Record.frx":000C
         Left            =   2280
         List            =   "Form_Settings_Record.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   800
         Width           =   3975
      End
      Begin VB.ComboBox Combox_MidiDevice 
         Height          =   315
         ItemData        =   "Form_Settings_Record.frx":0035
         Left            =   2280
         List            =   "Form_Settings_Record.frx":0042
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1180
         Width           =   3975
      End
      Begin VB.CheckBox Checkbox_AnimationDat 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Show the application DAT animation"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   3120
         Value           =   2  'Grayed
         Width           =   3615
      End
      Begin Audiostation.Hyperlink Hyperlink_ChangeRecorderSettings 
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   2220
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   344
         Caption         =   "T(1041)"
         ClickResponse   =   1
         BackColor       =   12632256
         ColorNormal     =   16711680
         ColorHot        =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CheckBox Checkbox_ShowSpectrum 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Show the spectrum analyzer"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   3600
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.CheckBox Checkbox_AnimationCD 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Show the application CD animation"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   3360
         Value           =   2  'Grayed
         Width           =   3375
      End
      Begin Audiostation.Hyperlink Hyperlink_ChangeAutoVolumeSettings 
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   2400
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   344
         Caption         =   "T(1042)"
         BackColor       =   12632256
         ColorNormal     =   8421504
         ColorHot        =   8421504
         ColorDown       =   8421504
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T(1044)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   4080
         Width           =   780
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T(1043)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   2880
         UseMnemonic     =   0   'False
         Width           =   780
      End
      Begin VB.Label Label_CDRomCombox 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CD-Rom Drive:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T(1038)"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   855
         Width           =   675
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T(1039)"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   1245
         Width           =   675
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T(1002)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   2025
         Width           =   780
      End
   End
   Begin Audiostation.ButtonBig Button_Cancel 
      Height          =   390
      Left            =   3390
      TabIndex        =   1
      Top             =   5760
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   688
      Caption         =   "T(1037)"
      TextAlignment   =   0
   End
   Begin Audiostation.ButtonBig Button_Save 
      Height          =   390
      Left            =   2190
      TabIndex        =   2
      Top             =   5760
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   688
      Caption         =   "T(1017)"
      TextAlignment   =   0
   End
End
Attribute VB_Name = "Form_Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private CurrentLanguage As String
Private Sub Button_Cancel_Click()
Unload Me
End Sub

Private Sub Button_Save_Click()
If Combox_Language.Text <> CurrentLanguage Then: MsgBox GetTranslation(1065), vbExclamation + vbOKOnly

Call Extensions.INIWrite("main", "CDDeviceId", Combox_CDRomDrive.ListIndex, ConfigFile)
Call Extensions.INIWrite("main", "MidiPlaybackDeviceId", Combox_MidiDevice.ListIndex, ConfigFile)
Call Extensions.INIWrite("main", "UseSpectrumAnalyzer", Checkbox_ShowSpectrum.Value, ConfigFile)
Call Extensions.INIWrite("main", "Langauge", Combox_Language.Text, ConfigFile)
Call Extensions.INIWrite("main", "Animation-CD", Checkbox_AnimationCD.Value, ConfigFile)
Call Extensions.INIWrite("main", "Animation-DAT", Checkbox_AnimationDat.Value, ConfigFile)
Call Extensions.INIWrite("main", "AutoStop", Checkbox_AutoStop.Value, ConfigFile)
Call Extensions.INIWrite("main", "ShowRemainingTime", OptionButton_RemainingTime.Value, ConfigFile)
Call Extensions.INIWrite("main", "ShowRemainingTimeForMidi", OptionButton_RemainingTimeMidi.Value, ConfigFile)
Call Extensions.INIWrite("main", "AutoStopForMidi", Checkbox_AutoStopMidi.Value, ConfigFile)

Form_Main.menu_Popup_AutoStop.Checked = Checkbox_AutoStop.Value
Form_Main.menu_Popup_ShowSpectrum.Checked = Checkbox_ShowSpectrum.Value
Form_Main.ShowRemaining = OptionButton_RemainingTime.Value

Form_Main.SettingsChanged

Unload Me
End Sub

Private Sub Form_Load()
' Get CD-Rom devices
Dim CdDevice As mdlAdioCdDrive

Combox_CDRomDrive.Clear
For Each CdDevice In Form_Main.AdioCDPlayer.GetListOfDrives
    Combox_CDRomDrive.AddItem CdDevice.cdLetter & ":"
Next

' Get MIDI devices
Dim MidiDevice As mdlAdioMidiDevice

Combox_MidiDevice.Clear
For Each MidiDevice In Form_Main.AdioMidiPlayer.GetListOfMidiDevices
    Combox_MidiDevice.AddItem MidiDevice.mName
Next

'If Combox_PlaybackDevice.ListCount > 0 Then: Combox_PlaybackDevice.ListIndex = Extensions.INIRead("main", "PlaybackDeviceId", ConfigFile, 0)
If Combox_MidiDevice.ListCount > 0 Then: Combox_MidiDevice.ListIndex = Extensions.INIRead("main", "MidiPlaybackDeviceId", ConfigFile, 0)
If Combox_CDRomDrive.ListCount > 0 Then
    Combox_CDRomDrive.ListIndex = Extensions.INIRead("main", "CDDeviceId", ConfigFile, 0)
Else
    Combox_CDRomDrive.Enabled = False
    Label_CDRomCombox.Enabled = False
End If

Combox_Language.Text = LCase(Extensions.INIRead("main", "Langauge", ConfigFile, "English"))

Checkbox_AutoStop.Value = Extensions.INIRead("main", "AutoStop", ConfigFile, 1)
Checkbox_ShowSpectrum.Value = Extensions.INIRead("main", "UseSpectrumAnalyzer", ConfigFile, 1)
Checkbox_AnimationCD.Value = Extensions.INIRead("main", "Animation-CD", ConfigFile, 1)
Checkbox_AnimationDat.Value = Extensions.INIRead("main", "Animation-DAT", ConfigFile, 1)
Checkbox_AutoStopMidi.Value = Extensions.INIRead("main", "AutoStopForMidi", ConfigFile, 1)

OptionButton_RemainingTime.Value = Extensions.INIRead("main", "ShowRemainingTime", ConfigFile, 0)
OptionButton_RemainingTimeMidi.Value = Extensions.INIRead("main", "ShowRemainingTimeForMidi", ConfigFile, 0)

' Set the current selected language
CurrentLanguage = Combox_Language.Text

Call TranslateFormAndControls(Me)
End Sub

Private Sub Hyperlink_ChangeAutoVolumeSettings_Click()
Form_Settings_AutoVolume.Show vbModal, Me
End Sub

Private Sub Hyperlink_ChangeRecorderSettings_Click()
Form_Settings_Recorder.Show vbModal, Me
End Sub
