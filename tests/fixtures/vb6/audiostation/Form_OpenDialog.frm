VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.OCX"
Begin VB.Form Form_OpenStream 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Open Stream"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8295
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form_OpenDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab Tab_Strip 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   8493
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabMaxWidth     =   2117
      BackColor       =   12632256
      TabCaption(0)   =   "Existing"
      TabPicture(0)   =   "Form_OpenDialog.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Listview_Streams"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Imagelist_Listview"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin MSComctlLib.ImageList Imagelist_Listview 
         Left            =   7320
         Top             =   480
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_OpenDialog.frx":0028
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_OpenDialog.frx":0183
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView Listview_Streams 
         Height          =   4275
         Left            =   120
         TabIndex        =   3
         Top             =   400
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   7541
         View            =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "Imagelist_Listview"
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         Appearance      =   0
         NumItems        =   0
      End
   End
   Begin Audiostation.ButtonBig Button_Cancel 
      Height          =   390
      Left            =   6840
      TabIndex        =   1
      Top             =   5040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   688
      Caption         =   "Cancel"
      TextAlignment   =   0
   End
   Begin Audiostation.ButtonBig Button_Open 
      Height          =   390
      Left            =   5400
      TabIndex        =   2
      Top             =   5040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   688
      Caption         =   "Open"
      TextAlignment   =   0
   End
End
Attribute VB_Name = "Form_OpenStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CurrentRowIndex As Integer
Private Function CheckIfListviewExists(value As String) As Boolean
Dim Item As ListItem

For Each Item In Listview_Streams.ListItems
    If Item.Text = value Then
        CheckIfListviewExists = True
        Exit Function
    End If
Next

CheckIfListviewExists = False
End Function
Private Sub OpenStreamDatabase(RowIndex As Integer)
Dim Streams As String
Dim DbValue As String

Streams = Extensions.FileGetContents(App.path & "\streams.db")
Listview_Streams.ListItems.Clear

For Each StreamItem In Split(Streams, vbNewLine)
    If StreamItem <> vbNullString Then
        DbValue = Extensions.Explode(CStr(StreamItem), ",", RowIndex)
        
        If InStr(1, DbValue, ";") > 0 Then
            If Not CheckIfListviewExists(DbValue) Then
                Dim lstItem As ListItem
                
                Set lstItem = Listview_Streams.ListItems.Add(, , Trim(Extensions.Explode(DbValue, ";", 0)))
                lstItem.Tag = Trim(Extensions.Explode(DbValue, ";", 1))
                lstItem.SmallIcon = 2
            End If
        Else
            If Not CheckIfListviewExists(DbValue) Then Listview_Streams.ListItems.Add , , Trim(DbValue), , 1
        End If
    End If
Next

CurrentRowIndex = RowIndex
End Sub
Private Sub Button_Cancel_Click()
Unload Me
End Sub

Private Sub Button_Open_Click()
AudioStaStreamer.Name = Listview_Streams.SelectedItem.Text
AudioStaStreamer.Url = Listview_Streams.SelectedItem.Tag

Unload Me
End Sub

Private Sub Form_Load()
CurrentRowIndex = 0
Button_Open.Enabled = False

Call OpenStreamDatabase(CurrentRowIndex)
End Sub
Private Sub Listview_Streams_DblClick()
If Listview_Streams.SelectedItem.SmallIcon = 1 Then
    Call OpenStreamDatabase(CurrentRowIndex + 1)
Else
    AudioStaStreamer.Name = Listview_Streams.SelectedItem.Text
    AudioStaStreamer.Url = Listview_Streams.SelectedItem.Tag
    
    Unload Me
End If
End Sub

