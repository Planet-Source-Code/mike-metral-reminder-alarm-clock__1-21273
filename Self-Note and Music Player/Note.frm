VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form frmNoteAlarm 
   BorderStyle     =   0  'None
   Caption         =   "Self-Notes And Digital Alarm Clock"
   ClientHeight    =   3345
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   6225
   HasDC           =   0   'False
   Icon            =   "Note.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   6225
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkSetNote 
      BackColor       =   &H00FFFF80&
      Caption         =   "Set Note"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   1200
      MouseIcon       =   "Note.frx":030A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   720
      Width           =   1215
   End
   Begin VB.Timer tmrAlarm 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   1080
   End
   Begin VB.Timer tmrMessage 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   2040
   End
   Begin VB.Timer tmrTime 
      Interval        =   1000
      Left            =   120
      Top             =   1560
   End
   Begin VB.OptionButton optPm 
      Caption         =   "PM"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   4920
      TabIndex        =   9
      Top             =   2160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.OptionButton optAm 
      Caption         =   "AM"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4200
      TabIndex        =   8
      Top             =   2160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtTime 
      Height          =   285
      Left            =   2640
      TabIndex        =   7
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtNote 
      Height          =   285
      Left            =   2640
      TabIndex        =   6
      Top             =   2760
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CheckBox chkNote 
      BackColor       =   &H00FFFF80&
      Caption         =   "Note"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   1200
      MouseIcon       =   "Note.frx":0614
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox chkTime 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Set Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   1200
      MouseIcon       =   "Note.frx":091E
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox chkControls 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Show Controls"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   3840
      MouseIcon       =   "Note.frx":0C28
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.CheckBox chkMusic 
      BackColor       =   &H0080C0FF&
      Caption         =   "Listen to Music"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   2520
      MouseIcon       =   "Note.frx":0F32
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.CheckBox chkOpenMusic 
      BackColor       =   &H0080C0FF&
      Caption         =   "Music"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   1200
      MouseIcon       =   "Note.frx":123C
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MCI.MMControl MMControl1 
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   661
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   4080
      TabIndex        =   12
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblTime 
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   10
      Top             =   240
      Width           =   1575
   End
   Begin VB.Menu mnuTimeColor 
      Caption         =   "&TimeColor"
   End
   Begin VB.Menu mnuStart 
      Caption         =   "&Start Alarm"
   End
   Begin VB.Menu mnuExit 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "frmNoteAlarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkControls_Click()
 If chkControls.Caption = "Show Controls" And chkControls.Value = 1 Then
  MMControl1.Visible = True
  chkControls.Caption = "Hide Controls"
 ElseIf chkControls.Caption = "Hide Controls" And chkControls.Value = 0 Then
  MMControl1.Visible = False
  chkControls.Caption = "Show Controls"
 End If

End Sub

Private Sub chkMusic_Click()

 If chkMusic.Value = 1 Then
  chkSetNote.Enabled = False
  chkOpenMusic.Visible = True
 Else
  chkSetNote.Enabled = True
  chkOpenMusic.Visible = False
End If

End Sub


Private Sub chkNote_Click()
 If chkNote.Value = 1 Then
   txtNote.Visible = True

 ElseIf chkNote.Value = 0 Then
   txtNote.Visible = False
 End If
 
End Sub

Private Sub chkOpenMusic_Click()
 Dim strFilename As String
 tmrAlarm.Enabled = False
 If chkMusic.Value = 1 And chkOpenMusic.Value = 1 Then
  If MMControl1.Command = "Play" Then
  MMControl1.Command = "Close"
  End If
  CommonDialog1.Filter = "Media Files |*.mp3; *.wav; *.mid; *.midi; *.cda"
  CommonDialog1.ShowOpen
  strFilename = CommonDialog1.FileName
  If Right(strFilename, 3) = "mp3" Then
  MMControl1.DeviceType = "MPEGVideo"
  ElseIf Right(strFilename, 3) = "wav" Then
  MMControl1.DeviceType = "WaveAudio"
  ElseIf Right(strFilename, 3) = "wav" Or Right(strFilename, 4) = "midi" Then
  MMControl1.DeviceType = "Sequencer"
  ElseIf Right(strFilename, 3) = "CDA" Then
  MMControl1.DeviceType = "CDAudio"
  End If
  MMControl1.FileName = CommonDialog1.FileName
  MMControl1.Command = "Open"
  MMControl1.Command = "Play"
  tmrMessage.Enabled = True
 ElseIf chkSetNote.Value = 1 And chkOpenMusic.Value = 1 Then
  If MMControl1.Command = "Play" Then
  MMControl1.Command = "Close"
  End If
  CommonDialog1.Filter = "Media Files |*.mp3; *.wav; *.mid; *.midi; *.cda"
  CommonDialog1.ShowOpen
  strFilename = CommonDialog1.FileName
  If Right(strFilename, 3) = "mp3" Then
  MMControl1.DeviceType = "MPEGVideo"
  ElseIf Right(strFilename, 3) = "wav" Then
  MMControl1.DeviceType = "WaveAudio"
  ElseIf Right(strFilename, 3) = "wav" Or Right(strFilename, 4) = "midi" Then
  MMControl1.DeviceType = "Sequencer"
  ElseIf Right(strFilename, 3) = "CDA" Then
  MMControl1.DeviceType = "CDAudio"
  End If
  MMControl1.FileName = CommonDialog1.FileName
  MMControl1.Command = "Open"
End If


End Sub

Private Sub chkSetNote_Click()
 If chkSetNote.Value = 1 Then
  chkOpenMusic.Visible = True
  chkTime.Visible = True
  chkNote.Visible = True
  chkMusic.Enabled = False
 Else
  chkOpenMusic.Visible = False
  chkTime.Visible = False
  chkNote.Visible = False
  chkMusic.Enabled = True
  txtTime.Visible = False
  optAm.Visible = False
  optPm.Visible = False
  txtNote.Visible = False
  chkNote.Value = 0
  chkTime.Value = 0
 End If
End Sub

Private Sub chkTime_Click()
 If chkTime.Value = 1 Then
   txtTime.Visible = True
   optAm.Visible = True
   optPm.Visible = True
 ElseIf chkTime.Value = 0 Then
   txtTime.Visible = False
   optAm.Visible = False
   optPm.Visible = False
 End If
End Sub




Private Sub Form_Unload(Cancel As Integer)
MMControl1.Command = "Close"
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuStart_Click()

 If CommonDialog1.FileName = "" Or txtTime.Text = "" Or txtNote.Text = "" Then
  MsgBox ("You're missing some type of information for your note"), , "Missing information"
  frmNoteAlarm.WindowState = 0
 Else
  frmNoteAlarm.WindowState = 1
  tmrAlarm.Enabled = True
  chkSetNote.Value = 0
  chkTime.Value = 0
  chkNote.Value = 0
  chkMusic.Value = 0
  chkControls.Value = 0
  txtNote.Visible = False
End If

End Sub

Private Sub mnuTimeColor_Click()
 CommonDialog1.Flags = &H4&
 CommonDialog1.ShowColor
 lblTime.ForeColor = CommonDialog1.Color
End Sub

Private Sub tmrAlarm_Timer()
Dim strTime As String
strTime = txtTime.Text
CurrentTime = Format(Time, "hh:mm")

 If optAm.Value = True Then
  strTime = strTime
 ElseIf optPm.Value = True Then
  strTime = Left(strTime, 1) + 12 & Right(strTime, 3)
 End If

 If CurrentTime = strTime Then
  tmrAlarm.Interval = 1
  MMControl1.Command = "Play"
  MsgBox (txtNote.Text), , "Your Note"
  tmrAlarm.Enabled = False
   If vbOK Then
   frmNoteAlarm.WindowState = 0
   End If
 End If
  
  
End Sub

Private Sub tmrTime_Timer()
lblTime.Caption = Time
End Sub
