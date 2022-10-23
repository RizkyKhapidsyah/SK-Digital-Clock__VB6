VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   " DigitalClock"
   ClientHeight    =   2880
   ClientLeft      =   3495
   ClientTop       =   3390
   ClientWidth     =   9330
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   192
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   622
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrOnTop 
      Left            =   5160
      Top             =   1680
   End
   Begin VB.Timer tmrZingSet 
      Interval        =   1000
      Left            =   4200
      Top             =   1080
   End
   Begin VB.Timer tmrshpAlarm 
      Left            =   4680
      Top             =   1080
   End
   Begin VB.Timer tmrQuit 
      Left            =   5160
      Top             =   1080
   End
   Begin MSComDlg.CommonDialog D1 
      Left            =   3120
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Select Music"
      Filter          =   "All Supported Formats|*.wav;*.mid;*.wma;*.mp3|wav|*.wav|mp3|*.mp3|wma|*.wma|mid|*.mid"
   End
   Begin VB.Timer tmrZing 
      Left            =   6000
      Top             =   5400
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4320
      TabIndex        =   13
      Text            =   "00 : 00 : 00"
      Top             =   6240
      Width           =   1935
   End
   Begin VB.TextBox txtActiveTime 
      Height          =   285
      Left            =   120
      MaxLength       =   12
      TabIndex        =   0
      Text            =   "00 : 00 : 00"
      ToolTipText     =   "Set Activation Time with Tihs Formt HH : MM : SS"
      Top             =   2400
      Width           =   1935
   End
   Begin VB.PictureBox picBoxSS 
      BackColor       =   &H00000000&
      Height          =   615
      Left            =   120
      ScaleHeight     =   555
      ScaleWidth      =   1875
      TabIndex        =   10
      Top             =   720
      Width           =   1935
      Begin VB.Image Segment6 
         Height          =   135
         Index           =   6
         Left            =   4320
         Picture         =   "frmMain.frx":000C
         Stretch         =   -1  'True
         Top             =   480
         Width           =   30
      End
      Begin VB.Image Segment6 
         Height          =   135
         Index           =   5
         Left            =   3840
         Picture         =   "frmMain.frx":061E
         Stretch         =   -1  'True
         Top             =   480
         Width           =   30
      End
      Begin VB.Image Segment5 
         Height          =   135
         Index           =   6
         Left            =   4320
         Picture         =   "frmMain.frx":0C30
         Stretch         =   -1  'True
         Top             =   600
         Width           =   30
      End
      Begin VB.Image Segment5 
         Height          =   135
         Index           =   5
         Left            =   3840
         Picture         =   "frmMain.frx":1242
         Stretch         =   -1  'True
         Top             =   600
         Width           =   30
      End
      Begin VB.Image Segment0 
         Height          =   30
         Index           =   6
         Left            =   4320
         Picture         =   "frmMain.frx":1854
         Stretch         =   -1  'True
         Top             =   600
         Width           =   120
      End
      Begin VB.Image Segment0 
         Height          =   30
         Index           =   5
         Left            =   3840
         Picture         =   "frmMain.frx":1DFE
         Stretch         =   -1  'True
         Top             =   600
         Width           =   120
      End
      Begin VB.Image Segment3 
         Height          =   135
         Index           =   6
         Left            =   4440
         Picture         =   "frmMain.frx":23A8
         Stretch         =   -1  'True
         Top             =   600
         Width           =   30
      End
      Begin VB.Image Segment3 
         Height          =   135
         Index           =   5
         Left            =   3960
         Picture         =   "frmMain.frx":29BA
         Stretch         =   -1  'True
         Top             =   600
         Width           =   30
      End
      Begin VB.Image Segment2 
         Height          =   135
         Index           =   6
         Left            =   4440
         Picture         =   "frmMain.frx":2FCC
         Stretch         =   -1  'True
         Top             =   480
         Width           =   30
      End
      Begin VB.Image Segment2 
         Height          =   135
         Index           =   5
         Left            =   3960
         Picture         =   "frmMain.frx":35DE
         Stretch         =   -1  'True
         Top             =   480
         Width           =   30
      End
      Begin VB.Image Segment1 
         Height          =   30
         Index           =   6
         Left            =   4320
         Picture         =   "frmMain.frx":3BF0
         Stretch         =   -1  'True
         Top             =   480
         Width           =   135
      End
      Begin VB.Image Segment1 
         Height          =   30
         Index           =   5
         Left            =   3840
         Picture         =   "frmMain.frx":41BA
         Stretch         =   -1  'True
         Top             =   480
         Width           =   135
      End
      Begin VB.Image Segment4 
         Height          =   30
         Index           =   6
         Left            =   4320
         Picture         =   "frmMain.frx":4784
         Stretch         =   -1  'True
         Top             =   720
         Width           =   135
      End
      Begin VB.Image Segment4 
         Height          =   30
         Index           =   5
         Left            =   3840
         Picture         =   "frmMain.frx":4D4E
         Stretch         =   -1  'True
         Top             =   720
         Width           =   135
      End
      Begin VB.Image Segment6 
         Height          =   135
         Index           =   4
         Left            =   3360
         Picture         =   "frmMain.frx":5318
         Stretch         =   -1  'True
         Top             =   480
         Width           =   30
      End
      Begin VB.Image Segment6 
         Height          =   135
         Index           =   3
         Left            =   2880
         Picture         =   "frmMain.frx":592A
         Stretch         =   -1  'True
         Top             =   480
         Width           =   30
      End
      Begin VB.Image Segment6 
         Height          =   135
         Index           =   2
         Left            =   2400
         Picture         =   "frmMain.frx":5F3C
         Stretch         =   -1  'True
         Top             =   480
         Width           =   30
      End
      Begin VB.Image Segment6 
         Height          =   135
         Index           =   1
         Left            =   1920
         Picture         =   "frmMain.frx":654E
         Stretch         =   -1  'True
         Top             =   480
         Width           =   30
      End
      Begin VB.Image Segment5 
         Height          =   135
         Index           =   4
         Left            =   3360
         Picture         =   "frmMain.frx":6B60
         Stretch         =   -1  'True
         Top             =   600
         Width           =   30
      End
      Begin VB.Image Segment5 
         Height          =   135
         Index           =   3
         Left            =   2880
         Picture         =   "frmMain.frx":7172
         Stretch         =   -1  'True
         Top             =   600
         Width           =   30
      End
      Begin VB.Image Segment5 
         Height          =   135
         Index           =   2
         Left            =   2400
         Picture         =   "frmMain.frx":7784
         Stretch         =   -1  'True
         Top             =   600
         Width           =   30
      End
      Begin VB.Image Segment5 
         Height          =   135
         Index           =   1
         Left            =   1920
         Picture         =   "frmMain.frx":7D96
         Stretch         =   -1  'True
         Top             =   600
         Width           =   30
      End
      Begin VB.Image Segment4 
         Height          =   30
         Index           =   3
         Left            =   2880
         Picture         =   "frmMain.frx":83A8
         Stretch         =   -1  'True
         Top             =   720
         Width           =   135
      End
      Begin VB.Image Segment4 
         Height          =   30
         Index           =   2
         Left            =   2400
         Picture         =   "frmMain.frx":8972
         Stretch         =   -1  'True
         Top             =   720
         Width           =   135
      End
      Begin VB.Image Segment4 
         Height          =   30
         Index           =   1
         Left            =   1920
         Picture         =   "frmMain.frx":8F3C
         Stretch         =   -1  'True
         Top             =   720
         Width           =   135
      End
      Begin VB.Image Segment0 
         Height          =   30
         Index           =   4
         Left            =   3360
         Picture         =   "frmMain.frx":9506
         Stretch         =   -1  'True
         Top             =   600
         Width           =   120
      End
      Begin VB.Image Segment0 
         Height          =   30
         Index           =   3
         Left            =   2880
         Picture         =   "frmMain.frx":9AB0
         Stretch         =   -1  'True
         Top             =   600
         Width           =   120
      End
      Begin VB.Image Segment0 
         Height          =   30
         Index           =   2
         Left            =   2400
         Picture         =   "frmMain.frx":A05A
         Stretch         =   -1  'True
         Top             =   600
         Width           =   120
      End
      Begin VB.Image Segment0 
         Height          =   30
         Index           =   1
         Left            =   1920
         Picture         =   "frmMain.frx":A604
         Stretch         =   -1  'True
         Top             =   600
         Width           =   120
      End
      Begin VB.Image Segment3 
         Height          =   135
         Index           =   4
         Left            =   3480
         Picture         =   "frmMain.frx":ABAE
         Stretch         =   -1  'True
         Top             =   600
         Width           =   30
      End
      Begin VB.Image Segment3 
         Height          =   135
         Index           =   2
         Left            =   2520
         Picture         =   "frmMain.frx":B1C0
         Stretch         =   -1  'True
         Top             =   600
         Width           =   30
      End
      Begin VB.Image Segment3 
         Height          =   135
         Index           =   1
         Left            =   2040
         Picture         =   "frmMain.frx":B7D2
         Stretch         =   -1  'True
         Top             =   600
         Width           =   30
      End
      Begin VB.Image Segment2 
         Height          =   135
         Index           =   4
         Left            =   3480
         Picture         =   "frmMain.frx":BDE4
         Stretch         =   -1  'True
         Top             =   480
         Width           =   30
      End
      Begin VB.Image Segment2 
         Height          =   135
         Index           =   3
         Left            =   3000
         Picture         =   "frmMain.frx":C3F6
         Stretch         =   -1  'True
         Top             =   480
         Width           =   30
      End
      Begin VB.Image Segment2 
         Height          =   135
         Index           =   1
         Left            =   2040
         Picture         =   "frmMain.frx":CA08
         Stretch         =   -1  'True
         Top             =   480
         Width           =   30
      End
      Begin VB.Image Segment1 
         Height          =   30
         Index           =   4
         Left            =   3360
         Picture         =   "frmMain.frx":D01A
         Stretch         =   -1  'True
         Top             =   480
         Width           =   135
      End
      Begin VB.Image Segment1 
         Height          =   30
         Index           =   3
         Left            =   2880
         Picture         =   "frmMain.frx":D5E4
         Stretch         =   -1  'True
         Top             =   480
         Width           =   135
      End
      Begin VB.Image Segment1 
         Height          =   30
         Index           =   2
         Left            =   2400
         Picture         =   "frmMain.frx":DBAE
         Stretch         =   -1  'True
         Top             =   480
         Width           =   135
      End
      Begin VB.Image Segment4 
         Height          =   30
         Index           =   4
         Left            =   3360
         Picture         =   "frmMain.frx":E178
         Stretch         =   -1  'True
         Top             =   720
         Width           =   135
      End
      Begin VB.Image Segment3 
         Height          =   135
         Index           =   3
         Left            =   3000
         Picture         =   "frmMain.frx":E742
         Stretch         =   -1  'True
         Top             =   600
         Width           =   30
      End
      Begin VB.Image Segment2 
         Height          =   135
         Index           =   2
         Left            =   2520
         Picture         =   "frmMain.frx":ED54
         Stretch         =   -1  'True
         Top             =   480
         Width           =   30
      End
      Begin VB.Image Segment1 
         Height          =   30
         Index           =   1
         Left            =   1920
         Picture         =   "frmMain.frx":F366
         Stretch         =   -1  'True
         Top             =   480
         Width           =   135
      End
      Begin VB.Image Seven 
         Height          =   315
         Index           =   6
         Left            =   1560
         Picture         =   "frmMain.frx":F930
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Image Seven 
         Height          =   315
         Index           =   5
         Left            =   1320
         Picture         =   "frmMain.frx":11082
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Image Seven 
         Height          =   315
         Index           =   4
         Left            =   960
         Picture         =   "frmMain.frx":127D4
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Image Seven 
         Height          =   315
         Index           =   3
         Left            =   720
         Picture         =   "frmMain.frx":13F26
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Image Seven 
         Height          =   315
         Index           =   2
         Left            =   360
         Picture         =   "frmMain.frx":15678
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Image Seven 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   120
         Picture         =   "frmMain.frx":16DCA
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   180
      End
   End
   Begin VB.ComboBox cboDo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmMain.frx":1851C
      Left            =   120
      List            =   "frmMain.frx":18529
      TabIndex        =   9
      Text            =   "Work ?"
      ToolTipText     =   "Select With Up and Down Keys ."
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox txtNum 
      Height          =   285
      Index           =   6
      Left            =   5400
      MaxLength       =   1
      TabIndex        =   8
      Text            =   "0"
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtNum 
      Height          =   285
      Index           =   5
      Left            =   5160
      MaxLength       =   1
      TabIndex        =   7
      Text            =   "0"
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtNum 
      Height          =   285
      Index           =   4
      Left            =   4920
      MaxLength       =   1
      TabIndex        =   6
      Text            =   "0"
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtNum 
      Height          =   285
      Index           =   3
      Left            =   4680
      MaxLength       =   1
      TabIndex        =   5
      Text            =   "0"
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtNum 
      Height          =   285
      Index           =   2
      Left            =   4440
      MaxLength       =   1
      TabIndex        =   4
      Text            =   "0"
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtNum 
      Height          =   285
      Index           =   1
      Left            =   4200
      MaxLength       =   1
      TabIndex        =   3
      Text            =   "0"
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtSevenSegment 
      Height          =   285
      Index           =   2
      Left            =   4440
      MaxLength       =   1
      TabIndex        =   2
      Text            =   "2"
      Top             =   4440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtSevenSegment 
      Height          =   285
      Index           =   1
      Left            =   4200
      MaxLength       =   1
      TabIndex        =   1
      Text            =   "2"
      Top             =   4440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   5520
      Top             =   5400
   End
   Begin MSComDlg.CommonDialog D2 
      Left            =   3600
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Select Program"
      Filter          =   "Execute Files (*.exe)|*.exe"
   End
   Begin VB.Line Line1 
      X1              =   8
      X2              =   140
      Y1              =   184
      Y2              =   184
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000F&
      BorderWidth     =   2
      X1              =   8
      X2              =   140
      Y1              =   184
      Y2              =   184
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "^ With Up & Down Keys ."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Index           =   5
      Left            =   7320
      TabIndex        =   17
      Top             =   4320
      Width           =   1845
   End
   Begin VB.Shape shpAlarm 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H000000FF&
      BorderStyle     =   0  'Transparent
      Height          =   255
      Left            =   6120
      Shape           =   5  'Rounded Square
      Top             =   840
      Width           =   375
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Digital Clock"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Index           =   3
      Left            =   1080
      TabIndex        =   16
      Top             =   120
      Width           =   1050
   End
   Begin VB.Image imgIcon 
      Height          =   240
      Left            =   120
      Picture         =   "frmMain.frx":18554
      Stretch         =   -1  'True
      Top             =   120
      Width           =   255
   End
   Begin VB.Image TitleBar 
      Height          =   375
      Left            =   0
      Picture         =   "frmMain.frx":1895E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7935
   End
   Begin VB.Image imgZingoff 
      Height          =   300
      Left            =   1440
      Picture         =   "frmMain.frx":1E09F
      Stretch         =   -1  'True
      Top             =   6120
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image imgQuit 
      Height          =   300
      Left            =   1440
      Picture         =   "frmMain.frx":1EC06
      Stretch         =   -1  'True
      Top             =   6480
      Visible         =   0   'False
      Width           =   1500
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp 
      Height          =   2475
      Left            =   4320
      TabIndex        =   15
      Top             =   6120
      Width           =   2400
      URL             =   ""
      rate            =   1
      balance         =   1
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   999
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   100
      mute            =   0   'False
      uiMode          =   "mini"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   4233
      _cy             =   4366
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Activation Time :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Index           =   2
      Left            =   120
      TabIndex        =   14
      Top             =   2160
      Width           =   1395
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Activation Type :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   1380
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Width           =   510
   End
   Begin VB.Image BG 
      Height          =   4590
      Left            =   0
      Picture         =   "frmMain.frx":1F5AF
      Top             =   0
      Width           =   3765
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim strs(2) As String
Private Sub BG_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  On Error Resume Next

  If Len(txtActiveTime.Text) <> 12 Then
    
    MsgBox "Invalid Activation Time. Use This Format HH : MM : SS", vbCritical, "Error"
    txtActiveTime.SetFocus
  
  End If

  imgQuit.Picture = LoadPicture(App.Path & "\Digital Clock\btn\button1.jpg")
  imgZingoff.Picture = LoadPicture(App.Path & "\Digital Clock\btn\button11.jpg")
  
   Dim RV As Long
   
   If Button = 1 Then
      
      Call ReleaseCapture
      RV = SendMessage(frmMain.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
      Call SetPosinCorner
   
   End If

End Sub

Private Sub cboDo_Click()

 On Error GoTo err

   If cboDo.Text = "Play Music ..." Then
     D1.ShowOpen
     strs(1) = D1.FileName
   ElseIf cboDo.Text = "Run Program ..." Then
     D2.ShowOpen
     strs(2) = D2.FileName
   End If
err:

End Sub

Private Sub Form_Load()

  On Error Resume Next

  'lbl(4).Caption = "Hadi Samadzad"

  tmrOnTop.Interval = 0

  LanBar.mnuMin.Checked = GetSetting(App.EXEName, "mnu", "Mini", 1)
  
  
  If mnuMin.Checked = False Then
        frmMain.tmrOnTop.Interval = 10
  Else
        frmMain.tmrOnTop.Interval = 0
  End If
  
  If LanBar.mnuMin.Checked = True Then
    
    frmMain.Height = 675
    frmMain.Width = 1995
    frmMain.picBoxSS.Left = 2
    frmMain.picBoxSS.Top = 2
    frmMain.TitleBar.Visible = False
    frmMain.imgIcon.Visible = False
    frmMain.lbl(3).Visible = False
    frmMain.lbl(0).Visible = False
    mnuMin.Checked = True
    tmrOnTop.Interval = 10
  
  Else
      
    frmMain.Height = 2880
    frmMain.Width = 2205
    frmMain.picBoxSS.Left = 8
    frmMain.picBoxSS.Top = 48
    frmMain.TitleBar.Visible = True
    frmMain.imgIcon.Visible = True
    frmMain.lbl(3).Visible = True
    frmMain.lbl(0).Visible = True
    mnuMin.Checked = False
 
 End If
  
  tmrOnTop.Interval = 0
  
  i = 0
  
  SetPos
  SetSegmentPic

  LanBar.Show
  LanBar.Hide
  
End Sub

Private Sub Form_Unload(Cancel As Integer)

  SaveSetting App.EXEName, "mnu", "Mini", LanBar.mnuMin.Checked

End Sub

Private Sub imgIcon_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   Dim RV As Long
   
   If Button = 1 Then
      
      Call GoodForm.ReleaseCapture
      RV = GoodForm.SendMessage(frmMain.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
      Call SetPosinCorner
   
   End If

End Sub

Private Sub imgQuit_Click()

 On Error Resume Next
   
 imgQuit.Picture = LoadPicture(App.Path & "\Digital Clock\btn\button3.jpg")

  Me.Visible = False

End Sub

Private Sub imgQuit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

 On Error Resume Next

   imgQuit.Picture = LoadPicture(App.Path & "\Digital Clock\btn\button2.jpg")

End Sub

Private Sub imgZingoff_Click()

  On Error Resume Next
  
  imgZingoff.Picture = LoadPicture(App.Path & "\Digital Clock\btn\button13.jpg")
  tmrZing.Interval = 0
  wmp.URL = ""

End Sub

Private Sub imgZingoff_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
 On Error Resume Next

  imgZingoff.Picture = LoadPicture(App.Path & "\Digital Clock\btn\button12.jpg")

End Sub

Private Sub lbl_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

   Dim RV As Long
   
   If Button = 1 Then
      
      Call ReleaseCapture
      RV = SendMessage(frmMain.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
      Call SetPosinCorner
   
   End If
 
 End Sub
 
Private Sub picBoxSS_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
 Dim RV As Long
   
   If Button = 1 Then
      
      Call ReleaseCapture
      RV = SendMessage(frmMain.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
      Call SetPosinCorner
      
   End If

End Sub

Private Sub Timer1_Timer()

  Call TimeCommand(strs(1), strs(2))

End Sub

Private Sub TitleBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
   Dim RV As Long
   
   If Button = 1 Then
      
      Call ReleaseCapture
      RV = SendMessage(frmMain.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
      Call SetPosinCorner
   
   End If

End Sub

Private Sub tmrOnTop_Timer()
  
  frmx = Screen.Width / Screen.TwipsPerPixelX - Me.ScaleWidth - 10
  frmy = Screen.Height / Screen.TwipsPerPixelY - Me.ScaleHeight - 35

  tmpval = SetWindowPos(frmMain.hWnd, HWND_TOPMOST, frmx, frmy, Me.ScaleWidth, Me.ScaleHeight, SWP_SHOWME)   ' show form1

End Sub

Private Sub tmrshpAlarm_Timer()

 If frmMain.tmrZing.Interval <> 0 Then
   
   shpAlarm.BackColor = Rnd * 1000000

   If shpAlarm.Visible = True Then
     shpAlarm.Visible = False
   Else
     shpAlarm.Visible = True
   End If
 
 End If

End Sub

Private Sub tmrZing_Timer()
  
  i = i + 1

  Beep
  
  If i = 24 Then
   tmrZing.Interval = 250
  ElseIf i = 60 Then
   tmrZing.Interval = 230
  ElseIf i = 120 Then
   tmrZing.Interval = 200
  ElseIf i = 120 Then
   tmrZing.Interval = 160
  ElseIf i = 180 Then
   tmrZing.Interval = 130
  ElseIf i = 260 Then
   tmrZing.Interval = 90
  ElseIf i = 360 Then
   tmrZing.Interval = 60
  ElseIf i = 460 Then
   tmrZing.Interval = 45
  ElseIf i = 560 Then
   tmrZing.Interval = 30
  ElseIf i = 660 Then
   tmrZing.Interval = 20
  
  End If
  
End Sub

Private Sub tmrZingSet_Timer()

  If Me.txtActiveTime.Text <> "00 : 00 : 00" And Len(Me.txtActiveTime.Text) = 12 Then
     shpAlarm.Visible = True
  Else
     shpAlarm.Visible = False
  End If
  
End Sub
