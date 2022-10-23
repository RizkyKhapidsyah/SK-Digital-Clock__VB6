VERSION 5.00
Begin VB.Form LanBar 
   Caption         =   "Digital Clock"
   ClientHeight    =   2790
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   2325
   LinkTopic       =   "Form1"
   ScaleHeight     =   2790
   ScaleWidth      =   2325
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picIcon 
      Height          =   855
      Left            =   0
      Picture         =   "LanBar.frx":0000
      ScaleHeight     =   795
      ScaleWidth      =   1275
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   960
      Top             =   1680
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuAlOff 
         Caption         =   "Alarm Off"
      End
      Begin VB.Menu Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSound 
         Caption         =   "Sound Setting ..."
      End
      Begin VB.Menu Sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMin 
         Caption         =   "Mini Size"
      End
      Begin VB.Menu mnuShow 
         Caption         =   "Show Clock"
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About ..."
      End
      Begin VB.Menu Sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit"
      End
   End
End
Attribute VB_Name = "LanBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Bar
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As Bar) As Boolean
Dim ThisForm As Bar

Private Declare Function mciSendString Lib "winmm.dll" Alias _
"mciSendStringA" (ByVal lpstrCommand As String, ByVal _
lpstrReturnString As String, ByVal uReturnLength As Long, _
ByVal hwndCallback As Long) As Long

Private Sub Form_Load()
    
    ThisForm.cbSize = Len(ThisForm)
    ThisForm.hWnd = picIcon.hWnd
    ThisForm.uId = 1&
    ThisForm.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    ThisForm.uCallbackMessage = WM_MOUSEMOVE
    ThisForm.hIcon = picIcon.Picture
    ThisForm.szTip = "Digital Clock" & Chr$(0)
    Shell_NotifyIcon NIM_ADD, ThisForm
    Me.Hide
    App.TaskVisible = False
    
End Sub

Private Sub mnuAbout_Click()

  frmAbout.Show

End Sub

Private Sub mnuQuit_Click()
  
  Unload frmMain
  End

End Sub

Private Sub mnuShow_Click()

 If LanBar.mnuShow.Checked = False Then
   frmMain.Visible = True
 Else
   frmMain.Visible = False
 End If
 
End Sub

Private Sub mnuSound_Click()
 Shell "sndvol32", vbNormalFocus
End Sub

Private Sub picIcon_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  If Button = 2 Then
   
     PopupMenu Me.mnuFile
   
  End If

End Sub

Private Sub Timer1_Timer()

   mnuShow.Checked = frmMain.Visible
   mnuMin.Enabled = mnuShow.Checked
 
End Sub
Private Sub mnuAlOff_Click()

  On Error Resume Next
  
  frmMain.imgZingoff.Picture = LoadPicture(App.Path & "\btn\button13.jpg")
  frmMain.tmrZing.Interval = 0
  frmMain.wmp.URL = ""
  
  frmMain.tmrshpAlarm.Interval = 0

End Sub

Private Sub mnuMin_Click()

  If mnuMin.Checked = False Then
    
    frmMain.Height = 675
    frmMain.Width = 1995
    frmMain.picBoxSS.Left = 2
    frmMain.picBoxSS.Top = 2
    frmMain.TitleBar.Visible = False
    frmMain.imgIcon.Visible = False
    frmMain.lbl(3).Visible = False
    frmMain.lbl(0).Visible = False
    mnuMin.Checked = True
    frmMain.tmrOnTop.Interval = 0
  
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
    frmMain.tmrOnTop.Interval = 0
  
  End If

 frmMain.tmrOnTop.Interval = 0

End Sub
