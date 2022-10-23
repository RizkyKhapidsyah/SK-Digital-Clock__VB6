VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   4245
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   2040
      Top             =   2880
   End
   Begin VB.PictureBox imgMain 
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Form_Load()

    Dim WindowRegion As Long

    imgMain.ScaleMode = vbPixels
    imgMain.AutoRedraw = True
    imgMain.AutoSize = True
    imgMain.BorderStyle = vbBSNone
    Me.BorderStyle = vbBSNone
        
    Set imgMain.Picture = LoadPicture(App.Path & "\Digital Clock\Splash.dglc")
    
    Me.Width = imgMain.Width
    Me.Height = imgMain.Height
    
    WindowRegion = MakeRegion(imgMain)
    SetWindowRgn Me.hWnd, WindowRegion, True

End Sub

Private Sub imgMain_Click()

 Unload Me
 frmMain.Show

End Sub

Private Sub Timer1_Timer()

 Unload Me
 frmMain.Show

End Sub
