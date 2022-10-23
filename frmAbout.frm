VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   3495
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":0000
   ScaleHeight     =   1200
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblMe 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "by Hadi Samadzad"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   840
      Width           =   2145
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Digital Clock"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   540
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2700
   End
   Begin VB.Image BG 
      Height          =   3075
      Left            =   0
      Picture         =   "frmAbout.frx":0557
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4665
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BG_Click()

  Unload Me

End Sub

Private Sub Form_Load()

  lblName.Caption = "Digital Clock"
  lblMe.Caption = "by Hadi Samadzad"

End Sub
