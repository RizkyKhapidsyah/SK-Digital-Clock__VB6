Attribute VB_Name = "SetTimer"
'## ## ## ## ## ## ## ## ## ## ## ## ## ## ##
Public Sub Clear()

 Dim Counter As Integer
 
  For Counter = 1 To 6
    With frmMain
     
     frmMain.Segment0(Counter).Visible = False
     frmMain.Segment1(Counter).Visible = False
     frmMain.Segment2(Counter).Visible = False
     frmMain.Segment3(Counter).Visible = False
     frmMain.Segment4(Counter).Visible = False
     frmMain.Segment5(Counter).Visible = False
     frmMain.Segment6(Counter).Visible = False
    
    End With
  Next
 
End Sub
'## ## ## ## ## ## ## ## ## ## ## ## ## ## ##
Public Sub SetPos()

 Dim Counter1 As Integer
 
   Clear '{Off All Segment}
 
  With frmMain
 
  For Counter1 = 1 To 6
    .Seven(Counter1).Top = 100
  Next
  
  For Counter1 = 1 To 6
  
    .Segment0(Counter1).Left = .Seven(Counter1).Left + 45 / 2
    .Segment0(Counter1).Top = .Seven(Counter1).Top + 285 / 2
 
    .Segment1(Counter1).Left = .Seven(Counter1).Left + 27 / 2
    .Segment1(Counter1).Top = .Seven(Counter1).Top + 0 / 2
    
    .Segment2(Counter1).Left = .Seven(Counter1).Left + 285 / 2
    .Segment2(Counter1).Top = .Seven(Counter1).Top + 30 / 2
  
    .Segment3(Counter1).Left = .Seven(Counter1).Left + 285 / 2
    .Segment3(Counter1).Top = .Seven(Counter1).Top + 330 / 2

    .Segment4(Counter1).Left = .Seven(Counter1).Left + 32.5 / 2
    .Segment4(Counter1).Top = .Seven(Counter1).Top + 585 / 2
        
    .Segment5(Counter1).Left = .Seven(Counter1).Left + 0 / 2
    .Segment5(Counter1).Top = .Seven(Counter1).Top + 330 / 2
    
    .Segment6(Counter1).Left = .Seven(Counter1).Left + 0 / 2
    .Segment6(Counter1).Top = .Seven(Counter1).Top + 30 / 2
    
  Next

  End With

End Sub
'## ## ## ## ## ## ## ## ## ## ## ## ## ## ##
Public Sub SetSegmentPic()

 Dim Counter2 As Integer

  With frmMain

  For Counter2 = 1 To 6
  
    .Segment0(Counter2).Picture = LoadPicture(App.Path & "\Digital Clock\Segment\7.ico")
    .Segment1(Counter2).Picture = LoadPicture(App.Path & "\Digital Clock\Segment\1.ico")
    .Segment2(Counter2).Picture = LoadPicture(App.Path & "\Digital Clock\Segment\2.ico")
    .Segment3(Counter2).Picture = LoadPicture(App.Path & "\Digital Clock\Segment\2.ico")
    .Segment4(Counter2).Picture = LoadPicture(App.Path & "\Digital Clock\Segment\4.ico")
    .Segment5(Counter2).Picture = LoadPicture(App.Path & "\Digital Clock\Segment\5.ico")
    .Segment6(Counter2).Picture = LoadPicture(App.Path & "\Digital Clock\Segment\5.ico")
    
    .Seven(Counter2).Picture = LoadPicture(App.Path & "\Digital Clock\Segment\Seven Segment.ico")
    
  Next


  End With

End Sub
