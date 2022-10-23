Attribute VB_Name = "SetTime"
Public Sub TimeCommand(Com, Com2)

 With frmMain
 
  Dim intHour1, intHour2 As Integer
  Dim intMinute1, intMinute2 As Integer
  Dim intSecond1, intSecond2 As Integer

  intHour1 = Hour(Time) / 10
  .txtNum(1).Text = intHour1
  
  intHour2 = Hour(Time) Mod 10
  .txtNum(2).Text = intHour2
  
  intMinute1 = Minute(Time) / 10
  .txtNum(3).Text = intMinute1
  
  intMinute2 = Minute(Time) Mod 10
  .txtNum(4).Text = intMinute2
  
  intSecond1 = Second(Time) / 10
  .txtNum(5).Text = intSecond1
  
  intSecond2 = Second(Time) Mod 10
  .txtNum(6).Text = intSecond2
  
  .Text1.Text = .txtNum(1).Text & .txtNum(2).Text & " : " & .txtNum(3).Text & .txtNum(4).Text & " : " & .txtNum(5).Text & .txtNum(6).Text
  If .txtActiveTime.Text = .txtNum(1).Text & .txtNum(2).Text & " : " & .txtNum(3).Text & .txtNum(4).Text & " : " & .txtNum(5).Text & .txtNum(6).Text Then
      Call Active(Com, Com2)
  End If
  
  ShowCommand

 End With

End Sub

Function Active(Com3, Com4)

  Select Case frmMain.cboDo.Text
   
   Case "Default Zing":
     frmMain.tmrZing.Interval = 300
     Com3 = ""
     Com4 = ""
   Case "Play Music ...":
     frmMain.wmp.URL = Com3
     Com4 = ""
     frmMain.tmrZing.Interval = 0
   Case "Run Program ...":
     ii = Shell(Com4, vbNormalFocus)
     frmMain.wmp.URL = ""
     frmMain.tmrZing.Interval = 0
  End Select
 
  frmMain.shpAlarm.Visible = True
  frmMain.tmrshpAlarm.Interval = 500
 
End Function
