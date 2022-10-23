Attribute VB_Name = "Show"
Public Sub ShowCommand()

 Dim Number As Double
 Dim SevenSegments As Integer
 Dim i As Integer

 With frmMain

 For i = 1 To 6

   Number = .txtNum(i).Text
   .txtSevenSegment(1).Text = i

   If Number < 10 And Number >= 0 Then
   
     ShowNum (Number)
   
   End If
 
 Next

 End With

End Sub

Public Sub ShowNum(Numbers As Double)
  
  Dim SevenSegment As Integer
  
   SevenSegment = frmMain.txtSevenSegment(1).Text
  
  Select Case Numbers
      Case 0: Zero (SevenSegment)
      Case 1: One (SevenSegment)
      Case 2: Two (SevenSegment)
      Case 3: Three (SevenSegment)
      Case 4: Four (SevenSegment)
      Case 5: Five (SevenSegment)
      Case 6: Six (SevenSegment)
      Case 7: Seven (SevenSegment)
      Case 8: Eight (SevenSegment)
      Case 9: Nine (SevenSegment)
  End Select
 
End Sub

Public Sub Zero(SevenSegment As Integer)

  With frmMain
    .Segment0(SevenSegment).Visible = False
    .Segment1(SevenSegment).Visible = True
    .Segment2(SevenSegment).Visible = True
    .Segment3(SevenSegment).Visible = True
    .Segment4(SevenSegment).Visible = True
    .Segment5(SevenSegment).Visible = True
    .Segment6(SevenSegment).Visible = True
  End With

End Sub

Public Sub One(SevenSegment As Integer)

  With frmMain
    .Segment0(SevenSegment).Visible = False
    .Segment1(SevenSegment).Visible = False
    .Segment2(SevenSegment).Visible = True
    .Segment3(SevenSegment).Visible = True
    .Segment4(SevenSegment).Visible = False
    .Segment5(SevenSegment).Visible = False
    .Segment6(SevenSegment).Visible = False
  End With

End Sub

Public Sub Two(SevenSegment As Integer)

  With frmMain
    .Segment0(SevenSegment).Visible = True
    .Segment1(SevenSegment).Visible = True
    .Segment2(SevenSegment).Visible = True
    .Segment3(SevenSegment).Visible = False
    .Segment4(SevenSegment).Visible = True
    .Segment5(SevenSegment).Visible = True
    .Segment6(SevenSegment).Visible = False
  End With

End Sub

Public Sub Three(SevenSegment As Integer)

  With frmMain
    .Segment0(SevenSegment).Visible = True
    .Segment1(SevenSegment).Visible = True
    .Segment2(SevenSegment).Visible = True
    .Segment3(SevenSegment).Visible = True
    .Segment4(SevenSegment).Visible = True
    .Segment5(SevenSegment).Visible = False
    .Segment6(SevenSegment).Visible = False
  End With

End Sub

Public Sub Four(SevenSegment As Integer)

  With frmMain
    .Segment0(SevenSegment).Visible = True
    .Segment1(SevenSegment).Visible = False
    .Segment2(SevenSegment).Visible = True
    .Segment3(SevenSegment).Visible = True
    .Segment4(SevenSegment).Visible = False
    .Segment5(SevenSegment).Visible = False
    .Segment6(SevenSegment).Visible = True
  End With

End Sub

Public Sub Five(SevenSegment As Integer)

  With frmMain
    .Segment0(SevenSegment).Visible = True
    .Segment1(SevenSegment).Visible = True
    .Segment2(SevenSegment).Visible = False
    .Segment3(SevenSegment).Visible = True
    .Segment4(SevenSegment).Visible = True
    .Segment5(SevenSegment).Visible = False
    .Segment6(SevenSegment).Visible = True
  End With

End Sub

Public Sub Six(SevenSegment As Integer)

  With frmMain
    .Segment0(SevenSegment).Visible = True
    .Segment1(SevenSegment).Visible = True
    .Segment2(SevenSegment).Visible = False
    .Segment3(SevenSegment).Visible = True
    .Segment4(SevenSegment).Visible = True
    .Segment5(SevenSegment).Visible = True
    .Segment6(SevenSegment).Visible = True
  End With

End Sub

Public Sub Seven(SevenSegment As Integer)

  With frmMain
    .Segment0(SevenSegment).Visible = False
    .Segment1(SevenSegment).Visible = True
    .Segment2(SevenSegment).Visible = True
    .Segment3(SevenSegment).Visible = True
    .Segment4(SevenSegment).Visible = False
    .Segment5(SevenSegment).Visible = False
    .Segment6(SevenSegment).Visible = False
  End With

End Sub

Public Sub Eight(SevenSegment As Integer)

  With frmMain
    .Segment0(SevenSegment).Visible = True
    .Segment1(SevenSegment).Visible = True
    .Segment2(SevenSegment).Visible = True
    .Segment3(SevenSegment).Visible = True
    .Segment4(SevenSegment).Visible = True
    .Segment5(SevenSegment).Visible = True
    .Segment6(SevenSegment).Visible = True
  End With

End Sub

Public Sub Nine(SevenSegment As Integer)

  With frmMain
    .Segment0(SevenSegment).Visible = True
    .Segment1(SevenSegment).Visible = True
    .Segment2(SevenSegment).Visible = True
    .Segment3(SevenSegment).Visible = True
    .Segment4(SevenSegment).Visible = True
    .Segment5(SevenSegment).Visible = False
    .Segment6(SevenSegment).Visible = True
  End With

End Sub
'#########################
'IT IS NOT FOR THIS MODULE

Public Sub SetPosinCorner()
 
  If frmMain.Top < 350 Then
    frmMain.Top = 0
  End If

  If frmMain.Left < 300 Then
    frmMain.Left = 0
  End If

End Sub

