Attribute VB_Name = "Module1"
Public Sub oo()
 On Error GoTo Err
 With frmMain
   If .cboDo.Text = "Play Music ..." Then
     .D1.ShowOpen
     strs(1) = .D1.FileName
     'wmp.URL = D1.FileName
   ElseIf .cboDo.Text = "Run Program ..." Then
     .D2.ShowOpen
     strs(2) = .D2.FileName
   End If
Err:

End Sub
