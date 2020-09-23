Attribute VB_Name = "matrix"
Public Sub Main()
    If App.PrevInstance Then End
    If InStr(Command, "/s") > 0 Then
        form10.Show
    ElseIf InStr(Command, "/c") > 0 Then
        Form2.Show
        
    End If
End Sub




