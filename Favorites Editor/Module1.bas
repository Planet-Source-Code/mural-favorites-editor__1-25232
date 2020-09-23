Attribute VB_Name = "Module1"
Sub Activatee()
On Error Resume Next
Dim Title, Des, URL As String
Form1.File1.Path = App.Path
Form1.File1.Refresh
Form1.ListView1.ListItems.Clear
For i = 0 To Form1.File1.ListCount - 1
Open App.Path + "\" + Form1.File1.List(i) For Input As #1
Input #1, Title
Input #1, Des
Input #1, URL
Set li = Form1.ListView1.ListItems.add(, , Title)
li.SubItems(1) = URL
Close #1
Next
End Sub
