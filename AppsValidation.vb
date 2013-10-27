Sub Button1_Click()
Dim Max As Integer
Dim Counter As Integer
Dim RowValue As Integer
Dim CategoryContents As String
Dim RootName As String
Max = 5
Counter = 0
RowValue = 3
Do While RowValue < 10
    CategoryContents = Cells(RowValue, "C").Value
    If CategoryContents = "ROOT" Then
        Counter = (Counter + 1)
    ElseIf LCase(CategoryContents) = "root" Then
        Counter = (Counter + 1)
        Cells(RowValue, "C").Value = "ROOT"
    End If
    RootName = Cells(RowValue, "A").Text
    If Len(RootName) > 11 Then
        Cells(RowValue, "A").Interior.Color = RGB(255, 41, 41)
    Else
        Cells(RowValue, "A").Interior.ColorIndex = 0
    End If
    
    If Len(RootName) > 0 And Len(CategoryContents) < 1 Then
        Cells(RowValue, "C").Interior.Color = RGB(255, 41, 41)
    ElseIf Len(RootName) < 1 And Len(CategoryContents) > 0 Then
        Cells(RowValue, "A").Interior.Color = RGB(255, 41, 41)
    Else
    Cells(RowValue, "A").Interior.ColorIndex = 0
    Cells(RowValue, "C").Interior.ColorIndex = 0
    End If
    RowValue = (RowValue + 1)
    If Counter > Max Then
        MsgBox ("Too Many Roots")
        Exit Do
    End If
Loop
End Sub
