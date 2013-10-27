Private Sub Worksheet_Change(ByVal Target As Range)
Dim Max As Integer
Dim Counter As Integer
Dim RowValue As Integer
Dim CategoryContents As String
Dim RootName As String
Dim Roots(15) As Integer
Dim RootNameRow As String
Dim CategoryRow As String
Dim WarningColor As Long
Dim ErrorColor As Long
WarningColor = RGB(255, 255, 55)
ErrorColor = RGB(255, 41, 41)
CategoryRow = "C"
RootNameRow = "A"
Max = 5
Counter = 0
RowValue = 3
Do While RowValue < 10
    CategoryContents = Cells(RowValue, CategoryRow).Value
    If CategoryContents = "ROOT" Then
        Counter = (Counter + 1)
        Roots(Counter) = RowValue
    ElseIf LCase(CategoryContents) = "root" Then
        Counter = (Counter + 1)
        Cells(RowValue, CategoryRow).Value = "ROOT"
        Roots(Counter) = RowValue
    End If
    RootName = Cells(RowValue, RootNameRow).Text
    If Len(RootName) > 0 And Len(CategoryContents) < 1 Then
        Cells(RowValue, CategoryRow).Interior.Color = ErrorColor
    ElseIf Len(RootName) < 1 And Len(CategoryContents) > 0 Then
        Cells(RowValue, RootNameRow).Interior.Color = ErrorColor
    Else
         If Len(RootName) > 11 Then
            Cells(RowValue, RootNameRow).Interior.Color = ErrorColor
        Else
            Cells(RowValue, RootNameRow).Interior.ColorIndex = 0
            Cells(RowValue, CategoryRow).Interior.ColorIndex = 0
        End If
    End If
    RowValue = (RowValue + 1)
Loop
If Counter > Max Then
    
End If
End Sub



