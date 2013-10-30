Private Sub Worksheet_Change(ByVal Target As Range)
Dim Max As Integer
Dim RowCounter As Integer
Dim CategoryContents As String
Dim RootName As String
Dim RootNameRow As String
Dim CategoryRow As String
Dim WarningColor As Long
Dim ErrorColor As Long
Dim AdjustedCounter As Integer
Dim RootThreshold As Integer
Dim Counter As Integer
Dim StartRow As Integer
WarningColor = RGB(255, 255, 55)
ErrorColor = RGB(255, 41, 41)
CategoryRow = "C"
RootNameRow = "A"
Max = 5
Counter = 0
StartRow = 3
RootThreshold = 10
RowCounter = StartRow
Do While RowCounter < RootThreshold
    CategoryContents = Cells(RowCounter, CategoryRow).Text
    CategoryContents = Trim(CategoryContents)
    RootName = Cells(RowCounter, RootNameRow).Text
    RootName = Trim(RootName)
    If CategoryContents = "ROOT" Then
        Counter = (Counter + 1)
    ElseIf LCase(CategoryContents) = "root" Then
        Counter = (Counter + 1)
        Cells(RowCounter, CategoryRow).Value = "ROOT"
    End If
    If Len(RootName) > 0 And Len(CategoryContents) < 1 Then
        Cells(RowCounter, CategoryRow).Interior.Color = ErrorColor
    ElseIf Len(RootName) < 1 And Len(CategoryContents) > 0 Then
        Cells(RowCounter, RootNameRow).Interior.Color = ErrorColor
    ElseIf Len(RootName) > 11 And CategoryContents = "ROOT" Then
        Cells(RowCounter, RootNameRow).Interior.Color = ErrorColor
    Else
        Cells(RowCounter, RootNameRow).Interior.ColorIndex = 0
        Cells(RowCounter, CategoryRow).Interior.ColorIndex = 0
    End If
    RowCounter = (RowCounter + 1)
Loop
AdjustedCounter = Counter
RowCounter = StartRow
If Counter > Max Then
    Do While RowCounter < RootThreshold
        
        CategoryContents = Cells(RowCounter, CategoryRow).Value
        CategoryContents = Trim(CategoryContents)
        If Cells(RowCounter, CategoryRow).Interior.Color = ErrorColor Then
            RowCounter = RowCounter + 1
        Else
            If CategoryContents = "ROOT" Then
                Cells(RowCounter, CategoryRow).Interior.Color = WarningColor
                RowCounter = RowCounter + 1
            Else
                RowCounter = RowCounter + 1
            End If
        End If
    Loop
End If
End Sub




