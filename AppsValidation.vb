Public Sub Worksheet_Change(ByVal Target As Range)
Dim CategoryContent As String
Dim DescriptionContent As String
Dim CategoryName As String
Dim NameColumn As String
Dim CategoryColumn As String
Dim DescriptionColumn As String
Dim ROOTS() As String
Dim Max As Integer
Dim RowCounter As Integer
Dim AdjustedRootCounter As Integer
Dim RootThreshold As Integer
Dim RootCounter As Integer
Dim StartRow As Integer
Dim ErrorCounter As Integer
Dim LastRootPosition As Integer
Dim LastItemThreshold As Integer
Dim WarningColor As Long
Dim ErrorColor As Long
WarningColor = RGB(255, 255, 55)
ErrorColor = RGB(255, 41, 41)
CategoryColumn = "C"
NameColumn = "A"
DescriptionColumn = "B"
Application.EnableEvents = False

Max = 5
RootCounter = 0
StartRow = 3
RootThreshold = 15
LastItemThreshold = 50
RowCounter = StartRow
ErrorCounter = 0
Do While RowCounter < RootThreshold
    CategoryContent = Cells(RowCounter, CategoryColumn).Text
    CategoryContent = Trim(CategoryContent)
    CategoryName = Cells(RowCounter, NameColumn).Text
    CategoryName = Trim(CategoryName)
    DescriptionContent = Cells(RowCounter, DescriptionColumn).Text
    DescriptionContent = Trim(DescriptionContent)
    If CategoryContent = "ROOT" Then
        RootCounter = (RootCounter + 1)
        ReDim Preserve ROOTS(RowCounter - StartRow)
        ROOTS(RowCounter - StartRow) = CategoryName
        LastRootPosition = RowCounter
    ElseIf LCase(CategoryContent) = "root" Then
        RootCounter = (RootCounter + 1)
        Cells(RowCounter, CategoryColumn).Value = "ROOT"
        ReDim Preserve ROOTS(RowCounter - StartRow)
        ROOTS(RowCounter - StartRow) = CategoryName
        LastRootPosition = RowCounter
    End If
    If DescriptionContent = "" Then
        Cells(RowCounter, DescriptionColumn).Interior.Color = WarningColor
        ErrorCounter = (ErrorCounter + 1)
    Else
        Cells(RowCounter, DescriptionColumn).Interior.ColorIndex = 0
    End If
    If Len(CategoryName) > 0 And Len(CategoryContent) < 1 Then
        Cells(RowCounter, CategoryColumn).Interior.Color = ErrorColor
        ErrorCounter = (ErrorCounter + 1)
    ElseIf Len(CategoryName) < 1 And Len(CategoryContent) > 0 Then
        Cells(RowCounter, NameColumn).Interior.Color = ErrorColor
        ErrorCounter = (ErrorCounter + 1)
    ElseIf Len(CategoryName) > 11 And CategoryContent = "ROOT" Then
        Cells(RowCounter, NameColumn).Interior.Color = ErrorColor
        ErrorCounter = (ErrorCounter + 1)
    Else
        Cells(RowCounter, NameColumn).Interior.ColorIndex = 0
        Cells(RowCounter, CategoryColumn).Interior.ColorIndex = 0
    End If
    RowCounter = (RowCounter + 1)
Loop
AdjustedRootCounter = RootCounter

Dim TooManyRoots As Boolean
If RootCounter > Max Then
    TooManyRoots = True
Else
    TooManyRoots = False
End If
RowCounter = StartRow
If RootCounter > Max Then
    Do While RowCounter <= LastRootPosition
        CategoryContent = Cells(RowCounter, CategoryColumn).Value
        CategoryContent = Trim(CategoryContent)
        If Cells(RowCounter, CategoryColumn).Interior.Color = ErrorColor Then
            RowCounter = RowCounter + 1
        Else
            If CategoryContent = "ROOT" Then
                Cells(RowCounter, CategoryColumn).Interior.Color = WarningColor
                RowCounter = RowCounter + 1
                TooManyRoots = True
            Else
                RowCounter = RowCounter + 1
            End If
        End If
    Loop
End If
If TooManyRoots = True Then
ErrorCounter = (ErrorCounter + 1)
End If

For Each Root In ROOTS
    'DoNothing
Next
'LastPositionCheck
Dim LastItemPosition As Integer
RowCounter = StartRow
Do While RowCounter <= LastItemThreshold
    CategoryContent = Cells(RowCounter, CategoryColumn).Text
    CategoryContent = Trim(CategoryContent)
    CategoryName = Cells(RowCounter, NameColumn).Text
    CategoryName = Trim(CategoryName)
    DescriptionContent = Cells(RowCounter, DescriptionColumn).Text
    DescriptionContent = Trim(DescriptionContent)
    If CategoryContent <> "" Or DescriptionContent <> "" Or CategoryName <> "" Then
        LastItemPosition = RowCounter
    End If
    RowCounter = (RowCounter + 1)
Loop
'Content Check
RowCounter = RootCounter + StartRow
Do While RowCounter <= LastItemPosition
    CategoryContent = Cells(RowCounter, CategoryColumn).Text
    CategoryContent = Trim(CategoryContent)
    CategoryName = Cells(RowCounter, NameColumn).Text
    CategoryName = Trim(CategoryName)
    DescriptionContent = Cells(RowCounter, DescriptionColumn).Text
    DescriptionContent = Trim(DescriptionContent)
    If DescriptionContent = "" Then
        Cells(RowCounter, DescriptionColumn).Interior.Color = WarningColor
        ErrorCounter = (ErrorCounter + 1)
    Else
        Cells(RowCounter, DescriptionColumn).Interior.ColorIndex = 0
    End If
     If Len(CategoryName) > 0 And Len(CategoryContent) < 1 Then
        Cells(RowCounter, CategoryColumn).Interior.Color = ErrorColor
        ErrorCounter = (ErrorCounter + 1)
    ElseIf Len(CategoryName) < 1 And Len(CategoryContent) > 0 Then
        Cells(RowCounter, NameColumn).Interior.Color = ErrorColor
        ErrorCounter = (ErrorCounter + 1)
    Else
        Cells(RowCounter, NameColumn).Interior.ColorIndex = 0
        Cells(RowCounter, CategoryColumn).Interior.ColorIndex = 0
    End If
    RowCounter = (RowCounter + 1)
Loop
If ErrorCounter > 0 Then
Cells(1, "E").Value = CStr(ErrorCounter) & " Errors Found!"
Else
Cells(1, "E").Value = ""
End If
Application.EnableEvents = True
End Sub

