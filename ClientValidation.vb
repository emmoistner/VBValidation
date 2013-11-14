Public Sub Worksheet_Change(ByVal Target As range)
Application.EnableEvents = False

Dim StartRow As Integer
Dim RowCounter As Integer
Dim LastItemThreshold As Integer
Dim LastItemPosition As Integer
Dim ErrorCounter As Integer
Dim WarningCounter As Integer
Dim LastRowRequired As Integer

Dim RangeStart As String
Dim RangeEnd As String
Dim Root As String
Dim Category As String
Dim Name As String
Dim Description As String
Dim Address As String
Dim City As String
Dim State As String
Dim Zip As String
Dim Phone As String
Dim Website As String
Dim Email As String
Dim Latitude As String
Dim Longitude As String
Dim Facebook As String
Dim Twitter As String
Dim Tags As String
Dim Photo As String

Dim CategoryColumn As String
Dim NameColumn As String
Dim DescriptionColumn As String
Dim AddressColumn As String
Dim CityColumn As String
Dim StateColumn As String
Dim ZipColumn As String
Dim PhoneColumn As String
Dim WebsiteColumn As String
Dim EmailColumn As String
Dim LatitudeColumn As String
Dim LongitudeColumn As String
Dim FacebookColumn As String
Dim TwitterColumn As String
Dim TagsColumn As String
Dim PhotoColumn As String
Dim WarningColor As Long
Dim ErrorColor As Long
Dim AltRowColor As Long
AltRowColor = RGB(232, 238, 240)
WarningColor = RGB(255, 255, 55)
ErrorColor = RGB(255, 41, 41)

CategoryColumn = "B"
NameColumn = "C"
DescriptionColumn = "D"
AddressColumn = "E"
CityColumn = "F"
StateColumn = "G"
ZipColumn = "H"
PhoneColumn = "I"
WebsiteColumn = "J"
EmailColumn = "K"
LatitudeColumn = "L"
LongitudeColumn = "M"
FacebookColumn = "N"
TwitterColumn = "O"
TagsColumn = "P"
PhotoColumn = "Q"

StartRow = 3
RowCounter = StartRow
LastItemPosition = 3
WarningCounter = 0
ErrorCounter = 0

'Parse Desired Item Limit
LastItemRequired = ThisWorkbook.Sheets("Instructions").range("F2").Text

If IsNumeric(LastItemRequired) Then
    LastItemRequired = CInt(LastItemRequired)
Else
    LastItemRequired = -1
End If

If LastItemRequired < 1 Then
MsgBox ("Please use a number in cell F2 on the instructions sheet.")
End If

'LastPositionSearch
Do While RowCounter <= LastItemRequired
    Category = Cells(RowCounter, CategoryColumn).Text
    Category = Trim(Category)
    Name = Cells(RowCounter, NameColumn).Text
    Name = Trim(Name)
    Description = Cells(RowCounter, DescriptionColumn).Text
    Description = Trim(Description)
    Address = Cells(RowCounter, AddressColumn).Text
    Address = Trim(Address)
    City = Cells(RowCounter, CityColumn).Text
    City = Trim(City)
    State = Cells(RowCounter, StateColumn).Text
    State = Trim(State)
    Zip = Cells(RowCounter, ZipColumn).Text
    Zip = Trim(Zip)
    Phone = Cells(RowCounter, PhoneColumn).Text
    Phone = Trim(Phone)
    Website = Cells(RowCounter, WebsiteColumn).Text
    Website = Trim(Website)
    Email = Cells(RowCounter, EmailColumn).Text
    Email = Trim(Email)
    Latitude = Cells(RowCounter, LatitudeColumn).Text
    Latitude = Trim(Latitude)
    Longitude = Cells(RowCounter, LongitudeColumn).Text
    Longitude = Trim(Longitude)
    Facebook = Cells(RowCounter, FacebookColumn).Text
    Facebook = Trim(Facebook)
    Twitter = Cells(RowCounter, TwitterColumn).Text
    Twitter = Trim(Twitter)
    Tags = Cells(RowCounter, TagsColumn).Text
    Tags = Trim(Tags)
    Photo = Cells(RowCounter, PhotoColumn).Text
    Photo = Trim(Photo)
    If Category <> "" Or Name <> "" Or Description <> "" Or Address <> "" Or City <> "" Or State <> "" Or Zip <> "" Or Phone <> "" Or Website <> "" Or Email <> "" Or Latitude <> "" Or Longitude <> "" Or Facebook <> "" Or Twitter <> "" Or Tags <> "" Or Photo <> "" Then
        LastItemPosition = RowCounter
    End If
    RowCounter = (RowCounter + 1)
Loop

'Content Check
RowCounter = StartRow
Do While RowCounter <= LastItemPosition
    Category = Cells(RowCounter, CategoryColumn).Text
    Category = Trim(Category)
    Name = Cells(RowCounter, NameColumn).Text
    Name = Trim(Name)
    Description = Cells(RowCounter, DescriptionColumn).Text
    Description = Trim(Description)
    Address = Cells(RowCounter, AddressColumn).Text
    Address = Trim(Address)
    City = Cells(RowCounter, CityColumn).Text
    City = Trim(City)
    State = Cells(RowCounter, StateColumn).Text
    State = Trim(State)
    Zip = Cells(RowCounter, ZipColumn).Text
    Zip = Trim(Zip)
    Phone = Cells(RowCounter, PhoneColumn).Text
    Phone = Trim(Phone)
    Website = Cells(RowCounter, WebsiteColumn).Text
    Website = Trim(Website)
    Email = Cells(RowCounter, EmailColumn).Text
    Email = Trim(Email)
    Latitude = Cells(RowCounter, LatitudeColumn).Text
    Latitude = Trim(Latitude)
    Longitude = Cells(RowCounter, LongitudeColumn).Text
    Longitude = Trim(Longitude)
    Facebook = Cells(RowCounter, FacebookColumn).Text
    Facebook = Trim(Facebook)
    Twitter = Cells(RowCounter, TwitterColumn).Text
    Twitter = Trim(Twitter)
    Tags = Cells(RowCounter, TagsColumn).Text
    Tags = Trim(Tags)
    Photo = Cells(RowCounter, PhotoColumn).Text
    Photo = Trim(Photo)

'Empty Category or Name Cell, but other fields have been entered
    If Description <> "" Or Address <> "" Or City <> "" Or State <> "" Or Zip <> "" Or Phone <> "" Or Website <> "" Or Email <> "" Or Latitude <> "" Or Longitude <> "" Or Facebook <> "" Or Twitter <> "" Or Tags <> "" Or Photo <> "" Then
        If Category = "" And Name = "" Then
            Cells(RowCounter, CategoryColumn).Interior.Color = ErrorColor
            Cells(RowCounter, NameColumn).Interior.Color = ErrorColor
            ErrorCounter = (ErrorCounter + 2)
        End If
    End If
    
'Zip Verification

     If fncIsZip(Zip) Or Cells(RowCounter, ZipColumn).Text = "" Then
        Cells(RowCounter, ZipColumn).Interior.ColorIndex = 0
            If ((RowCounter Mod 2) = 0) Then
                Cells(RowCounter, ZipColumn).Interior.Color = AltRowColor
            End If
    Else
        Cells(RowCounter, ZipColumn).Interior.Color = WarningColor
        WarningCounter = WarningCounter + 1
    End If
    
'Email Verification
    If fncIsMail(Cells(RowCounter, EmailColumn).Text) Or Cells(RowCounter, EmailColumn).Text = "" Then
        Cells(RowCounter, EmailColumn).Interior.ColorIndex = 0
            If ((RowCounter Mod 2) = 0) Then
                Cells(RowCounter, EmailColumn).Interior.Color = AltRowColor
            End If
    Else
        Cells(RowCounter, EmailColumn).Interior.Color = WarningColor
        WarningCounter = WarningCounter + 1
    End If

'Phone Verification
   
    If Len(Phone) = 10 And fncIsPhoneNumber(Phone) = -1 Then
    Cells(RowCounter, PhoneColumn).Value = fncToPhone(Cells(RowCounter, PhoneColumn).Text)
    End If
    Phone = Cells(RowCounter, PhoneColumn).Text
    Phone = Trim(Phone)
    If fncIsPhoneNumber(Phone) = 0 And Phone = "" Or Len(Phone) = 10 And fncIsPhoneNumber(Phone) = -1 Then
        Cells(RowCounter, EmailColumn).Interior.ColorIndex = 0
        If ((RowCounter Mod 2) = 0) Then
            Cells(RowCounter, PhoneColumn).Interior.Color = AltRowColor
        End If
    Else
        Cells(RowCounter, PhoneColumn).Interior.Color = WarningColor
        WarningCounter = WarningCounter + 1
    End If
    If Len(Phone) = 12 And fncIsPhoneNumber(Phone) = 1 Then
        Cells(RowCounter, PhoneColumn).Interior.ColorIndex = 0
        If ((RowCounter Mod 2) = 0) Then
            Cells(RowCounter, PhoneColumn).Interior.Color = AltRowColor
        End If
    End If
     
'Category and Name Verification && Category and Name Cell Color Reset
    If Len(Category) < 1 And Len(Name) > 0 Then
        Cells(RowCounter, CategoryColumn).Interior.Color = ErrorColor
        ErrorCounter = (ErrorCounter + 1)
        If Category <> "" Then
                Cells(RowCounter, CategoryColumn).Interior.ColorIndex = 0
                If ((RowCounter Mod 2) = 0) Then
                    Cells(RowCounter, CategoryColumn).Interior.Color = AltRowColor
                End If
        End If
        If Name <> "" Then
            Cells(RowCounter, NameColumn).Interior.ColorIndex = 0
            If ((RowCounter Mod 2) = 0) Then
                Cells(RowCounter, NameColumn).Interior.Color = AltRowColor
            End If
        End If
    ElseIf Len(Name) < 1 And Len(Category) > 0 Then
        Cells(RowCounter, NameColumn).Interior.Color = ErrorColor
        ErrorCounter = (ErrorCounter + 1)
        If Category <> "" Then
            Cells(RowCounter, CategoryColumn).Interior.ColorIndex = 0
            If ((RowCounter Mod 2) = 0) Then
                Cells(RowCounter, CategoryColumn).Interior.Color = AltRowColor
            End If
        End If
            If Name <> "" Then
            Cells(RowCounter, NameColumn).Interior.ColorIndex = 0
                If ((RowCounter Mod 2) = 0) Then
                    Cells(RowCounter, NameColumn).Interior.Color = AltRowColor
            End If
        End If
    Else
        If Description <> "" Or Address <> "" Or City <> "" Or State <> "" Or Zip <> "" Or Phone <> "" Or Website <> "" Or Email <> "" Or Latitude <> "" Or Longitude <> "" Or Facebook <> "" Or Twitter <> "" Or Tags <> "" Or Photo <> "" Then
            'Do Nothing
        Else
            Cells(RowCounter, CategoryColumn).Interior.ColorIndex = 0
            Cells(RowCounter, NameColumn).Interior.ColorIndex = 0
            Cells(RowCounter, EmailColumn).Interior.ColorIndex = 0
            Cells(RowCounter, PhoneColumn).Interior.ColorIndex = 0
            Cells(RowCounter, ZipColumn).Interior.ColorIndex = 0
            If ((RowCounter Mod 2) = 0) Then
                Cells(RowCounter, CategoryColumn).Interior.Color = AltRowColor
                Cells(RowCounter, NameColumn).Interior.Color = AltRowColor
                Cells(RowCounter, EmailColumn).Interior.Color = AltRowColor
                Cells(RowCounter, PhoneColumn).Interior.Color = AltRowColor
                Cells(RowCounter, ZipColumn).Interior.Color = AltRowColor
            End If
        End If
    End If
    

    RowCounter = RowCounter + 1
Loop
'End Content Check

RowCounter = LastItemPosition + 1
'Beyond Last Item Format Reset
Do While RowCounter <= LastItemRequired
    RangeStart = CategoryColumn & CStr(RowCounter)
    RangeEnd = PhotoColumn & CStr(RowCounter)
    rangevariable = RangeStart & ":" & RangeEnd
    If ((RowCounter Mod 2) = 0) Then
        range(rangevariable).Interior.Color = AltRowColor
    Else
        range(rangevariable).Interior.ColorIndex = 0
    End If
    range(rangevariable).Borders.LineStyle = xlContinuous
    RowCounter = RowCounter + 1
Loop

'Blank Out Unused Spaces

Dim ItemCap As Integer
ItemCap = 75
RowCounter = LastItemRequired + 1
Do While RowCounter <= ItemCap
    RangeStart = CategoryColumn & CStr(RowCounter)
    RangeEnd = PhotoColumn & CStr(RowCounter)
    rangevariable = RangeStart & ":" & RangeEnd
    
    If RowCounter = LastItemRequired + 1 Then
        range(rangevariable).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlLineStyleNone
        range(rangevariable).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlLineStyleNone
        range(rangevariable).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlLineStyleNone
        range(rangevariable).Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlLineStyleNone
        range(rangevariable).Borders(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlLineStyleNone
        range(rangevariable).Interior.ColorIndex = 0
    Else
        range(rangevariable).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlLineStyleNone
        range(rangevariable).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlLineStyleNone
        range(rangevariable).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlLineStyleNone
        range(rangevariable).Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlLineStyleNone
        range(rangevariable).Borders(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlLineStyleNone
        range(rangevariable).Borders(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlLineStyleNone
        range(rangevariable).Interior.ColorIndex = 0
    End If

RowCounter = RowCounter + 1
Loop

If ErrorCounter > 0 Then
'MsgBox (CStr(ErrorCounter) & " Errors Found!")
End If
Application.EnableEvents = True
End Sub

Private Function fncIsMail(ByVal strEmail As String) As Boolean
    Const strRFC2822 = "[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!" & _
                        "#$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:" & _
                        "[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:" & _
                        "[a-z0-9-]*[a-z0-9])?"
    Dim objRegEx As Object
    On Error GoTo Fin
    Set objRegEx = CreateObject("Vbscript.Regexp")
    With objRegEx
        .Pattern = strRFC2822
        .IgnoreCase = True
        fncIsMail = .Test(strEmail)
    End With
Fin:
    Set objRegEx = Nothing
    If Err.Number <> 0 Then MsgBox "Error: " & _
        Err.Number & " " & Err.Description
End Function

Private Function fncToPhone(ByVal strPhoneNumber As String) As String
    Dim first As String
    Dim second As String
    Dim third As String
    first = Left(strPhoneNumber, 3)
    second = Mid$(strPhoneNumber, 4, 3)
    third = Right(strPhoneNumber, 4)
    fncToPhone = first + "-" + second + "-" + third
End Function

Private Function fncIsPhoneNumber(ByVal strPhone As String) As Integer
    fncIsPhoneNumber = 0
    If (strPhone Like "###[-]###[-]####") Then
        fncIsPhoneNumber = 1
    End If
    If (strPhone Like "##########") Then
        fncIsPhoneNumber = -1
    End If
End Function

Private Function fncIsZip(ByVal strZip As String) As Boolean
    fncIsZip = strZip Like "#####"
End Function


