Private Sub Worksheet_Activate()
Application.EnableEvents = False
Dim StartRow As Integer
Dim RowCounter As Integer
Dim LastItemThreshold As Integer
Dim LastItemPosition As Integer
Dim ErrorCounter As Integer
Dim WarningCounter As Integer
Dim LastRowRequired As Integer
Dim ItemCap As Integer

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
LastItemThreshold = 15
ItemCap = 500

'Parse Desired Item Limit
If IsNumeric(ThisWorkbook.Sheets("Instructions").range("F2").Text) Then
    LastItemRequired = CInt(ThisWorkbook.Sheets("Instructions").range("F2").Text)
Else
    LastItemRequired = -1
End If
'Error Check for Nonintegers parsed. Default Will be set to LastItemThreshold
If LastItemRequired < 1 Then
MsgBox ("Please use a number in cell F2 on the instructions sheet.")
LastItemRequired = LastItemThreshold
End If
'End Parse Desired Item Limit

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


RowCounter = LastItemPosition
    RangeStart = CategoryColumn & CStr(RowCounter)
    RangeEnd = PhotoColumn & CStr(RowCounter)
    rangevariable = RangeStart & ":" & RangeEnd
range(rangevariable).Borders.LineStyle = xlContinuous
RowCounter = RowCounter + 1
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


Public Sub Worksheet_Change(ByVal Target As range)
Application.EnableEvents = False

Dim StartRow As Integer
Dim RowCounter As Integer
Dim LastItemThreshold As Integer
Dim LastItemPosition As Integer
Dim ErrorCounter As Integer
Dim WarningCounter As Integer
Dim LastRowRequired As Integer
Dim ItemCap As Integer

Dim RangeStart As String
Dim RangeEnd As String
Dim WatchRange As range
Dim Intersectrange As range

Dim Values(15) As String


Dim Columns(15) As String
Columns(0) = "B"
Columns(1) = "C"
Columns(2) = "D"
Columns(3) = "E"
Columns(4) = "F"
Columns(5) = "G"
Columns(6) = "H"
Columns(7) = "I"
Columns(8) = "J"
Columns(9) = "K"
Columns(10) = "L"
Columns(11) = "M"
Columns(12) = "N"
Columns(13) = "O"
Columns(14) = "P"
Columns(15) = "Q"

Dim WarningColor As Long
Dim ErrorColor As Long
Dim AltRowColor As Long
AltRowColor = RGB(232, 238, 240)
WarningColor = RGB(255, 255, 55)
ErrorColor = RGB(255, 41, 41)


StartRow = 3
RowCounter = StartRow
LastItemPosition = 3
WarningCounter = 0
ErrorCounter = 0
LastItemThreshold = 15
ItemCap = 500

'Parse Desired Item Limit
If IsNumeric(ThisWorkbook.Sheets("Instructions").range("F2").Text) Then
    LastItemRequired = CInt(ThisWorkbook.Sheets("Instructions").range("F2").Text)
Else
    LastItemRequired = -1
End If
'Error Check for Nonintegers parsed. Default Will be set to LastItemThreshold
If LastItemRequired < 1 Then
MsgBox ("Please use a number in cell F2 on the instructions sheet.")
LastItemRequired = LastItemThreshold
End If
'End Parse Desired Item Limit

'LastPositionSearch
Dim i As Integer
Do While RowCounter <= LastItemRequired
    i = 0
    Do While i < 16
    Values(i) = Cells(RowCounter, Columns(i)).Text
    i = i + 1
    Loop
    If Values(0) <> "" Or Values(1) <> "" Or Values(2) <> "" Or Values(3) <> "" Or Values(4) <> "" Or Values(5) <> "" Or Values(6) <> "" Or Values(7) <> "" Or Values(8) <> "" Or Values(9) <> "" Or Values(10) <> "" Or Values(11) <> "" Or Values(12) <> "" Or Values(13) <> "" Or Values(14) <> "" Or Values(15) <> "" Then
        LastItemPosition = RowCounter
    End If
    
    RowCounter = (RowCounter + 1)
Loop




'Content Check
RowCounter = StartRow
Do While RowCounter <= LastItemPosition
    i = 0
    Do While i < 16
    Values(i) = Trim(Cells(RowCounter, Columns(i)).Text)
    i = i + 1
    Loop

'Empty Category or Name Cell, but other fields have been entered
    If Values(2) <> "" Or Values(3) <> "" Or Values(4) <> "" Or Values(5) <> "" Or Values(6) <> "" Or Values(7) <> "" Or Values(8) <> "" Or Values(9) <> "" Or Values(10) <> "" Or Values(11) <> "" Or Values(12) <> "" Or Values(13) <> "" Or Values(14) <> "" Or Values(15) <> "" Then
        If Values(0) = "" And Values(1) = "" Then
            Cells(RowCounter, Columns(0)).Interior.Color = ErrorColor
            Cells(RowCounter, Columns(1)).Interior.Color = ErrorColor
            ErrorCounter = (ErrorCounter + 2)
        End If
    End If
    
'Zip Verification

     If fncIsZip(Values(6)) Or Cells(RowCounter, Columns(6)).Text = "" Then
        Cells(RowCounter, Columns(6)).Interior.ColorIndex = 0
            If ((RowCounter Mod 2) = 0) Then
                Cells(RowCounter, Columns(6)).Interior.Color = AltRowColor
            End If
    Else
        Cells(RowCounter, Columns(6)).Interior.Color = WarningColor
        WarningCounter = WarningCounter + 1
    End If
    
'Email Verification
    If fncIsMail(Cells(RowCounter, Columns(9)).Text) Or Cells(RowCounter, Columns(9)).Text = "" Then
        Cells(RowCounter, Columns(9)).Interior.ColorIndex = 0
            If ((RowCounter Mod 2) = 0) Then
                Cells(RowCounter, Columns(9)).Interior.Color = AltRowColor
            End If
    Else
        Cells(RowCounter, Columns(9)).Interior.Color = WarningColor
        WarningCounter = WarningCounter + 1
    End If
    
'Website Verification
    If fncIsWebsite(Values(8)) Or Values(8) = "" Then
        Cells(RowCounter, Columns(8)).Interior.ColorIndex = 0
            If ((RowCounter Mod 2) = 0) Then
                Cells(RowCounter, Columns(8)).Interior.Color = AltRowColor
            End If
    Else
        Cells(RowCounter, Columns(8)).Interior.Color = WarningColor
        WarningCounter = WarningCounter + 1
    End If

'Facebook Validation
   If fncIsFacebook(Values(12)) Or Values(12) = "" Then
        Cells(RowCounter, Columns(12)).Interior.ColorIndex = 0
            If ((RowCounter Mod 2) = 0) Then
                Cells(RowCounter, Columns(12)).Interior.Color = AltRowColor
            End If
    Else
        Cells(RowCounter, Columns(12)).Interior.Color = WarningColor
        WarningCounter = WarningCounter + 1
    End If
    
'Facebook Validation
   If fncIsTwitter(Values(13)) Or Values(13) = "" Then
        Cells(RowCounter, Columns(13)).Interior.ColorIndex = 0
            If ((RowCounter Mod 2) = 0) Then
                Cells(RowCounter, Columns(13)).Interior.Color = AltRowColor
            End If
    Else
        Cells(RowCounter, Columns(13)).Interior.Color = WarningColor
        WarningCounter = WarningCounter + 1
    End If
    
'Picture URL Validation
   If fncIsPictureUrl(Values(15)) Or Values(15) = "" Then
        Cells(RowCounter, Columns(15)).Interior.ColorIndex = 0
            If ((RowCounter Mod 2) = 0) Then
                Cells(RowCounter, Columns(15)).Interior.Color = AltRowColor
            End If
    Else
        Cells(RowCounter, Columns(15)).Interior.Color = WarningColor
        WarningCounter = WarningCounter + 1
    End If

'Phone Verification
   
    If Len(Values(7)) = 10 And fncIsPhoneNumber(Values(7)) = -1 Then
    Cells(RowCounter, Columns(7)).Value = fncToPhone(Cells(RowCounter, Columns(7)).Text)
    End If
    Values(7) = Trim(Cells(RowCounter, Columns(7)).Text)
    If fncIsPhoneNumber(Values(7)) = 0 And Values(7) = "" Or Len(Values(7)) = 10 And fncIsPhoneNumber(Values(7)) = -1 Then
        Cells(RowCounter, Columns(7)).Interior.ColorIndex = 0
        If ((RowCounter Mod 2) = 0) Then
            Cells(RowCounter, Columns(7)).Interior.Color = AltRowColor
        End If
    Else
        Cells(RowCounter, Columns(7)).Interior.Color = WarningColor
        WarningCounter = WarningCounter + 1
    End If
    If Len(Values(7)) = 12 And fncIsPhoneNumber(Values(7)) = 1 Then
        Cells(RowCounter, Columns(7)).Interior.ColorIndex = 0
        If ((RowCounter Mod 2) = 0) Then
            Cells(RowCounter, Columns(7)).Interior.Color = AltRowColor
        End If
    End If
     
'Category and Name Verification && Category and Name Cell Color Reset
    If Len(Values(0)) < 1 And Len(Values(1)) > 0 Then
        Cells(RowCounter, Columns(0)).Interior.Color = ErrorColor
        ErrorCounter = (ErrorCounter + 1)
        If Values(0) <> "" Then
                Cells(RowCounter, Columns(0)).Interior.ColorIndex = 0
                If ((RowCounter Mod 2) = 0) Then
                    Cells(RowCounter, Columns(0)).Interior.Color = AltRowColor
                End If
        End If
        If Values(1) <> "" Then
            Cells(RowCounter, Columns(1)).Interior.ColorIndex = 0
            If ((RowCounter Mod 2) = 0) Then
                Cells(RowCounter, Columns(1)).Interior.Color = AltRowColor
            End If
        End If
    ElseIf Len(Values(1)) < 1 And Len(Values(0)) > 0 Then
        Cells(RowCounter, Columns(1)).Interior.Color = ErrorColor
        ErrorCounter = (ErrorCounter + 1)
        If Values(0) <> "" Then
            Cells(RowCounter, Columns(0)).Interior.ColorIndex = 0
            If ((RowCounter Mod 2) = 0) Then
                Cells(RowCounter, Columns(0)).Interior.Color = AltRowColor
            End If
        End If
            If Values(1) <> "" Then
            Cells(RowCounter, Columns(1)).Interior.ColorIndex = 0
                If ((RowCounter Mod 2) = 0) Then
                    Cells(RowCounter, Columns(1)).Interior.Color = AltRowColor
            End If
        End If
    Else
        If Values(2) <> "" Or Values(3) <> "" Or Values(4) <> "" Or Values(5) <> "" Or Values(6) <> "" Or Values(7) <> "" Or Values(8) <> "" Or Values(9) <> "" Or Values(10) <> "" Or Values(11) <> "" Or Values(12) <> "" Or Values(13) <> "" Or Values(14) <> "" Or Values(15) <> "" Then
            'Do Nothing
        Else
            Cells(RowCounter, Columns(0)).Interior.ColorIndex = 0
            Cells(RowCounter, Columns(1)).Interior.ColorIndex = 0
            Cells(RowCounter, Columns(9)).Interior.ColorIndex = 0
            Cells(RowCounter, Columns(6)).Interior.ColorIndex = 0
            Cells(RowCounter, Columns(7)).Interior.ColorIndex = 0
            If ((RowCounter Mod 2) = 0) Then
                Cells(RowCounter, Columns(0)).Interior.Color = AltRowColor
                Cells(RowCounter, Columns(1)).Interior.Color = AltRowColor
                Cells(RowCounter, Columns(9)).Interior.Color = AltRowColor
                Cells(RowCounter, Columns(6)).Interior.Color = AltRowColor
                Cells(RowCounter, Columns(7)).Interior.Color = AltRowColor
            End If
        End If
    End If
    

    RowCounter = RowCounter + 1
Loop
'End Content Check

RowCounter = LastItemPosition
    RangeStart = Columns(0) & CStr(RowCounter)
    RangeEnd = Columns(15) & CStr(RowCounter)
    rangevariable = RangeStart & ":" & RangeEnd
range(rangevariable).Borders.LineStyle = xlContinuous
RowCounter = RowCounter + 1
'Beyond Last Item Format Reset
Do While RowCounter <= LastItemRequired
    RangeStart = Columns(0) & CStr(RowCounter)
    RangeEnd = Columns(15) & CStr(RowCounter)
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


RowCounter = LastItemRequired + 1
Do While RowCounter <= ItemCap
    RangeStart = Columns(0) & CStr(RowCounter)
    RangeEnd = Columns(15) & CStr(RowCounter)
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

Private Function fncIsWebsite(ByVal strWebsite As String) As Boolean
    Const strWebsiteRegExp = "(https?://)?((?:(\w+-)*\w+)\.)+(?:com|org|net|edu|gov|biz|info|name|museum|[a-z]{2})(\/?\w?-?=?_?\??&?)+[\.]?[a-z0-9\?=&_\-%#]*"
    On Error GoTo Fin
    Set obj = CreateObject("Vbscript.Regexp")
    With obj
        .Pattern = strWebsiteRegExp
        .IgnoreCase = True
        fncIsWebsite = .Test(strWebsite)
    End With
Fin:
    Set objRegEx = Nothing
    If Err.Number <> 0 Then MsgBox "Error: " & _
        Err.Number & " " & Err.Description
End Function

Private Function fncIsFacebook(ByVal strFacebook As String) As Boolean
    Const strFacebookRegExp = "(?:(?:http|https):\/\/)?(?:www.)?facebook.com\/(?:(?:\w)*#!\/)?(?:pages\/)?(?:[?\w\-]*\/)?(?:profile.php\?id=(?=\d.*))?([\w\-]*)?"
    On Error GoTo Fin
    Set objFacebook = CreateObject("Vbscript.Regexp")
    With objFacebook
        .Pattern = strFacebookRegExp
        .IgnoreCase = True
        fncIsFacebook = .Test(strFacebook)
    End With
Fin:
    Set objRegEx = Nothing
    If Err.Number <> 0 Then MsgBox "Error: " & _
        Err.Number & " " & Err.Description
End Function

Private Function fncIsTwitter(ByVal strTwitter As String) As Boolean
    Const strTwitterRegExp = "(?:(?:http|https):\/\/)?(?:www.)?twitter\.com\/(#!\/)?[a-zA-Z0-9_]+/?"
    On Error GoTo Fin
    Set objTwitter = CreateObject("Vbscript.Regexp")
    With objTwitter
        .Pattern = strTwitterRegExp
        .IgnoreCase = True
        fncIsTwitter = .Test(strTwitter)
    End With
Fin:
    Set objRegEx = Nothing
    If Err.Number <> 0 Then MsgBox "Error: " & _
        Err.Number & " " & Err.Description
End Function

Private Function fncIsPictureUrl(ByVal strPicture As String) As Boolean
    Const strPictureRegExp = "^https?://(?:[a-z\-]+\.)+[a-z]{2,6}(?:/[^/#?]+)+\.(?:jpe?g|gif|bmp|png)$"
    On Error GoTo Fin
    Set objPicture = CreateObject("Vbscript.Regexp")
    With objPicture
        .Pattern = strPictureRegExp
        .IgnoreCase = True
        fncIsPictureUrl = .Test(strPicture)
    End With
Fin:
    Set objRegEx = Nothing
    If Err.Number <> 0 Then MsgBox "Error: " & _
        Err.Number & " " & Err.Description
End Function

