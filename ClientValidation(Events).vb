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
Dim StartDate As String
Dim StartTime As String
Dim EndDate As String
Dim EndTime As String
Dim LastColumnUsed As String

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
Dim StartDateColumn As String
Dim StartTimeColumn As String
Dim EndDateColumn As String
Dim EndTimeColumn As String
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
StartDateColumn = "Q"
StartTimeColumn = "R"
EndDateColumn = "S"
EndTimeColumn = "T"
LastColumnUsed = EndTimeColumn

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
    StartDate = Cells(RowCounter, StartDateColumn).Text
    StartDate = Trim(StartDate)
    StartTime = Cells(RowCounter, StartTimeColumn).Text
    StartTime = Trim(StartTime)
    EndDate = Cells(RowCounter, EndDateColumn).Text
    EndDate = Trim(EndDate)
    EndTime = Cells(RowCounter, EndTimeColumn).Text
    EndTime = Trim(EndTime)
    If Category <> "" Or Name <> "" Or Description <> "" Or Address <> "" Or City <> "" Or State <> "" Or Zip <> "" Or Phone <> "" Or Website <> "" Or Email <> "" Or Latitude <> "" Or Longitude <> "" Or Facebook <> "" Or Twitter <> "" Or Tags <> "" Or StartDate <> "" Or StartTime <> "" Or EndDate <> "" Or EndTime <> "" Then
        LastItemPosition = RowCounter
    End If
    RowCounter = (RowCounter + 1)
Loop

'Format Anything and everything. (Prevents copy and paste format errors)
'End Content Check

RowCounter = StartRow
    RangeStart = CategoryColumn & CStr(RowCounter)
    RangeEnd = EndTimeColumn & CStr(RowCounter)
    rangevariable = RangeStart & ":" & RangeEnd
range(rangevariable).Borders.LineStyle = xlContinuous

RowCounter = RowCounter + 1
'Beyond Last Item Format Reset
Do While RowCounter <= LastItemRequired
    RangeStart = CategoryColumn & CStr(RowCounter)
    RangeEnd = EndTimeColumn & CStr(RowCounter)
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
    RangeEnd = EndTimeColumn & CStr(RowCounter)
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
    StartDate = Cells(RowCounter, StartDateColumn).Text
    StartDate = Trim(StartDate)
    StartTime = Cells(RowCounter, StartTimeColumn).Text
    StartTime = Trim(StartTime)
    EndDate = Cells(RowCounter, EndDateColumn).Text
    EndDate = Trim(EndDate)
    EndTime = Cells(RowCounter, EndTimeColumn).Text
    EndTime = Trim(EndTime)

'Empty Category, Name, StartDate, StartTime, EndDate, or EndTime Cell, but other fields have been entered
    If Description <> "" Or Address <> "" Or City <> "" Or State <> "" Or Zip <> "" Or Phone <> "" Or Website <> "" Or Email <> "" Or Latitude <> "" Or Longitude <> "" Or Facebook <> "" Or Twitter <> "" Or Tags <> "" Then
        If Category = "" And Name = "" Then
            Cells(RowCounter, CategoryColumn).Interior.Color = ErrorColor
            Cells(RowCounter, NameColumn).Interior.Color = ErrorColor
            ErrorCounter = (ErrorCounter + 4)
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
    
'Website Verification
    If fncIsWebsite(Website) Or Website = "" Then
        Cells(RowCounter, WebsiteColumn).Interior.ColorIndex = 0
            If ((RowCounter Mod 2) = 0) Then
                Cells(RowCounter, WebsiteColumn).Interior.Color = AltRowColor
            End If
    Else
        Cells(RowCounter, WebsiteColumn).Interior.Color = WarningColor
        WarningCounter = WarningCounter + 1
    End If

'Facebook Validation
   If fncIsFacebook(Facebook) Or Facebook = "" Then
        Cells(RowCounter, FacebookColumn).Interior.ColorIndex = 0
            If ((RowCounter Mod 2) = 0) Then
                Cells(RowCounter, FacebookColumn).Interior.Color = AltRowColor
            End If
    Else
        Cells(RowCounter, FacebookColumn).Interior.Color = WarningColor
        WarningCounter = WarningCounter + 1
    End If
    
'Facebook Validation
   If fncIsTwitter(Twitter) Or Twitter = "" Then
        Cells(RowCounter, TwitterColumn).Interior.ColorIndex = 0
            If ((RowCounter Mod 2) = 0) Then
                Cells(RowCounter, TwitterColumn).Interior.Color = AltRowColor
            End If
    Else
        Cells(RowCounter, TwitterColumn).Interior.Color = WarningColor
        WarningCounter = WarningCounter + 1
    End If
    
'StartDate and EndDate Validation
   If fncIsDate(StartDate) Or StartDate = "" Then
        Cells(RowCounter, StartDateColumn).Interior.ColorIndex = 0
            If ((RowCounter Mod 2) = 0) Then
                Cells(RowCounter, StartDateColumn).Interior.Color = AltRowColor
            End If
    Else
        Cells(RowCounter, StartDateColumn).Interior.Color = ErrorColor
        ErrorCounter = ErrorCounter + 1
        
    
    End If

   If fncIsDate(EndDate) Or EndDate = "" Then
        Cells(RowCounter, EndDateColumn).Interior.ColorIndex = 0
            If ((RowCounter Mod 2) = 0) Then
                Cells(RowCounter, EndDateColumn).Interior.Color = AltRowColor
            End If
    Else
        Cells(RowCounter, EndDateColumn).Interior.Color = ErrorColor
        ErrorCounter = ErrorCounter + 1
        
    
    End If
    
'StartTime and EndTime Validation
'   If fncIsTime(StartTime) Or StartTime = "" Then
 '       Cells(RowCounter, StartTimeColumn).Interior.ColorIndex = 0
  '          If ((RowCounter Mod 2) = 0) Then
   '             Cells(RowCounter, StartTimeColumn).Interior.Color = AltRowColor
    '        End If
'    Else
 '       Cells(RowCounter, StartTimeColumn).Interior.Color = ErrorColor
  '      ErrorCounter = ErrorCounter + 1
   '
    '
'    End If
'
 '  If fncIsTime(EndTime) Or EndTime = "" Then
  '      Cells(RowCounter, EndTimeColumn).Interior.ColorIndex = 0
   '         If ((RowCounter Mod 2) = 0) Then
    '            Cells(RowCounter, EndTimeColumn).Interior.Color = AltRowColor
     '       End If
'    Else
 '       Cells(RowCounter, EndTimeColumn).Interior.Color = ErrorColor
  '      ErrorCounter = ErrorCounter + 1
   ' End If

'Phone Verification
   
    If Len(Phone) = 10 And fncIsPhoneNumber(Phone) = -1 Then
    Cells(RowCounter, PhoneColumn).Value = fncToPhone(Cells(RowCounter, PhoneColumn).Text)
    End If
    Phone = Cells(RowCounter, PhoneColumn).Text
    Phone = Trim(Phone)
    If fncIsPhoneNumber(Phone) = 0 And Phone = "" Or Len(Phone) = 10 And fncIsPhoneNumber(Phone) = -1 Then
        Cells(RowCounter, PhoneColumn).Interior.ColorIndex = 0
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

RowCounter = LastItemPosition
    RangeStart = CategoryColumn & CStr(RowCounter)
    RangeEnd = LastColumnUsed & CStr(RowCounter)
    rangevariable = RangeStart & ":" & RangeEnd
range(rangevariable).Borders.LineStyle = xlContinuous
RowCounter = RowCounter + 1
'Beyond Last Item Format Reset
Do While RowCounter <= LastItemRequired
    RangeStart = CategoryColumn & CStr(RowCounter)
    RangeEnd = LastColumnUsed & CStr(RowCounter)
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
    RangeEnd = LastColumnUsed & CStr(RowCounter)
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

Private Function fncIsDate(ByVal strDate As String) As Boolean
    If strDate Like ("#[/]#[/]####") Or strDate Like ("##[/]#[/]####") Or strDate Like ("#[/]##[/]####") Or strDate Like ("##[/]##[/]####") Then
        fncIsDate = True
    Else
        fncIsDate = False
    End If
End Function

'Private Function fncIsTime(ByVal strTime As String) As Boolean
 '   If strDate Like ("##:##:##[ ][APap][Mm]") Or strDate Like ("#:##:##[ ][APap][Mm]") Then
  '      fncIsTime = True
   ' Else
    '    fncIsTime = False
    'End If
'End Function



