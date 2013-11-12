Public Sub Worksheet_Change(ByVal Target As Range)
Application.EnableEvents = False
Dim StartRow As Integer
Dim RowCounter As Integer
Dim LastItemThreshold As Integer
Dim LastItemPosition As Integer
Dim ErrorCounter As Integer
Dim WarningCounter As Integer
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

StartRow = 2
LastItemThreshold = 15
RowCounter = StartRow
LastItemPosition = 25
'LastPositionSearch
Do While RowCounter <= LastItemThreshold
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
    If Name = "" Then
        Cells(RowCounter, NameColumn).Interior.Color = ErrorColor
        ErrorCounter = (ErrorCounter + 1)
    Else
        Cells(RowCounter, NameColumn).Interior.ColorIndex = 0
        If ((RowCounter Mod 2) = 0) Then
            Cells(RowCounter, NameColumn).Interior.Color = AltRowColor
            Else
        End If
    End If
    RowCounter = (RowCounter + 1)
Loop

RowCounter = LastItemPosition + 1
'Beyond Last Item Format Reset
Do While RowCounter <= LastItemThreshold
    Dim RangeStart As String
    Dim RangeEnd As String
    RangeStart = CategoryColumn & CStr(RowCounter)
    RangeEnd = PhotoColumn & CStr(RowCounter)
    RangeVariable = RangeStart & ":" & RangeEnd
    If ((RowCounter Mod 2) = 0) Then
        Range(RangeVariable).Interior.Color = AltRowColor
    Else
        Range(RangeVariable).Interior.ColorIndex = 0
    End If
    RowCounter = RowCounter + 1
Loop


If ErrorCounter > 0 Then
MsgBox (CStr(ErrorCounter) & " Errors Found!")
End If
Application.EnableEvents = True
End Sub
