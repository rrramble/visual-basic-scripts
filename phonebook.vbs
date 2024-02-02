' 1. The script does not allow Cyrillic characters in the text. Even if you see them they will go wrong eventually!
' 2. Requires 'Microsoft Offce Object Library' (in the menu: Tools/References, then check this option)
' 3. (They say) requires a component "Microsoft VBScript Regular Expression 5.5" (Tools->References...)
'
' https://learn.microsoft.com/en-us/office/vba/api/excel.publishobjects
' https://learn.microsoft.com/en-us/office/vba/api/excel.xlsourcetype
'
' "Sub Main()" must be called to start

' External API constant
Private Const EXTERNAL_API_URI = "https://api.qrserver.com/v1/create-qr-code/?data="

' Company-specific constants:
Private Const HTML_FILENAME_WITH_PATH = "\\portal.local\DavWWWRoot\SiteAssets\phonebook.html"
Private Const PHONEBOOK_SHEET_NAME = "Page1"
Private Const COMPANY_LEGAL_NAME = "SOME COMPANY LLP"


' File related constants:
Private Const QR_STARTING_ROW = 9
Private Const QR_LAST_ROW = 40
Private Const QR_HELP_TEXT = "Click to get a QR contact and scan it with your smartphone"
Private Const JOB_TITLE_COLUMN_LETTER = "B"
Private Const FULLNAME_COLUMN_LETTER = "C"
Private Const EMAIL_COLUMN_LETTER = "D"
Private Const LANDLINE_PHONE_NUMBER_COLUMN_LETTER = ""
Private Const MOBILE_PHONE_NUMBER_COLUMN_LETTER = "E"


Sub Main()
    Dim mhtname$
    Dim wb As Excel.Workbook
    Dim ws As Excel.Worksheet
    Set wb = ThisWorkbook

    AddQRLinksToSheet
    
    With wb.PublishObjects.Add(xlSourcePrintArea, HTML_FILENAME_WITH_PATH, PHONEBOOK_SHEET_NAME, "", xlHtmlStatic)
        .Publish (True)
        .Delete
    End With

    MsgBox "Phonebook is exported!"
End Sub


Sub AddQRLinksToSheet()
    Dim rowNumber As Integer
    rowNumber = QR_STARTING_ROW
    
    While rowNumber <= QR_LAST_ROW
        Name = Trim(ThisWorkbook.Sheets(PHONEBOOK_SHEET_NAME).Range(FULLNAME_COLUMN_LETTER & rowNumber).Text)
        If Name <> "" Then
            AddQRToRow (rowNumber)
        End If
        rowNumber = rowNumber + 1
    Wend
End Sub

Private Sub AddQRToRow(rowNumber As Integer)
        Name = Trim(ThisWorkbook.Sheets(PHONEBOOK_SHEET_NAME).Range(FULLNAME_COLUMN_LETTER & rowNumber).Text)

        fullNameParts = Split(Name)
        strFirstName = fullNameParts(1)
        strSureName = fullNameParts(0)
        
        If GetLength(fullNameParts) = 3 Then
           strFatherName = fullNameParts(2)
        Else
           strFatherName = ""
        End If
        
        mobilePhoneNumber = Trim(ThisWorkbook.Sheets(PHONEBOOK_SHEET_NAME).Range(MOBILE_PHONE_NUMBER_COLUMN_LETTER & rowNumber).Text)
        If InStr(1, mobilePhoneNumber, "8 7") = 1 Then
            mobilePhoneNumber = "+7 7" & Mid(mobilePhoneNumber, 4)
        End If

        If LANDLINE_PHONE_NUMBER_COLUMN_LETTER <> "" Then
            landlinePhoneNumber = Trim(ThisWorkbook.Sheets(PHONEBOOK_SHEET_NAME).Range(LANDLINE_PHONE_NUMBER_COLUMN_LETTER & rowNumber).Text)
            If landlinePhoneNumber <> "" Then
                landlinePhoneNumber = Replace("+77172" & landlinePhoneNumber, " ", "")
            Else
                landlinePhoneNumber = ""
            End If
        End If
        
        strTitle = Trim(ThisWorkbook.Sheets(PHONEBOOK_SHEET_NAME).Range(JOB_TITLE_COLUMN_LETTER & rowNumber).Text)
        strEmail = Trim(ThisWorkbook.Sheets(PHONEBOOK_SHEET_NAME).Range(EMAIL_COLUMN_LETTER & rowNumber).Text)
        
        strVCARD = "BEGIN:VCARD" & vbNewLine _
        & "VERSION:3.0" & vbNewLine _
        & "FN;CHARSET=UTF-8:" & strSureName & " " & strFirstName & " " & strFatherName & vbNewLine _
        & "N;CHARSET=UTF-8:" & strSureName & ";" & strFirstName & " " & strFatherName & ";;;" & vbNewLine _
        & "TEL;TYPE=CELL,VOICE:" & mobilePhoneNumber & vbNewLine _
        & "TEL;TYPE=WORK,VOICE:" & landlinePhoneNumber & vbNewLine _
        & "EMAIL;TYPE=WORK:" & strEmail & vbNewLine _
        & "TITLE;CHARSET=UTF-8:" & strTitle & vbNewLine _
        & "ORG;CHARSET=UTF-8:" & COMPANY_LEGAL_NAME & vbNewLine _
        & "REV:2023-08-20T00:00:00.000Z" & vbNewLine _
        & "END:VCARD"
                
        EncodedString = EXTERNAL_API_URI & Application.EncodeURL(strVCARD)
        If ThisWorkbook.Sheets(PHONEBOOK_SHEET_NAME).Range(MOBILE_PHONE_NUMBER_COLUMN_LETTER & rowNumber).Text <> "" Then
            Dim Range As Range
            ThisWorkbook.Sheets(PHONEBOOK_SHEET_NAME).Range(MOBILE_PHONE_NUMBER_COLUMN_LETTER & rowNumber).ClearHyperlinks
            With ThisWorkbook.Sheets(PHONEBOOK_SHEET_NAME)
                .Hyperlinks.Add _
                Anchor:=.Range(MOBILE_PHONE_NUMBER_COLUMN_LETTER & rowNumber), _
                Address:=EncodedString, _
                ScreenTip:="Scan QR with a smartphone to get the contact"
            End With
        End If
        
        If LANDLINE_PHONE_NUMBER_COLUMN_LETTER <> "" Then
            If ThisWorkbook.Sheets(PHONEBOOK_SHEET_NAME).Range(LANDLINE_PHONE_NUMBER_COLUMN_LETTER & rowNumber).Text <> "" Then
                With ThisWorkbook.Sheets(PHONEBOOK_SHEET_NAME)
                    .Hyperlinks.Add _
                    Anchor:=.Range(LANDLINE_PHONE_NUMBER_COLUMN_LETTER & rowNumber), _
                    Address:=EncodedString, _
                    ScreenTip:="Scan QR with a smartphone to get the contact"
                End With
            End If
        End If
        
        If ThisWorkbook.Sheets(PHONEBOOK_SHEET_NAME).Range(FULLNAME_COLUMN_LETTER & rowNumber).Text <> "" Then
            PhotoURL = ""
            With ThisWorkbook.Sheets(PHONEBOOK_SHEET_NAME)
                .Hyperlinks.Add _
                Anchor:=.Range(FULLNAME_COLUMN_LETTER & rowNumber), _
                Address:=PhotoURL, _
                ScreenTip:="Click to see a photo"
            End With
        End If
End Sub


Public Function GetLength(a As Variant) As Integer
   If IsEmpty(a) Then
      GetLength = 0
   Else
      GetLength = UBound(a) - LBound(a) + 1
   End If
End Function
