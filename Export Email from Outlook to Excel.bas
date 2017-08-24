Attribute VB_Name = "Module10"
Sub Export_email_to_excel()

'===============================================================
'Macro to export emails from Outlook to Excel for specified days
'Coder: Rajathithan Rajasekar
'GitHub: https://github.com/rajathithan/EXCEL-VBA
'WebSite: www.gadoth.com
'Facebook: https://www.facebook.com/gadoth/
'===============================================================

'Error Handler
On Error GoTo error_handler

'Excel File variable
Dim SFilename As String

'Outlook Variables
Dim oApp As Outlook.Application
Dim oNS As Outlook.NameSpace
Dim oInbox As Outlook.MAPIFolder
Dim itemPA As MailItem
Dim oINItems As Outlook.Items
Dim oINItemsRestrict As Outlook.Items

'Start Date and End Date
Dim datStartUTC As Date
Dim datEndUTC As Date

'Provides the ability to create, get, set, and
'delete properties on objects
Dim oPA As PropertyAccessor

'This namespace is used to access Messaging Application
'Programming Interface (MAPI) properties in the Exchange store
Const SchemaPropTag As String = _
    "http://schemas.microsoft.com/mapi/proptag/"
    

'Excel Variables
Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook

'you can give the destination of the filename as you like
SFilename = "D:\Mail-log.xlsx"

'row count
Count = 1

'Set Excel object
Set objExcel = CreateObject("Excel.Application")

'If no file available create one
If Dir(SFilename) <> "" Then
    Set objWorkbook = objExcel.Workbooks.Open(SFilename)
Else
    Set objWorkbook = objExcel.Workbooks.Add
    objWorkbook.SaveAs SFilename
    Set objWorkbook = objExcel.Workbooks.Open(SFilename)
End If

'Clear the contents of the excel and update the column headers
With objWorkbook.Sheets("Sheet1")
    .Cells.ClearContents
    .Cells(Count, 1).Value = "S.NO"
    .Cells(Count, 2).Value = "Mail From"
    .Cells(Count, 3).Value = "Mail received by"
    .Cells(Count, 4).Value = "Subject Line"
    .Cells(Count, 5).Value = "Mail Received Time"
End With

'Move to row number - 2
Count = Count + 1

'Set Outlook objects
Set oApp = New Outlook.Application
Set oNS = oApp.GetNamespace("MAPI")
'Get Inbox folder
Set oInbox = oNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox)
'Get Inbox Items
Set oINItems = oInbox.Items

'Access the mail item properties
Set itemPA = Application.CreateItem(olMailItem)
Set oPA = itemPA.PropertyAccessor

'Get the dates for the filter query
NeededDate = InputBox("Enter the number of days you need to backtrack from current date:" _
, "BackTrackDays")
datStartUTC = oPA.LocalTimeToUTC(Date - NeededDate)
datEndUTC = Now

'Create the SQL filter query, use add quotes function to add the quotes
'This filter uses http://schemas.microsoft.com/mapi/proptag
    strFilter = AddQuotes(SchemaPropTag & "0x0E060040") _
    & " > '" & datStartUTC & "' AND " _
    & AddQuotes(SchemaPropTag & "0x0E060040") _
    & " < '" & datEndUTC & "'"

'Filter the inbox.items
Set oINItemsRestrict = oINItems.Restrict("@SQL=" & strFilter)
'Debug.Print (oINItemsRestrict.Count)

'Iterate through the items
For Each item In oINItemsRestrict
    'Just in case if something is crappy in your inbox
    On Error GoTo nt
    
    'remember you should NOT give the exact email address here ,
    'copy/paste the display name for that email address,
    'That's how outlook recognizes the email address here
    If item.SenderName = "someone(companyname)" Then
        mailinsender = item.SenderName
        mailinSubject = item.Subject
        'You could also capture the body of the email by item.body
        mailinTime = item.ReceivedTime
                With objWorkbook.Sheets("Sheet1")
                            .Cells(Count, 1).Value = Count - 1
                            .Cells(Count, 2).Value = mailinsender
                            .Cells(Count, 3).Value = item.ReceivedByName
                            .Cells(Count, 4).Value = mailinSubject
                            .Cells(Count, 5).Value = mailinTime
                            'Autofit columns in excel sheet
                            .Columns.AutoFit
                 End With
                 'increase row count
                 Count = Count + 1
                 
    End If
nt:
Next


'Save and close everything
objWorkbook.Save
objWorkbook.Close
Set objWorkbook = Nothing
objExcel.Quit
Set objExcel = Nothing

'error handler for code before the iteration
error_handler:
If Err.Number <> 0 Then
    MsgBox "Error found: " & Err.Description
     objWorkbook.Save
     objWorkbook.Close
     Set objWorkbook = Nothing
     objExcel.Quit
     Set objExcel = Nothing
    Exit Sub
End If

MsgBox "Completed the email export from Outlook to Excel !!"

End Sub
Function AddQuotes(ByVal SchemaName As String) As String
    AddQuotes = Chr(34) & SchemaName & Chr(34)
End Function


