Attribute VB_Name = "Export_Emails"
Function GetSenderEmailAddress(mailItem As Outlook.mailItem) As String
    Dim sender As Outlook.AddressEntry
    Dim exchangeUser As Outlook.exchangeUser
    
    Set sender = mailItem.sender
    If sender.AddressEntryUserType = olExchangeUserAddressEntry Then
        Set exchangeUser = sender.GetExchangeUser
        GetSenderEmailAddress = exchangeUser.PrimarySmtpAddress
    Else
        GetSenderEmailAddress = sender.Address
    End If
End Function



Sub ExportEmailsToExcel()
    Dim olApp As Outlook.Application
    Dim olNamespace As Outlook.NameSpace
    Dim olFolder As Outlook.MAPIFolder
    Dim olMail As Outlook.mailItem
    Dim xlApp As Excel.Application
    Dim xlWorkbook As Excel.Workbook
    Dim xlWorksheet As Excel.Worksheet
    Dim xlRange As Excel.Range
    Dim row As Long
    Dim currentDate As Date
    Dim userDate As Date
    Dim verification As Boolean
    
    ' Prompt the user for a date
    userDate = InputBox("Enter the date (MM/DD/YYYY):")
    ' Convert the user input to a date value
    If IsDate(userDate) Then
        currentDate = CDate(userDate)
    Else
        MsgBox "Invalid date format. Exiting the macro.", vbExclamation
        Exit Sub
    End If
    
    ' Create Outlook application and get the Inbox folder
    Set olApp = New Outlook.Application
    Set olNamespace = olApp.GetNamespace("MAPI")
     ' Allow the user to pick the folder in which to start the search.
    Set olFolder = olNamespace.PickFolder
     ' Check to make sure user didn't cancel PickFolder dialog.
    If Not (olFolder Is Nothing) Then
        
        ' Start the search process.
        Debug.Print olFolder
    
        ' Create Excel application and open the workbook "Outlook Emails"
        Set xlApp = New Excel.Application
    
        ' Check for the Outlook Email workbook
        strPath = Environ("USERPROFILE") & "\Documents\Excel Workbooks\Outlook Emails.xlsx"
        Debug.Print strPath
        If Dir(strPath) <> "" Then
            ' Workbook exists
            verification = True
            Set xlWorkbook = xlApp.Workbooks.Open(strPath)
            ' Add a new worksheet labeled with the user-given date
            Set xlWorksheet = xlWorkbook.Worksheets.Add
        Else
            ' Workbook does not exist
            verification = False
            Set xlWorkbook = xlApp.Workbooks.Add
            Set xlWorksheet = xlWorkbook.Sheets(1)
            ' Save the workbook at the predefined directory
        End If
    
   
        xlWorksheet.Name = olFolder & " " & Format(currentDate, "MMDDYYYY")
    
        ' Set headers in Excel
        With xlWorksheet
            .Cells(1, 1).Value = "Sender Name"
            .Cells(1, 2).Value = "Sender Email Address"
            .Cells(1, 3).Value = "Subject"
            .Cells(1, 4).Value = "Content"
            .Cells(1, 5).Value = "Received Date"
        End With
    
        row = 2 ' Start from the second row
    
        ' Loop through each email in the Inbox folder
        For Each olMail In olFolder.Items
            ' Check if the item is a MailItem and matches the specified date
            If TypeOf olMail Is Outlook.mailItem And olMail.ReceivedTime >= currentDate Then
                ' Export email information to Excel
                With xlWorksheet
                    .Cells(row, 1).Value = olMail.SenderName
                    .Cells(row, 2).Value = GetSenderEmailAddress(olMail)
                    .Cells(row, 3).Value = olMail.Subject
                    .Cells(row, 4).Value = Replace(Replace(olMail.Body, vbCrLf, vbLf), vbLf & vbLf, vbLf)
                    .Cells(row, 5).Value = olMail.ReceivedTime
                End With
                row = row + 1 ' Move to the next row
            End If
        Next olMail
    
        ' Autofit columns in Excel
        xlWorksheet.Columns.AutoFit
    
        ' Save and close Excel workbook
        If verification = True Then
            xlWorkbook.Save
        Else
            xlWorkbook.SaveAs strPath
        End If
    
    End If
    ' Clean up objects
    xlWorkbook.Close
    Set xlWorksheet = Nothing
    Set xlWorkbook = Nothing
    Set xlApp = Nothing
    Set olFolder = Nothing
    Set olNamespace = Nothing
    Set olApp = Nothing
End Sub

