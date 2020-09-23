<div align="center">

## E\-mail Access Reports


</div>

### Description

Allows you to use Outlook to send bulk e-mails of Access Reports in Snapshot Format to users (ISO Printing, Faxing etc...)
 
### More Info
 
Uses an Access Database with User e-mail adresses

Used DAO, thus you need to set a Reference in Outlook to the MS DAO 3.51 Library

This code should go in an Oulook module. You can then run the code by itself or by creating a menuitem on your menus.

This would also require you to first save your files as Snapshot files in a specified direcory. You can of course do this in one step but we needed to send 6 files. That would have meant six e-mails, we prefferred to send 1 only and hence the effort to first create the files

No Error Trapping.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Miyagi](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/miyagi.md)
**Level**          |Intermediate
**User Rating**    |4.0 (12 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Microsoft Office Apps/VBA](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/microsoft-office-apps-vba__1-42.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/miyagi-e-mail-access-reports__1-11566/archive/master.zip)





### Source Code

```
'Option Explicit
Public Sub MailToUsers()
  Dim myOlApp As Application
  Dim myItem As MailItem
  Dim Path As String
  Dim myAttachments As Attachments
  Dim db As Database
  Dim rs As Recordset
  Dim BodyMsg As String
  On Error GoTo myErr
  'Set Database and Path to use to use
  Set db = OpenDatabase("z:\DatabasePath\dbDatabaseName.mdb")
  'Set Path to where Files are located
  Path = "z:\SnapshotFilesPath\"
  'Set Value for Body Message
  BodyMsg = "Type whatever bodymessage you might need"
  'Set Recordset to Users Table
  Set rs = db.OpenRecordset("tblUsers")
  'Open or use Outlook
  Set myOlApp = CreateObject("Outlook.Application")
  rs.MoveLast
  rs.MoveFirst
  Do Until rs.EOF
    'Creates a new Outlook MailItem
    Set myItem = myOlApp.CreateItem(olMailItem)
    With myItem
      .To = rs.Fields("[Email]")
      .Subject = "Supply your subject line here"
      .Body = BodyMsg
    End With
    'This Creates an Outlook attachment
    Set myAttachments = myItem.Attachments
    With myAttachments
      'Do for all reports
      .Add Path & "\rptReport1.snp"
      .Add Path & "\rptReport2.snp"
      '************************************
      'Additional Documents can be added
      'Supply full Path and File Name
      '.Add "c:\moc\Questionnaire Script Changes for Dealer Reports 2000_03.doc"
      '************************************
    'Use myItem.Save ISO myItem.Send to view before sending
    'myItem.Save
    myItem.Send
    End With
    'Go to the next user
    rs.MoveNext
  Loop
  Set myOlApp = Nothing
  Set rs = Nothing
  Set db = Nothing
  Exit Sub
myErr:
  Resume Next
End Sub
```

