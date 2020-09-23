<div align="center">

## Check your Outlook Inbox for new messages \. \. \.


</div>

### Description

Checks you Microsoft Outlook Inbox for new Mail Items.
 
### More Info
 
Need to set "References" to Microsoft Outlook.

The number of new messages


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Thirst Technologies](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/thirst-technologies.md)
**Level**          |Unknown
**User Rating**    |3.4 (24 globes from 7 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[OLE/ COM/ DCOM/ Active\-X](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/ole-com-dcom-active-x__1-29.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/thirst-technologies-check-your-outlook-inbox-for-new-messages__1-2215/archive/master.zip)





### Source Code

```
Dim objOutlook As Outlook.Application
Dim objMapiName As Outlook.NameSpace
Dim intCountUnRead As Integer
Private Sub Check_Mail_Click()
 Set objOutlook = New Outlook.Application
 Set objMapiName = objOutlook.GetNamespace("MAPI")
 For I = 1 To objMapiName.GetDefaultFolder(olFolderInbox).UnReadItemCount
  intCountUnRead = intCountUnRead + 1
 Next
  MsgBox "You have " & intCountUnRead & " new messages in your Inbox . .
  ",   vbInformation + vbOKOnly, "New Messages . . ."
  intCountUnRead = 0
 Set objMapiName = Nothing
 Set objOutlook = Nothing
End Sub
```

